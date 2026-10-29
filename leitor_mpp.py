"""Leitor de arquivos .mpp do Microsoft Project — parser puro Python (sem Java/COM).

Reimplementa o layout decifrado do binário MPP14+ (2010 em diante), incluindo o
mapeamento de campos das tabelas internas (TBkndTask, TBkndCons, TBkndRsc,
TBkndAssn). Não requer MS Project instalado nem MPXJ/Java — usa apenas `olefile`.

Referência de conceito (MIT): https://github.com/andrewbergsma/msp-export —
a ideia de ler o compound-document com olefile veio de lá, mas todos os offsets
e o mapeamento de campos abaixo foram decifrados empiricamente contra um .mpp real
("CRONO CTP-PSB - Atualizado - Cópia.mpp") e validados contra um export em Excel
do mesmo cronograma (11/12 linhas iguais; a linha-raiz difere por arredondamento
do Excel) e contra o marcador estrutural de hierarquia do próprio binário
(923/923 tarefas).

O que este leitor preenche no DataFrame de saída (padrão DB_Cronograma_MSProject):
  - ID_Tarefa (ID sequencial do MS Project), Nome_Tarefa, WBS, Inicio/Fim_Planejado,
    Duracao_dias, %_Concluido, Predecessora, Sucessora, Responsavel, Status.
O que NÃO vem do binário (fica vazio / derivado):
  - Caminho crítico (não é campo binário) e datas reais (derivadas do %).
  - Responsavel: o vínculo exato alocação↔recurso usa índices internos não
    decifrados; só preenche se houver um único recurso nomeado (nunca inventa).
"""

from __future__ import annotations

import io
import struct
from datetime import datetime, timedelta

import olefile
import pandas as pd

COLUNAS_CRONOGRAMA = [
    "ID_Tarefa",
    "Nome_Tarefa",
    "WBS",
    "Inicio_Planejado",
    "Fim_Planejado",
    "Inicio_Real",
    "Fim_Real",
    "Duracao_dias",
    "%_Concluido",
    "Predecessora",
    "Sucessora",
    "Caminho_Critico_SIM_NAO",
    "Responsavel",
    "Status",
    "Observacao",
]

# Layout do FixedData das TAREFAS (TBkndTask) — decifrado e validado empíricamente.
# O stream tem exatamente 48 + 923*202 bytes (resto 0) para o arquivo de teste:
# header de 48 B e registros de 202 B, com o ID sequencial em @0, o UID em @4 e o
# flag de tarefa-resumo em @140. (Ler com header 52 desloca tudo em 4 B e perde
# a última tarefa.)
_TASK_HDR = 48           # bytes de cabeçalho antes da 1a linha de registro
_TASK_REC = 202          # tamanho do registro fixo de tarefa
_OFF_ID = 0              # u32: ID sequencial do MS Project (ordem de exibição)
_OFF_UID = 4             # u32: Unique ID da tarefa
_OFF_IS_SUMMARY = 140    # u8: 1 = tarefa-resumo, 0 = folha

# Tipos de campo variável (VarMeta) das tarefas — validados contra gabarito
_TIPO_NOME = 14          # utf-16: nome da tarefa
_TIPO_DURACAO = 27       # u32: duracao em décimos de minutos (9h/dia = 5400)
_TIPO_INICIO = 43        # u32: inicio (data MPP)
_TIPO_FIM = 44           # u32: fim (data MPP)
_TIPO_PCT = 87           # f64: % concluido (0..100)

# Vínculos (TBkndCons)
_CONS_REC = 20
_CONS_OFF_PRED = 4       # i32: UID da predecessora
_CONS_OFF_SUCC = 8       # i32: UID da sucessora
_CONS_OFF_TYPE = 12      # i32: tipo de vínculo (1=FS, 3=SS, etc.)

# Recursos (TBkndRsc) — FixedData: 8 registros; uid@0, e um campo de índice @168
_RSC_REC = 172
_RSC_INI = 52
_RSC_OFF_UID = 0
_RSC_OFF_IDX = 168

# Alocações (TBkndAssn) — FixedData: @0 = id interno da alocação, @4 = UID da tarefa
_ASSN_REC = 110
_ASSN_OFF_TASK_UID = 4

# Época das datas do formato MPP: 31/12/1983. Valor u32 = dias * 65536 desde a época.
_EPOCA = datetime(1983, 12, 31)


def _u32(b: bytes, o: int) -> int:
    return struct.unpack_from("<I", b, o)[0] if o + 4 <= len(b) else 0


def _i32(b: bytes, o: int) -> int:
    return struct.unpack_from("<i", b, o)[0] if o + 4 <= len(b) else 0


def _f64(b: bytes, o: int) -> float:
    return struct.unpack_from("<d", b, o)[0] if o + 8 <= len(b) else 0.0


def _data_mpp(v: int):
    """Converte u32 MPP em date (apenas a data; a hora-de-dia é imprecisa)."""
    if not (300_000_000 < v < 1_500_000_000):
        return None
    try:
        return (_EPOCA + timedelta(days=v / 65536)).date()
    except (OverflowError, ValueError):
        return None


def _utf16(b: bytes) -> str:
    try:
        return b.decode("utf-16-le", errors="ignore").rstrip("\x00").strip()
    except Exception:
        return ""


def _minutos_por_dia(duracoes_tenths: list[int]) -> int:
    """Detecta o dia útil do arquivo (480=8h, 540=9h, 600=10h) por divisibilidade."""
    candidatos = (480, 540, 600, 720)
    melhor, melhor_nota = 480, -1.0
    reais = [d for d in duracoes_tenths if d and d > 0]
    if not reais:
        return 480
    for mpd in candidatos:
        inteiros = sum(1 for d in reais if abs((d / 10 / mpd) - round(d / 10 / mpd)) < 0.02)
        nota = inteiros / len(reais)
        if nota > melhor_nota:
            melhor_nota, melhor = nota, mpd
    return melhor


def _status_de_pct(pct: float) -> str:
    if pct >= 100:
        return "Concluído"
    if pct > 0:
        return "Em Andamento"
    return "Não Iniciado"


def _parse_varmeta(stream_varmeta: bytes, stream_var2data: bytes, tipos: set[int]):
    """Percorre VarMeta e devolve dict[uid][tipo] = payload(bytes) para os tipos pedidos."""
    itens = _u32(stream_varmeta, 8)
    out: dict[int, dict[int, bytes]] = {}
    for i in range(itens):
        o = 24 + i * 12
        if o + 12 > len(stream_varmeta):
            break
        a, off_v2, campo = struct.unpack_from("<III", stream_varmeta, o)
        tipo = campo & 0xFFFF
        if tipo not in tipos:
            continue
        if off_v2 + 4 > len(stream_var2data):
            continue
        n = _u32(stream_var2data, off_v2)
        if n > 50000 or off_v2 + 4 + n > len(stream_var2data):
            continue
        out.setdefault(a, {})[tipo] = stream_var2data[off_v2 + 4 : off_v2 + 4 + n]
    return out


def interpretar_mpp(conteudo: bytes) -> tuple[pd.DataFrame, list[str]]:
    """Lê um .mpp (bytes) e devolve (DataFrame no padrão Sheets, avisos).

    Lança ValueError com mensagem amigável se o arquivo não for MPP14+ legível.
    """
    try:
        ole = olefile.OleFileIO(io.BytesIO(conteudo))
    except Exception as exc:
        raise ValueError("Arquivo não é um .mpp válido do Microsoft Project (binário OLE).") from exc

    avisos: list[str] = []
    try:
        raiz = None
        for e in ole.listdir(streams=True, storages=True):
            if e and e[0].startswith("   114"):
                raiz = e[0]
                break
        if raiz is None:
            raise ValueError(
                "Formato não suportado: apenas arquivos .mpp do MS Project 2010 ou "
                "mais recentes (MPP14+). Para versões antigas, salve como XML ou CSV "
                "no MS Project e importe esse arquivo."
            )

        def st(*caminho):
            return ole.openstream([raiz] + list(caminho)).read()

        # --- tarefas: blocos fixos ---
        fixed = st("TBkndTask", "FixedData")
        vmeta = st("TBkndTask", "VarMeta")
        v2 = st("TBkndTask", "Var2Data")

        registros: list[dict] = []
        n_linhas = (len(fixed) - _TASK_HDR) // _TASK_REC
        for i in range(n_linhas):
            ini = _TASK_HDR + i * _TASK_REC
            if ini + _TASK_REC > len(fixed):
                break
            reg = fixed[ini : ini + _TASK_REC]
            uid = _u32(reg, _OFF_UID)
            if uid == 0xFFFFFFFF:
                continue  # sentinela / linha vazia
            registros.append(
                {
                    "uid": uid,
                    "reg": reg,
                    "id": _u32(reg, _OFF_ID),
                    "resumo": bool(reg[_OFF_IS_SUMMARY]) if _OFF_IS_SUMMARY < len(reg) else False,
                }
            )

        if not registros:
            raise ValueError("Nenhuma tarefa encontrada no arquivo .mpp.")

        por_uid = {r["uid"]: r for r in registros}

        # --- campos variáveis das tarefas ---
        tipos_task = {_TIPO_NOME, _TIPO_DURACAO, _TIPO_INICIO, _TIPO_FIM, _TIPO_PCT}
        campos = _parse_varmeta(vmeta, v2, tipos_task)

        # detecta dia útil p/ converter duração
        duracoes = []
        for r in registros:
            p = campos.get(r["uid"], {}).get(_TIPO_DURACAO)
            if p:
                duracoes.append(_u32(p, 0))
        mpd = _minutos_por_dia(duracoes)
        avisos.append(f"Durações convertidas com {mpd} min/dia (calendário do arquivo).")

        # --- vínculos ---
        pred_de: dict[int, list[int]] = {}   # succ_uid -> [pred_uid]
        succ_de: dict[int, list[int]] = {}   # pred_uid -> [succ_uid]
        try:
            cons = st("TBkndCons", "FixedData")
            n_cons = len(cons) // _CONS_REC
            for i in range(n_cons):
                off = i * _CONS_REC
                pred = _i32(cons, off + _CONS_OFF_PRED)
                succ = _i32(cons, off + _CONS_OFF_SUCC)
                if pred in por_uid and succ in por_uid:
                    pred_de.setdefault(succ, []).append(pred)
                    succ_de.setdefault(pred, []).append(succ)
        except Exception:
            avisos.append("Sem tabela de vínculos (TBkndCons); Predecessora ficará vazio.")

        # --- responsáveis via alocações + recursos ---
        responsavel = _ler_responsavel(ole, raiz, por_uid, avisos)

        # --- ordem de exibição (ID sequencial) e WBS pela hierarquia ---
        registros.sort(key=lambda r: r["id"])
        mapa_id, mapa_wbs = _gerar_wbs_e_id(registros, campos)

        linhas = []
        for r in registros:
            uid = r["uid"]
            c = campos.get(uid, {})
            nome = _utf16(c[_TIPO_NOME]) if _TIPO_NOME in c else ""
            dur_tenths = _u32(c[_TIPO_DURACAO], 0) if _TIPO_DURACAO in c else 0
            dur_dias = round(dur_tenths / 10.0 / mpd, 2) if mpd else 0.0
            inicio = _data_mpp(_u32(c[_TIPO_INICIO], 0)) if _TIPO_INICIO in c else None
            fim = _data_mpp(_u32(c[_TIPO_FIM], 0)) if _TIPO_FIM in c else None
            pct = _f64(c[_TIPO_PCT], 0) if _TIPO_PCT in c else 0.0
            pct = max(0.0, min(100.0, pct))

            # datas reais derivadas do % (o binário não traz Inicio/Fim_Real)
            if pct >= 100:
                ir = inicio.strftime("%d/%m/%Y") if inicio else ""
                fr = fim.strftime("%d/%m/%Y") if fim else ""
            elif pct > 0:
                ir = inicio.strftime("%d/%m/%Y") if inicio else ""
                fr = ""
            else:
                ir, fr = "", ""

            preds = pred_de.get(uid, [])
            pred_txt = "; ".join(str(mapa_id.get(p, "")) for p in preds if p in mapa_id)
            succs = succ_de.get(uid, [])
            succ_txt = "; ".join(str(mapa_id.get(s, "")) for s in succs if s in mapa_id)

            linhas.append(
                {
                    "ID_Tarefa": mapa_id.get(uid, ""),
                    "Nome_Tarefa": nome,
                    "WBS": mapa_wbs.get(uid, ""),
                    "Inicio_Planejado": inicio.strftime("%d/%m/%Y") if inicio else "",
                    "Fim_Planejado": fim.strftime("%d/%m/%Y") if fim else "",
                    "Inicio_Real": ir,
                    "Fim_Real": fr,
                    "Duracao_dias": dur_dias,
                    "%_Concluido": pct,
                    "Predecessora": pred_txt,
                    "Sucessora": succ_txt,
                    "Caminho_Critico_SIM_NAO": "NAO",
                    "Responsavel": responsavel.get(uid, ""),
                    "Status": _status_de_pct(pct),
                    "Observacao": "",
                }
            )

        avisos.append(
            "Caminho crítico não existe no binário .mpp: marcado como 'NAO'. "
            "Datas reais derivadas do % de conclusão (100% = fim real; parcial = só início)."
        )
        df = pd.DataFrame(linhas, columns=COLUNAS_CRONOGRAMA)
        return df, avisos
    finally:
        try:
            ole.close()
        except Exception:
            pass


def _gerar_wbs_e_id(ordem: list[dict], campos: dict) -> tuple[dict[int, str], dict[int, str]]:
    """Gera ID_Tarefa (ID sequencial do MS Project) e WBS pela hierarquia.

    O binário não expõe o "outline level" como campo simples, mas garante duas
    coisas que este algoritmo usa:

    1. As tarefas estão em pré-ordem (pai antes dos filhos) quando ordenadas pelo
       ID sequencial (registro @0) — a ordem de exibição do MS Project.
    2. As datas de uma tarefa-resumo são o rolamento dos filhas: todo filho cabe
       exatamente no intervalo [início, fim] do resumo que o contém.

    Logo, percorrendo em pré-ordem e mantendo uma pilha de resumos abertos, um
    filho fecha (desempilha) qualquer resumo do topo cujo intervalo não o contém.
    O WBS sai como '1', '2' (seções sob a raiz), '1.1', '2.3.1' etc. — a raiz do
    projeto (primeira tarefa, resumo) fica sem segmento, como no MS Project.

    Validado contra o marcador estrutural 0x0B4000B3 do VarMeta (21 = resumo,
    7 = folha de nível 3, 5 = folha de nível 4) em 923/923 tarefas do arquivo real.
    """
    mapa_id: dict[int, str] = {}
    mapa_wbs: dict[int, str] = {}
    # cada nó da pilha: {"ini": date|None, "fim": date|None, "n": contador}
    pilha: list[dict] = []

    for pos, r in enumerate(ordem):
        uid = r["uid"]
        c = campos.get(uid, {})
        ini = _data_mpp(_u32(c[_TIPO_INICIO], 0)) if _TIPO_INICIO in c else None
        fim = _data_mpp(_u32(c[_TIPO_FIM], 0)) if _TIPO_FIM in c else None

        if pos == 0:
            # tarefa-raiz (resumo do projeto): âncora da pilha, sem segmento WBS
            mapa_wbs[uid] = ""
            mapa_id[uid] = str(r["id"])
            pilha.append({"ini": ini, "fim": fim, "n": 0})
            continue

        # fecha qualquer resumo aberto que não contenha esta tarefa
        while len(pilha) > 1 and not _cabe_no_intervalo(ini, fim, pilha[-1]):
            pilha.pop()
        pilha[-1]["n"] += 1
        mapa_wbs[uid] = ".".join(str(no["n"]) for no in pilha)
        mapa_id[uid] = str(r["id"])

        if r["resumo"]:
            pilha.append({"ini": ini, "fim": fim, "n": 0})

    return mapa_id, mapa_wbs


def _cabe_no_intervalo(ini, fim, no: dict) -> bool:
    """True se [ini, fim] está contido no intervalo do resumo `no` (datas conhecidas)."""
    if ini is None or fim is None or no["ini"] is None or no["fim"] is None:
        return True  # sem datas não dá para recusar: mantém no nível corrente
    return no["ini"] <= ini and fim <= no["fim"]


def _ler_responsavel(ole, raiz: str, por_uid: dict, avisos: list[str]) -> dict[int, str]:
    """Preenche Responsavel por tarefa a partir das alocações (TBkndAssn).

    O binário liga alocação→tarefa pelo UID (FixedData @4), mas o nome do recurso
    está numa tabela interna ligada por índice, não por UID direto. Sem o vínculo
    exato 100% confirmado, preenchemos de forma conservadora: se houver um único
    recurso nomeado no arquivo, ele é o responsável de todas as tarefas alocadas;
    caso contrário deixamos vazio para não inventar dado (regra do AGENTS.md).
    """
    try:
        rsc_vmeta = ole.openstream([raiz, "TBkndRsc", "VarMeta"]).read()
        rsc_v2 = ole.openstream([raiz, "TBkndRsc", "Var2Data"]).read()
        assn_fixed = ole.openstream([raiz, "TBkndAssn", "FixedData"]).read()
    except Exception:
        return {}

    # nomes de recurso: varre o Var2Data procurando strings UTF-16 plausíveis
    nomes = _extrair_nomes_recurso(rsc_vmeta, rsc_v2)
    if not nomes:
        return {}

    # tarefas que têm alguma alocação
    alocadas: set[int] = set()
    n_assn = len(assn_fixed) // _ASSN_REC
    for i in range(n_assn):
        off = i * _ASSN_REC
        task_uid = _i32(assn_fixed, off + _ASSN_OFF_TASK_UID)
        if task_uid in por_uid:
            alocadas.add(task_uid)

    if len(nomes) == 1:
        unico = nomes[0]
        avisos.append(
            f"Responsável aplicado de {len(alocadas)} tarefa(s) alocada(s) "
            f"(recurso único no arquivo: '{unico}')."
        )
        return {uid: unico for uid in alocadas}

    avisos.append(
        "Responsável não preenchido: o binário não expõe o vínculo exato recurso↔tarefa "
        f"({len(nomes)} recursos no arquivo, vários campos). Preencha manualmente se necessário."
    )
    return {}


def _extrair_nomes_recurso(rsc_vmeta: bytes, rsc_v2: bytes) -> list[str]:
    """Extrai nomes de recurso do TBkndRsc (strings UTF-16 plausíveis)."""
    itens = _u32(rsc_vmeta, 8)
    nomes: list[str] = []
    vistos: set[str] = set()
    for i in range(itens):
        o = 24 + i * 12
        if o + 12 > len(rsc_vmeta):
            break
        a, off_v2, campo = struct.unpack_from("<III", rsc_vmeta, o)
        if off_v2 + 4 > len(rsc_v2):
            continue
        n = _u32(rsc_v2, off_v2)
        if not (4 <= n <= 200 and off_v2 + 4 + n <= len(rsc_v2)):
            continue
        payload = rsc_v2[off_v2 + 4 : off_v2 + 4 + n]
        # nome de recurso: UTF-16 limpo (bytes ímpares zerados) e com letras
        if len(payload) >= 4 and payload[1] == 0 and payload[3] == 0:
            s = _utf16(payload)
            if s and any(ch.isalpha() for ch in s) and s not in vistos:
                vistos.add(s)
                nomes.append(s)
    return nomes
