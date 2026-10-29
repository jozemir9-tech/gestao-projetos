"""Client HTTP da API do Google Sheets (Google Apps Script).

A API roda como Web App no Apps Script e devolve as 6 abas do projeto em JSON,
protegidas por API Key. Este módulo baixa cada aba e devolve DataFrames no
mesmo formato dos CSVs locais, para o app funcionar igual. Também grava
(create/update/delete via POST) usando a MESMA API Key da leitura.

Configuração (barra lateral do app, ou variáveis de ambiente):
    SHEETS_API_URL  URL do Web App terminando em /exec
    SHEETS_API_KEY  a chave gerada pela função gerarEInstalarChave

Exemplo de URL:
    https://script.google.com/macros/s/AKfy.../exec
"""
from __future__ import annotations

import os
from typing import Optional

import pandas as pd
import requests

TIMEOUT = 60  # Apps Script pode ser lento na 1ª chamada (cold start)

# Nome de cada aba (igual ao CONFIG.ABAS do script) e seus DataFrames.
ABAS = {
    "cronograma": "DB_Cronograma_MSProject",
    "hist_mo": "DB_Histograma_MO",
    "hist_equip": "DB_Histograma_Equip",
    "rh": "DB_Recursos_Humanos",
    "contratos": "DB_Contratos",
    "timeline": "Timeline_Visao",
}


class ErroSheetsApi(RuntimeError):
    """Erro de rede ou da própria API (mensagem amigável para o app)."""


def _cabecalhos(url: str, chave: str) -> dict:
    return {"x-api-key": chave} if chave else {}


def _get(url: str, chave: str, params: Optional[dict] = None) -> dict:
    """Faz um GET na API e valida a resposta JSON."""
    if not url:
        raise ErroSheetsApi("URL da API não informada.")
    q = dict(params or {})
    # Se não houver header de API key, manda na query string.
    if not chave and "api_key" not in q:
        raise ErroSheetsApi("API Key não informada.")
    if not chave:
        q["api_key"] = chave
    try:
        r = requests.get(url, params=q, headers=_cabecalhos(url, chave), timeout=TIMEOUT)
    except requests.RequestException as exc:
        raise ErroSheetsApi(f"Falha de rede ao chamar a API ({exc}).") from exc
    try:
        corpo = r.json()
    except ValueError as exc:
        raise ErroSheetsApi(
            f"A API não devolveu JSON (status {r.url and r.status_code}). "
            "Confira se a URL termina em /exec e se o app foi publicado como Web App."
        ) from exc
    if not corpo.get("success"):
        raise ErroSheetsApi(corpo.get("error") or f"Erro da API (status {r.status_code}).")
    return corpo


def _post(url: str, chave: str, payload: dict) -> dict:
    """Faz um POST na API (create/update/delete) e valida a resposta JSON."""
    if not url:
        raise ErroSheetsApi("URL da API não informada.")
    if not chave:
        raise ErroSheetsApi("API Key não informada (obrigatória para gravar).")
    corpo = dict(payload)
    corpo["api_key"] = chave  # redundante com o header; o script aceita os dois.
    try:
        r = requests.post(url, json=corpo, headers=_cabecalhos(url, chave), timeout=TIMEOUT)
    except requests.RequestException as exc:
        raise ErroSheetsApi(f"Falha de rede ao chamar a API ({exc}).") from exc
    try:
        resposta = r.json()
    except ValueError as exc:
        raise ErroSheetsApi(
            f"A API não devolveu JSON (status {r.status_code}). "
            "Confira se a URL termina em /exec e se o app foi publicado como Web App."
        ) from exc
    if not resposta.get("success"):
        raise ErroSheetsApi(resposta.get("error") or f"Erro da API (status {r.status_code}).")
    return resposta


def criar_linha(url: str, chave: str, aba: str, campos: dict) -> dict:
    """Cria uma linha na aba. Sem coluna-chave no payload, a API gera uma."""
    return _post(url, chave, {"action": "create", "aba": aba, **campos}).get("data", {})


def atualizar_linha(url: str, chave: str, aba: str, campos: dict) -> dict:
    """Atualiza a linha cuja coluna-chave da aba está em `campos` (ex.: ID_Tarefa)."""
    return _post(url, chave, {"action": "update", "aba": aba, **campos}).get("data", {})


def excluir_linha(url: str, chave: str, aba: str, chave_linha) -> dict:
    """Exclui a linha identificada pelo valor da coluna-chave da aba."""
    return _post(
        url, chave, {"action": "delete", "aba": aba, "chave": str(chave_linha)}
    ).get("data", {})


def ping(url: str, chave: str) -> bool:
    """Testa se a API está no ar e a chave é válida."""
    try:
        _get(url, chave, {"action": "ping"})
        return True
    except ErroSheetsApi:
        return False


def info(url: str, chave: str) -> dict:
    """Estrutura da planilha: abas encontradas, colunas e nº de linhas."""
    return _get(url, chave, {"action": "info"}).get("sheets", {})


def carregar_aba(url: str, chave: str, aba: str) -> pd.DataFrame:
    """Baixa uma aba inteira como DataFrame (mesmo formato do CSV local)."""
    corpo = _get(url, chave, {"action": "list", "aba": aba})
    registros = corpo.get("data", [])
    if not registros:
        return pd.DataFrame()
    df = pd.DataFrame(registros)
    # A coluna interna _linha é detalhe da API; o app não usa.
    df = df.drop(columns=["_linha"], errors="ignore").fillna("")
    return df


def carregar_tudo(url: str, chave: str) -> tuple[dict, list[str]]:
    """Baixa as 6 abas. Devolve ({chave: DataFrame}, [avisos de aba que falhou])."""
    dados: dict[str, pd.DataFrame] = {}
    avisos: list[str] = []
    for chave_aba, nome_aba in ABAS.items():
        try:
            df = carregar_aba(url, chave, nome_aba)
        except ErroSheetsApi as exc:
            dados[chave_aba] = pd.DataFrame()
            avisos.append(f"Aba {nome_aba} indisponível na API ({exc}) — usando CSV local para ela.")
            continue
        if df.empty:
            avisos.append(f"Aba {nome_aba} vazia na planilha — usando CSV local para ela.")
            dados[chave_aba] = pd.DataFrame()
        else:
            dados[chave_aba] = df
    return dados, avisos
