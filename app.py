#!/usr/bin/env python3
"""
App Gestão de Projetos — painel Streamlit para cronograma, histogramas,
recursos humanos e contratos.

Fonte de dados: Google Sheets (as 6 abas DB_*/Timeline_Visao) com fallback
automático para os CSVs locais de exemplo em data/exemplos/.

Uso:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import contextlib
import importlib.util
import io
import re
import tempfile
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

st.set_page_config(page_title="App Gestão de Projetos", page_icon="📋", layout="wide")

BASE = Path(__file__).resolve().parent
PASTA_EXEMPLOS = BASE / "data" / "exemplos"

# Reusa o conversor que já existe na raiz (o nome tem hífen → import dinâmico).
_spec = importlib.util.spec_from_file_location(
    "conversor_cronograma", BASE / "conversor-cronograma.py"
)
conversor = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(conversor)

# Client da API do Google Sheets (Web App do Apps Script) — import dinâmico
# pelo mesmo motivo do conversor (o nome do arquivo tem underscore e é simples,
# mas mantemos o spec_from_file_location para carregar do diretório base).
_spec_api = importlib.util.spec_from_file_location("sheets_api", BASE / "sheets_api.py")
sheets_api = importlib.util.module_from_spec(_spec_api)
_spec_api.loader.exec_module(sheets_api)

COLUNAS_CRONOGRAMA: list[str] = list(conversor.SHEETS_HEADER)
STATUS_VALIDOS = {"não iniciado", "em andamento", "concluído", "atrasado"}

ABAS = {
    "cronograma": "DB_Cronograma_MSProject",
    "hist_mo": "DB_Histograma_MO",
    "hist_equip": "DB_Histograma_Equip",
    "rh": "DB_Recursos_Humanos",
    "contratos": "DB_Contratos",
    "timeline": "Timeline_Visao",
}

PAGINAS = [
    # Ordem = fluxo de trabalho: cadastrar/importar → analisar → detalhar → editar.
    "📑 Contratos",
    "🔁 Importar / Exportar",
    "📊 Dashboard",
    "📅 Cronograma Timeline",
    "👷 Histograma MO + Equip",
    "👥 Recursos Humanos",
    "📝 Relatório (IA)",
    "✏️ Editar Dados",
]


# ------------------------------------------------------------ estilo visual
# Tema escuro moderno que espelha .streamlit/config.toml. Os seletores usam os
# data-testid do Streamlit 1.66 (verificados no navegador com DOM real) — se o
# Streamlit for atualizado, confira os seletores de novo antes de mexer aqui.
# Tema claro/escuro: só as variáveis abaixo mudam; o corpo do CSS é único e
# referencia tudo por var(--...). Por que CSS e não o tema nativo do Streamlit:
# o config.toml só é lido na abertura da sessão (NewSession) — trocar
# st._config.set_option em runtime não reenvia o tema na sessão atual.
TEMAS: dict[str, dict[str, str]] = {
    "Escuro": {
        "modo": "escuro",
        "fundo": "#0B1220",
        "superficie": "#101A2E",
        "borda": "rgba(148, 163, 184, 0.16)",
        "texto": "#E2E8F0",
        "suave": "#94A3B8",
        "acento": "#818CF8",
        "acento_forte": "#6366F1",
        "acento2": "#22D3EE",  # 2ª cor dos gradientes (ícone, progresso, aba)
        "acento_sombra": "rgba(99, 102, 241, 0.35)",
        # gradiente da barra lateral
        "sidebar": "linear-gradient(178deg, #101A2E 0%, #0C1526 100%)",
        "marca_cor": "#FFFFFF",
        # texto sobre o fundo de destaque (rádio/aba/opção selecionada)
        "sel_cor": "#FFFFFF",
        "sel_fundo": "linear-gradient(90deg, rgba(99, 102, 241, 0.30), rgba(34, 211, 238, 0.10))",
        "prim_cor": "#FFFFFF",  # texto do botão primário
        # overlays que dão "corpo" a botões, cartões e alertas
        "overlay_1": "rgba(255, 255, 255, 0.04)",
        "overlay_2": "rgba(255, 255, 255, 0.03)",
        "card_grad": "linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.015))",
        "metric_grad": (
            "linear-gradient(180deg, rgba(99,102,241,0.14), rgba(99,102,241,0.03) 65%, "
            "rgba(255,255,255,0.02))"
        ),
        "sombra_1": "rgba(0, 0, 0, 0.22)",
        "sombra_2": "rgba(0, 0, 0, 0.3)",
        "sombra_3": "rgba(0, 0, 0, 0.5)",
        "sombra_hover": "rgba(0, 0, 0, 0.35)",
        # cartão translúcido (glass) + luzes de fundo
        "glass": "rgba(16, 26, 46, 0.66)",
        "blob_1": "rgba(99, 102, 241, 0.16)",
        "blob_2": "rgba(34, 211, 238, 0.12)",
        # cores dos gráficos Plotly
        "grade": "rgba(148, 163, 184, 0.12)",
        "linha_eixo": "rgba(148, 163, 184, 0.28)",
    },
    "Claro": {
        "modo": "claro",
        "fundo": "#F1F5F9",
        "superficie": "#FFFFFF",
        "borda": "rgba(15, 23, 42, 0.14)",
        "texto": "#0F172A",
        "suave": "#64748B",
        "acento": "#6366F1",
        "acento_forte": "#4F46E5",
        "acento2": "#22D3EE",
        "acento_sombra": "rgba(99, 102, 241, 0.28)",
        "sidebar": "linear-gradient(178deg, #FFFFFF 0%, #E9EEF5 100%)",
        "marca_cor": "#0F172A",
        "sel_cor": "#0F172A",
        "sel_fundo": "linear-gradient(90deg, rgba(99, 102, 241, 0.16), rgba(34, 211, 238, 0.10))",
        "prim_cor": "#FFFFFF",
        "overlay_1": "rgba(255, 255, 255, 0.9)",
        "overlay_2": "rgba(255, 255, 255, 0.75)",
        "card_grad": "linear-gradient(180deg, #FFFFFF, #FFFFFF)",
        "metric_grad": (
            "linear-gradient(180deg, rgba(99,102,241,0.10), rgba(99,102,241,0.02) 65%, "
            "rgba(255,255,255,0.9))"
        ),
        "sombra_1": "rgba(15, 23, 42, 0.08)",
        "sombra_2": "rgba(15, 23, 42, 0.14)",
        "sombra_3": "rgba(15, 23, 42, 0.18)",
        "sombra_hover": "rgba(15, 23, 42, 0.16)",
        "glass": "rgba(255, 255, 255, 0.78)",
        "blob_1": "rgba(99, 102, 241, 0.10)",
        "blob_2": "rgba(34, 211, 238, 0.10)",
        "grade": "rgba(15, 23, 42, 0.10)",
        "linha_eixo": "rgba(15, 23, 42, 0.18)",
        # paleta do dataframe nos temas claros (--gdg-*)
        "gdg_header": "#EEF2F7",
        "gdg_focus": "#E2E8F0",
        "gdg_hover": "#E9EEF5",
    },
    "Azul Profundo": {
        "modo": "escuro",
        "fundo": "#061527",
        "superficie": "#0C2340",
        "borda": "rgba(147, 197, 253, 0.18)",
        "texto": "#E8F3FF",
        "suave": "#93B8DC",
        "acento": "#38BDF8",
        "acento_forte": "#0EA5E9",
        "acento2": "#818CF8",
        "acento_sombra": "rgba(56, 189, 248, 0.35)",
        "sidebar": "linear-gradient(178deg, #0C2340 0%, #061A2E 100%)",
        "marca_cor": "#FFFFFF",
        "sel_cor": "#FFFFFF",
        "sel_fundo": "linear-gradient(90deg, rgba(56, 189, 248, 0.28), rgba(129, 140, 248, 0.14))",
        "prim_cor": "#FFFFFF",
        "overlay_1": "rgba(255, 255, 255, 0.05)",
        "overlay_2": "rgba(255, 255, 255, 0.04)",
        "card_grad": "linear-gradient(180deg, rgba(255,255,255,0.06), rgba(255,255,255,0.02))",
        "metric_grad": (
            "linear-gradient(180deg, rgba(56,189,248,0.16), rgba(56,189,248,0.04) 65%, "
            "rgba(255,255,255,0.03))"
        ),
        "sombra_1": "rgba(2, 8, 20, 0.35)",
        "sombra_2": "rgba(2, 8, 20, 0.45)",
        "sombra_3": "rgba(2, 8, 20, 0.55)",
        "sombra_hover": "rgba(2, 8, 20, 0.5)",
        "glass": "rgba(12, 35, 64, 0.66)",
        "blob_1": "rgba(56, 189, 248, 0.15)",
        "blob_2": "rgba(129, 140, 248, 0.14)",
        "grade": "rgba(147, 197, 253, 0.12)",
        "linha_eixo": "rgba(147, 197, 253, 0.3)",
    },
    "Verde Engenharia": {
        "modo": "escuro",
        "fundo": "#071711",
        "superficie": "#0E2A1C",
        "borda": "rgba(110, 231, 183, 0.16)",
        "texto": "#E9F7F0",
        "suave": "#8FC2AE",
        "acento": "#34D399",
        "acento_forte": "#10B981",
        "acento2": "#A3E635",
        "acento_sombra": "rgba(52, 211, 153, 0.35)",
        "sidebar": "linear-gradient(178deg, #0E2A1C 0%, #071811 100%)",
        "marca_cor": "#FFFFFF",
        "sel_cor": "#FFFFFF",
        "sel_fundo": "linear-gradient(90deg, rgba(52, 211, 153, 0.28), rgba(163, 230, 53, 0.14))",
        "prim_cor": "#FFFFFF",
        "overlay_1": "rgba(255, 255, 255, 0.05)",
        "overlay_2": "rgba(255, 255, 255, 0.04)",
        "card_grad": "linear-gradient(180deg, rgba(255,255,255,0.055), rgba(255,255,255,0.018))",
        "metric_grad": (
            "linear-gradient(180deg, rgba(52,211,153,0.15), rgba(52,211,153,0.04) 65%, "
            "rgba(255,255,255,0.03))"
        ),
        "sombra_1": "rgba(3, 14, 9, 0.32)",
        "sombra_2": "rgba(3, 14, 9, 0.42)",
        "sombra_3": "rgba(3, 14, 9, 0.55)",
        "sombra_hover": "rgba(3, 14, 9, 0.48)",
        "glass": "rgba(14, 42, 28, 0.66)",
        "blob_1": "rgba(52, 211, 153, 0.14)",
        "blob_2": "rgba(163, 230, 53, 0.10)",
        "grade": "rgba(134, 239, 172, 0.12)",
        "linha_eixo": "rgba(134, 239, 172, 0.3)",
    },
    "Sepia Documento": {
        "modo": "claro",
        "fundo": "#F4EEE2",
        "superficie": "#FFFCF5",
        "borda": "rgba(120, 95, 60, 0.22)",
        "texto": "#2F2A20",
        "suave": "#7C6F5A",
        "acento": "#B45309",
        "acento_forte": "#92400E",
        "acento2": "#D97706",
        "acento_sombra": "rgba(180, 83, 9, 0.28)",
        "sidebar": "linear-gradient(178deg, #FFFCF5 0%, #EDE5D5 100%)",
        "marca_cor": "#2F2A20",
        "sel_cor": "#2F2A20",
        "sel_fundo": "linear-gradient(90deg, rgba(180, 83, 9, 0.16), rgba(217, 119, 6, 0.10))",
        "prim_cor": "#FFFFFF",
        "overlay_1": "rgba(255, 255, 255, 0.85)",
        "overlay_2": "rgba(255, 255, 255, 0.7)",
        "card_grad": "linear-gradient(180deg, #FFFFFF, #FFFDF6)",
        "metric_grad": (
            "linear-gradient(180deg, rgba(180,83,9,0.12), rgba(180,83,9,0.03) 65%, "
            "rgba(255,255,255,0.85))"
        ),
        "sombra_1": "rgba(63, 48, 23, 0.10)",
        "sombra_2": "rgba(63, 48, 23, 0.16)",
        "sombra_3": "rgba(63, 48, 23, 0.20)",
        "sombra_hover": "rgba(63, 48, 23, 0.18)",
        "glass": "rgba(255, 252, 245, 0.78)",
        "blob_1": "rgba(180, 83, 9, 0.10)",
        "blob_2": "rgba(217, 119, 6, 0.08)",
        "grade": "rgba(63, 48, 23, 0.10)",
        "linha_eixo": "rgba(63, 48, 23, 0.20)",
        "gdg_header": "#F1E9DA",
        "gdg_focus": "#E7DCC8",
        "gdg_hover": "#EDE5D5",
    },
    "Alto Contraste": {
        "modo": "escuro",
        "fundo": "#000000",
        "superficie": "#0B0B0B",
        "borda": "rgba(255, 255, 255, 0.55)",
        "texto": "#FFFFFF",
        "suave": "#E4E4E4",
        "acento": "#FFD400",
        "acento_forte": "#FFC400",
        "acento2": "#00E5FF",
        "acento_sombra": "rgba(255, 212, 0, 0.35)",
        "sidebar": "linear-gradient(178deg, #0B0B0B 0%, #000000 100%)",
        "marca_cor": "#FFFFFF",
        "sel_cor": "#000000",
        "sel_fundo": "linear-gradient(90deg, #FFD400, #FFD400)",
        "prim_cor": "#000000",  # botão primário amarelo com texto preto
        "overlay_1": "rgba(255, 255, 255, 0.08)",
        "overlay_2": "rgba(255, 255, 255, 0.06)",
        "card_grad": "linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.02))",
        "metric_grad": (
            "linear-gradient(180deg, rgba(255,212,0,0.14), rgba(255,212,0,0.04) 65%, "
            "rgba(255,255,255,0.03))"
        ),
        "sombra_1": "rgba(0, 0, 0, 0)",
        "sombra_2": "rgba(0, 0, 0, 0)",
        "sombra_3": "rgba(0, 0, 0, 0)",
        "sombra_hover": "rgba(255, 255, 255, 0.16)",
        "glass": "rgba(11, 11, 11, 0.92)",
        "blob_1": "rgba(255, 212, 0, 0.10)",
        "blob_2": "rgba(0, 229, 255, 0.08)",
        "grade": "rgba(255, 255, 255, 0.22)",
        "linha_eixo": "rgba(255, 255, 255, 0.5)",
    },
}


def tema_ativo() -> str:
    """Tema escolhido na barra lateral (padrão Escuro)."""
    t = str(st.session_state.get("tema", "Escuro"))
    return t if t in TEMAS else "Escuro"


def cores_tema() -> dict[str, str]:
    """Paleta do Plotly (CORES) alinhada ao tema ativo."""
    return TEMAS[tema_ativo()]


def _css_raiz(tema: str) -> str:
    t = TEMAS[tema]
    return (
        ":root {\n"
        f"  --fundo: {t['fundo']};\n"
        f"  --superficie: {t['superficie']};\n"
        f"  --borda: {t['borda']};\n"
        f"  --texto: {t['texto']};\n"
        f"  --suave: {t['suave']};\n"
        f"  --acento: {t['acento']};\n"
        f"  --acento-forte: {t['acento_forte']};\n"
        f"  --acento2: {t['acento2']};\n"
        f"  --acento-sombra: {t['acento_sombra']};\n"
        f"  --sidebar-grad: {t['sidebar']};\n"
        f"  --marca-cor: {t['marca_cor']};\n"
        f"  --sel-cor: {t['sel_cor']};\n"
        f"  --sel-fundo: {t['sel_fundo']};\n"
        f"  --prim-cor: {t['prim_cor']};\n"
        f"  --overlay-1: {t['overlay_1']};\n"
        f"  --overlay-2: {t['overlay_2']};\n"
        f"  --card-grad: {t['card_grad']};\n"
        f"  --metric-grad: {t['metric_grad']};\n"
        f"  --sombra-1: {t['sombra_1']};\n"
        f"  --sombra-2: {t['sombra_2']};\n"
        f"  --sombra-3: {t['sombra_3']};\n"
        f"  --sombra-hover: {t['sombra_hover']};\n"
        f"  --glass: {t['glass']};\n"
        f"  --blob-1: {t['blob_1']};\n"
        f"  --blob-2: {t['blob_2']};\n"
        "}\n\n"
    )


_CSS_ABERTURA = (
    "<style>\n"
    "@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800"
    "&display=swap');\n\n"
)

_CSS_CORPO = """

/* ---------- base ---------- */
/* O Streamlit pinta o .stApp com a cor fixa do config.toml (só vale na
   abertura da sessão) — aqui o fundo SEMPRE segue a paleta escolhida, senão
   tema claro fica texto escuro sobre fundo escuro. */
html, body, .stApp { background: var(--fundo) !important; }
/* header/toolbar do Streamlit também vêm pintados com a cor do config.toml */
[data-testid="stHeader"], [data-testid="stToolbar"] { background: var(--fundo) !important; }
.stApp, .stApp input, .stApp button, .stApp textarea {
  font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;
}
/* luzes de fundo que dão profundidade atrás do glass dos cartões */
.stApp::before {
  content: ""; position: fixed; inset: 0; z-index: -1; pointer-events: none;
  background:
    radial-gradient(620px 420px at 10% -8%, var(--blob-1), transparent 62%),
    radial-gradient(720px 520px at 96% 6%, var(--blob-2), transparent 58%),
    radial-gradient(840px 620px at 48% 112%, var(--blob-1), transparent 64%);
}
/* entrada suave do cabeçalho e dos KPIs */
@keyframes surge {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: none; }
}
.pag-cab, [data-testid="stMetric"] { animation: surge 0.34s ease-out both; }
/* drill-down dos KPIs: o cartão inteiro é o gatilho — o botão do popover fica
   invisível por cima do metric; hover acende um anel e o clique abre a lista */
[data-testid="stColumn"]:has([data-testid="stPopover"]) { position: relative; }
[data-testid="stColumn"]:has([data-testid="stPopover"]) [data-testid="stPopover"] {
  position: absolute; inset: 0; z-index: 1;
}
[data-testid="stColumn"]:has([data-testid="stPopover"]) [data-testid="stPopover"] button {
  position: absolute; inset: 0; width: 100%; height: 100%;
  background: transparent; border: 0; opacity: 0; cursor: pointer;
}
[data-testid="stColumn"]:has([data-testid="stPopover"]):hover [data-testid="stMetric"] {
  border-color: var(--acento) !important;
  box-shadow: 0 0 0 1px var(--acento), 0 14px 32px var(--acento-sombra);
  transform: translateY(-2px);
  transition: border-color .18s ease, box-shadow .18s ease, transform .18s ease;
}
[data-testid="stColumn"]:has([data-testid="stPopover"] button[aria-expanded="true"]) [data-testid="stMetric"] {
  border-color: var(--acento) !important;
}
[data-testid="stColumn"]:has([data-testid="stPopover"] button:focus-visible) [data-testid="stMetric"] {
  outline: 2px solid var(--acento); outline-offset: 3px;
}
/* gatilho do tema 🎨 na árvore lateral: item de largura total, estilo glass */
[data-testid="stSidebar"] [data-testid="stPopover"] > button {
  width: 100%; justify-content: space-between; border-radius: 12px;
  border: 1px solid var(--borda); background: var(--superficie);
  color: var(--texto); font-weight: 600; padding: 8px 14px; margin-bottom: 4px;
}
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after { animation: none !important; transition: none !important; }
}
::selection { background: var(--acento-forte); color: var(--prim-cor); }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(148, 163, 184, 0.3); border-radius: 8px; }
::-webkit-scrollbar-thumb:hover { background: rgba(148, 163, 184, 0.5); }
footer { visibility: hidden; }

/* ---------- tipografia ---------- */
.stApp h1 { font-size: 28px; font-weight: 800; letter-spacing: -0.6px; color: var(--texto); }
.stApp h2 { font-size: 23px; font-weight: 800; letter-spacing: -0.5px; color: var(--texto); }
.stApp h3 { font-size: 18.5px; font-weight: 700; color: var(--texto); }
.stApp h4 {
  position: relative; display: inline-block;
  font-size: 16px; font-weight: 700; letter-spacing: -0.2px;
  color: var(--texto); padding-bottom: 7px; margin: 16px 0 8px;
}
.stApp h4::after {
  content: ""; position: absolute; left: 0; bottom: 0;
  height: 3px; width: 44px; border-radius: 3px;
  background: linear-gradient(90deg, var(--acento-forte), var(--acento2));
}
.stApp [data-testid="stMarkdownContainer"] p { line-height: 1.6; }
[data-testid="stCaptionContainer"] p { color: var(--suave) !important; font-size: 12.5px !important; }
[data-testid="stWidgetLabel"] p {
  color: var(--suave) !important; font-size: 12.5px !important;
  font-weight: 600 !important; letter-spacing: 0.3px;
}

/* ---------- cabeçalho de página ---------- */
.pag-cab { display: flex; align-items: center; gap: 14px; margin: 2px 0 14px; }
.pag-ico {
  flex: 0 0 auto; width: 46px; height: 46px; border-radius: 13px;
  display: flex; align-items: center; justify-content: center; font-size: 22px;
  background: linear-gradient(135deg, var(--acento), var(--acento2));
  box-shadow: 0 8px 20px var(--acento-sombra);
  transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.pag-ico:hover { transform: rotate(-5deg) scale(1.08); box-shadow: 0 10px 26px var(--acento-sombra); }
.pag-tit h2 { margin: 0; font-size: 25px; line-height: 1.15; }
.pag-tit p { margin: 3px 0 0; color: var(--suave); font-size: 13.5px; line-height: 1.4; }
.pag-tit .pag-ass { font-style: italic; font-size: 12.5px; color: var(--acento); opacity: .85; margin-top: 1px; }

/* ---------- barra lateral ---------- */
[data-testid="stSidebar"] {
  background: var(--sidebar-grad) !important;
  border-right: 1px solid var(--borda);
}
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p { font-size: 12px; }
.marca {
  display: flex; align-items: center; gap: 12px;
  padding: 2px 0 14px; margin-bottom: 10px; border-bottom: 1px solid var(--borda);
}
.marca-ico {
  flex: 0 0 auto; width: 42px; height: 42px; border-radius: 12px;
  display: flex; align-items: center; justify-content: center; font-size: 20px;
  background: linear-gradient(135deg, var(--acento), var(--acento2));
  box-shadow: 0 6px 18px var(--acento-sombra);
  transition: transform 0.2s ease;
}
.marca-ico:hover { transform: rotate(5deg) scale(1.08); }
.marca-nome { font-size: 16.5px; font-weight: 800; letter-spacing: -0.3px; color: var(--marca-cor); line-height: 1.15; }
.marca-sub {
  margin-top: 2px; font-size: 10.5px; font-weight: 600;
  letter-spacing: 1.2px; text-transform: uppercase; color: var(--suave);
}
[data-testid="stRadioGroup"] { gap: 4px; }
[data-testid="stRadioOption"] {
  display: flex; align-items: center; gap: 10px;
  padding: 9px 12px !important; border: 1px solid transparent; border-radius: 10px;
  color: var(--texto);
  transition: background 0.15s ease, border-color 0.15s ease, box-shadow 0.15s ease;
}
[data-testid="stRadioOption"]:hover {
  background: color-mix(in srgb, var(--acento) 12%, transparent);
  border-color: color-mix(in srgb, var(--acento) 35%, transparent);
}
[data-testid="stRadioOption"][data-selected="true"] {
  background: var(--sel-fundo);
  border-color: var(--acento);
  box-shadow: 0 4px 14px var(--acento-sombra);
}
[data-testid="stRadioOption"][data-selected="true"] p { color: var(--sel-cor) !important; font-weight: 600; }

/* ---------- botões ---------- */
button[data-testid="stBaseButton-secondary"],
button[data-testid="stBaseButton-tertiary"] {
  border: 1px solid var(--borda) !important;
  background: var(--overlay-1) !important;
  color: var(--texto) !important;
  border-radius: 10px !important; font-weight: 600 !important;
  transition: all 0.15s ease !important;
}
button[data-testid="stBaseButton-secondary"]:hover,
button[data-testid="stBaseButton-tertiary"]:hover {
  background: color-mix(in srgb, var(--acento) 14%, transparent) !important;
  border-color: color-mix(in srgb, var(--acento) 55%, transparent) !important;
  transform: translateY(-1px); box-shadow: 0 6px 16px var(--sombra-2);
}
button[data-testid="stBaseButton-primary"] {
  background: linear-gradient(135deg, var(--acento-forte), var(--acento)) !important;
  border: none !important; color: var(--prim-cor) !important;
  border-radius: 10px !important; font-weight: 700 !important;
  box-shadow: 0 6px 18px var(--acento-sombra);
  transition: all 0.15s ease !important;
}
button[data-testid="stBaseButton-primary"]:hover {
  transform: translateY(-2px) scale(1.015);
  box-shadow: 0 10px 26px var(--acento-sombra);
  filter: brightness(1.07) saturate(1.1);
}
button:disabled { opacity: 0.45; transform: none !important; box-shadow: none !important; }

/* ---------- campos ---------- */
[data-testid="stSelectbox"] [role="group"] {
  background: var(--superficie) !important;
  border-color: var(--borda) !important; border-radius: 10px !important;
}
[data-testid="stSelectbox"] input[role="combobox"] {
  background: transparent !important; border: none !important; color: var(--texto) !important;
}
[data-testid="stSelectbox"] button { color: var(--suave) !important; }
[data-testid="stSelectbox"]:focus-within [role="group"] {
  border-color: var(--acento) !important;
  box-shadow: 0 0 0 3px color-mix(in srgb, var(--acento) 25%, transparent);
}
[data-testid="stTextInput"] div:has(> input) {
  background: var(--superficie) !important;
  border: 1px solid var(--borda) !important; border-radius: 10px !important;
}
[data-testid="stTextInput"] input { background: transparent !important; border: none !important; color: var(--texto) !important; }
[data-testid="stTextInput"] div:has(> input:focus) {
  border-color: var(--acento) !important;
  box-shadow: 0 0 0 3px color-mix(in srgb, var(--acento) 25%, transparent);
}
[data-testid="portal"] [role="listbox"], [role="listbox"] {
  background: var(--superficie) !important; border: 1px solid var(--borda);
  border-radius: 10px; box-shadow: 0 14px 34px var(--sombra-3); padding: 4px;
}
[role="option"] { border-radius: 7px; padding: 7px 10px; color: var(--texto); font-size: 14px; }
[role="option"][data-hovered="true"] {
  background: color-mix(in srgb, var(--acento) 16%, transparent);
}
[role="option"][data-selected="true"] {
  background: var(--sel-fundo);
  font-weight: 600;
}

/* ---------- abas ---------- */
[data-testid="stTab"] {
  border-radius: 10px 10px 0 0; color: var(--suave) !important;
  font-weight: 600 !important; font-size: 14.5px !important; padding: 10px 16px !important;
  transition: color 0.15s ease, background 0.15s ease;
}
[data-testid="stTab"]:hover {
  color: var(--texto) !important;
  background: color-mix(in srgb, var(--acento) 10%, transparent);
}
[data-testid="stTab"][aria-selected="true"] {
  color: var(--sel-cor) !important;
  background: var(--sel-fundo);
}
[data-testid="stTab"] .react-aria-SelectionIndicator {
  background: linear-gradient(90deg, var(--acento), var(--acento2)) !important;
  height: 3px !important; border-radius: 3px !important;
}
div[role="tablist"] { box-shadow: inset 0 -1px 0 var(--borda); }

/* ---------- KPIs e progresso ---------- */
[data-testid="stMetric"] {
  background: var(--metric-grad) !important;
  border: 1px solid var(--borda) !important;
  border-radius: 14px !important; padding: 16px 18px 14px !important;
  box-shadow: 0 2px 10px var(--sombra-1);
  transition: transform 0.15s ease, border-color 0.15s ease;
}
[data-testid="stMetric"]:hover {
  transform: translateY(-3px);
  border-color: var(--acento) !important;
  box-shadow: 0 10px 26px var(--acento-sombra);
}
[data-testid="stMetric"] > div { background: transparent !important; border: none !important; padding: 0 !important; }
[data-testid="stElementContainer"]:has(> [data-testid="stMetric"]) {
  background: transparent !important; border: none !important; padding: 0 !important;
}
[data-testid="stMetricLabel"] p {
  color: var(--suave) !important; font-size: 12px !important;
  font-weight: 600 !important; letter-spacing: 0.6px !important; text-transform: uppercase;
}
[data-testid="stMetricValue"] p {
  color: var(--texto) !important; font-size: 30px !important;
  font-weight: 800 !important; letter-spacing: -0.7px !important;
}
[data-testid="stMetricDelta"] p { font-size: 13px !important; font-weight: 600 !important; }
[data-testid="stProgressBarTrack"] { background: rgba(148, 163, 184, 0.18) !important; border-radius: 999px !important; }
[data-testid="stProgressBarTrack"] > div {
  background: linear-gradient(90deg, var(--acento), var(--acento2)) !important;
  border-radius: 999px !important;
}

/* ---------- gráficos, cartões e expansores ---------- */
[data-testid="stPlotlyChart"] {
  background: var(--card-grad) !important;
  border: 1px solid var(--borda) !important;
  border-radius: 16px !important; padding: 8px 12px 2px !important;
  box-shadow: 0 6px 22px var(--sombra-1);
}
[data-testid="stLayoutWrapper"] > [data-testid="stVerticalBlock"] {
  background: var(--glass) !important;
  backdrop-filter: blur(16px) saturate(150%);
  -webkit-backdrop-filter: blur(16px) saturate(150%);
  border: 1px solid var(--borda) !important;
  border-radius: 16px !important; padding: 20px 22px !important;
  box-shadow: 0 8px 26px var(--sombra-1);
  transition: border-color 0.15s ease, transform 0.15s ease, box-shadow 0.15s ease;
}
[data-testid="stLayoutWrapper"] > [data-testid="stVerticalBlock"]:hover {
  border-color: var(--acento) !important;
  transform: translateY(-3px); box-shadow: 0 14px 34px var(--sombra-hover);
}
[data-testid="stExpander"] details {
  background: var(--overlay-2) !important;
  border: 1px solid var(--borda) !important; border-radius: 12px !important;
}
[data-testid="stExpander"] summary { color: var(--texto) !important; font-weight: 600 !important; font-size: 14.5px !important; }
[data-testid="stExpander"] summary:hover { color: var(--acento) !important; }
[data-testid="stFileUploaderDropzone"] {
  border: 1.5px dashed color-mix(in srgb, var(--acento) 55%, transparent) !important;
  border-radius: 14px !important;
  background: color-mix(in srgb, var(--acento) 7%, transparent) !important;
  transition: border-color 0.15s ease, background 0.15s ease, transform 0.15s ease;
}
[data-testid="stFileUploaderDropzone"]:hover {
  border-color: var(--acento) !important;
  background: color-mix(in srgb, var(--acento) 13%, transparent) !important;
  transform: translateY(-1px);
}
[data-testid="stDataFrame"], [data-testid="stArrowDataFrame"] {
  border: 1px solid var(--borda); border-radius: 12px; overflow: hidden;
}

/* ---------- alertas ---------- */
[data-testid="stAlertContainer"] {
  background: var(--overlay-1) !important;
  border: 1px solid var(--borda) !important;
  border-left-width: 3px !important; border-radius: 12px !important;
}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentError"]) {
  border-left-color: #F87171 !important; background: rgba(248, 113, 113, 0.08) !important;
}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentWarning"]) {
  border-left-color: #FBBF24 !important; background: rgba(251, 191, 36, 0.08) !important;
}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentInfo"]) {
  border-left-color: #60A5FA !important; background: rgba(96, 165, 250, 0.08) !important;
}
[data-testid="stAlertContainer"]:has([data-testid="stAlertContentSuccess"]) {
  border-left-color: #34D399 !important; background: rgba(52, 211, 153, 0.08) !important;
}

/* ---------- celular ---------- */
@media (max-width: 768px) {
  .pag-ico { width: 38px; height: 38px; font-size: 18px; border-radius: 11px; }
  .pag-tit h2 { font-size: 20px; }
  [data-testid="stMetric"] { padding: 12px 14px !important; }
  [data-testid="stMetricValue"] p { font-size: 24px !important; }
  [data-testid="stLayoutWrapper"] > [data-testid="stVerticalBlock"] { padding: 14px !important; }
}
"""

# Ajustes extras aplicados NOS TEMAS CLAROS: o config.toml do Streamlit fica
# escuro (lido só na abertura da sessão), então a cor nativa de texto parágrafo
# e a paleta do dataframe (--gdg-*) precisam ser corrigidas por aqui.
def _css_claro(tema: str) -> str:
    t = TEMAS[tema]
    return f"""
<style>
.stApp {{ color: var(--texto); }}
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] td {{ color: var(--texto) !important; }}
[data-testid="stCodeBlock"] {{
  background: var(--superficie) !important; color: var(--texto) !important;
  border: 1px solid var(--borda);
}}
[data-testid="stDataFrame"], [data-testid="stArrowDataFrame"] {{
  --gdg-bg-cell: {t['superficie']};
  --gdg-bg-cell-medium: {t['fundo']};
  --gdg-bg-header: {t.get('gdg_header', t['fundo'])};
  --gdg-bg-header-has-focus: {t.get('gdg_focus', t['borda'])};
  --gdg-bg-header-hovered: {t.get('gdg_hover', t['fundo'])};
  --gdg-text-header: {t['texto']};
  --gdg-text-group-header: {t['texto']};
  --gdg-bg-group-header: {t.get('gdg_header', t['fundo'])};
  --gdg-bg-bubble: {t['acento_sombra']};
  --gdg-border-color: {t['borda']};
  --gdg-shadow-none: none;
}}
</style>
"""


def css_moderno(tema: str = "Escuro") -> str:
    """CSS completo do tema pedido (uma das paletas de TEMAS)."""
    base = _CSS_ABERTURA + _css_raiz(tema) + _CSS_CORPO + "</style>\n"
    if TEMAS[tema]["modo"] == "claro":
        base += _css_claro(tema)
    return base


# Alias mantido por compatibilidade com a documentação (CSS do tema escuro).
CSS_MODERNO = css_moderno("Escuro")


def cabecalho(icone: str, titulo: str, subtitulo: str) -> None:
    """Cabeçalho de página: ícone em bloco com gradiente + título e descrição."""
    st.markdown(
        f'<div class="pag-cab"><div class="pag-ico">{icone}</div>'
        f'<div class="pag-tit"><h2>{titulo}</h2><p>{subtitulo}</p>'
        f'<p class="pag-ass">por Jozemir Leite</p></div></div>',
        unsafe_allow_html=True,
    )


def estilizar_fig(fig: go.Figure) -> go.Figure:
    """Estilo comum de todos os gráficos Plotly (tema ativo, fundo transparente)."""
    cores = cores_tema()
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", size=13, color=cores["texto"]),
        hoverlabel=dict(
            bgcolor=cores["superficie"],
            bordercolor=cores["linha_eixo"],
            font=dict(family="Inter, sans-serif", size=12, color=cores["texto"]),
        ),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=12.5, color=cores["texto"])),
    )
    fig.update_xaxes(gridcolor=cores["grade"], zeroline=False, linecolor=cores["linha_eixo"])
    fig.update_yaxes(gridcolor=cores["grade"], zeroline=False, linecolor=cores["linha_eixo"])
    if fig.data and isinstance(fig.data[0], go.Bar):
        fig.update_traces(marker=dict(cornerradius=6))
    return fig


# ----------------------------------------------------------------- formatação
def seg(valor, padrao):
    """Widgets devolvem None fora do servidor (modo bare, usado em testes)."""
    return padrao if valor is None else valor


def num_pt(valor, casas: int = 0) -> str:
    try:
        n = float(valor)
    except (TypeError, ValueError):
        return "[FALTANTE]"
    if pd.isna(n):
        return "[FALTANTE]"
    txt = f"{n:,.{casas}f}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


def brl(valor) -> str:
    try:
        n = float(valor)
    except (TypeError, ValueError):
        return "[FALTANTE]"
    if pd.isna(n):
        return "[FALTANTE]"
    return "R$ " + num_pt(n, 2)


def pct_txt(valor, casas: int = 0) -> str:
    try:
        n = float(valor)
    except (TypeError, ValueError):
        return "[FALTANTE]"
    if pd.isna(n):
        return "[FALTANTE]"
    return num_pt(n, casas) + "%"


def parse_datas(valores: pd.Series) -> pd.Series:
    """DD/MM/AAAA é o padrão; aceita também ISO e outros com dia primeiro."""
    texto = valores.astype(str).str.strip()
    vazio = texto.str.lower().isin(("", "nan", "none", "nat", "<na>", "nonetype"))
    texto = texto.mask(vazio)
    d = pd.to_datetime(texto, format="%d/%m/%Y", errors="coerce")
    pendente = d.isna() & texto.notna()
    if bool(pendente.any()):
        resto = texto[pendente]
        # ISO (AAAA-MM-DD, com ou sem hora): dayfirst puro trocaria mês/dia no pandas.
        m_iso = resto.str.match(r"^\d{4}-\d{2}-\d{2}")
        parte_iso = pd.to_datetime(resto.where(m_iso).str[:10], format="%Y-%m-%d", errors="coerce")
        parte_out = pd.to_datetime(resto.where(~m_iso), dayfirst=True, errors="coerce")
        d[pendente] = parte_iso.fillna(parte_out)
    return d


def parse_num(valores: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(valores):
        return valores.astype(float)
    texto = (
        valores.astype(str)
        .str.strip()
        .str.replace("R$", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
    )
    if bool(texto.str.contains(",").any()):
        texto = texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(texto, errors="coerce")


def fmt_data(valor) -> str:
    ts = parse_datas(pd.Series([valor])).iloc[0]
    return "[FALTANTE]" if pd.isna(ts) else ts.strftime("%d/%m/%Y")


def fmt_data_opcional(valor) -> str:
    texto = "" if valor is None else str(valor).strip()
    if texto.lower() in ("", "nan", "none", "nat"):
        return ""
    return fmt_data(valor)


def hoje() -> pd.Timestamp:
    return pd.Timestamp(datetime.now().date())


def chave_mes(valor):
    for fmt in ("%m/%Y", "%Y-%m", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(valor).strip(), fmt)
        except ValueError:
            continue
    return datetime.max


def ordenar_meses(valores) -> list[str]:
    unicos = list(
        dict.fromkeys(
            str(v).strip()
            for v in valores
            if str(v).strip() and str(v).strip().lower() not in ("nan", "none")
        )
    )
    return sorted(unicos, key=chave_mes)


# -------------------------------------------------------------- carga de dados
@st.cache_data(ttl=300, show_spinner=False)
def carregar_csv_local(aba: str) -> pd.DataFrame:
    caminho = PASTA_EXEMPLOS / f"{aba}.csv"
    if not caminho.exists():
        return pd.DataFrame()
    return pd.read_csv(caminho, sep=";", dtype=str, encoding="utf-8-sig").fillna("")


def _extrair_id_planilha(texto: str) -> str:
    texto = (texto or "").strip()
    m = re.search(r"/d/([A-Za-z0-9\-_]+)", texto)
    return m.group(1) if m else texto


def _abrir_sheets(id_planilha: str) -> tuple[dict, list[str]]:
    import gspread

    try:
        info = dict(st.secrets["gcp_service_account"])
    except Exception as exc:
        raise RuntimeError(
            "credenciais ausentes — defina [gcp_service_account] em .streamlit/secrets.toml (ver README)"
        ) from exc
    gc = gspread.service_account(info=info)
    planilha = gc.open_by_key(id_planilha)
    brutos: dict = {}
    erros: list[str] = []
    for chave, nome_aba in ABAS.items():
        try:
            brutos[chave] = pd.DataFrame(planilha.worksheet(nome_aba).get_all_records())
        except Exception as exc:
            brutos[chave] = None
            erros.append(f"Aba {nome_aba} indisponível no Sheets ({exc}) — usando CSV local para ela.")
    return brutos, erros


def carregar_dados(
    fonte: str, id_planilha: str = "", api_url: str = "", api_key: str = ""
) -> tuple[dict, dict, list[str]]:
    """Retorna (dados, origem_por_aba, avisos). Cada aba cai para o CSV local se falhar."""
    dados: dict[str, pd.DataFrame] = {}
    origem: dict[str, str] = {}
    avisos: list[str] = []
    if fonte == "Google Sheets (API)":
        if not api_url or not api_key:
            avisos.append(
                "API do Sheets selecionada, mas URL ou API Key ausente — usando CSV local."
            )
        else:
            try:
                brutos, erros = sheets_api.carregar_tudo(api_url, api_key)
                avisos.extend(erros)
                for chave, df in brutos.items():
                    if df is not None and not df.empty:
                        dados[chave] = df
                        origem[chave] = "Google Sheets (API)"
            except Exception as exc:
                avisos.append(f"Falha ao acessar a API do Sheets ({exc}) — usando CSV local.")
    elif fonte == "Google Sheets":
        if not id_planilha:
            avisos.append("Google Sheets selecionado, mas o ID/URL não foi informado — usando CSV local.")
        else:
            try:
                brutos, erros = _abrir_sheets(id_planilha)
                avisos.extend(erros)
                for chave, df in brutos.items():
                    if df is not None and not df.empty:
                        dados[chave] = df
                        origem[chave] = "Google Sheets"
            except Exception as exc:
                avisos.append(f"Falha ao acessar o Google Sheets ({exc}) — usando CSV local.")
    for chave, aba in ABAS.items():
        if chave not in dados:
            dados[chave] = carregar_csv_local(aba)
            origem[chave] = "CSV local"
            if dados[chave].empty:
                avisos.append(f"[FALTANTE] arquivo data/exemplos/{aba}.csv ausente ou vazio.")
    return dados, origem, avisos


# ---------------------------------------------------------- gravação de dados
# Coluna-chave de cada aba na API (igual CONFIG.ABAS do apps_script_API.gs).
# ID_Linha/Matricula/ID_Marco podem não existir nos CSVs locais: sem elas a API
# gera a chave sozinha no create, mas update/delete exigem a coluna presente.
CHAVES_ABA = {
    "cronograma": "ID_Tarefa",
    "hist_mo": "ID_Linha",
    "hist_equip": "ID_Linha",
    "rh": "Matricula",
    "contratos": "ID_Contrato",
    "timeline": "ID_Marco",
}

# Coluna que carrega o vínculo com o contrato em cada aba (usada nas
# importações/edições — nos histogramas ela se chama Contrato_Vinculado).
COLUNA_CONTRATO = {
    "cronograma": "Contrato",
    "hist_mo": "Contrato_Vinculado",
    "hist_equip": "Contrato_Vinculado",
    "rh": "Contrato",
    "contratos": "ID_Contrato",
    "timeline": "Projeto",
}

# Colunas exatas de cada aba (fonte: gestao-projetos (1).md).
SCHEMAS: dict[str, list[str]] = {
    "cronograma": [
        "ID_Tarefa", "Nome_Tarefa", "WBS", "Inicio_Planejado", "Fim_Planejado",
        "Inicio_Real", "Fim_Real", "Duracao_dias", "%_Concluido", "Predecessora",
        "Sucessora", "Caminho_Critico_SIM_NAO", "Responsavel", "Status", "Observacao",
    ],
    "hist_mo": [
        "Mes_Semana", "Funcao_Cargo", "Quantidade_Prevista", "Quantidade_Real",
        "HH_Previsto", "HH_Real", "Custo_Hora", "Custo_Total_Previsto",
        "Custo_Total_Real", "Projeto_Frente", "Contrato_Vinculado",
    ],
    "hist_equip": [
        "Mes_Semana", "Equipamento", "Qtd_Prevista", "Qtd_Real", "Horas_Previstas",
        "Horas_Reais", "Custo_Mensal", "Contrato_Vinculado", "Fornecedor", "Status",
    ],
    "rh": [
        "Matricula_Nome", "Funcao", "Equipe_Frente", "Contrato", "Data_Admissao",
        "Data_Desligamento", "Turno", "Custo_Mes", "Alocacao_%", "Status_Ativo_Inativo",
        "Certificacoes_NR",
    ],
    "contratos": [
        "ID_Contrato", "Cliente_Fornecedor", "Objeto_Escopo", "Valor_Total",
        "Valor_Medido_Acumulado", "Saldo", "Data_Inicio", "Data_Fim", "Reajuste",
        "Aditivos", "Responsavel_Gestao", "Status",
    ],
    "timeline": [
        "Projeto", "Frente", "Marco_Milestone", "Data_Marco", "Status", "%_Avanço",
        "Desvio_dias",
    ],
}


def validar_aba(chave_aba: str, df: pd.DataFrame) -> tuple[list[str], list[str]]:
    """(erros, avisos) antes de gravar. Cronograma usa o validador completo;
    as demais abas conferem o schema e a unicidade da coluna-chave."""
    if chave_aba == "cronograma":
        return validar_cronograma(df)
    erros: list[str] = []
    avisos: list[str] = []
    ausentes = [c for c in SCHEMAS.get(chave_aba, []) if c not in df.columns]
    if ausentes:
        avisos.append(f"[FALTANTE] coluna(s) do schema ausente(s): {', '.join(ausentes)}.")
    if df.empty:
        erros.append("Nenhuma linha para gravar.")
    chave_col = CHAVES_ABA.get(chave_aba)
    if chave_col and chave_col in df.columns:
        chaves = df[chave_col].astype(str).str.strip()
        repetidas = sorted(
            {c for c in chaves[(chaves != "") & chaves.duplicated(keep=False)]}
        )
        if repetidas:
            erros.append(
                f"Coluna-chave {chave_col} com valor repetido: "
                f"{', '.join(repetidas[:5])}{' …' if len(repetidas) > 5 else ''}."
            )
    return erros, avisos


def _valor_planilha(v):
    """Valor seguro para CSV/Sheets/JSON: NaN/NaT→'', data→DD/MM/AAAA, numpy→python."""
    try:
        if v is None or pd.isna(v):
            return ""
    except (TypeError, ValueError):  # listas/objetos sem isna definido
        pass
    if isinstance(v, (pd.Timestamp, datetime, date)):
        com_hora = (getattr(v, "hour", 0), getattr(v, "minute", 0), getattr(v, "second", 0))
        return str(v) if com_hora != (0, 0, 0) else v.strftime("%d/%m/%Y")
    if isinstance(v, str) and v.strip().lower() in ("nan", "none", "nat", "<na>", "nonetype"):
        return ""
    if hasattr(v, "item") and not isinstance(v, (str, bytes)):  # numpy.int64 etc.
        try:
            v = v.item()
        except Exception:  # pragma: no cover - defensivo
            pass
    if isinstance(v, float) and float(v).is_integer():
        return int(v)
    return v


def _registros_planilha(df: pd.DataFrame) -> list[dict]:
    """Linhas do DataFrame como dicts com valores já normalizados."""
    return [
        {c: _valor_planilha(v) for c, v in zip(df.columns, linha)}
        for linha in df.itertuples(index=False, name=None)
    ]


def _difere(orig: dict, novo: dict) -> bool:
    """True se algum campo de `novo` difere do original (comparação textual)."""
    return any(str(orig.get(c, "")) != str(v) for c, v in novo.items())


def _salvar_api(
    api_url: str,
    api_key: str,
    chave_aba: str,
    df: pd.DataFrame,
    original: pd.DataFrame | None,
) -> str:
    """Grava via POST (create/update/delete) comparando por coluna-chave."""
    nome_aba = ABAS[chave_aba]
    chave_col = CHAVES_ABA[chave_aba]
    novos = _registros_planilha(df)
    tem_chave_df = chave_col in df.columns
    tem_chave_orig = bool(
        original is not None and not original.empty and chave_col in original.columns
    )

    if not tem_chave_df:
        if tem_chave_orig:
            raise ValueError(
                f"A coluna-chave '{chave_col}' sumiu dos dados editados — sem ela não dá "
                "para identificar o que atualizar/excluir. Recarregue a página e edite de novo."
            )
        for registro in novos:  # base vazia/nova: tudo é linha nova
            sheets_api.criar_linha(api_url, api_key, nome_aba, registro)
        return f"{nome_aba}: {len(novos)} linha(s) criada(s)."

    originais: dict[str, dict] = {}
    for registro in _registros_planilha(original):
        k = str(registro.get(chave_col, "")).strip()
        if k and k not in originais:
            originais[k] = registro

    atuais: dict[str, dict] = {}
    for registro in novos:
        k = str(registro.get(chave_col, "")).strip()
        if k and k in atuais:
            raise ValueError(f"Coluna-chave {chave_col} com valor repetido: '{k}'.")
        atuais[k] = registro

    criar = [r for k, r in atuais.items() if not k or k not in originais]
    atualizar = [
        r for k, r in atuais.items() if k and k in originais and _difere(originais[k], r)
    ]
    excluir = [k for k in originais if k not in atuais]

    for registro in criar:
        sheets_api.criar_linha(api_url, api_key, nome_aba, registro)
    for registro in atualizar:
        sheets_api.atualizar_linha(api_url, api_key, nome_aba, registro)
    for k in excluir:
        sheets_api.excluir_linha(api_url, api_key, nome_aba, k)
    return (
        f"{nome_aba}: {len(criar)} criada(s), {len(atualizar)} atualizada(s), "
        f"{len(excluir)} excluída(s)."
    )


def _salvar_gspread(id_planilha: str, chave_aba: str, df: pd.DataFrame) -> str:
    """Reescreve a aba inteira na conta de serviço (clear + update em A1)."""
    import gspread

    try:
        info = dict(st.secrets["gcp_service_account"])
    except Exception as exc:
        raise RuntimeError(
            "credenciais ausentes — defina [gcp_service_account] em .streamlit/secrets.toml (ver README)"
        ) from exc
    gc = gspread.service_account(info=info)
    ws = gc.open_by_key(id_planilha).worksheet(ABAS[chave_aba])
    registros = _registros_planilha(df)
    valores = [list(df.columns)] + [[r[c] for c in df.columns] for r in registros]
    ws.clear()
    ws.update(valores, "A1")
    return f"{ABAS[chave_aba]}: {len(registros)} linha(s) regravada(s) no Google Sheets."


def _sem_linhas_vazias(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas 100% em branco (linhas novas que o data_editor deixa abertas)."""
    vazio = df.apply(lambda s: s.astype(str).str.strip() == "").all(axis=1)
    return df[~vazio].reset_index(drop=True)


def salvar_aba(
    chave_aba: str,
    df: pd.DataFrame,
    fonte: str,
    id_planilha: str = "",
    api_url: str = "",
    api_key: str = "",
    original: pd.DataFrame | None = None,
) -> str:
    """Grava a aba na fonte ativa e devolve mensagem de sucesso.

    - CSV local: reescreve data/exemplos/<aba>.csv (`;`, UTF-8 com BOM).
    - Google Sheets (API): diff por coluna-chave → POST create/update/delete.
    - Google Sheets: clear + update pela conta de serviço.
    `original` (dados como carregados) é exigido só pela API, para o diff.
    """
    if chave_aba not in ABAS:
        raise ValueError(f"Aba desconhecida: {chave_aba}.")
    nome_aba = ABAS[chave_aba]
    # Normaliza valores (NaN→'', datas→DD/MM/AAAA) e descarta linhas em branco.
    base = _sem_linhas_vazias(df.copy().apply(lambda s: s.map(_valor_planilha)))
    if base.empty:
        raise ValueError("Nada para gravar: todas as linhas estão vazias.")

    if fonte == "CSV local (exemplos)":
        caminho = PASTA_EXEMPLOS / f"{nome_aba}.csv"
        base.to_csv(caminho, sep=";", index=False, encoding="utf-8-sig")
        carregar_csv_local.clear()
        return f"Salvo em {caminho.name} ({len(base)} linha(s))."
    if fonte == "Google Sheets (API)":
        if not api_url or not api_key:
            raise ValueError("URL da API e API Key são obrigatórias para gravar no Sheets.")
        return _salvar_api(api_url, api_key, chave_aba, base, original)
    if fonte == "Google Sheets":
        if not id_planilha:
            raise ValueError("ID ou URL da planilha é obrigatório para gravar no Sheets.")
        return _salvar_gspread(id_planilha, chave_aba, base)
    raise ValueError(f"Fonte desconhecida para gravação: {fonte}.")


def salvar_tudo_na_fonte(
    dados: dict,
    fonte: str,
    id_planilha: str = "",
    api_url: str = "",
    api_key: str = "",
) -> tuple[list[str], list[str]]:
    """Grava TODAS as abas efetivas na fonte ativa (Sheets/CSV).

    Itera as abas que têm alteração em `session_state` (imports, avanços,
    exclusões, cadastros de contrato) e grava cada uma via `salvar_aba`.
    Depois de gravar com sucesso, limpa `{chave}_importado` para a fonte
    passar a ser a origem de novo (propaga em todas as páginas).
    Devolve (sucessos, erros).
    """
    sucessos: list[str] = []
    erros: list[str] = []
    originais = {
        "cronograma": dados.get("cronograma", pd.DataFrame()),
        "hist_mo": dados.get("hist_mo", pd.DataFrame()),
        "hist_equip": dados.get("hist_equip", pd.DataFrame()),
        "rh": dados.get("rh", pd.DataFrame()),
        "contratos": dados.get("contratos", pd.DataFrame()),
        "timeline": dados.get("timeline", pd.DataFrame()),
    }
    for chave, original in originais.items():
        if f"{chave}_importado" not in st.session_state:
            continue  # aba sem alteração — não precisa gravar
        efetiva = aba_efetiva(dados, chave)
        if efetiva.empty:
            erros.append(f"{ABAS.get(chave, chave)}: base efetiva vazia — nada para gravar.")
            continue
        try:
            msg = salvar_aba(
                chave, efetiva, fonte, id_planilha, api_url, api_key, original=original
            )
            sucessos.append(msg)
            # gravação ok → a fonte volta a ser a origem desta aba
            st.session_state.pop(f"{chave}_importado", None)
            st.session_state.pop(f"{chave}_importado_nome", None)
        except Exception as exc:
            erros.append(f"{ABAS.get(chave, chave)}: {exc}")
    if sucessos:
        # força o data_editor a re-renderizar com a base já gravada
        st.session_state["versao_avancos"] = int(st.session_state.get("versao_avancos", 0)) + 1
    return sucessos, erros


# ---------------------------------------------------------- cronograma (base)
def aba_efetiva(dados: dict, chave: str) -> pd.DataFrame:
    """Importação aplicada (session_state) substitui a fonte da aba inteira."""
    importado = st.session_state.get(f"{chave}_importado")
    return importado if importado is not None else dados.get(chave, pd.DataFrame())


def cronograma_efetiva(dados: dict) -> pd.DataFrame:
    return aba_efetiva(dados, "cronograma")


def ids_contratos(dados: dict) -> list[str]:
    """IDs de contrato da base efetiva (DB_Contratos) para os seletores."""
    ct = aba_efetiva(dados, "contratos")
    if ct.empty or "ID_Contrato" not in ct.columns:
        return []
    ids = [str(x).strip() for x in ct["ID_Contrato"].dropna().unique()]
    return sorted({x for x in ids if x})


def acrescentar_cronograma(
    base: pd.DataFrame, novo: pd.DataFrame, projeto: str, contrato: str
) -> tuple[pd.DataFrame, int]:
    """Inclui o cronograma importado no grupo atual (sem substituir).

    Etiqueta as tarefas novas com Projeto_Nome/Contrato e renumera ID_Tarefa
    quando colide com a base (validar_cronograma reprova ID duplicado).
    """
    n = novo.copy()
    if projeto:
        n["Projeto_Nome"] = projeto
    if contrato:
        n["Contrato"] = contrato
    renumeradas = 0
    if "ID_Tarefa" in n.columns and "ID_Tarefa" in base.columns:
        ids_base = set(base["ID_Tarefa"].astype(str).str.strip()) - {""}
        atuais = set(n["ID_Tarefa"].astype(str).str.strip())
        if ids_base & atuais:
            sufixos = [
                int(m.group(1))
                for x in ids_base
                for m in [re.search(r"(\d+)$", x)]
                if m
            ]
            proximo = (max(sufixos) + 1) if sufixos else 1
            comuns = {re.match(r"^\D*", x).group(0) for x in ids_base}
            prefixo = max(comuns, key=len) if len(comuns) == 1 else ""
            n["ID_Tarefa"] = [
                f"{prefixo}{i:03d}" for i in range(proximo, proximo + len(n))
            ]
            renumeradas = len(n)
    mescladas = pd.concat([base, n], ignore_index=True, sort=False)
    return mescladas, renumeradas


def incluir_contrato(dados: dict, id_contrato: str, extras: dict | None = None) -> bool:
    """Cadastra um novo contrato na base efetiva (session_state). False se já existe."""
    base = aba_efetiva(dados, "contratos")
    ids = set(base.get("ID_Contrato", pd.Series(dtype=str)).astype(str).str.strip())
    if not id_contrato or id_contrato in ids:
        return False
    colunas = list(
        dict.fromkeys(SCHEMAS.get("contratos", ["ID_Contrato"]) + list(base.columns))
    )
    linha = {c: "" for c in colunas}
    linha["ID_Contrato"] = id_contrato
    linha["Status"] = "Em execução"
    for k, v in (extras or {}).items():
        if k in linha and v not in ("", None):
            linha[k] = v
    st.session_state["contratos_importado"] = pd.concat(
        [base, pd.DataFrame([linha])], ignore_index=True, sort=False
    )
    st.session_state["contratos_importado_nome"] = "cadastro no app"
    return True


def acrescentar_aba(
    dados: dict, chave: str, novo: pd.DataFrame, contrato: str = ""
) -> tuple[pd.DataFrame, int]:
    """Inclui linhas de uma aba (session) sem substituir.

    Etiqueta as linhas novas com o contrato e renumera a coluna-chave quando
    colide com a base (validar_aba reprova chave duplicada).
    """
    base = aba_efetiva(dados, chave)
    n = novo.copy()
    if contrato:
        n[COLUNA_CONTRATO.get(chave, "Contrato")] = contrato
    renumeradas = 0
    chave_col = CHAVES_ABA.get(chave, "")
    if (
        chave_col
        and chave_col in n.columns
        and not base.empty
        and chave_col in base.columns
    ):
        ids_base = set(base[chave_col].astype(str).str.strip()) - {""}
        if ids_base & set(n[chave_col].astype(str).str.strip()):
            sufixos = [
                int(m.group(1))
                for x in ids_base
                for m in [re.search(r"(\d+)$", x)]
                if m
            ]
            proximo = (max(sufixos) + 1) if sufixos else 1
            comuns = {re.match(r"^\D*", x).group(0) for x in ids_base}
            prefixo = max(comuns, key=len) if len(comuns) == 1 else ""
            n[chave_col] = [
                f"{prefixo}{i:03d}" if prefixo else str(i)
                for i in range(proximo, proximo + len(n))
            ]
            renumeradas = len(n)
    mescladas = pd.concat([base, n], ignore_index=True, sort=False)
    st.session_state[f"{chave}_importado"] = mescladas
    st.session_state[f"{chave}_importado_nome"] = (
        f"{st.session_state.get(f'{chave}_importado_nome') or ABAS[chave]} (+{len(n)})"
    )
    return mescladas, renumeradas


def aplicar_avancos(dados: dict, avancos: dict[str, float], excluir: list[str]) -> int:
    """Grava avanço rápido (% por ID_Tarefa) e exclusões no cronograma efetivo.

    Devolve quantas tarefas tiveram o % alterado. Exclusões são removidas da
    base efetiva (session_state) — a fonte original não é tocada até salvar.
    """
    base = cronograma_efetiva(dados)
    if base.empty:
        return 0
    df = base.copy()
    ids = df["ID_Tarefa"].astype(str).str.strip()
    if excluir:
        df = df[~ids.isin(set(excluir))].reset_index(drop=True)
        ids = df["ID_Tarefa"].astype(str).str.strip()
    alteradas = 0
    if avancos and "%_Concluido" in df.columns:
        mascara = ids.isin(set(avancos))
        # a coluna vem como texto do CSV (dtype str/Arrow): converte antes de gravar número
        df["%_Concluido"] = pd.to_numeric(df["%_Concluido"], errors="coerce").fillna(0.0)
        valores = df.loc[mascara, "ID_Tarefa"].map(
            lambda i: min(100.0, max(0.0, avancos.get(str(i).strip(), 0.0)))
        )
        alteradas = int((valores != df.loc[mascara, "%_Concluido"]).sum())
        df.loc[mascara, "%_Concluido"] = valores
        if "Status" in df.columns:
            # Status deriva do % no round-trip MS Project — mantemos coerência
            # sem destruir rótulos manuais como "Atrasado".
            concl = mascara & (pd.to_numeric(df["%_Concluido"], errors="coerce").fillna(0) >= 100)
            andando = mascara & (pd.to_numeric(df["%_Concluido"], errors="coerce").fillna(0) > 0) & ~concl
            df.loc[concl, "Status"] = "Concluído"
            em_branco = df["Status"].astype(str).str.strip().isin(["", "nan", "None"])
            df.loc[andando & em_branco, "Status"] = "Em andamento"
    st.session_state["cronograma_importado"] = df
    nome = st.session_state.get("cronograma_importado_nome") or ABAS["cronograma"]
    st.session_state["cronograma_importado_nome"] = nome
    return alteradas


def excluir_contrato(dados: dict, id_contrato: str) -> dict[str, int]:
    """Remove o contrato e limpa os vínculos nas abas efetivas (session).

    Cronograma/histogramas/RH: só a coluna de vínculo fica vazia (as linhas
    permanecem). Timeline: remove as linhas do projeto (cada marco pertence a
    um contrato). Devolve {aba: linhas_afetadas} para o resumo na tela.
    """
    impacto: dict[str, int] = {}
    ct = aba_efetiva(dados, "contratos")
    if ct.empty or id_contrato not in set(ct.get("ID_Contrato", pd.Series(dtype=str)).astype(str)):
        return impacto
    st.session_state["contratos_importado"] = ct[
        ct["ID_Contrato"].astype(str).str.strip() != id_contrato
    ].reset_index(drop=True)
    st.session_state["contratos_importado_nome"] = (
        st.session_state.get("contratos_importado_nome") or ABAS["contratos"]
    ) + " (exclusão)"
    impacto["contratos"] = 1
    for chave in ("cronograma", "hist_mo", "hist_equip", "rh"):
        base = aba_efetiva(dados, chave)
        col = COLUNA_CONTRATO.get(chave, "Contrato")
        if base.empty or col not in base.columns:
            continue
        mascara = base[col].astype(str).str.strip() == id_contrato
        n = int(mascara.sum())
        if n:
            df = base.copy()
            df.loc[mascara, col] = ""
            st.session_state[f"{chave}_importado"] = df
            st.session_state[f"{chave}_importado_nome"] = (
                st.session_state.get(f"{chave}_importado_nome") or ABAS[chave]
            )
            impacto[chave] = n
    tl = aba_efetiva(dados, "timeline")
    if not tl.empty and "Projeto" in tl.columns:
        mascara = tl["Projeto"].astype(str).str.strip() == id_contrato
        n = int(mascara.sum())
        if n:
            st.session_state["timeline_importado"] = tl[~mascara].reset_index(drop=True)
            st.session_state["timeline_importado_nome"] = (
                st.session_state.get("timeline_importado_nome") or ABAS["timeline"]
            )
            impacto["timeline"] = n
    return impacto


def fig_gantt(base: pd.DataFrame, origem_contrato: str = "") -> go.Figure | None:
    """Gantt (px.timeline) das tarefas com datas completas — usado na Timeline
    e na aba "Por Cronograma" do Dashboard. `base` já vem de preparar_cronograma."""
    com_datas = base.dropna(subset=["Inicio_Planejado_dt", "Fim_Planejado_dt"])
    invertidas = com_datas[com_datas["Fim_Planejado_dt"] < com_datas["Inicio_Planejado_dt"]]
    com_datas = com_datas[~com_datas.index.isin(invertidas.index)]
    if com_datas.empty:
        return None
    g = com_datas.copy()
    g["Tarefa"] = g["ID_Tarefa"].astype(str) + " — " + g["Nome_Tarefa"].astype(str)
    g["Caminho crítico"] = g["_critico"].map({True: "Sim", False: "Não"})
    g["Início"] = g["Inicio_Planejado_dt"].dt.strftime("%d/%m/%Y")
    g["Fim"] = g["Fim_Planejado_dt"].dt.strftime("%d/%m/%Y")
    g["% Concluído"] = g["_pct"].map(lambda v: f"{v:.0f}%")
    g["Frente/Contrato"] = (
        g["Frente"].fillna("—").astype(str) + " / " + g["_contrato"].fillna("—").astype(str)
    )
    ordem = g.sort_values(["Inicio_Planejado_dt", "ID_Tarefa"])["Tarefa"].tolist()
    fig = px.timeline(
        g,
        x_start="Inicio_Planejado_dt",
        x_end="Fim_Planejado_dt",
        y="Tarefa",
        color="Caminho crítico",
        color_discrete_map={"Sim": "#EF4444", "Não": "#3B82F6"},
        category_orders={"Tarefa": ordem},
        hover_data=["Início", "Fim", "% Concluído", "Status", "Responsavel", "Frente/Contrato"],
    )
    # px.timeline já inverte o eixo Y (1ª tarefa no topo) — não passar autorange aqui.
    fig.update_yaxes(title="")
    fig.update_xaxes(title="", tickformat="%d/%m/%Y")
    fig.add_vline(x=hoje(), line_color="#FB923C", line_dash="dot", annotation_text="hoje")
    # margem esquerda calculada: rótulos do eixo Y são nomes longos de tarefa
    # (com l=10 fixo o Plotly corta o texto na borda do SVG).
    maior_rot = max((len(s) for s in g["Tarefa"]), default=10)
    margem_e = min(500, max(90, int(maior_rot * 7.2) + 16))
    fig.update_layout(
        height=max(380, 70 + 46 * len(g)),
        legend_title_text="Caminho crítico",
        legend=dict(x=0.99, y=0.99, xanchor="right", yanchor="top"),
        margin=dict(l=margem_e, r=14, t=30, b=34),
    )
    return estilizar_fig(fig)


def preparar_cronograma(base: pd.DataFrame) -> pd.DataFrame:
    """Colunas derivadas: datas datetime, números e Frente (1º nível do WBS)."""
    t = base.copy()
    for col in ("Inicio_Planejado", "Fim_Planejado", "Inicio_Real", "Fim_Real"):
        if col not in t.columns:
            t[col] = ""
        t[col + "_dt"] = parse_datas(t[col])
    for col, destino in (("Duracao_dias", "_dur"), ("%_Concluido", "_pct")):
        if col not in t.columns:
            t[col] = 0
        t[destino] = parse_num(t[col]).fillna(0.0)
    t["_pct"] = t["_pct"].clip(0, 100)
    if "WBS" not in t.columns:
        t["WBS"] = ""
    t["Frente"] = (
        t["WBS"].astype(str).str.split(".").str[0].str.strip().replace("", pd.NA)
    )
    if "Caminho_Critico_SIM_NAO" not in t.columns:
        t["Caminho_Critico_SIM_NAO"] = "NAO"
    t["_critico"] = (
        t["Caminho_Critico_SIM_NAO"]
        .astype(str)
        .str.strip()
        .str.upper()
        .isin({"SIM", "S", "YES", "Y", "TRUE", "1"})
    )
    return t


def garantir_colunas(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    ausentes = [c for c in COLUNAS_CRONOGRAMA if c not in df.columns]
    out = df.copy()
    for c in ausentes:
        out[c] = ""
    return out, ausentes


def tabela_cronograma(t: pd.DataFrame) -> pd.DataFrame:
    """Cópia só com as colunas padrão e datas em DD/MM/AAAA, para exibição."""
    df = t.reindex(columns=COLUNAS_CRONOGRAMA).copy()
    for c in df.columns:
        df[c] = df[c].fillna("").astype(str).str.strip()
    for c in ("Inicio_Planejado", "Fim_Planejado", "Inicio_Real", "Fim_Real"):
        df[c] = df[c].map(fmt_data_opcional)
    return df


def validar_cronograma(df: pd.DataFrame) -> tuple[list[str], list[str]]:
    erros: list[str] = []
    avisos: list[str] = []
    colunas = set(df.columns)
    ausentes = [c for c in COLUNAS_CRONOGRAMA if c not in colunas]
    if ausentes:
        avisos.append(f"[FALTANTE] coluna(s) ausente(s), criada(s) vazia(s): {', '.join(ausentes)}.")
    vazio_mask = df.apply(
        lambda s: s.astype(str).str.strip().str.lower().isin(("", "nan", "none", "nan%"))
    )
    for col in ("ID_Tarefa", "Nome_Tarefa", "Inicio_Planejado", "Fim_Planejado"):
        if col in colunas and bool(vazio_mask[col].any()):
            avisos.append(f"[FALTANTE] {int(vazio_mask[col].sum())} linha(s) sem valor em {col}.")
    for col in ("Inicio_Planejado", "Fim_Planejado", "Inicio_Real", "Fim_Real"):
        if col not in colunas:
            continue
        bruto = df[col].astype(str).str.strip()
        preenchido = ~bruto.str.lower().isin(("", "nan", "none"))
        d = parse_datas(df[col])
        invalidas = sorted(set(bruto[preenchido & d.isna()]))
        if invalidas:
            erros.append(
                f"Data inválida em {col}: {len(invalidas)} valor(es) — ex.: "
                f"{', '.join(invalidas[:5])}. Use DD/MM/AAAA."
            )
    if "%_Concluido" in colunas:
        bruto = df["%_Concluido"].astype(str).str.strip()
        preenchido = ~bruto.str.lower().isin(("", "nan", "none"))
        v = parse_num(df["%_Concluido"])
        fora = sorted(set(bruto[preenchido & (v.isna() | (v < 0) | (v > 100))]))
        if fora:
            erros.append(f"%_Concluido fora de 0–100 ou não numérico: {', '.join(fora[:5])}.")
    if "Duracao_dias" in colunas:
        bruto = df["Duracao_dias"].astype(str).str.strip()
        preenchido = ~bruto.str.lower().isin(("", "nan", "none"))
        v = parse_num(df["Duracao_dias"])
        ruins = sorted(set(bruto[preenchido & v.isna()]))
        if ruins:
            avisos.append(f"Duração não numérica em {len(ruins)} linha(s): {', '.join(ruins[:5])}.")
    if "ID_Tarefa" in colunas:
        ids = df["ID_Tarefa"].astype(str).str.strip()
        dup = sorted(set(ids[ids.ne("") & ids.duplicated(keep=False)]))
        if dup:
            erros.append(f"ID_Tarefa duplicado: {', '.join(dup[:10])}.")
    if "Status" in colunas:
        stts = df["Status"].astype(str).str.strip()
        fora = sorted(
            {
                s
                for s in stts
                if s and s.lower() not in ("nan", "none") and s.lower() not in STATUS_VALIDOS
            }
        )
        if fora:
            avisos.append(
                "Status fora do padrão (Não Iniciado/Em Andamento/Concluído/Atrasado): "
                + ", ".join(fora[:5])
                + "."
            )
    if {"Inicio_Planejado", "Fim_Planejado"} <= colunas:
        ini = parse_datas(df["Inicio_Planejado"])
        fim = parse_datas(df["Fim_Planejado"])
        invertidas = int((fim < ini).sum())
        if invertidas:
            erros.append(f"{invertidas} tarefa(s) com Fim_Planejado anterior ao Inicio_Planejado.")
    return erros, avisos


# -------------------------------------------------------------------- curva S
def calcular_curva_s(t: pd.DataFrame) -> pd.DataFrame:
    total = float(t["_dur"].sum())
    if total <= 0:
        return pd.DataFrame(columns=["Data", "Planejado", "Real"])
    ini = t["Inicio_Planejado_dt"].min()
    if pd.isna(ini):
        return pd.DataFrame(columns=["Data", "Planejado", "Real"])
    fim = t["Fim_Planejado_dt"].max()
    fim = hoje() if pd.isna(fim) else max(fim, hoje())
    ultimo_real = t["Fim_Real_dt"].max()
    if pd.notna(ultimo_real) and ultimo_real > fim:
        fim = ultimo_real
    registros = t[
        ["_dur", "_pct", "Inicio_Planejado_dt", "Fim_Planejado_dt", "Inicio_Real_dt", "Fim_Real_dt"]
    ].to_dict("records")

    def fracao(dia, ini_t, fim_t) -> float:
        if pd.isna(ini_t):
            return 0.0
        if pd.isna(fim_t) or fim_t <= ini_t:
            return 1.0 if dia >= ini_t else 0.0
        return min(max((dia - ini_t) / (fim_t - ini_t), 0.0), 1.0)

    linhas = []
    for dia in pd.date_range(pd.Timestamp(ini), pd.Timestamp(fim), freq="D"):
        p = 0.0
        r = 0.0
        for x in registros:
            dur = float(x["_dur"])
            if dur <= 0:
                continue
            p += dur * fracao(dia, x["Inicio_Planejado_dt"], x["Fim_Planejado_dt"])
            pct = float(x["_pct"])
            if pct <= 0:
                continue
            ir = x["Inicio_Real_dt"]
            if pd.isna(ir):
                ir = x["Inicio_Planejado_dt"]
            if pd.isna(ir):
                continue
            fr = x["Fim_Real_dt"]
            if pct >= 100:
                fim_r = fr if pd.notna(fr) else ir
            else:
                fim_r = fr if pd.notna(fr) else hoje()
            if pd.isna(fim_r) or fim_r < ir:
                fim_r = max(ir, hoje())
            r += dur * (pct / 100.0) * fracao(dia, ir, fim_r)
        linhas.append((dia, round(100.0 * p / total, 2), round(100.0 * r / total, 2)))
    return pd.DataFrame(linhas, columns=["Data", "Planejado", "Real"])


# -------------------------------------------------------------------- alertas
def alertas_alocacao(rh: pd.DataFrame) -> list[str]:
    if rh.empty or "Alocacao_%" not in rh.columns:
        return []
    a = rh.copy()
    a["_aloc"] = parse_num(a["Alocacao_%"]).fillna(0)
    msgs = []
    for x in a[a["_aloc"] > 100].to_dict("records"):
        msgs.append(
            f"🔔 **{x.get('Matricula_Nome', '')}** com alocação {x['_aloc']:.0f}% (acima de 100%) — "
            f"{x.get('Funcao', '')} / {x.get('Equipe_Frente', '')}."
        )
    return msgs


def coletar_alertas(t: pd.DataFrame, dados: dict) -> tuple[list[str], list[str], list[str]]:
    erros: list[str] = []
    avisos: list[str] = []
    infos: list[str] = []

    atrasadas = t[
        t["Fim_Planejado_dt"].notna()
        & (t["Fim_Planejado_dt"] < hoje())
        & (t["_pct"] < 100)
    ].sort_values(["_critico", "Fim_Planejado_dt"], ascending=[False, True])
    for x in atrasadas.to_dict("records"):
        rotulo = f"{x.get('ID_Tarefa', '')} {x.get('Nome_Tarefa', '')}".strip()
        base = f"fim planejado em {fmt_data(x['Fim_Planejado_dt'])}, {x['_pct']:.0f}% concluído"
        if x["_critico"]:
            erros.append(f"🔴 **CRÍTICA** — tarefa **{rotulo}**: {base}.")
        else:
            avisos.append(f"⚠️ Tarefa **{rotulo}**: {base}.")

    tarde = t[
        t["Fim_Real_dt"].notna()
        & t["Fim_Planejado_dt"].notna()
        & (t["Fim_Real_dt"] > t["Fim_Planejado_dt"])
    ]
    for x in tarde.to_dict("records"):
        dias = int((x["Fim_Real_dt"] - x["Fim_Planejado_dt"]).days)
        rotulo = f"{x.get('ID_Tarefa', '')} {x.get('Nome_Tarefa', '')}".strip()
        infos.append(
            f"⏱ **{rotulo}** concluída {dias} dia(s) após o fim planejado "
            f"({fmt_data(x['Fim_Planejado_dt'])} → {fmt_data(x['Fim_Real_dt'])})."
        )

    erros.extend(alertas_alocacao(aba_efetiva(dados, "rh")))

    ct = aba_efetiva(dados, "contratos")
    if not ct.empty and "Data_Fim" in ct.columns:
        ct2 = ct.copy()
        ct2["_fim"] = parse_datas(ct2["Data_Fim"])
        for x in ct2.to_dict("records"):
            if pd.isna(x["_fim"]):
                continue
            stt = str(x.get("Status", "")).lower()
            encerrado = any(p in stt for p in ("conclu", "encerrad", "cancelad"))
            if encerrado:
                continue
            dias = int((x["_fim"] - hoje()).days)
            if dias < 0:
                erros.append(
                    f"📕 Contrato **{x.get('ID_Contrato', '')}** vencido há {abs(dias)} dia(s) "
                    f"(fim {fmt_data(x['_fim'])}) e status {x.get('Status', '')}."
                )
            elif dias <= 15:
                avisos.append(
                    f"📆 Contrato **{x.get('ID_Contrato', '')}** vence em {dias} dia(s) "
                    f"({fmt_data(x['_fim'])})."
                )

    stat = (
        t["Status"].astype(str).str.strip()
        if "Status" in t.columns
        else pd.Series(dtype=str)
    )
    faltantes = (
        int(t["Inicio_Planejado_dt"].isna().sum())
        + int(t["Fim_Planejado_dt"].isna().sum())
        + int(stat.isin(("", "nan", "None")).sum())
    )
    if faltantes:
        infos.append(
            f"[FALTANTE] {faltantes} campo(s) obrigatório(s) vazio(s) no cronograma "
            "(datas planejadas ou status). Preencha na fonte de dados."
        )
    return erros, avisos, infos


def renderizar_alertas(erros: list[str], avisos: list[str], infos: list[str]) -> None:
    if not (erros or avisos or infos):
        st.success("✅ Nenhum alerta no momento.")
        return
    for msg in erros:
        st.error(msg)
    for msg in avisos:
        st.warning(msg)
    for msg in infos:
        st.info(msg)


# ------------------------------------------------------- importação/exportação
def duracao_para_dias(valor) -> str:
    """'PT72H' / '5 dias' / '8h' → número de dias (8h = 1 dia)."""
    texto = str(valor or "").strip().lower()
    if not texto:
        return ""
    m = re.search(r"pt(\d+(?:[.,]\d+)?)\s*h", texto)
    if m:
        return f"{float(m.group(1).replace(',', '.')) / 8:g}"
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:dias?|days?|d)\b", texto)
    if m:
        return f"{float(m.group(1).replace(',', '.')):g}"
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*h", texto)
    if m:
        return f"{float(m.group(1).replace(',', '.')) / 8:g}"
    m = re.search(r"\d+", texto)
    return m.group(0) if m else ""


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def interpretar_msproject_xml(conteudo: bytes) -> pd.DataFrame:
    """Lê XML do MS Project (Arquivo > Salvar Como > XML) no padrão Sheets."""
    try:
        raiz = ET.fromstring(conteudo)
    except ET.ParseError as exc:
        raise ValueError(f"XML inválido ({exc}).") from exc
    tarefas = [el for el in raiz.iter() if _local(el.tag) == "Task"]
    if not tarefas:
        raise ValueError("XML sem elementos <Task>. No MS Project use Arquivo > Salvar Como > XML.")

    preparadas = []
    for seq, el in enumerate(tarefas, start=1):
        campos = {}
        preds_uid: list[str] = []
        for filho in el:
            nome = _local(filho.tag)
            if nome == "PredecessorLinks":
                preds_uid = [
                    (n.text or "").strip()
                    for n in filho.iter()
                    if _local(n.tag) == "TaskUID"
                ]
                continue
            if len(filho):
                continue
            campos[nome] = (filho.text or "").strip()
        if not campos.get("Name"):
            continue
        uid = campos.get("UID") or campos.get("ID") or str(seq)
        preparadas.append((seq, uid, campos, preds_uid))

    uid_para_id = {uid: (campos.get("ID") or uid) for _, uid, campos, _ in preparadas}
    linhas = []
    for seq, uid, campos, preds_uid in preparadas:
        pct = campos.get("PercentComplete")
        if pct in (None, ""):
            pct = campos.get("PhysicalPercentComplete") or campos.get("PercentWorkComplete") or ""
        try:
            pv = float(pct)
            pct_txt_valor = str(int(pv)) if pv == int(pv) else f"{pv:g}"
        except (TypeError, ValueError):
            pct_txt_valor = ""
        preds = ", ".join(
            uid_para_id.get(u, u) for u in preds_uid if u
        )
        critico = str(campos.get("Critical", "")).lower() in ("1", "true", "yes", "y")
        pct_num = float(pct_txt_valor) if pct_txt_valor else 0.0
        linhas.append(
            {
                "ID_Tarefa": campos.get("ID") or uid,
                "Nome_Tarefa": campos.get("Name", ""),
                "WBS": campos.get("WBS") or campos.get("OutlineNumber", ""),
                "Inicio_Planejado": _data_xml(campos.get("Start")),
                "Fim_Planejado": _data_xml(campos.get("Finish")),
                "Inicio_Real": _data_xml(campos.get("ActualStart")),
                "Fim_Real": _data_xml(campos.get("ActualFinish")),
                "Duracao_dias": duracao_para_dias(campos.get("Duration")),
                "%_Concluido": pct_txt_valor,
                "Predecessora": preds,
                "Sucessora": "",
                "Caminho_Critico_SIM_NAO": "SIM" if critico else "NAO",
                "Responsavel": campos.get("ResourceNames", ""),
                "Status": (
                    "Concluído" if pct_num >= 100 else "Em Andamento" if pct_num > 0 else "Não Iniciado"
                ),
                "Observacao": campos.get("Notes", ""),
            }
        )
    if not linhas:
        raise ValueError("XML não contém tarefas com nome.")
    return pd.DataFrame(linhas, columns=COLUNAS_CRONOGRAMA)


def _data_xml(valor) -> str:
    """'2026-09-01T00:00:00' → '01/09/2026' (usa o conversor para converter)."""
    texto = str(valor or "").strip()
    if not texto:
        return ""
    return conversor.parse_date_br(texto.split("T")[0])


def _eh_padrao(linha: str) -> bool:
    return "ID_Tarefa" in linha and "Nome_Tarefa" in linha


def _importar_via_conversor(texto: str) -> pd.DataFrame:
    with tempfile.TemporaryDirectory() as td:
        entrada = Path(td) / "entrada.csv"
        saida = Path(td) / "saida.csv"
        entrada.write_text(texto, encoding="utf-8")
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            conversor.msproject_para_sheets(str(entrada), str(saida))
        df = pd.read_csv(saida, sep=";", dtype=str, encoding="utf-8-sig").fillna("")
    # O conversor não tem Status (a origem MS Project/P6 não tem) → deriva do %.
    if "%_Concluido" in df.columns:
        v = parse_num(df["%_Concluido"]).fillna(0)
        df["Status"] = [
            "Concluído" if x >= 100 else "Em Andamento" if x > 0 else "Não Iniciado" for x in v
        ]
    return df


def interpretar_upload(nome: str, conteudo: bytes) -> tuple[pd.DataFrame | None, list[str]]:
    ext = nome.lower().rsplit(".", 1)[-1]
    avisos: list[str] = []
    if ext == "xml":
        return interpretar_msproject_xml(conteudo), avisos
    if ext == "mpp":
        try:
            from leitor_mpp import interpretar_mpp
        except ImportError as exc:  # olefile ausente
            avisos.append(
                f"[FALTANTE] leitura de .mpp indisponível ({exc}). "
                "Instale a dependência: pip install olefile."
            )
            return None, avisos
        df_mpp, avisos_mpp = interpretar_mpp(conteudo)
        avisos.extend(avisos_mpp)
        return df_mpp, avisos
    if ext == "xlsx":
        bruto = pd.read_excel(io.BytesIO(conteudo))
        bruto.columns = [str(c).strip() for c in bruto.columns]
        texto = bruto.to_csv(sep=",", index=False)
    else:
        texto = conteudo.decode("utf-8-sig", errors="replace")
    primeira = texto.splitlines()[0] if texto else ""
    if _eh_padrao(primeira):
        sep = ";" if primeira.count(";") >= primeira.count(",") else ","
        df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str).fillna("")
        return df, avisos
    df = _importar_via_conversor(texto)
    avisos.append(
        "Origem MS Project/P6 convertida para o padrão DB_Cronograma_MSProject pelo conversor; "
        "Status derivado de %_Concluido."
    )
    return df, avisos


def interpretar_histograma(
    nome: str, conteudo: bytes
) -> tuple[pd.DataFrame | None, str, list[str]]:
    """Lê CSV/XLSX de histograma e devolve (df, chave da aba, avisos).

    A aba é detectada pelo cabeçalho: Funcao_Cargo → DB_Histograma_MO;
    Equipamento → DB_Histograma_Equip.
    """
    avisos: list[str] = []
    ext = nome.lower().rsplit(".", 1)[-1]
    if ext == "xlsx":
        bruto = pd.read_excel(io.BytesIO(conteudo), dtype=str)
        bruto.columns = [str(c).strip() for c in bruto.columns]
        df = bruto.fillna("")
    else:
        texto = conteudo.decode("utf-8-sig", errors="replace")
        primeira = texto.splitlines()[0] if texto else ""
        sep = ";" if primeira.count(";") >= primeira.count(",") else ","
        df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
    if "Funcao_Cargo" in df.columns:
        return df, "hist_mo", avisos
    if "Equipamento" in df.columns:
        return df, "hist_equip", avisos
    avisos.append(
        "[FALTANTE] não reconheci o arquivo: esperado um histograma com a coluna "
        "Funcao_Cargo (DB_Histograma_MO) ou Equipamento (DB_Histograma_Equip)."
    )
    return None, "", avisos


def normalizar_para_export(df: pd.DataFrame) -> pd.DataFrame:
    """Garante as 15 colunas padrão, datas DD/MM/AAAA e números sem casas decimais."""
    p = df.reindex(columns=COLUNAS_CRONOGRAMA).copy()
    for col in ("Inicio_Planejado", "Fim_Planejado", "Inicio_Real", "Fim_Real"):
        p[col] = parse_datas(p[col]).map(
            lambda v: "" if pd.isna(v) else pd.Timestamp(v).strftime("%d/%m/%Y")
        )
    for col in ("Duracao_dias", "%_Concluido"):
        num = parse_num(p[col])
        p[col] = num.map(
            lambda v: "" if pd.isna(v) else (f"{int(round(v))}" if float(v) == int(v) else f"{v:g}")
        )
    return p.fillna("")


def exportar_via_conversor(df: pd.DataFrame, modo: str) -> tuple[bytes, str]:
    padrao = normalizar_para_export(df)
    with tempfile.TemporaryDirectory() as td:
        entrada = Path(td) / "entrada.csv"
        saida = Path(td) / "saida.csv"
        padrao.to_csv(entrada, sep=";", index=False, encoding="utf-8-sig")
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            getattr(conversor, modo)(str(entrada), str(saida))
        texto = buf.getvalue().strip()
        msg = texto.splitlines()[0].split(" ->")[0] if texto else ""
        return saida.read_bytes(), msg


def exportar_xlsx(df: pd.DataFrame) -> bytes:
    padrao = normalizar_para_export(df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        padrao.to_excel(writer, index=False, sheet_name="DB_Cronograma_MSProject")
    return buf.getvalue()


# ----------------------------------------------------------------- 1 Dashboard
def vincular_contrato(t: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    """Coluna _contrato: a do próprio cronograma, senão a Frente → Projeto (Timeline)."""
    if "Contrato" in t.columns and bool(t["Contrato"].astype(str).str.strip().ne("").any()):
        t["_contrato"] = t["Contrato"].astype(str).str.strip()
    else:
        t["_contrato"] = t["Frente"].map(mapa)
    return t


def _tabela_tarefas(t: pd.DataFrame, somente_atrasadas: bool = False,
                    somente_criticas: bool = False) -> pd.DataFrame:
    """Tabela densa das tarefas (drill-down dos KPIs e do template Analítico)."""
    ver = t.sort_values("Fim_Planejado_dt").copy()
    if somente_atrasadas:
        ver = ver[
            ver["Fim_Planejado_dt"].notna()
            & (ver["Fim_Planejado_dt"] < hoje())
            & (ver["_pct"] < 100)
        ].copy()
        ver["Dias_Atraso"] = (hoje() - ver["Fim_Planejado_dt"]).dt.days
    if somente_criticas:
        ver = ver[ver["_critico"]].copy()
    id_col = (
        ver["ID_Tarefa"].astype(str)
        if "ID_Tarefa" in ver.columns
        else pd.Series([""] * len(ver), index=ver.index)
    )
    nome_col = (
        ver["Nome_Tarefa"].astype(str)
        if "Nome_Tarefa" in ver.columns
        else pd.Series([""] * len(ver), index=ver.index)
    )
    ver["Tarefa"] = id_col + " — " + nome_col
    cols = [
        "Tarefa", "WBS", "Frente", "_contrato", "Inicio_Planejado", "Fim_Planejado",
        "Duracao_dias", "%_Concluido", "Dias_Atraso", "Caminho_Critico_SIM_NAO",
        "Responsavel", "Status",
    ]
    return ver[[c for c in cols if c in ver.columns]]


def _drill_tarefas(rotulo: str, tabela: pd.DataFrame, vazio: str) -> None:
    """Popover de drill-down: abre a lista exata por trás de um KPI."""
    with st.popover(rotulo):
        if tabela.empty:
            st.info(vazio)
        else:
            st.dataframe(tabela, hide_index=True, height=320, width="stretch")


def _curva_s_fig(cv: pd.DataFrame, altura: int) -> go.Figure:
    """Curva S (planejado × real) — mesma figura nas duas templates."""
    fig = go.Figure()
    fig.add_scatter(
        x=cv["Data"], y=cv["Planejado"], name="Planejado", mode="lines",
        line=dict(color="#94A3B8", dash="dash", width=2),
    )
    fig.add_scatter(
        x=cv["Data"], y=cv["Real"], name="Real", mode="lines",
        line=dict(color="#6366F1", width=3),
        fill="tozeroy", fillcolor="rgba(99,102,241,0.16)",
    )
    fig.add_vline(x=hoje(), line_color="#FB923C", line_dash="dot", annotation_text="hoje")
    fig.update_layout(
        hovermode="x unified",
        height=altura,
        margin=dict(t=30, b=40),
        legend=dict(orientation="h", y=1.12),
        yaxis=dict(title="% concluído", range=[0, 105]),
        xaxis=dict(title="", tickformat="%d/%m/%Y"),
    )
    return estilizar_fig(fig)


def dashboard_executivo(t: pd.DataFrame, dados: dict, ct: pd.DataFrame,
                        rh: pd.DataFrame, mapa: dict) -> None:
    """Template Executivo: KPIs grandes, gráficos amplos e pouca tabela."""
    total_dur = float(t["_dur"].sum())
    avanco = float((t["_dur"] * t["_pct"]).sum() / total_dur) if total_dur else 0.0
    atrasadas = t[
        t["Fim_Planejado_dt"].notna() & (t["Fim_Planejado_dt"] < hoje()) & (t["_pct"] < 100)
    ]
    criticas_atrasadas = int(atrasadas["_critico"].sum())
    ativos = 0
    if not rh.empty and "Status_Ativo_Inativo" in rh.columns:
        ativos = int((rh["Status_Ativo_Inativo"].astype(str).str.strip() == "Ativo").sum())
    em_execucao = 0
    if not ct.empty and "Status" in ct.columns:
        em_execucao = int(
            (~ct["Status"].astype(str).str.lower().str.contains("conclu|encerrad|cancelad")).sum()
        )

    # avanço ponderado por frente (base do drill-down do KPI e do gráfico)
    linhas_frente = []
    for frente, grupo in t.groupby("Frente", dropna=False):
        dur = float(grupo["_dur"].sum())
        av = float((grupo["_dur"] * grupo["_pct"]).sum() / dur) if dur else 0.0
        linhas_frente.append(
            {
                "Frente": frente if pd.notna(frente) else "(sem frente)",
                "Tarefas": len(grupo),
                "Avanço_%": round(av, 1),
            }
        )
    avanco_frente = pd.DataFrame(linhas_frente)

    r1 = st.columns(3)
    with r1[0]:
        st.metric("Tarefas", len(t), border=True)
        _drill_tarefas(
            "Ver tarefas",
            _tabela_tarefas(t),
            "[FALTANTE] Sem tarefas no recorte atual.",
        )
    with r1[1]:
        st.metric("Avanço físico", pct_txt(avanco), border=True)
        _drill_tarefas("Ver por frente", avanco_frente, "[FALTANTE] Sem Frente no cronograma.")
    with r1[2]:
        st.metric(
            "Atrasadas",
            len(atrasadas),
            delta=f"{criticas_atrasadas} crítica(s)" if criticas_atrasadas else None,
            delta_color="inverse",
            border=True,
        )
        _drill_tarefas(
            "Ver atrasadas",
            _tabela_tarefas(t, somente_atrasadas=True),
            "✅ Nenhuma tarefa atrasada no recorte atual.",
        )
    r2 = st.columns(3)
    with r2[0]:
        st.metric("No caminho crítico", int(t["_critico"].sum()), border=True)
        _drill_tarefas(
            "Ver críticas",
            _tabela_tarefas(t, somente_criticas=True),
            "[FALTANTE] Sem caminho crítico identificado (coluna Caminho_Critico_SIM_NAO).",
        )
    with r2[1]:
        st.metric("Efetivo ativo", ativos, border=True)
        if not rh.empty and "Matricula_Nome" in rh.columns:
            ativos_df = rh[rh["Status_Ativo_Inativo"].astype(str).str.strip() == "Ativo"] \
                if "Status_Ativo_Inativo" in rh.columns else rh
            col_rh = [c for c in ("Matricula_Nome", "Funcao", "Equipe_Frente", "Contrato", "Alocacao_%")
                      if c in rh.columns]
        else:
            ativos_df, col_rh = pd.DataFrame(), []
        _drill_tarefas(
            "Ver equipe",
            ativos_df[col_rh].reset_index(drop=True) if col_rh else pd.DataFrame(),
            "[FALTANTE] Sem banco de RH com essas colunas.",
        )
    with r2[2]:
        st.metric("Contratos em execução", em_execucao, border=True)
        if not ct.empty and "ID_Contrato" in ct.columns:
            em_exec_df = ct[~ct["Status"].astype(str).str.lower().str.contains("conclu|encerrad|cancelad")] \
                if "Status" in ct.columns else ct
            col_ct = [c for c in ("ID_Contrato", "Cliente_Fornecedor", "Objeto_Escopo", "Valor_Total", "Data_Fim", "Status")
                      if c in ct.columns]
        else:
            em_exec_df, col_ct = pd.DataFrame(), []
        _drill_tarefas(
            "Ver contratos",
            em_exec_df[col_ct].reset_index(drop=True) if col_ct else pd.DataFrame(),
            "[FALTANTE] Sem contratos.",
        )

    st.markdown("#### Curva S")
    cv = calcular_curva_s(t)
    if cv.empty:
        st.info("[FALTANTE] Sem datas planejadas para montar a curva S.")
    else:
        st.plotly_chart(_curva_s_fig(cv, 420), theme=None, config={"displaylogo": False})

    # avanço por contrato (Frente → Projeto via Timeline_Visao)
    if mapa:
        st.markdown("#### Avanço por contrato")
        linhas = []
        for frente, grupo in t.groupby("Frente"):
            cid = mapa.get(frente)
            if not cid:
                continue
            dur = float(grupo["_dur"].sum())
            av = float((grupo["_dur"] * grupo["_pct"]).sum() / dur) if dur else 0.0
            linhas.append({"Contrato": cid, "Avanço": round(av, 1)})
        if linhas:
            av_ct = (
                pd.DataFrame(linhas)
                .groupby("Contrato", as_index=False)["Avanço"]
                .max()
                .sort_values("Contrato")
            )
            figb = px.bar(
                av_ct, x="Contrato", y="Avanço", text=av_ct["Avanço"].map(lambda v: f"{v:.1f}%"),
                color_discrete_sequence=["#6366F1"],
            )
            figb.update_traces(textposition="outside", marker_line_width=0)
            figb.update_layout(
                height=280, margin=dict(t=24, b=36), showlegend=False,
                yaxis=dict(title="% concluído", range=[0, 112]),
                xaxis=dict(title=""),
            )
            estilizar_fig(figb)
            st.plotly_chart(figb, theme=None, config={"displaylogo": False})

    st.markdown("#### Alertas")
    erros, avisos, infos = coletar_alertas(t, dados)
    renderizar_alertas(erros, avisos, infos)


def dashboard_analitico(t: pd.DataFrame, dados: dict, ct: pd.DataFrame,
                        rh: pd.DataFrame, mapa: dict) -> None:
    """Template Analítico: tabelas densas, métricas compactas e gráficos menores."""
    total_dur = float(t["_dur"].sum())
    avanco = float((t["_dur"] * t["_pct"]).sum() / total_dur) if total_dur else 0.0
    atrasadas = t[
        t["Fim_Planejado_dt"].notna() & (t["Fim_Planejado_dt"] < hoje()) & (t["_pct"] < 100)
    ]
    r1 = st.columns(4)
    with r1[0]:
        st.metric("Avanço", pct_txt(avanco), border=True)
    with r1[1]:
        st.metric("Tarefas", len(t), border=True)
    with r1[2]:
        st.metric("Atrasadas", len(atrasadas), border=True)
        _drill_tarefas(
            "Ver atrasadas",
            _tabela_tarefas(t, somente_atrasadas=True),
            "✅ Nenhuma tarefa atrasada no recorte atual.",
        )
    with r1[3]:
        st.metric("Críticas atrasadas", int(atrasadas["_critico"].sum()), border=True)
        _drill_tarefas(
            "Ver críticas atrasadas",
            _tabela_tarefas(t, somente_atrasadas=True, somente_criticas=True),
            "✅ Nenhuma crítica atrasada no recorte atual.",
        )

    # tabela densa do cronograma (só colunas existentes)
    st.markdown("#### 📋 Tarefas")
    ver = t.sort_values("Fim_Planejado_dt").copy()
    id_col = (
        ver["ID_Tarefa"].astype(str)
        if "ID_Tarefa" in ver.columns
        else pd.Series([""] * len(ver), index=ver.index)
    )
    nome_col = (
        ver["Nome_Tarefa"].astype(str)
        if "Nome_Tarefa" in ver.columns
        else pd.Series([""] * len(ver), index=ver.index)
    )
    ver["Tarefa"] = id_col + " — " + nome_col
    cols = [
        "Tarefa", "WBS", "Frente", "_contrato", "Inicio_Planejado", "Fim_Planejado",
        "Duracao_dias", "%_Concluido", "Caminho_Critico_SIM_NAO", "Responsavel", "Status",
    ]
    st.dataframe(ver[[c for c in cols if c in ver.columns]], hide_index=True, height=320)

    cv = calcular_curva_s(t)
    if cv.empty:
        st.info("[FALTANTE] Sem datas planejadas para montar a curva S.")
    else:
        st.plotly_chart(_curva_s_fig(cv, 320), theme=None, config={"displaylogo": False})

    # equipe por contrato (banco de RH traz a coluna Contrato)
    if not rh.empty and "Contrato" in rh.columns:
        st.markdown("#### 👥 Equipe por contrato")
        eq = rh.copy()
        eq["_c"] = eq["Contrato"].astype(str).str.strip()
        eq = eq[eq["_c"] != ""]
        if not eq.empty:
            tab_eq = (
                eq.groupby("_c")
                .agg(Efetivo=("Matricula_Nome", "size"), Custo_Mês=("Custo_Mes", "sum"))
                .rename_axis("Contrato")
                .reset_index()
            )
            tab_eq["Custo_Mês"] = tab_eq["Custo_Mês"].map(brl)
            st.dataframe(tab_eq, hide_index=True)

    st.markdown("#### ⚠️ Alertas")
    erros, avisos, infos = coletar_alertas(t, dados)
    renderizar_alertas(erros, avisos, infos)


def indicadores_visao_geral(t: pd.DataFrame, dados: dict, ct: pd.DataFrame) -> None:
    """Indicadores cruzados no fim da Visão Geral: marcos atrasados, frentes
    paradas com prazo próximo, físico × financeiro e divergência Timeline."""
    st.markdown("#### 🎯 Marcos, frentes e físico × financeiro")
    tl = aba_efetiva(dados, "timeline")
    cols = st.columns(3)

    # 1) marcos atrasados (Timeline_Visao: Data_Marco vencida sem conclusão)
    marcos_atr = pd.DataFrame()
    if not tl.empty and "Data_Marco" in tl.columns:
        tl2 = tl.copy()
        tl2["_data"] = parse_datas(tl2["Data_Marco"])
        tl2["_av"] = (
            parse_num(tl2["_Avanço"]).fillna(0)
            if "_Avanço" in tl2.columns
            else parse_num(tl2.get("%_Avanço", pd.Series(0, index=tl2.index))).fillna(0)
        )
        status = (
            tl2["Status"].astype(str).str.lower()
            if "Status" in tl2.columns
            else pd.Series("", index=tl2.index)
        )
        concluido = status.str.contains("conclu") | (tl2["_av"] >= 100)
        marcos_atr = tl2[
            tl2["_data"].notna() & (tl2["_data"] < hoje()) & ~concluido
        ].sort_values("_data")
    with cols[0]:
        st.metric("Marcos atrasados", len(marcos_atr), delta_color="inverse", border=True)
        if not marcos_atr.empty:
            _drill_tarefas(
                "Ver marcos",
                marcos_atr[
                    [
                        c
                        for c in ("Projeto", "Frente", "Marco_Milestone", "Data_Marco", "Status", "%_Avanço")
                        if c in marcos_atr.columns
                    ]
                ],
                "",
            )

    # 2) frentes com avanço ~0 e compromisso (marco/fim) em até 45 dias
    paradas = []
    for frente, grupo in t.groupby("Frente", dropna=False):
        dur = float(grupo["_dur"].sum())
        av = float((grupo["_dur"] * grupo["_pct"]).sum() / dur) if dur else 0.0
        if av >= 5:
            continue
        fr = str(frente).strip() if pd.notna(frente) else ""
        prazo = pd.NaT
        if not tl.empty and "Frente" in tl.columns and "Data_Marco" in tl.columns:
            m = tl[tl["Frente"].astype(str).str.strip() == fr]
            if not m.empty:
                datas = parse_datas(m["Data_Marco"]).dropna()
                futuras = datas[datas >= hoje()]
                prazo = (
                    futuras.min()
                    if not futuras.empty
                    else (datas.max() if not datas.empty else pd.NaT)
                )
        if pd.isna(prazo):
            fims = grupo["Fim_Planejado_dt"].dropna()
            fims_futuros = fims[fims >= hoje()]
            prazo = fims_futuros.min() if not fims_futuros.empty else pd.NaT
        dias = int((prazo - hoje()).days) if pd.notna(prazo) else None
        if dias is not None and dias <= 45:
            paradas.append(
                {
                    "Frente": fr or "(sem frente)",
                    "Avanço_%": round(av, 1),
                    "Próximo prazo": fmt_data_opcional(prazo),
                    "Dias": dias,
                }
            )
    paradas_df = pd.DataFrame(paradas).sort_values("Dias") if paradas else pd.DataFrame()
    with cols[1]:
        st.metric(
            "Frentes paradas (prazo ≤ 45 d)",
            len(paradas_df),
            delta_color="inverse",
            border=True,
        )
        _drill_tarefas(
            "Ver frentes", paradas_df, "✅ Nenhuma frente parada com prazo próximo."
        )

    # 3) físico (cronograma) × financeiro (medido/total) por contrato
    ff = pd.DataFrame()
    if not ct.empty and "Valor_Total" in ct.columns and "_contrato" in t.columns:
        linhas = []
        for cid, grupo in t.groupby("_contrato"):
            if not str(cid).strip():
                continue
            dur = float(grupo["_dur"].sum())
            fis = float((grupo["_dur"] * grupo["_pct"]).sum() / dur) if dur else 0.0
            linha = ct[ct["ID_Contrato"].astype(str).str.strip() == str(cid)]
            if linha.empty:
                continue
            total = parse_num(pd.Series([linha["Valor_Total"].iloc[0]])).iloc[0]
            med_serie = (
                linha["Valor_Medido_Acumulado"]
                if "Valor_Medido_Acumulado" in linha.columns
                else pd.Series([0])
            )
            med = parse_num(pd.Series([med_serie.iloc[0]])).iloc[0]
            fin = float(med) / float(total) * 100 if pd.notna(total) and total else 0.0
            linhas.append(
                {
                    "Contrato": str(cid),
                    "Físico_%": round(fis, 1),
                    "Financeiro_%": round(fin, 1),
                    "Desvio_pp": round(fis - fin, 1),
                }
            )
        ff = pd.DataFrame(linhas)
    with cols[2]:
        st.metric(
            "Contratos com desvio > 10 p.p.",
            int((ff["Desvio_pp"].abs() > 10).sum()) if not ff.empty else 0,
            delta_color="inverse",
            border=True,
            help="Diferença em pontos percentuais entre o avanço físico e o % medido financeiramente.",
        )
        _drill_tarefas(
            "Ver físico × financeiro",
            ff,
            "[FALTANTE] Sem valores de contrato para comparar.",
        )
    if not ff.empty:
        longo = ff.melt(
            id_vars="Contrato",
            value_vars=["Físico_%", "Financeiro_%"],
            var_name="Indicador",
            value_name="%",
        )
        figff = px.bar(
            longo,
            x="Contrato",
            y="%",
            color="Indicador",
            barmode="group",
            color_discrete_map={"Físico_%": "#6366F1", "Financeiro_%": "#22C55E"},
            text=longo["%"].map(lambda v: f"{v:.1f}%"),
        )
        figff.update_traces(textposition="outside", marker_line_width=0)
        figff.update_layout(
            height=300,
            margin=dict(t=24, b=36),
            legend=dict(x=0.99, y=1.12, xanchor="right", orientation="h"),
            yaxis=dict(title="% (0–100)", range=[0, 112]),
            xaxis=dict(title=""),
        )
        estilizar_fig(figff)
        st.plotly_chart(figff, theme=None, config={"displaylogo": False})

    # divergência de avanço entre Timeline_Visao e o cronograma (mesma frente)
    if not tl.empty and "Frente" in tl.columns:
        divergs = []
        for frente, grupo in t.groupby("Frente", dropna=False):
            fr = str(frente).strip() if pd.notna(frente) else ""
            m = tl[tl["Frente"].astype(str).str.strip() == fr]
            if m.empty:
                continue
            av_tl = parse_num(m["_Avanço"] if "_Avanço" in m.columns else m.get("%_Avanço", pd.Series(0))).fillna(0).max()
            dur = float(grupo["_dur"].sum())
            av_cr = float((grupo["_dur"] * grupo["_pct"]).sum() / dur) if dur else 0.0
            if abs(float(av_tl) - av_cr) > 10:
                divergs.append(
                    f"Frente {fr}: Timeline {float(av_tl):.0f}% × cronograma {av_cr:.0f}%"
                )
        if divergs:
            st.warning(
                "📊 **Avanço divergente entre Timeline_Visao e o cronograma** — "
                + " • ".join(divergs[:5])
                + (" …" if len(divergs) > 5 else "")
                + " Atualize o % no cronograma (aba *Por Cronograma*) ou na Timeline."
            )


def secao_cronograma_rapido(dados: dict, t: pd.DataFrame) -> None:
    """Gantt + avanço rápido: editar o % na linha e excluir tarefas (estilo MS Project)."""
    fig = fig_gantt(t)
    if fig is None:
        st.info("[FALTANTE] Sem tarefas com datas planejadas completas no recorte.")
    else:
        st.plotly_chart(fig, theme=None, config={"displaylogo": False})
        st.caption("Barra vermelha = caminho crítico • linha laranja = hoje.")
    if t.empty:
        return

    msg = st.session_state.pop("aviso_avancos", None)
    if msg:
        st.success(msg)

    st.markdown("#### ✏️ Avanço rápido")
    st.caption(
        "Duplo clique na coluna **% de avanço** para digitar o novo valor (0–100); "
        "a ✕ na lateral esquerda de cada linha marca a tarefa para exclusão. "
        "Depois clique em **Aplicar** — o Gantt e os KPIs atualizam na hora."
    )
    versao = int(st.session_state.get("versao_avancos", 0))
    orig = (
        t.set_index("ID_Tarefa")["_pct"].round(1).to_dict()
    )
    cols_editor = [
        c
        for c in (
            "ID_Tarefa", "Nome_Tarefa", "Frente", "Responsavel",
            "Fim_Planejado", "%_Concluido", "Status", "Caminho_Critico_SIM_NAO",
        )
        if c in t.columns
    ]
    ed = t[cols_editor].copy()
    ed["%_Concluido"] = (
        pd.to_numeric(ed["%_Concluido"], errors="coerce").fillna(0).clip(0, 100)
    )
    editado = st.data_editor(
        ed,
        num_rows="dynamic",
        hide_index=True,
        key=f"editor_avanco_{versao}",
        disabled=[c for c in cols_editor if c != "%_Concluido"],
        column_config={
            "%_Concluido": st.column_config.NumberColumn(
                "% de avanço", min_value=0, max_value=100, step=5, format="%d"
            ),
        },
    )
    editado = _sem_linhas_vazias_ids(editado)
    ids_orig = {str(x).strip() for x in orig}
    ids_edit = set(editado["ID_Tarefa"].astype(str).str.strip())
    excluidas = sorted(ids_orig - ids_edit)
    avancos: dict[str, float] = {}
    for x in editado.to_dict("records"):
        iid = str(x.get("ID_Tarefa", "")).strip()
        if not iid or iid not in orig:
            continue
        novo = float(x.get("%_Concluido") or 0)
        if abs(novo - orig[iid]) > 0.01:
            avancos[iid] = min(100.0, max(0.0, novo))
    if not avancos and not excluidas:
        return
    if avancos:
        st.markdown("**Avanços pendentes:**")
        for iid, novo in list(avancos.items())[:20]:
            st.write(f"- `{iid}`: {orig[iid]:.0f}% → **{novo:.0f}%**")
        if len(avancos) > 20:
            st.caption(f"... e mais {len(avancos) - 20} tarefa(s).")
    confirmar = True
    if excluidas:
        st.warning(
            f"**{len(excluidas)} tarefa(s) serão EXCLUÍDAS do cronograma:** "
            + ", ".join(f"`{i}`" for i in excluidas[:10])
            + (" …" if len(excluidas) > 10 else "")
        )
        confirmar = st.checkbox(
            f"Confirmo a exclusão de {len(excluidas)} tarefa(s)",
            key=f"confirma_excl_{versao}",
        )
    if st.button(
        "✅ Aplicar avanços e exclusões",
        type="primary",
        disabled=not confirmar or (not avancos and not excluidas),
        key=f"btn_aplicar_avancos_{versao}",
    ):
        alteradas = aplicar_avancos(dados, avancos, excluidas)
        st.session_state["versao_avancos"] = versao + 1
        partes = []
        if alteradas:
            partes.append(f"{alteradas} tarefa(s) com % atualizado")
        if excluidas:
            partes.append(f"{len(excluidas)} excluída(s)")
        st.session_state["aviso_avancos"] = " • ".join(partes) + "."
        st.rerun()


def _sem_linhas_vazias_ids(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas sem ID_Tarefa (＋ acidental no data_editor não vira tarefa)."""
    if df.empty or "ID_Tarefa" not in df.columns:
        return df
    mascara = df["ID_Tarefa"].astype(str).str.strip().isin(["", "nan", "None"])
    return df[~mascara].reset_index(drop=True) if mascara.any() else df


def pagina_dashboard(dados: dict, template: str | None = None) -> None:
    template = template or seg(
        st.sidebar.radio(
            "Template do Dashboard",
            ["Executivo", "Analítico"],
            horizontal=True,
            key="template_dash",
        ),
        "Executivo",
    )
    cabecalho(
        "📊",
        "Gestão de Projetos",
        "Visão geral do projeto e cronograma com avanço rápido — igual ao MS Project",
    )
    with st.expander("ℹ️ Como usar este Dashboard"):
        st.markdown(
            f"""
**Você está no template {template}** (troque na barra lateral, em *Template do Dashboard*).

| Aba | Para que serve |
|---|---|
| **🌍 Visão Geral** | Uma tela só com o projeto inteiro: KPIs, curva S, alertas + marcos atrasados, frentes paradas com prazo próximo e físico × financeiro por contrato |
| **📅 Por Cronograma** | Gantt com caminho crítico em vermelho + **avanço rápido**: digite o % na linha da tarefa (como no MS Project), exclua tarefas com a ✕ e aplique — tudo recalcula na hora |

**Como usar**
1. **Filtre por contrato** no seletor do topo — vale para as duas abas.
2. **Visão Geral**: clique nos cartões dos KPIs para ver a lista por trás do número.
3. **Por Cronograma**: duplo clique na coluna *% de avanço*, digite o valor, **✅ Aplicar**. Exclusões pedem confirmação.
4. **Alertas** (🔴 erro / 🟠 atenção / 🔵 informação) são sempre do projeto inteiro.
"""
        )
    t = preparar_cronograma(cronograma_efetiva(dados))
    if t.empty:
        st.info("[FALTANTE] Sem dados de cronograma.")
        return

    # filtro por contrato (mesma regra da Timeline: coluna Contrato, senão Frente → Projeto)
    mapa = mapa_frente_contrato(
        aba_efetiva(dados, "timeline"), aba_efetiva(dados, "contratos")
    )
    t = vincular_contrato(t, mapa)
    op_contratos = ["Todos"] + sorted(set(t["_contrato"].dropna()) - {""})
    contrato = seg(st.selectbox("Contrato", op_contratos, key="contrato_dash"), "Todos")
    rh = aba_efetiva(dados, "rh")
    ct = aba_efetiva(dados, "contratos")
    if contrato != "Todos":
        t = t[t["_contrato"] == contrato]
        if not rh.empty and "Contrato" in rh.columns:
            rh = rh[rh["Contrato"].astype(str).str.strip() == contrato]
        if not ct.empty and "ID_Contrato" in ct.columns:
            ct = ct[ct["ID_Contrato"].astype(str).str.strip() == contrato]
        if t.empty:
            st.info(f"[FALTANTE] Nenhuma tarefa vinculada ao contrato {contrato}.")
            return

    aba_vg, aba_cr = st.tabs(["🌍 Visão Geral", "📅 Por Cronograma (avanço rápido)"])
    with aba_vg:
        if template == "Analítico":
            dashboard_analitico(t, dados, ct, rh, mapa)
        else:
            dashboard_executivo(t, dados, ct, rh, mapa)
        indicadores_visao_geral(t, dados, ct)
    with aba_cr:
        secao_cronograma_rapido(dados, t)


# ------------------------------------------------------- 2 Cronograma Timeline
def mapa_frente_contrato(timeline: pd.DataFrame, contratos: pd.DataFrame) -> dict[str, str]:
    """Timeline_Visao liga Frente (1º nível do WBS) → Projeto (= ID_Contrato)."""
    if timeline.empty or "Frente" not in timeline.columns or "Projeto" not in timeline.columns:
        return {}
    ids = (
        set(contratos["ID_Contrato"].astype(str))
        if not contratos.empty and "ID_Contrato" in contratos.columns
        else set()
    )
    mapa = {}
    for frente, grupo in timeline.groupby(timeline["Frente"].astype(str).str.strip()):
        projeto = str(grupo["Projeto"].iloc[0]).strip()
        if not frente or frente.lower() == "nan":
            continue
        if ids and projeto not in ids:
            continue
        mapa[frente] = projeto
    return mapa


def pagina_timeline(dados: dict) -> None:
    cabecalho(
        "📅",
        "Cronograma Timeline",
        "Gantt com caminho crítico em vermelho e filtros por contrato, frente, responsável e status",
    )
    t = preparar_cronograma(cronograma_efetiva(dados))
    if t.empty:
        st.info("[FALTANTE] Sem dados de cronograma.")
        return

    mapa = mapa_frente_contrato(aba_efetiva(dados, "timeline"), aba_efetiva(dados, "contratos"))
    if "Contrato" in t.columns and bool(t["Contrato"].astype(str).str.strip().ne("").any()):
        t["_contrato"] = t["Contrato"].astype(str).str.strip()
        origem_contrato = "coluna Contrato do próprio cronograma"
    else:
        t["_contrato"] = t["Frente"].map(mapa)
        origem_contrato = "Timeline_Visao (Frente → Projeto)"

    c1, c2, c3, c4 = st.columns(4)
    op_contratos = ["Todos"] + sorted(set(t["_contrato"].dropna()) - {""})
    contrato = seg(c1.selectbox("Contrato", op_contratos), "Todos")
    op_frentes = ["Todas"] + sorted(set(t["Frente"].dropna()) - {""})
    frente = seg(c2.selectbox("Frente", op_frentes), "Todas")
    responsaveis = ["Todos"] + sorted(set(t.get("Responsavel", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    responsavel = seg(c3.selectbox("Responsável", responsaveis), "Todos")
    status_disp = ["Todos"] + sorted(set(t.get("Status", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    status = seg(c4.selectbox("Status", status_disp), "Todos")

    base = t
    if contrato != "Todos":
        base = base[base["_contrato"] == contrato]
    if frente != "Todas":
        base = base[base["Frente"] == frente]
    if responsavel != "Todos":
        base = base[base["Responsavel"].astype(str) == responsavel]
    if status != "Todos":
        base = base[base["Status"].astype(str) == status]

    com_datas = base.dropna(subset=["Inicio_Planejado_dt", "Fim_Planejado_dt"])
    invertidas = com_datas[com_datas["Fim_Planejado_dt"] < com_datas["Inicio_Planejado_dt"]]
    descartadas = len(base) - len(com_datas) + len(invertidas)
    if descartadas:
        st.warning(
            f"[FALTANTE] {descartadas} tarefa(s) fora do gráfico: sem datas planejadas completas "
            "ou com fim anterior ao início."
        )
    fig = fig_gantt(base, origem_contrato)
    if fig is None:
        st.info("Nenhuma tarefa no recorte.")
        return
    st.plotly_chart(fig, theme=None, config={"displaylogo": False})
    st.caption(f"Barra vermelha = caminho crítico. Contratos via {origem_contrato}; Frente = 1º nível do WBS.")

    with st.expander("Ver dados do recorte"):
        g = com_datas.copy()
        g["Tarefa"] = g["ID_Tarefa"].astype(str) + " — " + g["Nome_Tarefa"].astype(str)
        st.dataframe(tabela_cronograma(g), hide_index=True)


# ------------------------------------------------------- 3 Histograma MO + Equip
def _hist_kpis(f: pd.DataFrame, col_prev: str, col_real: str, metrica: str) -> None:
    tprev = float(f[col_prev].sum())
    treal = float(f[col_real].sum())
    fmt = brl if "Custo" in metrica else (lambda v: num_pt(v, 0))
    desvio = ((treal - tprev) / tprev * 100) if tprev else None
    k = st.columns(3)
    k[0].metric(f"Total previsto", fmt(tprev), border=True)
    k[1].metric("Total real", fmt(treal), border=True)
    k[2].metric(
        "Desvio",
        pct_txt(desvio, 1) if desvio is not None else "—",
        delta=f"{treal - tprev:+,.0f}".replace(",", ".") if "Custo" not in metrica else None,
        delta_color="inverse" if desvio is not None else "off",
        border=True,
    )


def _grafico_previsto_real(longo: pd.DataFrame, ordem_meses: list[str], rotulo_y: str, chave: str) -> None:
    fig = px.bar(
        longo,
        x="Mes_Semana",
        y="Valor",
        color="Variável",
        barmode="group",
        color_discrete_map={"Previsto": "#94A3B8", "Real": "#6366F1"},
        category_orders={"Mes_Semana": ordem_meses},
        labels={"Mes_Semana": "Mês", "Valor": rotulo_y},
    )
    fig.update_layout(
        height=420,
        margin=dict(t=30, b=46),
        legend_title_text="",
        # legenda horizontal acima da área de plotagem (a posição padrão fica
        # colada na borda direita e o texto "Previsto" saía cortado).
        legend=dict(orientation="h", y=1.12),
    )
    estilizar_fig(fig)
    # chave: o plotly_chart registra ID a cada execução; MO e Equip rodam juntas.
    st.plotly_chart(fig, theme=None, config={"displaylogo": False}, key=f"hist_graf_{chave}")


def _tabela_previsto_real(agg: pd.DataFrame, col_prev: str, col_real: str, metrica: str) -> None:
    tab = agg[["Mes_Semana", col_prev, col_real]].copy()
    tab["Desvio"] = tab[col_real] - tab[col_prev]
    tab = tab.rename(columns={col_prev: "Previsto", col_real: "Real"})
    fmt = brl if "Custo" in metrica else (lambda v: num_pt(v, 0))
    for col in ("Previsto", "Real", "Desvio"):
        tab[col] = tab[col].map(fmt)
    st.dataframe(tab, hide_index=True)


def _hist_mo(mo: pd.DataFrame) -> None:
    if mo.empty:
        st.info("[FALTANTE] Sem dados em DB_Histograma_MO.")
        return
    mo = mo.copy()
    # As abas MO e Equip rodam na MESMA execução: sem key único, dois
    # selectbox "Contrato" com opções iguais disparam StreamlitDuplicateElementId.
    c1, c2, c3 = st.columns(3)
    contratos = ["Todos"] + sorted(set(mo.get("Contrato_Vinculado", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    contrato = seg(c1.selectbox("Contrato", contratos, key="hist_mo_contrato"), "Todos")
    frentes = ["Todas"] + sorted(set(mo.get("Projeto_Frente", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    frente = seg(c2.selectbox("Frente", frentes, key="hist_mo_frente"), "Todas")
    funcoes = ["Todas"] + sorted(set(mo.get("Funcao_Cargo", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    funcao = seg(c3.selectbox("Função", funcoes, key="hist_mo_funcao"), "Todas")
    metrica = seg(
        st.radio("Métrica", ["Horas (HH)", "Quantidade", "Custo (R$)"], horizontal=True, key="hist_mo_metrica"),
        "Horas (HH)",
    )

    f = mo
    if contrato != "Todos":
        f = f[f["Contrato_Vinculado"].astype(str) == contrato]
    if frente != "Todas":
        f = f[f["Projeto_Frente"].astype(str) == frente]
    if funcao != "Todas":
        f = f[f["Funcao_Cargo"].astype(str) == funcao]
    if f.empty:
        st.info("Nenhum dado no recorte.")
        return

    col_prev, col_real = {
        "Horas (HH)": ("HH_Previsto", "HH_Real"),
        "Quantidade": ("Quantidade_Prevista", "Quantidade_Real"),
        "Custo (R$)": ("Custo_Total_Previsto", "Custo_Total_Real"),
    }[metrica]
    f = f.copy()
    f[col_prev] = parse_num(f[col_prev]).fillna(0)
    f[col_real] = parse_num(f[col_real]).fillna(0)
    agg = f.groupby("Mes_Semana", sort=False)[[col_prev, col_real]].sum().reset_index()
    agg["_ord"] = agg["Mes_Semana"].map(chave_mes)
    agg = agg.sort_values("_ord")
    if agg.empty:
        st.info("Nenhum dado no recorte.")
        return

    _hist_kpis(f, col_prev, col_real, metrica)
    longo = agg.melt(
        id_vars=["Mes_Semana"],
        value_vars=[col_prev, col_real],
        var_name="Variável",
        value_name="Valor",
    )
    longo["Variável"] = longo["Variável"].map({col_prev: "Previsto", col_real: "Real"})
    _grafico_previsto_real(
        longo, agg["Mes_Semana"].tolist(), f"{metrica} (previsto x real)", "mo"
    )
    _tabela_previsto_real(agg, col_prev, col_real, metrica)
    st.caption("Desvio = Real − Previsto. Cruze com o cronograma para explicar picos e ociosidade.")


def _hist_equip(equip: pd.DataFrame) -> None:
    if equip.empty:
        st.info("[FALTANTE] Sem dados em DB_Histograma_Equip.")
        return
    equip = equip.copy()
    c1, c2 = st.columns(2)
    nomes = ["Todos"] + sorted(set(equip.get("Equipamento", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    nome = seg(c1.selectbox("Equipamento", nomes, key="hist_equip_equipamento"), "Todos")
    contratos = ["Todos"] + sorted(set(equip.get("Contrato_Vinculado", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    contrato = seg(c2.selectbox("Contrato", contratos, key="hist_equip_contrato"), "Todos")
    metrica = seg(
        st.radio("Métrica", ["Horas", "Quantidade"], horizontal=True, key="hist_equip_metrica"),
        "Horas",
    )

    f = equip
    if nome != "Todos":
        f = f[f["Equipamento"].astype(str) == nome]
    if contrato != "Todos":
        f = f[f["Contrato_Vinculado"].astype(str) == contrato]
    if f.empty:
        st.info("Nenhum dado no recorte.")
        return

    col_prev, col_real = {
        "Horas": ("Horas_Previstas", "Horas_Reais"),
        "Quantidade": ("Qtd_Prevista", "Qtd_Real"),
    }[metrica]
    f = f.copy()
    f[col_prev] = parse_num(f[col_prev]).fillna(0)
    f[col_real] = parse_num(f[col_real]).fillna(0)
    agg = f.groupby("Mes_Semana", sort=False)[[col_prev, col_real]].sum().reset_index()
    agg["_ord"] = agg["Mes_Semana"].map(chave_mes)
    agg = agg.sort_values("_ord")
    if agg.empty:
        st.info("Nenhum dado no recorte.")
        return

    _hist_kpis(f, col_prev, col_real, metrica)
    longo = agg.melt(
        id_vars=["Mes_Semana"],
        value_vars=[col_prev, col_real],
        var_name="Variável",
        value_name="Valor",
    )
    longo["Variável"] = longo["Variável"].map({col_prev: "Previsto", col_real: "Real"})
    _grafico_previsto_real(longo, agg["Mes_Semana"].tolist(), f"{metrica} (previsto x real)", "equip")
    _tabela_previsto_real(agg, col_prev, col_real, metrica)
    if "Custo_Mensal" in f.columns:
        custo = parse_num(f["Custo_Mensal"]).fillna(0).sum()
        st.caption(f"Custo mensal de locação no recorte (soma dos meses exibidos): {brl(custo)}.")


def pagina_histograma(dados: dict) -> None:
    cabecalho("👷", "Histograma MO + Equip", "Previsto x real mês a mês para mão de obra e equipamentos")
    tab_mo, tab_eq = st.tabs(["Mão de obra", "Equipamentos"])
    with tab_mo:
        _hist_mo(aba_efetiva(dados, "hist_mo"))
    with tab_eq:
        _hist_equip(aba_efetiva(dados, "hist_equip"))


# ---------------------------------------------------------- 4 Recursos Humanos
def pagina_rh(dados: dict) -> None:
    cabecalho("👥", "Recursos Humanos", "Efetivo, custo mensal e alocação por frente")
    rh = aba_efetiva(dados, "rh")
    if rh.empty:
        st.info("[FALTANTE] Sem dados em DB_Recursos_Humanos.")
        return
    rh = rh.copy()
    rh["_aloc"] = parse_num(rh.get("Alocacao_%", pd.Series(0, index=rh.index))).fillna(0)
    rh["_custo"] = parse_num(rh.get("Custo_Mes", pd.Series(0, index=rh.index))).fillna(0)

    c1, c2, c3 = st.columns(3)
    situacoes = ["Todos"] + sorted(
        set(rh.get("Status_Ativo_Inativo", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"}
    )
    situacao = seg(c1.selectbox("Situação", situacoes), "Todos")
    equipes = ["Todas"] + sorted(set(rh.get("Equipe_Frente", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    equipe = seg(c2.selectbox("Equipe / Frente", equipes), "Todas")
    contratos = ["Todos"] + sorted(set(rh.get("Contrato", pd.Series(dtype=str)).fillna("").astype(str)) - {"", "nan"})
    contrato = seg(c3.selectbox("Contrato", contratos), "Todos")

    f = rh
    if situacao != "Todos":
        f = f[f["Status_Ativo_Inativo"].astype(str) == situacao]
    if equipe != "Todas":
        f = f[f["Equipe_Frente"].astype(str) == equipe]
    if contrato != "Todos":
        f = f[f["Contrato"].astype(str) == contrato]
    if f.empty:
        st.info("Nenhum recurso no recorte.")
        return

    k = st.columns(3)
    k[0].metric("Pessoas no recorte", len(f), border=True)
    k[1].metric("Custo mensal total", brl(f["_custo"].sum()), border=True)
    k[2].metric("Alocação média", pct_txt(f["_aloc"].mean(), 1), border=True)

    alertas = alertas_alocacao(f)
    if alertas:
        for msg in alertas:
            st.error(msg)
    else:
        st.success("✅ Nenhuma alocação acima de 100% no recorte.")

    tabela = f.copy()
    for col in ("Data_Admissao", "Data_Desligamento"):
        if col in tabela.columns:
            tabela[col] = tabela[col].map(fmt_data_opcional)
    if "Custo_Mes" in tabela.columns:
        tabela["Custo_Mes"] = tabela["_custo"].map(brl)
    if "Alocacao_%" in tabela.columns:
        tabela["Alocacao_%"] = tabela["_aloc"].map(lambda v: pct_txt(v, 0))
    tabela = tabela.drop(columns=["_aloc", "_custo"], errors="ignore")
    st.dataframe(tabela, hide_index=True)
    st.caption("Coluna Alocacao_% acima de 100% sinaliza sobreposição de escalas — ajuste a escala ou a alocação.")


# --------------------------------------------------------------- 5 Contratos
def _vincular_cronograma(dados: dict, alvo: str) -> None:
    """Anexa um arquivo de cronograma ao contrato (inclui, não substitui)."""
    arq = st.file_uploader(
        "Arquivo de cronograma (.mpp, CSV, XLSX ou XML)",
        type=["csv", "xlsx", "xml", "mpp"],
        key="ctr_up_cron",
        help="As tarefas entram já etiquetadas com o contrato escolhido.",
    )
    if arq is None:
        return
    try:
        novo, avisos = interpretar_upload(arq.name, arq.getvalue())
    except Exception as exc:
        st.error(f"Não foi possível ler o arquivo: {exc}")
        return
    if novo is None:
        return
    novo, _ = garantir_colunas(novo)
    erros, avisos_v = validar_cronograma(novo)
    renderizar_alertas(erros, avisos + avisos_v, [])
    st.dataframe(tabela_cronograma(novo).head(20), hide_index=True)
    if st.button(
        f"➕ Acrescentar {len(novo)} tarefa(s) ao contrato {alvo}",
        type="primary",
        disabled=bool(erros),
        key="ctr_btn_cron",
    ):
        _, renumeradas = acrescentar_aba(dados, "cronograma", novo, contrato=alvo)
        extra = f" ({renumeradas} renumerada(s) para não colidir)" if renumeradas else ""
        st.session_state["aviso_ctr"] = (
            f"{len(novo)} tarefa(s) incluída(s) no cronograma de {alvo}{extra}."
        )
        st.rerun()


def _vincular_histograma(dados: dict, alvo: str) -> None:
    """Anexa um arquivo de histograma (MO/Equip) ao contrato."""
    arq = st.file_uploader(
        "Arquivo de histograma (CSV ou XLSX)",
        type=["csv", "xlsx"],
        key="ctr_up_hist",
        help="A aba é detectada pelo cabeçalho; as linhas entram com o contrato escolhido.",
    )
    if arq is None:
        return
    try:
        df_h, chave_h, avisos_h = interpretar_histograma(arq.name, arq.getvalue())
    except Exception as exc:
        st.error(f"Não foi possível ler o arquivo: {exc}")
        return
    if df_h is None:
        for msg in avisos_h:
            st.error(msg)
        return
    erros_h, avisos_v = validar_aba(chave_h, df_h)
    renderizar_alertas(erros_h, avisos_h + avisos_v, [])
    st.caption(f"Reconhecido como **{ABAS[chave_h]}**: {len(df_h)} linha(s).")
    st.dataframe(df_h.head(20), hide_index=True)
    if st.button(
        f"➕ Acrescentar {len(df_h)} linha(s) ao contrato {alvo}",
        type="primary",
        disabled=bool(erros_h),
        key="ctr_btn_hist",
    ):
        _, renumeradas = acrescentar_aba(dados, chave_h, df_h, contrato=alvo)
        extra = f" ({renumeradas} renumerada(s))" if renumeradas else ""
        st.session_state["aviso_ctr"] = (
            f"{len(df_h)} linha(s) de {ABAS[chave_h]} incluídas em {alvo}{extra}."
        )
        st.rerun()


def ir_para_pagina(nome: str) -> None:
    """Navega para outra página trocando o rádio da barra lateral (via callback)."""
    st.session_state["paginas"] = nome


def pagina_contratos(dados: dict) -> None:
    cabecalho(
        "📑",
        "Contratos",
        "Cadastro, valor, avanço e vínculo de cronograma/histograma por contrato",
    )
    ct = aba_efetiva(dados, "contratos")
    tl = aba_efetiva(dados, "timeline")

    with st.expander("➕ Cadastrar novo contrato"):
        f1, f2 = st.columns(2)
        novo_id = (
            f1.text_input("ID do contrato", key="ctr_novo_id", placeholder="ex.: CTR-04")
            or ""
        ).strip()
        novo_cliente = (
            f2.text_input("Cliente / Fornecedor", key="ctr_novo_cliente") or ""
        ).strip()
        f3, f4 = st.columns(2)
        novo_objeto = (
            f3.text_input("Objeto / Escopo", key="ctr_novo_objeto") or ""
        ).strip()
        novo_valor = float(
            f4.number_input(
                "Valor total (R$)", min_value=0.0, step=1000.0, key="ctr_novo_valor"
            )
        )
        f5, f6, f7 = st.columns(3)
        novo_ini = (
            f5.text_input("Data de início (DD/MM/AAAA)", key="ctr_novo_ini") or ""
        ).strip()
        novo_fim = (
            f6.text_input("Data de fim (DD/MM/AAAA)", key="ctr_novo_fim") or ""
        ).strip()
        novo_status = seg(
            f7.selectbox(
                "Status",
                ["Em execução", "Não Iniciado", "Concluído", "Suspenso"],
                key="ctr_novo_status",
            ),
            "Em execução",
        )
        salvar_agora = st.checkbox(
            "💾 Salvar na fonte agora (gravar no Sheets/CSV)",
            value=True,
            key="ctr_novo_salvar",
            help=(
                "Marcado = o contrato é gravado já na origem (Sheets ou CSV local) "
                "e vale para qualquer usuário/sessão. Desmarcado = vale só nesta "
                "sessão até você usar 💾 Salvar tudo na fonte na barra lateral."
            ),
        )
        datas_ok = all(
            v == "" or pd.notna(parse_datas(pd.Series([v])).iloc[0])
            for v in (novo_ini, novo_fim)
        )
        if not datas_ok:
            st.warning("[FALTANTE] Data em formato inválido — use DD/MM/AAAA.")
        if st.button(
            "➕ Cadastrar contrato",
            type="primary",
            disabled=not novo_id or not datas_ok,
            key="ctr_btn_novo",
        ):
            incluir_contrato(
                dados,
                novo_id,
                {
                    "Cliente_Fornecedor": novo_cliente,
                    "Objeto_Escopo": novo_objeto,
                    "Valor_Total": novo_valor,
                    "Data_Inicio": novo_ini,
                    "Data_Fim": novo_fim,
                    "Status": novo_status,
                },
            )
            if salvar_agora:
                # grava já na origem (Sheets/CSV) e limpa a pendência desta aba
                ctx = st.session_state.get("ctx_gravacao", {})
                try:
                    msg = salvar_aba(
                        "contratos",
                        aba_efetiva(dados, "contratos"),
                        ctx.get("fonte", "CSV local (exemplos)"),
                        ctx.get("id_planilha", ""),
                        ctx.get("api_url", ""),
                        ctx.get("api_key", ""),
                        original=dados.get("contratos", pd.DataFrame()),
                    )
                    st.session_state.pop("contratos_importado", None)
                    st.session_state.pop("contratos_importado_nome", None)
                    st.session_state["aviso_ctr"] = (
                        f"Contrato **{novo_id}** cadastrado e salvo na fonte. {msg}"
                    )
                except Exception as exc:
                    st.session_state["aviso_ctr"] = (
                        f"Contrato **{novo_id}** cadastrado só na sessão — "
                        f"falha ao salvar na fonte: {exc}"
                    )
            else:
                st.session_state["aviso_ctr"] = (
                    f"Contrato **{novo_id}** cadastrado na sessão "
                    "(use 💾 Salvar tudo na fonte para gravar)."
                )
            st.rerun()

    msg = st.session_state.pop("aviso_ctr", None)
    if msg:
        st.success(msg)
        st.button(
            "✏️ Ir para Editar Dados",
            on_click=ir_para_pagina,
            args=("✏️ Editar Dados",),
            key="ctr_link_editar",
        )

    st.markdown("#### 📎 Vincular cronograma / histograma a um contrato")
    ids = ids_contratos(dados)
    if not ids:
        st.info("[FALTANTE] Cadastre um contrato acima para anexar cronograma/histograma.")
    else:
        alvo = seg(st.selectbox("Contrato de destino", ids, key="ctr_vinc_alvo"), ids[0])
        aba_cron, aba_hist = st.tabs(["📅 Cronograma", "📊 Histograma (MO / Equip)"])
        with aba_cron:
            _vincular_cronograma(dados, alvo)
        with aba_hist:
            _vincular_histograma(dados, alvo)

    with st.expander("🗑️ Excluir contrato"):
        if not ids:
            st.info("Nenhum contrato para excluir.")
        else:
            alvo_exc = seg(
                st.selectbox("Contrato a excluir", ids, key="ctr_exc_alvo"), ids[0]
            )
            ct_base = aba_efetiva(dados, "contratos")
            linha = ct_base[ct_base["ID_Contrato"].astype(str).str.strip() == alvo_exc]
            if not linha.empty:
                total_exc = parse_num(pd.Series([linha["Valor_Total"].iloc[0]])).iloc[0]
                st.caption(
                    f"**{alvo_exc}** — {linha.get('Cliente_Fornecedor', pd.Series([''])).iloc[0]} • "
                    f"valor {brl(total_exc) if pd.notna(total_exc) else '—'}"
                )
            impacto_prev: dict[str, int] = {}
            for chave, rotulo in (
                ("cronograma", "tarefa(s) do cronograma"),
                ("hist_mo", "linha(s) do histograma de MO"),
                ("hist_equip", "linha(s) do histograma de Equip"),
                ("rh", "pessoa(s) de RH"),
            ):
                base_k = aba_efetiva(dados, chave)
                col = COLUNA_CONTRATO.get(chave, "Contrato")
                if not base_k.empty and col in base_k.columns:
                    n = int((base_k[col].astype(str).str.strip() == alvo_exc).sum())
                    if n:
                        impacto_prev[chave] = n
            tl_base = aba_efetiva(dados, "timeline")
            if not tl_base.empty and "Projeto" in tl_base.columns:
                n = int((tl_base["Projeto"].astype(str).str.strip() == alvo_exc).sum())
                if n:
                    impacto_prev["timeline"] = n
            if impacto_prev:
                st.warning(
                    "Impacto da exclusão: **1 contrato** será removido e os vínculos ficarão "
                    "vazios em "
                    + ", ".join(
                        f"{v} {rot}"
                        for k, v in impacto_prev.items()
                        for rot in [
                            {
                                "cronograma": "tarefa(s) do cronograma",
                                "hist_mo": "linha(s) do histograma de MO",
                                "hist_equip": "linha(s) do histograma de Equip",
                                "rh": "pessoa(s) de RH",
                                "timeline": "marco(s) da Timeline_Visao",
                            }.get(k, k)
                        ]
                    )
                    + ". As linhas em si permanecem — só perdem o vínculo "
                    "(a Timeline tem os marcos removidos, pois cada marco pertence a um contrato)."
                )
            else:
                st.caption("Este contrato não tem vínculos nas outras abas — só ele será removido.")
            confirmado = st.checkbox(
                f"Confirmo a exclusão do contrato {alvo_exc}", key="ctr_exc_confirma"
            )
            if st.button(
                f"🗑️ Excluir {alvo_exc}",
                type="primary",
                disabled=not confirmado,
                key="ctr_exc_btn",
            ):
                impacto = excluir_contrato(dados, alvo_exc)
                st.session_state["aviso_ctr"] = (
                    f"Contrato **{alvo_exc}** excluído. Vínculos ajustados: "
                    + (
                        ", ".join(f"{v} em {k}" for k, v in impacto.items())
                        if len(impacto) > 1
                        else "nenhum outro"
                    )
                    + "."
                )
                st.rerun()

    if ct.empty:
        st.info("[FALTANTE] Sem dados em DB_Contratos.")
        return
    ct = ct.copy()
    ct["_total"] = parse_num(ct.get("Valor_Total", pd.Series(0, index=ct.index)))
    ct["_medido"] = parse_num(ct.get("Valor_Medido_Acumulado", pd.Series(0, index=ct.index)))
    ct["_saldo"] = parse_num(ct.get("Saldo", pd.Series(0, index=ct.index)))
    ct["_ini"] = parse_datas(ct.get("Data_Inicio", pd.Series("", index=ct.index)))
    ct["_fim"] = parse_datas(ct.get("Data_Fim", pd.Series("", index=ct.index)))

    for i, x in enumerate(ct.to_dict("records")):
        if i:
            st.write("")
        with st.container(border=True):
            st.markdown(f"#### {x.get('ID_Contrato', '')} — {x.get('Cliente_Fornecedor', '')}")
            st.caption(f"{x.get('Objeto_Escopo', '')} • Status: **{x.get('Status', '')}**")
            total = float(x["_total"]) if pd.notna(x["_total"]) else None
            medido = float(x["_medido"]) if pd.notna(x["_medido"]) else None
            saldo = float(x["_saldo"]) if pd.notna(x["_saldo"]) else None
            avanco = (medido / total * 100) if total and medido is not None else None
            k = st.columns(4)
            k[0].metric("Valor total", brl(total), border=True)
            k[1].metric("Medido acumulado", brl(medido), border=True)
            k[2].metric("Saldo", brl(saldo), border=True)
            k[3].metric("% avanço financeiro", pct_txt(avanco, 1) if avanco is not None else "[FALTANTE]", border=True)
            if avanco is not None:
                st.progress(min(max(avanco / 100.0, 0.0), 1.0))

            prazo = f"{fmt_data(x['_ini'])} → {fmt_data(x['_fim'])}"
            if pd.notna(x["_fim"]):
                dias = int((x["_fim"] - hoje()).days)
                if dias < 0:
                    prazo += f" — vencido há {abs(dias)} dia(s) ⚠️"
                elif dias <= 15:
                    prazo += f" — vence em {dias} dia(s) ⚠️"
            st.caption(
                f"Período: {prazo} • Reajuste: {x.get('Reajuste', '')} • "
                f"Aditivos: {x.get('Aditivos', '')} • Gestão: {x.get('Responsavel_Gestao', '')}"
            )

            st.markdown("**Marcos vinculados (Timeline_Visao):**")
            marcos = (
                tl[tl["Projeto"].astype(str).str.strip() == str(x.get("ID_Contrato", "")).strip()]
                if not tl.empty and "Projeto" in tl.columns
                else tl.iloc[0:0]
            )
            if marcos.empty:
                st.markdown("- [FALTANTE] nenhum marco vinculado a este contrato.")
            else:
                for m in marcos.to_dict("records"):
                    av = parse_num(pd.Series([m.get("%_Avanço", 0)])).iloc[0]
                    desvio = parse_num(pd.Series([m.get("Desvio_dias", 0)])).iloc[0]
                    desvio_txt = "" if pd.isna(desvio) else f", desvio {int(desvio):+d} d"
                    st.markdown(
                        f"- {fmt_data(m.get('Data_Marco'))} — **{m.get('Marco_Milestone', '')}** "
                        f"({m.get('Status', '')}, {pct_txt(av, 0)}{desvio_txt})"
                    )


# ------------------------------------------------------- 6 Importar / Exportar
def pagina_importar_exportar(dados: dict) -> None:
    cabecalho(
        "🔁",
        "Importar / Exportar",
        "CSV, XLSX, XML e .mpp do MS Project ⇄ MS Project, P6 e XLSX",
    )
    t = cronograma_efetiva(dados)
    importado = st.session_state.get("cronograma_importado")

    st.markdown("#### Importar (software → Sheets)")
    if importado is not None:
        st.success(
            f"Cronograma importado ativo: **{st.session_state.get('cronograma_importado_nome', '')}** "
            "(use 'Restaurar dados de origem' na barra lateral para voltar)."
        )
    arquivo = st.file_uploader(
        "Arquivo de cronograma",
        type=["csv", "xlsx", "xml", "mpp"],
        help=(
            "CSV/XLSX no padrão DB_Cronograma_MSProject ou exportado do MS Project/P6; "
            "XML do MS Project (Arquivo > Salvar Como > XML); "
            "ou o próprio arquivo .mpp do MS Project 2010 ou superior "
            "(Caminho crítico e Responsável não vêm do binário)."
        ),
    )
    if arquivo is not None:
        try:
            novo, avisos = interpretar_upload(arquivo.name, arquivo.getvalue())
        except Exception as exc:
            st.error(f"Não foi possível ler o arquivo: {exc}")
            novo, avisos = None, []
        if novo is not None:
            novo, ausentes = garantir_colunas(novo)
            erros, avisos_valid = validar_cronograma(novo)
            avisos = avisos + avisos_valid
            st.caption(f"Pré-visualização: {len(novo)} tarefa(s), {len(novo.columns)} coluna(s).")
            st.dataframe(tabela_cronograma(novo).head(50), hide_index=True)
            renderizar_alertas(erros, avisos, [])
            st.markdown("##### Para onde vai este cronograma?")
            w1, w2 = st.columns(2)
            nome_projeto = w1.text_input(
                "Nome do projeto a incluir",
                key="imp_nome_projeto",
                placeholder="ex.: Ponte Rio Verde",
                help="Etiqueta gravada na coluna Projeto_Nome das tarefas importadas.",
            )
            contrato_imp = w2.selectbox(
                "Qual contrato deste projeto?",
                ["(sem contrato)"] + ids_contratos(dados) + ["➕ Novo contrato..."],
                key="imp_contrato",
                help=(
                    "Grava a coluna Contrato das tarefas importadas — é ela que o "
                    "Dashboard usa no filtro por contrato."
                ),
            )
            novo_ctr_id = novo_ctr_cliente = ""
            if contrato_imp == "➕ Novo contrato...":
                novo_ctr_id = (
                    w2.text_input(
                        "ID do novo contrato",
                        key="imp_novo_ctr_id",
                        placeholder="ex.: CTR-04",
                    )
                    or ""
                ).strip()
                novo_ctr_cliente = (
                    w2.text_input(
                        "Cliente / Fornecedor (opcional)",
                        key="imp_novo_ctr_cliente",
                    )
                    or ""
                ).strip()
            w3, w4 = st.columns(2)
            frentes_existentes = (
                sorted(set(preparar_cronograma(t)["Frente"].dropna()) - {""})
                if not t.empty
                else []
            )
            frente_imp = w3.selectbox(
                "Frente das tarefas",
                ["(usar a do arquivo)"] + frentes_existentes + ["➕ Nova frente..."],
                key="imp_frente",
                help=(
                    "A Frente é o 1º nível do WBS. Escolher aqui reescreve esse "
                    "primeiro nível das tarefas importadas (ex.: tudo passa a FUND.1, FUND.2...)."
                ),
            )
            novo_frente = ""
            if frente_imp == "➕ Nova frente...":
                novo_frente = (
                    w3.text_input(
                        "Nome da nova frente",
                        key="imp_novo_frente",
                        placeholder="ex.: FUND",
                    )
                    or ""
                ).strip()
            frentes_txt = (
                novo_frente
                if frente_imp == "➕ Nova frente..."
                else (frente_imp if frente_imp != "(usar a do arquivo)" else "")
            )
            responsaveis_existentes = (
                sorted(
                    set(
                        t.get("Responsavel", pd.Series(dtype=str))
                        .fillna("")
                        .astype(str)
                    )
                    - {"", "nan"}
                )
                if not t.empty
                else []
            )
            resp_imp = w4.selectbox(
                "Responsável das tarefas",
                ["(manter os do arquivo)"]
                + responsaveis_existentes
                + ["➕ Novo responsável..."],
                key="imp_responsavel",
                help="Preenche apenas as tarefas importadas que estão sem responsável.",
            )
            novo_resp = ""
            if resp_imp == "➕ Novo responsável...":
                novo_resp = (
                    w4.text_input(
                        "Nome do novo responsável",
                        key="imp_novo_resp",
                        placeholder="ex.: Maria Costa",
                    )
                    or ""
                ).strip()
            resp_txt = (
                novo_resp
                if resp_imp == "➕ Novo responsável..."
                else (resp_imp if resp_imp != "(manter os do arquivo)" else "")
            )
            contrato_escolhido = contrato_imp not in ("(sem contrato)", "")
            if contrato_imp == "➕ Novo contrato...":
                contrato_escolhido = bool(novo_ctr_id)
            vinc_tl = st.checkbox(
                "Vincular as frentes à Cronograma Timeline (Timeline_Visao)",
                key="imp_vinc_tl",
                disabled=not contrato_escolhido,
                help=(
                    "Cada Frente importada ganha uma linha em Timeline_Visao ligada ao "
                    "contrato escolhido (Frente → Projeto): é o vínculo que faz o filtro "
                    "por contrato funcionar e aparece na página Contratos."
                ),
            )
            if vinc_tl and not contrato_escolhido:
                st.warning("Escolha (ou cadastre) um contrato para vincular à Timeline.")
            modo = seg(
                st.radio(
                    "Como aplicar?",
                    [
                        "➕ Acrescentar ao cronograma atual (incluir no grupo)",
                        "🔄 Substituir o cronograma atual",
                    ],
                    horizontal=True,
                    key="imp_modo",
                ),
                "➕ Acrescentar ao cronograma atual (incluir no grupo)",
            )
            acrescentar = modo.startswith("➕")
            etiqueta = (nome_projeto or "").strip() or arquivo.name
            chave_aplicao = f"{arquivo.name}|{etiqueta}|{contrato_imp}:{novo_ctr_id}|{modo}"
            ja_aplicado = st.session_state.get("imp_ultima_chave") == chave_aplicao
            faltando_id = contrato_imp == "➕ Novo contrato..." and not novo_ctr_id
            c1, c2 = st.columns(2)
            aplicar = c1.button(
                f"✅ Acrescentar “{etiqueta}” ao cronograma"
                if acrescentar
                else "✅ Substituir o cronograma atual",
                disabled=bool(erros) or ja_aplicado or faltando_id,
                type="primary",
            )
            if faltando_id:
                c1.caption("Informe o ID do novo contrato para aplicar.")
            if ja_aplicado:
                c1.caption("✅ Já aplicado — troque o modo, o projeto ou envie outro arquivo.")
            csv_padrao = normalizar_para_export(novo).to_csv(
                sep=";", index=False, encoding="utf-8-sig"
            ).encode("utf-8-sig")
            c2.download_button(
                "⬇️ CSV no padrão Sheets (para colar na aba)",
                data=csv_padrao,
                file_name="DB_Cronograma_MSProject.csv",
                mime="text/csv",
                width="stretch",
            )
            if aplicar:
                contrato_txt = (
                    novo_ctr_id
                    if contrato_imp == "➕ Novo contrato..."
                    else (contrato_imp if contrato_imp != "(sem contrato)" else "")
                )
                if contrato_imp == "➕ Novo contrato..." and contrato_txt:
                    incluir_contrato(
                        dados, contrato_txt, {"Cliente_Fornecedor": novo_ctr_cliente}
                    )
                if frentes_txt:
                    restos = novo["WBS"].astype(str).str.split(".").str[1:].str.join(".")
                    novo["WBS"] = [
                        f"{frentes_txt}.{r}" if r and r != "nan" else frentes_txt
                        for r in restos
                    ]
                if resp_txt and "Responsavel" in novo.columns:
                    vazias = novo["Responsavel"].astype(str).str.strip().isin(
                        ["", "nan", "None"]
                    )
                    novo.loc[vazias, "Responsavel"] = resp_txt
                if acrescentar and not t.empty:
                    final, renumeradas = acrescentar_cronograma(
                        t, novo, (nome_projeto or "").strip(), contrato_txt
                    )
                    if renumeradas:
                        st.info(
                            f"{renumeradas} tarefa(s) renumerada(s) para não colidir "
                            "com os IDs já existentes."
                        )
                else:
                    final = novo
                    if (nome_projeto or "").strip():
                        final["Projeto_Nome"] = (nome_projeto or "").strip()
                    if contrato_txt:
                        final["Contrato"] = contrato_txt
                anterior = st.session_state.get("cronograma_importado_nome", "")
                st.session_state["cronograma_importado"] = final
                st.session_state["cronograma_importado_nome"] = (
                    f"{anterior} + {etiqueta}" if acrescentar and anterior else etiqueta
                )
                st.session_state["imp_ultima_chave"] = chave_aplicao
                if vinc_tl and contrato_txt:
                    frentes_novas = sorted(
                        set(novo["WBS"].astype(str).str.split(".").str[0].str.strip())
                        - {"", "nan"}
                    )
                    tl_base = aba_efetiva(dados, "timeline")
                    ja_vinc = (
                        set(tl_base["Frente"].astype(str).str.strip())
                        if not tl_base.empty and "Frente" in tl_base.columns
                        else set()
                    )
                    linhas_tl = pd.DataFrame(
                        [
                            {
                                "ID_Marco": str(i),
                                "Projeto": contrato_txt,
                                "Frente": fr,
                                "Marco_Milestone": "",
                                "Data_Marco": "",
                                "Status": "",
                                "%_Avanço": 0,
                                "Desvio_dias": 0,
                            }
                            for i, fr in enumerate(frentes_novas, start=1)
                            if fr not in ja_vinc
                        ]
                    )
                    if not linhas_tl.empty:
                        acrescentar_aba(dados, "timeline", linhas_tl, contrato=contrato_txt)
                st.rerun()

    st.markdown("#### Importar histogramas (MO / Equip)")
    ativos_h = [
        st.session_state.get(f"{k}_importado_nome", ABAS[k])
        for k in ("hist_mo", "hist_equip")
        if st.session_state.get(f"{k}_importado") is not None
    ]
    if ativos_h:
        st.success(
            f"Histograma importado ativo: **{', '.join(ativos_h)}** "
            "(use 'Restaurar dados de origem' na barra lateral para voltar)."
        )
    arq_hist = st.file_uploader(
        "Arquivo de histograma",
        type=["csv", "xlsx"],
        key="uploader_histograma",
        help=(
            "CSV (; ou ,) ou XLSX com as colunas de DB_Histograma_MO (Funcao_Cargo) "
            "ou DB_Histograma_Equip (Equipamento); a aba é detectada pelo cabeçalho."
        ),
    )
    if arq_hist is not None:
        try:
            df_h, chave_h, avisos_h = interpretar_histograma(arq_hist.name, arq_hist.getvalue())
        except Exception as exc:
            st.error(f"Não foi possível ler o arquivo: {exc}")
            df_h, chave_h, avisos_h = None, "", []
        if df_h is None:
            for msg in avisos_h:
                st.error(msg)
        else:
            contrato_h = st.selectbox(
                "Qual contrato este histograma representa?",
                ["(manter o que vier no arquivo)"] + ids_contratos(dados) + ["➕ Novo contrato..."],
                key="imp_hist_contrato",
                help=(
                    "Grava a coluna Contrato_Vinculado de todas as linhas — "
                    "os filtros do histograma usam isso."
                ),
            )
            novo_h_id = novo_h_fornecedor = ""
            if contrato_h == "➕ Novo contrato...":
                novo_h_id = (
                    st.text_input(
                        "ID do novo contrato",
                        key="imp_hist_novo_id",
                        placeholder="ex.: CTR-04",
                    )
                    or ""
                ).strip()
                novo_h_fornecedor = (
                    st.text_input(
                        "Fornecedor (opcional)", key="imp_hist_novo_fornecedor"
                    )
                    or ""
                ).strip()
            contrato_h_txt = (
                novo_h_id
                if contrato_h == "➕ Novo contrato..."
                else (contrato_h if contrato_h != "(manter o que vier no arquivo)" else "")
            )
            if contrato_h_txt:
                df_h[COLUNA_CONTRATO.get(chave_h, "Contrato_Vinculado")] = contrato_h_txt
            erros_h, avisos_val = validar_aba(chave_h, df_h)
            avisos_h = avisos_h + avisos_val
            nome_h = ABAS[chave_h]
            st.caption(
                f"Reconhecido como **{nome_h}**: {len(df_h)} linha(s), {len(df_h.columns)} coluna(s)."
            )
            st.dataframe(df_h.head(50), hide_index=True)
            renderizar_alertas(erros_h, avisos_h, [])
            c1, c2 = st.columns(2)
            aplicar_h = c1.button(
                "✅ Aplicar ao app (substitui o histograma)",
                disabled=bool(erros_h)
                or (contrato_h == "➕ Novo contrato..." and not novo_h_id),
                type="primary",
                key="btn_aplicar_histograma",
            )
            csv_h = df_h.to_csv(sep=";", index=False, encoding="utf-8-sig").encode("utf-8-sig")
            c2.download_button(
                "⬇️ CSV no padrão Sheets (para colar na aba)",
                data=csv_h,
                file_name=f"{nome_h}.csv",
                mime="text/csv",
                width="stretch",
                key="dl_histograma",
            )
            if aplicar_h:
                if contrato_h == "➕ Novo contrato..." and novo_h_id:
                    incluir_contrato(
                        dados, novo_h_id, {"Cliente_Fornecedor": novo_h_fornecedor}
                    )
                st.session_state[f"{chave_h}_importado"] = df_h
                st.session_state[f"{chave_h}_importado_nome"] = arq_hist.name
                st.rerun()

    st.markdown("#### Exportar (Sheets → software)")
    origem = "cronograma importado" if importado is not None else "fonte de dados atual"
    nome_exp = st.text_input(
        "Nome do projeto (para nomear os arquivos exportados)",
        key="exp_nome_projeto",
        placeholder="ex.: Ponte Rio Verde — vira Ponte_Rio_Verde_msproject.csv",
    )
    slug = re.sub(r"[^0-9A-Za-zÀ-ÿ_-]+", "_", (nome_exp or "").strip()).strip("_")
    st.caption(f"Origem do export: {origem} • datas DD/MM/AAAA • separador do Sheets: ;")
    if t.empty:
        st.info("[FALTANTE] Sem cronograma para exportar.")
        return
    try:
        csv_msp, msg_msp = exportar_via_conversor(t, "sheets_para_msproject")
        csv_p6, msg_p6 = exportar_via_conversor(t, "sheets_para_p6")
        xlsx = exportar_xlsx(t)
    except Exception as exc:
        st.error(f"Falha ao gerar arquivos de exportação: {exc}")
        return
    c1, c2, c3 = st.columns(3)
    c1.download_button(
        "⬇️ MS Project CSV (;)",
        data=csv_msp,
        file_name=f"{slug}_msproject.csv" if slug else "cronograma_msproject.csv",
        mime="text/csv",
        width="stretch",
        help=msg_msp,
    )
    c2.download_button(
        "⬇️ P6 CSV (,)",
        data=csv_p6,
        file_name=f"{slug}_p6.csv" if slug else "cronograma_p6.csv",
        mime="text/csv",
        width="stretch",
        help=msg_p6,
    )
    c3.download_button(
        "⬇️ XLSX",
        data=xlsx,
        file_name=f"{slug}_cronograma.xlsx" if slug else "cronograma.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    st.caption(f"Conversor: {msg_msp} • {msg_p6}")

    with st.expander("Passo a passo de importação no software"):
        st.markdown(
            "**MS Project:** Arquivo > Abrir > selecione o CSV > Assistente de Importação > "
            'Novo mapa > marque "O projeto contém cabeçalho" > mapeie os campos > Concluir.\n\n'
            "**Primavera P6:** File > Import > Spreadsheet (ou XER/XML) > escolha o arquivo gerado.\n\n"
            "**Google Sheets:** Arquivo > Importar > enviar o CSV baixado > separador `;` > "
            "substituir a aba DB_Cronograma_MSProject."
        )
    with st.expander("Formatos aceitos na importação"):
        st.markdown(
            "- **CSV/XLSX no padrão Sheets** (cabeçalho ID_Tarefa...): importa direto, preserva Status.\n"
            "- **CSV do MS Project ou P6** (cabeçalho em inglês): convertido pelo "
            "`conversor-cronograma.py`; Status é derivado de %_Concluido.\n"
            "- **XML do MS Project**: lido diretamente (Task/Name/Start/Finish/Duration/Notes).\n"
            "- **`.mpp` do MS Project 2010+ (MPP14)**: lido pelo `leitor_mpp.py` "
            "(Caminho crítico e Responsável não vêm do binário).\n"
            "- `.xer` não é suportado: exporte antes para CSV ou XML no software."
        )


# ------------------------------------------------------------- 7 Editar Dados
def pagina_editar(dados: dict) -> None:
    cabecalho(
        "✏️",
        "Editar Dados",
        "Edite as 6 abas na tabela e grave na fonte ativa (CSV local, API ou Sheets)",
    )
    ctx = st.session_state.get("ctx_gravacao") or {}
    fonte = str(ctx.get("fonte", "CSV local (exemplos)"))

    aviso_grav = st.session_state.pop("aviso_gravacao", None)
    if aviso_grav:
        st.success(aviso_grav)

    escolha = seg(
        st.selectbox(
            "Aba para editar", list(ABAS), format_func=lambda k: ABAS[k], key="sel_aba_edicao"
        ),
        "cronograma",
    )
    nome_aba = ABAS[escolha]

    # O editor parte do que o app mostra (importação ativa, se houver); o
    # `original` da fonte é quem entra no diff da API.
    mostrado = _sem_linhas_vazias(
        aba_efetiva(dados, escolha).apply(lambda s: s.map(_valor_planilha))
    )
    if mostrado.empty:
        st.info(
            f"[FALTANTE] A aba {nome_aba} está vazia na fonte atual — "
            "use o editor abaixo com as colunas do schema para criar as primeiras linhas."
        )
        mostrado = pd.DataFrame(columns=SCHEMAS.get(escolha, []))

    st.caption(
        f"Fonte de gravação: **{fonte}** • duplo clique na célula para editar • "
        "＋ adiciona linha e ✕ na lateral remove. Linhas novas ficam vazias até ser preenchidas."
    )
    editado = st.data_editor(
        mostrado,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"editor_{escolha}",
    )
    editado = _sem_linhas_vazias(editado.apply(lambda s: s.map(_valor_planilha)))

    erros, avisos = validar_aba(escolha, editado)
    renderizar_alertas(erros, avisos, [])

    # Pré-condições da fonte de gravação.
    contexto_ok = True
    if fonte == "Google Sheets (API)" and not (ctx.get("api_url") and ctx.get("api_key")):
        contexto_ok = False
        st.warning("Informe a URL da API e a API Key na barra lateral para gravar no Sheets.")
    elif fonte == "Google Sheets" and not ctx.get("id_planilha"):
        contexto_ok = False
        st.warning("Informe o ID/URL da planilha na barra lateral para gravar no Sheets.")

    confirmado = st.checkbox(
        "Confirmo a gravação na fonte ativa (a alteração é permanente)",
        key=f"confirmar_gravacao_{escolha}",
    )
    if st.button(
        "💾 Gravar alterações",
        type="primary",
        disabled=bool(erros) or not confirmado or not contexto_ok,
        key=f"botao_gravar_{escolha}",
    ):
        try:
            msg = salvar_aba(
                escolha,
                editado,
                fonte,
                id_planilha=str(ctx.get("id_planilha", "")),
                api_url=str(ctx.get("api_url", "")),
                api_key=str(ctx.get("api_key", "")),
                original=dados.get(escolha, pd.DataFrame()),
            )
        except Exception as exc:
            st.error(f"Falha ao gravar: {exc}")
        else:
            # Limpa o estado do editor para que a próxima run traga a fonte já gravada.
            st.session_state.pop(f"editor_{escolha}", None)
            st.session_state["aviso_gravacao"] = msg
            st.rerun()


# ------------------------------------------------------- 8 Relatório com IA
def config_ia() -> dict:
    """Credenciais [ia] do secrets.toml — API compatível com OpenAI."""
    cfg = {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o-mini",
        "api_key": "",
    }
    try:
        if "ia" in st.secrets:
            bloco = st.secrets["ia"]
            for k in cfg:
                if bloco.get(k):
                    cfg[k] = str(bloco[k]).strip()
    except Exception:
        pass  # sem secrets.toml (ex.: só CSV local) — usa o padrão
    return cfg


def resumo_dados(dados: dict) -> dict:
    """Indicadores consolidados que alimentam o relatório (local e IA)."""
    t = preparar_cronograma(cronograma_efetiva(dados))
    if t.empty:
        return {"total": 0}
    mapa = mapa_frente_contrato(
        dados.get("timeline", pd.DataFrame()), dados.get("contratos", pd.DataFrame())
    )
    t = vincular_contrato(t, mapa)
    dur = float(t["_dur"].sum())
    avanco = float((t["_dur"] * t["_pct"]).sum() / dur) if dur else 0.0
    atrasadas = t[
        t["Fim_Planejado_dt"].notna()
        & (t["Fim_Planejado_dt"] < hoje())
        & (t["_pct"] < 100)
    ].copy()
    atrasadas["Dias_Atraso"] = (hoje() - atrasadas["Fim_Planejado_dt"]).dt.days
    atrasadas = atrasadas.sort_values("Dias_Atraso", ascending=False)
    rh = aba_efetiva(dados, "rh")
    ativos = (
        rh[rh["Status_Ativo_Inativo"].astype(str).str.strip() == "Ativo"]
        if not rh.empty and "Status_Ativo_Inativo" in rh.columns
        else rh
    )
    custo = (
        float(parse_num(ativos["Custo_Mes"]).sum())
        if not ativos.empty and "Custo_Mes" in ativos.columns
        else 0.0
    )
    ct = aba_efetiva(dados, "contratos")
    contratos = []
    # avanço físico por contrato: média ponderada de TODAS as tarefas do
    # contrato (agrupa por _contrato, não o máximo das frentes).
    fisico_contrato: dict[str, float] = {}
    if "_contrato" in t.columns:
        for cid, grupo in t.groupby(t["_contrato"].astype(str).str.strip()):
            dur = float(grupo["_dur"].sum())
            if cid and dur:
                fisico_contrato[cid] = float((grupo["_dur"] * grupo["_pct"]).sum() / dur)
    for x in (ct.to_dict("records") if not ct.empty else []):
        def num(col: str) -> float:
            v = parse_num(pd.Series([x.get(col, 0)])).iloc[0]
            return float(v) if pd.notna(v) else 0.0

        cid = str(x.get("ID_Contrato", "")).strip()
        total, medido = num("Valor_Total"), num("Valor_Medido_Acumulado")
        contratos.append(
            {
                "id": cid,
                "cliente": str(x.get("Cliente_Fornecedor", "")).strip(),
                "objeto": str(x.get("Objeto_Escopo", "")).strip(),
                "total": total,
                "medido": medido,
                "saldo": num("Saldo"),
                "fim": fmt_data_opcional(parse_datas(pd.Series([x.get("Data_Fim", "")])).iloc[0]),
                "status": str(x.get("Status", "")).strip(),
                "financeiro": (medido / total * 100) if total else 0.0,
                "fisico": fisico_contrato.get(cid, 0.0),
            }
        )
    # frentes paradas: avanço ~0 e compromisso (marco/fim) em até 45 dias
    tl = aba_efetiva(dados, "timeline")
    paradas = []
    for frente, grupo in t.groupby("Frente", dropna=False):
        dur = float(grupo["_dur"].sum())
        av = float((grupo["_dur"] * grupo["_pct"]).sum() / dur) if dur else 0.0
        if av >= 5:
            continue
        fr = str(frente).strip() if pd.notna(frente) else ""
        prazo = pd.NaT
        if not tl.empty and "Frente" in tl.columns and "Data_Marco" in tl.columns:
            m = tl[tl["Frente"].astype(str).str.strip() == fr]
            if not m.empty:
                datas = parse_datas(m["Data_Marco"]).dropna()
                futuras = datas[datas >= hoje()]
                prazo = (
                    futuras.min()
                    if not futuras.empty
                    else (datas.max() if not datas.empty else pd.NaT)
                )
        if pd.isna(prazo):
            fims = grupo["Fim_Planejado_dt"].dropna()
            fims_futuros = fims[fims >= hoje()]
            prazo = fims_futuros.min() if not fims_futuros.empty else pd.NaT
        dias = int((prazo - hoje()).days) if pd.notna(prazo) else None
        if dias is not None and dias <= 45:
            paradas.append(
                {"frente": fr or "(sem frente)", "prazo": fmt_data_opcional(prazo), "dias": dias}
            )
    # marcos atrasados da Timeline_Visao
    marcos_atrasados = []
    if not tl.empty and "Data_Marco" in tl.columns:
        tl2 = tl.copy()
        tl2["_data"] = parse_datas(tl2["Data_Marco"])
        tl2["_av"] = parse_num(tl2.get("%_Avanço", pd.Series(0, index=tl2.index))).fillna(0)
        status_tl = (
            tl2["Status"].astype(str).str.lower()
            if "Status" in tl2.columns
            else pd.Series("", index=tl2.index)
        )
        concluidos = status_tl.str.contains("conclu") | (tl2["_av"] >= 100)
        for x in tl2[tl2["_data"].notna() & (tl2["_data"] < hoje()) & ~concluidos].sort_values("_data").to_dict("records"):
            marcos_atrasados.append(
                {
                    "frente": str(x.get("Frente", "")).strip(),
                    "contrato": str(x.get("Projeto", "")).strip(),
                    "marco": str(x.get("Marco_Milestone", "")).strip(),
                    "data": str(x.get("Data_Marco", "")).strip(),
                    "dias": int((hoje() - x["_data"]).days),
                }
            )
    return {
        "total": len(t),
        "avanco": avanco,
        "criticas": int(atrasadas["_critico"].sum()),
        "atrasadas": atrasadas,
        "frentes": sorted(set(t["Frente"].dropna()) - {""}),
        "ativos": len(ativos),
        "custo": custo,
        "contratos": contratos,
        "frentes_paradas": paradas,
        "marcos_atrasados": marcos_atrasados,
    }


def bloco_resumo(r: dict) -> str:
    """Texto-base com os números do projeto (o que a IA lê)."""
    linhas = [
        f"Data-base: {hoje():%d/%m/%Y}",
        f"Total de tarefas: {r['total']}",
        f"Avanço físico ponderado: {r['avanco']:.0f}%",
        f"Frentes: {', '.join(r['frentes']) or '—'}",
        f"Tarefas atrasadas: {len(r['atrasadas'])} (críticas entre elas: {r['criticas']})",
        f"Efetivo ativo: {r['ativos']} • custo mensal estimado: {brl(r['custo'])}",
    ]
    if r.get("marcos_atrasados"):
        linhas.append(
            f"Marcos atrasados na Timeline: {len(r['marcos_atrasados'])} — "
            + "; ".join(
                f"{m['marco']} ({m['frente']}, {m['dias']} dia(s), {m['contrato']})"
                for m in r["marcos_atrasados"][:8]
            )
        )
    if r.get("frentes_paradas"):
        linhas.append(
            "Frentes sem avanço com compromisso em até 45 dias: "
            + "; ".join(
                f"{p['frente']} (prazo {p['prazo']}, {p['dias']} dia(s))"
                for p in r["frentes_paradas"]
            )
        )
    if not r["atrasadas"].empty:
        linhas.append("Atrasadas (mais antigas primeiro):")
        for x in r["atrasadas"].head(12).to_dict("records"):
            linhas.append(
                f"- {x.get('ID_Tarefa', '')} — {x.get('Nome_Tarefa', '')} | Frente {x.get('Frente', '')} "
                f"| contrato {x.get('_contrato', '') or '—'} | fim {x.get('Fim_Planejado', '')} "
                f"| {int(x['Dias_Atraso'])} dia(s) de atraso | {int(x.get('_pct', 0))}% concluído "
                f"| crítico: {'SIM' if x.get('_critico') else 'NAO'} | responsável {x.get('Responsavel', '') or '—'}"
            )
    if r["contratos"]:
        linhas.append("Contratos (físico = avanço do cronograma do contrato; financeiro = medido/total):")
        for c in r["contratos"]:
            linhas.append(
                f"- {c['id']} | {c['cliente']} | {c['objeto'][:80]} | total {brl(c['total'])} "
                f"| medido {brl(c['medido'])} ({c['financeiro']:.0f}% financeiro) | físico {c['fisico']:.0f}% "
                f"| saldo {brl(c['saldo'])} | fim {c['fim']} | {c['status']}"
            )
    return "\n".join(linhas)


def gerar_relatorio_ia(base: str, estilo: str, instrucoes: str, cfg: dict) -> str:
    """Chama API compatível com OpenAI e devolve o relatório em markdown."""
    sistema = (
        "Você é um analista de projetos de engenharia civil. Escreva em português "
        "(pt-BR), em markdown, usando APENAS os números fornecidos — nunca invente "
        "dados; se faltar algo essencial, marque [FALTANTE]. Datas em DD/MM/AAAA, "
        "moeda R$, percentuais 0-100. Estrutura: '## Panorama geral', '## Atrasos e "
        "riscos' (inclua marcos atrasados e frentes paradas com prazo próximo), "
        "'## Contratos' (quando houver, comente o desvio físico × financeiro), "
        "'## Equipe', '## Próximos passos' (3 a 5 recomendações)."
    )
    if estilo == "Técnico":
        sistema += " Tom técnico: cite frentes, WBS e caminho crítico."
    elif estilo == "Ata de reunião":
        sistema = (
            "Você é o secretário de uma reunião de projeto de engenharia civil. Escreva "
            "em markdown, em pt-BR, usando APENAS os dados fornecidos (marque [FALTANTE] "
            "se faltar). Estrutura: '## Presentes' (escreva 'equipe do projeto' — não "
            "invente nomes), '## Andamento', '## Pendências e riscos', "
            "'## Decisões e próximos passos'."
        )
    if instrucoes:
        sistema += f" Observações do usuário: {instrucoes}"
    resp = requests.post(
        cfg["base_url"].rstrip("/") + "/chat/completions",
        headers={"Authorization": f"Bearer {cfg['api_key']}"},
        json={
            "model": cfg["model"],
            "messages": [
                {"role": "system", "content": sistema},
                {"role": "user", "content": base},
            ],
            "temperature": 0.3,
        },
        timeout=120,
    )
    resp.raise_for_status()
    return str(resp.json()["choices"][0]["message"]["content"]).strip()


def relatorio_local(r: dict, estilo: str) -> str:
    """Relatório determinístico sem IA: mesmos números, narrativa de template."""
    linhas = [
        f"# Relatório do Projeto — {hoje():%d/%m/%Y}",
        "",
        f"*Estilo: {estilo} • gerado localmente (sem IA)*",
        "",
        "## Panorama geral",
        "",
        f"- **{r['total']}** tarefas no cronograma; avanço físico ponderado de **{r['avanco']:.0f} %**.",
        f"- **{len(r['atrasadas'])}** tarefa(s) atrasada(s), das quais **{r['criticas']}** no caminho crítico.",
        f"- Efetivo ativo: **{r['ativos']}** pessoa(s); custo mensal estimado: **{brl(r['custo'])}**.",
        f"- Frentes: {', '.join(r['frentes']) or '—'}.",
        "",
        "## Atrasos e riscos",
        "",
    ]
    if r["atrasadas"].empty:
        linhas.append("- ✅ Nenhuma tarefa atrasada na data-base.")
    else:
        for x in r["atrasadas"].head(15).to_dict("records"):
            crit = " ⚠️ crítico" if x.get("_critico") else ""
            linhas.append(
                f"- **{x.get('ID_Tarefa', '')} — {x.get('Nome_Tarefa', '')}**{crit}: "
                f"{int(x['Dias_Atraso'])} dia(s) de atraso (fim planejado {x.get('Fim_Planejado', '')}), "
                f"{int(x.get('_pct', 0))}% concluído, frente {x.get('Frente', '') or '—'}."
            )
    if r.get("marcos_atrasados"):
        linhas.append("")
        linhas.append("### Marcos atrasados (Timeline_Visao)")
        linhas.append("")
        for m in r["marcos_atrasados"]:
            linhas.append(
                f"- **{m['marco'] or '—'}** (frente {m['frente'] or '—'}, contrato {m['contrato'] or '—'}): "
                f"venceu em {m['data']} — {m['dias']} dia(s) de atraso."
            )
    if r.get("frentes_paradas"):
        linhas.append("")
        linhas.append("### Frentes sem avanço com compromisso próximo (≤ 45 dias)")
        linhas.append("")
        for p in r["frentes_paradas"]:
            linhas.append(
                f"- **{p['frente']}**: 0–5 % de avanço e prazo em {p['prazo']} "
                f"({p['dias']} dia(s)) — mobilizar equipe ou revisar o plano."
            )
    linhas += ["", "## Contratos (físico × financeiro)", ""]
    if not r["contratos"]:
        linhas.append("- [FALTANTE] Sem dados em DB_Contratos.")
    for c in r["contratos"]:
        desvio = c["fisico"] - c["financeiro"]
        alerta_desvio = (
            " ⚠️ desvio físico-financeiro acima de 10 p.p."
            if abs(desvio) > 10
            else ""
        )
        linhas.append(
            f"- **{c['id']} — {c['cliente']}**: total {brl(c['total'])}, medido {brl(c['medido'])} "
            f"({c['financeiro']:.0f} % financeiro) × {c['fisico']:.0f} % físico"
            f"{alerta_desvio}, saldo {brl(c['saldo'])}, fim em {c['fim']}, status {c['status'] or '—'}."
        )
    linhas += [
        "",
        "## Próximos passos",
        "",
        "- Priorizar as tarefas críticas atrasadas — no caminho crítico, cada dia empurra o fim do projeto.",
        "- Revisar as alocações das frentes com atraso na próxima reunião de acompanhamento.",
        "- Gere este relatório **com IA** para recomendações personalizadas a partir destes números.",
    ]
    return "\n".join(linhas)


def pagina_relatorio(dados: dict) -> None:
    cabecalho(
        "📝",
        "Relatório do Projeto",
        "Texto executivo gerado por IA — com modo local para nunca ficar sem relatório",
    )
    r = resumo_dados(dados)
    if r["total"] == 0:
        st.info("[FALTANTE] Sem cronograma para gerar relatório.")
        return
    cfg = config_ia()
    estilo = seg(
        st.radio(
            "Estilo", ["Executivo", "Técnico", "Ata de reunião"], horizontal=True, key="rel_estilo"
        ),
        "Executivo",
    )
    instrucoes = (
        st.text_area(
            "Observações para o relatório (opcional)",
            key="rel_instr",
            placeholder="ex.: destacar o atraso da impermeabilização e o risco da vistoria",
            height=80,
        )
        or ""
    )
    if not cfg["api_key"]:
        cfg["api_key"] = (
            st.text_input(
                "API Key da IA (ou configure [ia] no secrets.toml)",
                type="password",
                key="rel_key",
                help=(
                    "Funciona com OpenAI, Groq, DeepSeek, OpenRouter e Gemini "
                    "(compatível com OpenAI) — a chave fica só nesta sessão."
                ),
            )
            or ""
        ).strip()
    c1, c2 = st.columns(2)
    gerar = c1.button(
        "✨ Gerar com IA", type="primary", disabled=not cfg["api_key"], width="stretch"
    )
    local = c2.button("📄 Gerar resumo local (sem IA)", width="stretch")
    if gerar:
        with st.spinner("Escrevendo o relatório..."):
            try:
                st.session_state["relatorio_md"] = gerar_relatorio_ia(
                    bloco_resumo(r), estilo, instrucoes, cfg
                )
                st.session_state["relatorio_fonte"] = f"IA · {cfg['model']}"
            except Exception as exc:
                st.error(
                    f"Falha na IA: {exc} — confira base_url/model/api_key no secrets.toml "
                    "ou use o resumo local."
                )
    if local:
        st.session_state["relatorio_md"] = relatorio_local(r, estilo)
        st.session_state["relatorio_fonte"] = "local (sem IA)"
    md = st.session_state.get("relatorio_md")
    if md:
        st.success(f"Relatório gerado ({st.session_state.get('relatorio_fonte', '')})")
        st.markdown(md)
        st.download_button(
            "⬇️ Baixar relatório (.md)",
            data=md,
            file_name=f"relatorio_projeto_{hoje():%d%m%Y}.md",
            mime="text/markdown",
            key="dl_relatorio",
        )
    with st.expander("⚙️ Configurar a IA (secrets.toml)"):
        st.markdown(
            "Funciona com qualquer API compatível com OpenAI — OpenAI, Groq (gratuito), "
            "DeepSeek, OpenRouter e Gemini. Crie a chave no provedor e acrescente em "
            "`.streamlit/secrets.toml` (no Community Cloud: *App settings → Secrets*):\n\n"
            "```toml\n"
            "[ia]\n"
            'base_url = "https://api.openai.com/v1"\n'
            'model = "gpt-4o-mini"\n'
            'api_key = "sk-..."\n'
            "```\n\n"
            "Exemplo Groq gratuito: `base_url = \"https://api.groq.com/openai/v1\"` e "
            '`model = "llama-3.3-70b-versatile"`. Sem chave, o **resumo local** sai sempre.'
        )


# ------------------------------------------------------------------ execução
MAPEIO = {
    "📊 Dashboard": pagina_dashboard,
    "📅 Cronograma Timeline": pagina_timeline,
    "👷 Histograma MO + Equip": pagina_histograma,
    "👥 Recursos Humanos": pagina_rh,
    "📑 Contratos": pagina_contratos,
    "📝 Relatório (IA)": pagina_relatorio,
    "🔁 Importar / Exportar": pagina_importar_exportar,
    "✏️ Editar Dados": pagina_editar,
}


# Preencha depois de gerar o .apk no PWABuilder (ex.:
# "https://github.com/SEU-USER/SEU-REPO/releases/download/v1/app.apk").
# Com valor, o botão 📲 de instalação ganha também o link "Baixar .apk".
URL_APK = ""


def injetar_pwa() -> None:
    """Manifest, meta do iOS, service worker e botão de instalação do PWA.

    Precisa rodar em todo run: o iframe é destruído a cada rerun do Streamlit e
    o `<head>` da página é montado pelo React (o config.toml não serve manifest).
    O botão flutuante 📲 é injetado no documento principal (fora do React, então
    sobrevive aos reruns): detecta Android/iOS/Windows e, quando o navegador
    oferece (`beforeinstallprompt`), dispara a instalação real num clique.
    """
    url_apk_js = URL_APK.replace("\\", "\\\\").replace('"', '\\"')
    st.components.v1.html(
        """
<script>
(() => {
  const w = window.parent;
  const d = w.document;
  if (!d.head) return;

  // ---------- head: manifest, iOS, tema e service worker ----------
  if (!d.getElementById('pwa-manifest')) {
    const add = (tag, attrs) => {
      const el = d.createElement(tag);
      Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
      d.head.appendChild(el);
    };
    const mark = d.createElement('meta');
    mark.id = 'pwa-manifest';
    mark.name = 'pwa-instalado';
    mark.content = '1';
    d.head.appendChild(mark);
    add('link', {rel: 'manifest', href: '/app/static/manifest.json'});
    add('link', {rel: 'apple-touch-icon', href: '/app/static/icon-192.png'});
    add('meta', {name: 'theme-color', content: '#0B1220'});
    add('meta', {name: 'apple-mobile-web-app-capable', content: 'yes'});
    add('meta', {name: 'apple-mobile-web-app-status-bar-style', content: 'black-translucent'});
    add('meta', {name: 'mobile-web-app-capable', content: 'yes'});
    if ('serviceWorker' in w.navigator) {
      w.navigator.serviceWorker.register('/app/static/sw.js').catch(() => {});
    }
  }

  // ---------- barra lateral (parte inferior): instalar + compartilhar ----------
  const URL_APK = "__URL_APK__";
  w.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    w.__pwaPrompt = e;
    const b = d.getElementById('pwa-instalar');
    if (b) b.style.display = 'inline-flex';
  });
  const ua = w.navigator.userAgent || '';
  const android = /Android/i.test(ua);
  const ios = /iPhone|iPad|iPod/i.test(ua);
  const win = /Windows/i.test(ua);
  const passos = ios
    ? 'No <b>Safari</b>: toque no ícone de <b>Compartilhar</b> (quadrado com seta) → <b>Adicionar à Tela de Início</b> → <b>Adicionar</b>.'
    : android
      ? 'No <b>Chrome</b>: toque nos <b>⋮ três pontos</b> → <b>Instalar app</b> (ou <b>Adicionar à tela inicial</b>).'
      : win
        ? 'No <b>Chrome</b> ou <b>Edge</b>: menu <b>⋮</b> → <b>Instalar app</b> — o app entra no Menu Iniciar como aplicativo de verdade.'
        : 'No navegador: menu <b>⋮</b> → <b>Instalar app</b> ou <b>Criar atalho</b>.';
  const estilo = d.createElement('style');
  estilo.textContent = `
    #pwa-bloco { margin: 14px 4px 6px; padding: 10px; border-radius: 14px;
      border: 1px solid var(--borda, rgba(255,255,255,.12));
      background: color-mix(in srgb, var(--superficie, #141C2C) 65%, transparent); }
    #pwa-botoes { display: flex; gap: 8px; }
    #pwa-btn, #pwa-share { flex: 1; display: inline-flex; align-items: center;
      justify-content: center; gap: 6px; cursor: pointer;
      border: 1px solid var(--acento, #818CF8); border-radius: 10px;
      background: transparent; color: var(--acento, #818CF8);
      font: 600 12px/1 Inter, system-ui, sans-serif; padding: 8px 6px;
      transition: background .18s ease, transform .18s ease; }
    #pwa-btn:hover, #pwa-share:hover {
      background: color-mix(in srgb, var(--acento, #818CF8) 12%, transparent);
      transform: translateY(-1px); }
    #pwa-painel { display: none; margin-top: 10px;
      font: 400 12px/1.5 Inter, system-ui, sans-serif; color: var(--texto, #E6EDF6); }
    #pwa-painel h4 { margin: 0 0 6px; font-size: 13px; }
    #pwa-painel p { margin: 0 0 10px; color: var(--suave, #93A3B8); }
    #pwa-painel a { color: var(--acento, #818CF8); font-weight: 600; }
    #pwa-instalar { display: none; align-items: center; gap: 6px; cursor: pointer;
      border: 0; border-radius: 10px; padding: 8px 12px; margin-right: 10px;
      background: var(--acento, #818CF8); color: #fff;
      font: 700 12px/1 Inter, system-ui, sans-serif; }
  `;
  d.head.appendChild(estilo);
  const compartilhar = async () => {
    const url = w.location.href;
    const dados = {title: d.title || 'Gestão de Projetos', url};
    if (w.navigator.share) {
      try { await w.navigator.share(dados); return; } catch (err) { /* cancelou: copia */ }
    }
    let ok = false;
    try { await w.navigator.clipboard.writeText(url); ok = true; } catch (err) {}
    if (!ok) {
      const ta = d.createElement('textarea');
      ta.value = url;
      d.body.appendChild(ta);
      ta.select();
      try { ok = d.execCommand('copy'); } catch (err) {}
      ta.remove();
    }
    const rot = d.getElementById('pwa-share-txt');
    if (rot) {
      rot.textContent = ok ? 'Copiado ✓' : 'Copie da URL';
      setTimeout(() => { rot.textContent = 'Compartilhar'; }, 2500);
    }
  };
  const injetar = () => {
    if (d.getElementById('pwa-bloco')) return true;
    const alvo = d.querySelector('[data-testid="stSidebar"] [data-testid="stSidebarContent"]')
      || d.querySelector('[data-testid="stSidebar"]');
    if (!alvo) return false;
    const bloco = d.createElement('div');
    bloco.id = 'pwa-bloco';
    bloco.innerHTML =
      '<div id="pwa-botoes">' +
        '<button id="pwa-btn" aria-label="Instalar o app neste aparelho">📲 <span>Instalar</span></button>' +
        '<button id="pwa-share" aria-label="Compartilhar o link do app">🔗 <span id="pwa-share-txt">Compartilhar</span></button>' +
      '</div>' +
      '<div id="pwa-painel"><h4>📲 Instalar o app</h4><p>' + passos + '</p>' +
        '<button id="pwa-instalar">Instalar agora</button>' +
        (URL_APK ? '<a href="' + URL_APK + '" target="_blank" rel="noopener">⬇ Baixar .apk</a>' : '') +
      '</div>';
    alvo.appendChild(bloco);
    d.getElementById('pwa-btn').addEventListener('click', () => {
      const p = d.getElementById('pwa-painel');
      p.style.display = p.style.display === 'block' ? 'none' : 'block';
    });
    d.getElementById('pwa-share').addEventListener('click', compartilhar);
    d.getElementById('pwa-instalar').addEventListener('click', () => {
      const p = w.__pwaPrompt;
      if (!p) return;
      p.prompt();
      p.userChoice.then(() => { w.__pwaPrompt = null; }).catch(() => {});
    });
    if (w.__pwaPrompt) d.getElementById('pwa-instalar').style.display = 'inline-flex';
    return true;
  };
  let tentativas = 0;
  const espera = setInterval(() => {
    if (injetar() || ++tentativas > 120) clearInterval(espera);
  }, 500);
})();
</script>
""".replace("__URL_APK__", url_apk_js),
        height=0,
    )


def main() -> None:
    # Tema claro/escuro (ver TEMAS): o config.toml do Streamlit só vale na
    # abertura da sessão, então o toggle é por CSS próprio injetado a cada
    # execução. O valor fica em session_state["tema"] (key do radio).
    # Fica atrás do ícone 🎨 no topo da árvore lateral: clicar desoculta as paletas.
    with st.sidebar:
        with st.popover("🎨 Tema", help="Clique para escolher a paleta"):
            seg(st.radio("Tema", list(TEMAS), key="tema"), "Escuro")
    st.markdown(css_moderno(tema_ativo()), unsafe_allow_html=True)
    st.sidebar.markdown(
        '<div class="marca"><div class="marca-ico">📋</div><div>'
        '<div class="marca-nome">Gestão de Projetos</div>'
        '<div class="marca-sub">Painel executivo</div></div></div>',
        unsafe_allow_html=True,
    )
    fonte = seg(
        st.sidebar.radio(
            "Fonte de dados",
            ["CSV local (exemplos)", "Google Sheets (API)", "Google Sheets"],
        ),
        "CSV local (exemplos)",
    )
    id_planilha = ""
    api_url = ""
    api_key = ""
    if fonte == "Google Sheets (API)":
        api_url = seg(
            st.sidebar.text_input(
                "URL da API (Web App)",
                placeholder="https://script.google.com/macros/s/.../exec",
            ),
            "",
        ).strip()
        api_key = seg(st.sidebar.text_input("API Key", type="password"), "").strip()
        if api_url and api_key:
            if sheets_api.ping(api_url, api_key):
                st.sidebar.success("✅ API conectada.")
            else:
                st.sidebar.error(
                    "Falha ao conectar. Confira a URL (/exec), a chave e se o app foi publicado."
                )
    elif fonte == "Google Sheets":
        bruto = seg(
            st.sidebar.text_input(
                "ID ou URL da planilha",
                placeholder="https://docs.google.com/spreadsheets/d/...",
            ),
            "",
        )
        id_planilha = _extrair_id_planilha(bruto)
        st.sidebar.caption(
            "Credenciais em .streamlit/secrets.toml → [gcp_service_account]. Ver README."
        )
    pagina = seg(st.sidebar.radio("Páginas", PAGINAS, key="paginas"), PAGINAS[0])
    template_dash = seg(
        st.sidebar.radio(
            "Template do Dashboard",
            ["Executivo", "Analítico"],
            horizontal=True,
            key="template_dash",
        ),
        "Executivo",
    )

    importados = [k for k in ABAS if st.session_state.get(f"{k}_importado") is not None]
    if importados:
        st.sidebar.warning(
            "Importado ativo: "
            + ", ".join(
                str(st.session_state.get(f"{k}_importado_nome", ABAS[k])) for k in importados
            )
        )
        if st.sidebar.button("↩ Restaurar dados de origem"):
            for k in importados:
                st.session_state.pop(f"{k}_importado", None)
                st.session_state.pop(f"{k}_importado_nome", None)
            st.rerun()

    # Contexto de gravação da página "Editar Dados" (fonte escolhida acima).
    st.session_state["ctx_gravacao"] = {
        "fonte": fonte,
        "id_planilha": id_planilha,
        "api_url": api_url,
        "api_key": api_key,
    }
    dados, origem, avisos = carregar_dados(fonte, id_planilha, api_url, api_key)
    for msg in avisos:
        st.warning(msg)
    st.sidebar.caption("Origem: " + ", ".join(sorted(set(origem.values()))))

    # 💾 Salvar tudo na fonte — grava em uma tacada as abas alteradas nesta
    # sessão (avanços/exclusões do Dashboard, cadastros/vínculos/exclusões de
    # contrato, imports) no Sheets/CSV ativo, para valer em todas as páginas.
    pendentes = [k for k in ABAS if st.session_state.get(f"{k}_importado") is not None]
    if pendentes:
        with st.sidebar.expander(
            f"💾 Salvar tudo na fonte ({len(pendentes)} aba(s) alterada(s))",
            expanded=False,
        ):
            st.caption(
                "Alterações ainda NÃO gravadas na origem: "
                + ", ".join(ABAS[k] for k in pendentes)
                + ". Gravar faz elas valerem em todas as páginas e no arquivo/planilha."
            )
            if st.button("💾 Salvar agora", type="primary", key="btn_salvar_tudo"):
                sucessos, erros = salvar_tudo_na_fonte(
                    dados, fonte, id_planilha, api_url, api_key
                )
                for s in sucessos:
                    st.success(s)
                for e in erros:
                    st.error(e)
                if sucessos and not erros:
                    st.session_state["aviso_salvar_tudo"] = (
                        f"{len(sucessos)} aba(s) gravada(s) na fonte."
                    )
                    st.rerun()
    aviso_salvar = st.session_state.pop("aviso_salvar_tudo", None)
    if aviso_salvar:
        st.sidebar.success(aviso_salvar)

    injetar_pwa()
    if pagina == "📊 Dashboard":
        MAPEIO[pagina](dados, template=template_dash)
    else:
        MAPEIO[pagina](dados)


if __name__ == "__main__":
    main()
