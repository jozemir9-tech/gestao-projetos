#!/usr/bin/env python3
"""
conversor-cronograma.py
Converte cronograma entre Google Sheets (padrão DB_Cronograma_MSProject) <-> MS Project / Primavera P6

Uso:
  python conversor-cronograma.py sheets-para-msproject --input DB_Cronograma.csv --output cronograma_msproject.csv
  python conversor-cronograma.py sheets-para-p6 --input DB_Cronograma.csv --output cronograma_p6.csv
  python conversor-cronograma.py msproject-para-sheets --input export_msproject.csv --output DB_Cronograma.csv

Formatos:
  - Sheets espera ; como separador e datas DD/MM/AAAA
  - MS Project CSV importa melhor com cabeçalho em inglês (Task Name, Start, Finish...)
"""
import csv
import argparse
import sys
from datetime import datetime

# Cabeçalho padrão Sheets (nossa base de dados)
SHEETS_HEADER = ["ID_Tarefa","Nome_Tarefa","WBS","Inicio_Planejado","Fim_Planejado","Inicio_Real","Fim_Real","Duracao_dias","%_Concluido","Predecessora","Sucessora","Caminho_Critico_SIM_NAO","Responsavel","Status","Observacao"]

# Cabeçalho MS Project CSV para importação direta (Arquivo > Abrir > CSV)
MSPROJECT_HEADER = ["ID","Task Name","Outline Number","Duration","Start","Finish","Actual Start","Actual Finish","% Complete","Predecessors","Successors","Critical","Resource Names","Notes"]

# Cabeçalho Primavera P6 via spreadsheet
P6_HEADER = ["Activity ID","Activity Name","WBS Code","Planned Start","Planned Finish","Actual Start","Actual Finish","Activity % Complete","Predecessors","Successors","Critical","Resource IDs","Notes"]

def parse_date_br(s):
    s = (s or "").strip()
    if not s: return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except: pass
    return s

def sheets_para_msproject(inp, out):
    with open(inp, encoding="utf-8-sig") as f:
        sample = f.read(2048); f.seek(0)
        delim = ";" if sample.count(";") > sample.count(",") else ","
        reader = csv.DictReader(f, delimiter=delim)
        rows = list(reader)
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=MSPROJECT_HEADER, delimiter=";")
        w.writeheader()
        for r in rows:
            crit = (r.get("Caminho_Critico_SIM_NAO") or "").strip().upper()
            w.writerow({
                "ID": r.get("ID_Tarefa",""),
                "Task Name": r.get("Nome_Tarefa",""),
                "Outline Number": r.get("WBS",""),
                "Duration": f'{r.get("Duracao_dias","")} dias' if r.get("Duracao_dias") else "",
                "Start": parse_date_br(r.get("Inicio_Planejado","")),
                "Finish": parse_date_br(r.get("Fim_Planejado","")),
                "Actual Start": parse_date_br(r.get("Inicio_Real","")),
                "Actual Finish": parse_date_br(r.get("Fim_Real","")),
                "% Complete": r.get("%_Concluido",""),
                "Predecessors": r.get("Predecessora",""),
                "Successors": r.get("Sucessora",""),
                "Critical": "Yes" if crit in ("SIM","S","YES","Y") else "No",
                "Resource Names": r.get("Responsavel",""),
                "Notes": f'{r.get("Status","")} | {r.get("Observacao","")}'.strip(" |"),
            })
    print(f"OK: {len(rows)} tarefas exportadas para MS Project -> {out}")
    print("No MS Project: Arquivo > Abrir > selecione o CSV > Assistente de Importação > Novo mapa > marque cabeçalho > Concluir.")

def sheets_para_p6(inp, out):
    with open(inp, encoding="utf-8-sig") as f:
        sample = f.read(2048); f.seek(0)
        delim = ";" if sample.count(";") > sample.count(",") else ","
        reader = csv.DictReader(f, delimiter=delim)
        rows = list(reader)
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=P6_HEADER, delimiter=",")
        w.writeheader()
        for r in rows:
            w.writerow({
                "Activity ID": r.get("ID_Tarefa",""),
                "Activity Name": r.get("Nome_Tarefa",""),
                "WBS Code": r.get("WBS",""),
                "Planned Start": parse_date_br(r.get("Inicio_Planejado","")),
                "Planned Finish": parse_date_br(r.get("Fim_Planejado","")),
                "Actual Start": parse_date_br(r.get("Inicio_Real","")),
                "Actual Finish": parse_date_br(r.get("Fim_Real","")),
                "Activity % Complete": r.get("%_Concluido",""),
                "Predecessors": r.get("Predecessora",""),
                "Successors": r.get("Sucessora",""),
                "Critical": r.get("Caminho_Critico_SIM_NAO",""),
                "Resource IDs": r.get("Responsavel",""),
                "Notes": r.get("Observacao",""),
            })
    print(f"OK: {len(rows)} atividades exportadas para P6 -> {out}")
    print("No P6: File > Import > Spreadsheet > selecione o arquivo.")

def msproject_para_sheets(inp, out):
    with open(inp, encoding="utf-8-sig") as f:
        sample = f.read(2048); f.seek(0)
        delim = ";" if sample.count(";") > sample.count(",") else ","
        reader = csv.DictReader(f, delimiter=delim)
        # normaliza chaves
        rows_in = list(reader)
        field_map = {k.lower().strip(): k for k in (reader.fieldnames or [])}
        def g(r, *names):
            for n in names:
                if n.lower() in field_map:
                    return r[field_map[n.lower()]] or ""
            return ""
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=SHEETS_HEADER, delimiter=";")
        w.writeheader()
        for r in rows_in:
            dur = g(r, "Duration","Duracao")
            dur_num = "".join(c for c in dur if c.isdigit())
            crit = g(r, "Critical","Critica")
            pct = g(r, "% Complete","% Complete","% Concluido")
            notes = g(r, "Notes","Observacao")
            w.writerow({
                "ID_Tarefa": g(r, "ID","ID_Tarefa","Activity ID"),
                "Nome_Tarefa": g(r, "Task Name","Activity Name","Nome_Tarefa"),
                "WBS": g(r, "Outline Number","WBS Code","WBS"),
                "Inicio_Planejado": parse_date_br(g(r, "Start","Planned Start")),
                "Fim_Planejado": parse_date_br(g(r, "Finish","Planned Finish")),
                "Inicio_Real": parse_date_br(g(r, "Actual Start")),
                "Fim_Real": parse_date_br(g(r, "Actual Finish")),
                "Duracao_dias": dur_num,
                "%_Concluido": pct,
                "Predecessora": g(r, "Predecessors","Predecessora"),
                "Sucessora": g(r, "Successors","Sucessora"),
                "Caminho_Critico_SIM_NAO": "SIM" if crit.strip().lower() in ("yes","sim","y","s","true") else "NAO",
                "Responsavel": g(r, "Resource Names","Resource IDs"),
                "Status": "",
                "Observacao": notes,
            })
    print(f"OK: {len(rows_in)} tarefas importadas para Sheets -> {out}")
    print("No Sheets: Arquivo > Importar > colar na aba DB_Cronograma_MSProject.")

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("modo", choices=["sheets-para-msproject","sheets-para-p6","msproject-para-sheets"])
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    a = ap.parse_args()
    try:
        if a.modo == "sheets-para-msproject": sheets_para_msproject(a.input, a.output)
        elif a.modo == "sheets-para-p6": sheets_para_p6(a.input, a.output)
        else: msproject_para_sheets(a.input, a.output)
    except FileNotFoundError as e:
        print(f"Erro: arquivo não encontrado: {e}", file=sys.stderr); sys.exit(1)
