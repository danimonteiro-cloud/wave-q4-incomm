"""Wave Q4 2026 - gera os arquivos do dashboard Incomm a partir da planilha oficial.

Uso: python3 gerar_dados.py <planilha.xlsx> <pasta_saida> [historico_anterior.json]

Saídas (JSON {rows, meta}):
  pdvs.json         - 1 linha por PDV da aba "Wave Q42026_Total" (sem nome/CPF de promotor)
  historico.json    - 1 linha por dia com os totais; o dia atual é substituído a cada execução
Dados pessoais (nome e CPF de promotores) nunca são lidos para a saída.
"""
import json, sys, os
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl

src, out = sys.argv[1], sys.argv[2]
prev = sys.argv[3] if len(sys.argv) > 3 else None
code = open(__file__, encoding="utf-8").read()
now = datetime.now(ZoneInfo("America/Sao_Paulo"))
stamp = now.strftime("%Y-%m-%dT%H:%M:%S%z")

def clean(v):
    if v is None:
        return None
    return str(v).strip() or None

def norm_varejo(v):
    v = clean(v) or "Sem informação"
    return {"Lojas Americanas S.A.": "Americanas", "Pague Menos Brazil": "Pague Menos",
            "Extra farma": "Extrafarma"}.get(v, v)

wb = openpyxl.load_workbook(src, read_only=True, data_only=True)

def sheet(prefix):
    for ws in wb.worksheets:
        if ws.title.startswith(prefix):
            rows = [r for r in ws.iter_rows(values_only=True)]
            head = [clean(h) for h in rows[0]]
            return [dict(zip(head, r)) for r in rows[1:] if any(c is not None for c in r)]
    raise SystemExit(f"Aba não encontrada: {prefix}")

# PDVs cuja região ainda está com promotor em contratação (aba Regiões_Em_Contratação)
em_contrat = {clean(r.get("Código da Loja")) for r in sheet("Regiões_Em_Contrata")}

pdvs = []
for r in sheet("Wave Q4"):
    visit = clean(r.get("PDVs Visitados")) or "Sem informação"
    mat = clean(r.get("STATUS DE MATERIAIS")) or "Sem informação"
    pdvs.append({
        "codigo_loja": clean(r.get("Código da Loja")),
        "cnpj": clean(r.get("CNPJ")),
        "varejo": norm_varejo(r.get("Varejo")),
        "bandeira": clean(r.get("Bandeira")),
        "cidade": clean(r.get("Cidade")),
        "uf": clean(r.get("Estado")),
        "visitado": visit,
        "materiais": mat,
        "testeira_netflix": "Sim" if clean(r.get("Testeira Netflix")) else "Não",
        "promotor": "Em contratação" if clean(r.get("Código da Loja")) in em_contrat else "Alocado",
    })

def count(key, val):
    return sum(1 for p in pdvs if p[key] == val)

today = {
    "data": now.strftime("%Y-%m-%d"),
    "atualizado_em": stamp,
    "total_pdvs": len(pdvs),
    "visitados_sim": count("visitado", "Sim"),
    "visitados_nao": count("visitado", "Não"),
    "pendentes": count("visitado", "Pendente"),
    "mat_entregue": count("materiais", "Entregue"),
    "mat_em_transito": count("materiais", "Em trânsito"),
    "mat_aguardando_incomm": count("materiais", "Aguardando atualização da Incomm"),
    "mat_aguardando_retirada": count("materiais", "Aguardando retirada"),
    "mat_ocorrencia": count("materiais", "Ocorrência operacional"),
    "em_contratacao": count("promotor", "Em contratação"),
}
hist = []
if prev and os.path.exists(prev):
    hist = [h for h in json.load(open(prev, encoding="utf-8"))["rows"] if h["data"] != today["data"]]
hist.append(today)
hist.sort(key=lambda h: h["data"])

meta = [{"label": "gerar_dados.py", "value": code, "type": "code", "lang": "python"},
        {"label": "Entrada", "value": os.path.basename(src)},
        {"label": "Gerado em", "value": stamp}]
os.makedirs(out, exist_ok=True)
for name, rows in (("pdvs", pdvs), ("historico", hist)):
    with open(os.path.join(out, name + ".json"), "w", encoding="utf-8") as f:
        json.dump({"rows": rows, "meta": meta}, f, ensure_ascii=False)
print(json.dumps(today, ensure_ascii=False))
