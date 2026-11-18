"""Monta o site da Wave Q4 a partir de pdvs.json e historico.json (formato {rows, meta}).
Uso: python3 montar_site.py <pdvs.json> <historico.json> <logo.png> <pasta_saida>
Saída: <pasta>/site/ (index.html + dados/ + logo.png, para hospedar)
       <pasta>/Wave_Q4_Incomm.html (arquivo único com os dados dentro, para enviar)"""
import json, sys, os, base64, shutil
pdvs_f, hist_f, logo_f, out = sys.argv[1:5]
here = os.path.dirname(os.path.abspath(__file__))
tpl = open(os.path.join(here, "template.html"), encoding="utf-8").read()
cols = ["codigo_loja", "cnpj", "varejo", "bandeira", "cidade", "uf", "visitado", "materiais", "promotor"]
pdvs = [{k: r.get(k) for k in cols} for r in json.load(open(pdvs_f, encoding="utf-8"))["rows"]]
hist = json.load(open(hist_f, encoding="utf-8"))["rows"]
site = os.path.join(out, "site"); os.makedirs(os.path.join(site, "dados"), exist_ok=True)
for name, rows in (("pdvs", pdvs), ("historico", hist)):
    json.dump({"rows": rows}, open(os.path.join(site, "dados", name + ".json"), "w", encoding="utf-8"), ensure_ascii=False)
shutil.copy(logo_f, os.path.join(site, "logo.png"))
shutil.copy(os.path.join(here, "d3.min.js"), os.path.join(site, "d3.min.js"))
shutil.copy(os.path.join(here, "exceljs.min.js"), os.path.join(site, "exceljs.min.js"))
open(os.path.join(site, "index.html"), "w", encoding="utf-8").write(
    tpl.replace("__LOGO__", "logo.png").replace("<!--__D3__-->", '<script src="d3.min.js"></script>'))
d3 = open(os.path.join(here, "d3.min.js"), encoding="utf-8").read().replace("</script", "<\\/script")
logo64 = "data:image/png;base64," + base64.b64encode(open(logo_f, "rb").read()).decode()
emb = json.dumps({"pdvs": pdvs, "historico": hist}, ensure_ascii=False).replace("</", "<\\/")
single = tpl.replace("<!--__D3__-->", "<script>" + d3 + "</script>").replace("__LOGO__", logo64).replace("/*__DADOS__*/null", emb)
open(os.path.join(out, "Wave_Q4_Incomm.html"), "w", encoding="utf-8").write(single)
print("ok", len(pdvs), "PDVs,", len(hist), "dia(s)")
