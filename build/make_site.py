"""Rebuild the whole site from families.py.
Usage:  python3 build/make_site.py   (run from the repo root or anywhere)
Steps: build.py -> people.json/xlsx/outline ; then index.html + data/*.csv in the repo root."""
import os, sys, json, csv, subprocess
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
subprocess.run([sys.executable, os.path.join(HERE, "build.py")], check=True, cwd=HERE)
tpl = open(os.path.join(HERE, "page_template.html"), encoding="utf-8").read()
data = open(os.path.join(HERE, "people.json"), encoding="utf-8").read().replace("</", "<\\/")
page = tpl.replace("__DATA__", data)
head = """<!doctype html>
<html lang="om">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="description" content="Hidda Latinsa fi Hortee Qaalluu Shariifaa — roots and descendants of the Qaalluu and Shariifaa lineages">
<style>html{color-scheme:light dark}body{margin:0;font:14px system-ui}img{max-width:100%}[hidden]{display:none!important}</style>
</head>
<body>
"""
open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8").write(head + page + "\n</body>\n</html>\n")
import openpyxl
wb = openpyxl.load_workbook(os.path.join(HERE, "Qaalluu_Sharifa_family_lookup.xlsx"), read_only=True)
os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
S = lambda v: "" if v is None else str(v)
rows = list(wb["People"].iter_rows(values_only=True)); hdr = rows[0]; ix = {h: i for i, h in enumerate(hdr)}
cols = ['#','Name','Lineage','Generation (Ali Shabel = 0)','Father','Mother','Spouse(s)','Children','Unnamed children (dots)','Other clan via FATHER','Other clan via MOTHER','Notes / location','Photo #']
out = [["#","Name","Lineage","Generation","Father","Mother","Spouse(s)","Children","Unnamed children (dots)","Other clan via FATHER","Other clan via MOTHER","Notes / location","Photo #"]]
out += [[S(r[ix[c]]) for c in cols] for r in rows[1:]]
fam = [[S(v) for v in r] for r in wb["Families"].iter_rows(values_only=True)]; fam[0][5] = "Status"
oq = [[S(v) for v in r] for r in wb["Open questions"].iter_rows(values_only=True)]
for name, d in [("people.csv", out), ("families.csv", fam), ("open_questions.csv", oq)]:
    with open(os.path.join(ROOT, "data", name), "w", newline="", encoding="utf-8") as f: csv.writer(f).writerows(d)
print("site rebuilt:", len(out)-1, "people;", len(fam)-1, "families")
