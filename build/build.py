import sys, re, collections
import os
OUT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, "/home/claude")
from families import FAMILIES, ALI_SHABEL, SHARIF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

P = {}  # name -> dict
def person(n):
    if n not in P:
        P[n] = dict(name=n, father=None, mother=None, spouses=[], children=[], notes=[], src=set(),
                    dots=0, nochildren=False, uncertain=("?" in n))
    return P[n]

def split(c):
    if "|" in c:
        n, note = c.split("|", 1); return n.strip(), note.strip()
    return c.strip(), ""

for fa, mo, kids, src, note, dots, noch in FAMILIES:
    for par in (fa, mo):
        if par:
            p = person(par); p["src"].add(src)
            if dots: p["dots"] += dots
            if noch: p["nochildren"] = True
            if note and (not kids or dots or noch): p["notes"].append(note)
    if fa and mo:
        if mo not in person(fa)["spouses"]: P[fa]["spouses"].append(mo)
        if fa not in person(mo)["spouses"]: P[mo]["spouses"].append(fa)
    for c in kids:
        n, cn = split(c); ch = person(n); ch["src"].add(src)
        if fa: ch["father"] = fa
        if mo: ch["mother"] = mo
        if cn: ch["notes"].append(cn)
        for par in (fa, mo):
            if par and n not in P[par]["children"]: P[par]["children"].append(n)
    if note and kids:
        for par in (fa, mo):
            if par and note not in P[par]["notes"]: P[par]["notes"].append(note)

# ancestry: generation from Ali Shabel
gen = {ALI_SHABEL: 0, "Abdulli Nuur": -1, "Nuur Hille": -2, "Hille Kalaf": -3, "Kalaf Diini": -4, "Diini Haaji": -5, "Haaji Umar": -6, "Umar Ziyad": -7, "Ziyad": -8}
changed = True
while changed:
    changed = False
    for n, p in P.items():
        for par in (p["father"], p["mother"]):
            if par in gen and (n not in gen or gen[par] + 1 < gen[n]):
                gen[n] = gen[par] + 1; changed = True

def desc_of(root):
    out = set(); stack = [root]
    while stack:
        x = stack.pop()
        for c in P[x]["children"]:
            if c not in out: out.add(c); stack.append(c)
    return out

sharifa = desc_of(SHARIF) | {SHARIF}
qaalluu_all = desc_of(ALI_SHABEL) | {ALI_SHABEL}
# branch = which child of Ali Shabel / which child of Sharif
branch = {}
for top in P[ALI_SHABEL]["children"]:
    for d in desc_of(top) | {top}:
        branch.setdefault(d, set()).add(top)
for top in P[SHARIF]["children"]:
    for d in desc_of(top) | {top}:
        branch.setdefault(d, set()).add(top)

# other roots (lineages not descended from Ali Shabel). label, root person
ROOTS = [("Sharifa of Dhummuga (Warra Hasan)", "Hasan (Dhummuga)"),
         ("Qaalluu - Warra Abdisabur", "Abdusabur"), ("Qaalluu - Mohammadnur", "Mohammadnur"),
         ("Sharif - Abrahim Hassan family", "Abrahim Hassan"),
         ("Noolee - Dulal family", "Dulal") ]
ROOTS = [(l, r) for l, r in ROOTS if r in P]
root_desc = {l: desc_of(r) | {r} for l, r in ROOTS}
root_gen = {}
for l, r in ROOTS:
    g = {r: 0}; ch = True
    while ch:
        ch = False
        for n, p in P.items():
            for par in (p["father"], p["mother"]):
                if par in g and (n not in g or g[par] + 1 < g[n]): g[n] = g[par] + 1; ch = True
    root_gen[l] = g
def other_roots(n): return [l for l, _ in ROOTS if n in root_desc[l]]
def root_generation(n):
    """(label, gen) for the first other root n descends from, if any"""
    for l, _ in ROOTS:
        if n in root_desc[l]: return l, root_gen[l][n]
    return None, None

def lineage(n):
    s = n in sharifa; q = n in qaalluu_all
    base = _lineage_main(n, s, q)
    others = other_roots(n)
    if others:
        return (base + " + " if base != "Married in / other" else "") + " + ".join(others)
    return base
def _lineage_main(n, s, q):
    if s and q:
        # Sharifa all descend from Ammo; call it Sharifa unless also from another Ali Shabel child
        other_q = branch.get(n, set()) - {"Ammo Ali Shabel"} - set(P[SHARIF]["children"])
        return "Sharifa + Qaalluu (both)" if other_q else "Sharifa"
    if q: return "Qaalluu (Ali Shabel)"
    if n == SHARIF: return "Sharifa (founder)"
    return "Married in / other"

from clans import CLAN
# add spouse-note persons whose marker sits on a child note (e.g. "& Abdalla Ali Abdi (Ituu)") are handled in CLAN directly
inh = {}
def inherited(n, stack=()):
    if n in inh: return inh[n]
    if n in stack: return set()
    p = P[n]; s = set()
    for par in (p["father"], p["mother"]):
        if par:
            s |= inherited(par, stack + (n,))
            if par in CLAN: s.add(CLAN[par])
    inh[n] = s; return s
def side(n, which):
    par = P[n][which]
    if not par: return ""
    s = set(inherited(par))
    if par in CLAN: s.add(CLAN[par])
    return "; ".join(sorted(s))

def children_status(p):
    if p["nochildren"]: return "No children (recorded)"
    if p["children"] and p["dots"]: return f"Yes - {len(p['children'])} named + {p['dots']} unnamed"
    if p["children"]: return f"Yes - {len(p['children'])} named"
    if p["dots"]: return f"Yes - {p['dots']} unnamed (dots)"
    return "Not recorded"

def ancestor_line(n):
    out = []; x = n
    while x:
        p = P[x]
        nxt = p["father"] if p["father"] in gen or p["mother"] not in gen else p["mother"]
        if nxt is None or nxt not in gen: break
        out.append(nxt); x = nxt
    return out

# ---------- Spreadsheet ----------
wb = Workbook()
hdr_font = Font(name="Arial", bold=True, color="FFFFFF"); hdr_fill = PatternFill("solid", fgColor="2F4F4F")
body = Font(name="Arial", size=10)
def write_sheet(ws, headers, rows, widths):
    ws.append(headers)
    for c in ws[1]: c.font = hdr_font; c.fill = hdr_fill; c.alignment = Alignment(vertical="center", wrap_text=True)
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font = body; c.alignment = Alignment(vertical="top", wrap_text=True)
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"; ws.auto_filter.ref = ws.dimensions

ws = wb.active; ws.title = "People"
RANK={"Qaalluu (Ali Shabel)":0,"Sharifa (founder)":1,"Sharifa":2,"Sharifa + Qaalluu (both)":3,"Married in / other":9}
order = sorted(P, key=lambda n: (RANK.get(lineage(n), 5), gen.get(n, 99), n))
rows = []
for i, n in enumerate(order, 1):
    p = P[n]
    rows.append([i, n, lineage(n), gen.get(n, ""), ", ".join(sorted(branch.get(n, []))), CLAN.get(n, ""), side(n, "father"), side(n, "mother"),
                 p["father"] or "", p["mother"] or "", "; ".join(p["spouses"]), len(p["children"]), p["dots"],
                 children_status(p), "; ".join(p["children"]), " > ".join(ancestor_line(n)),
                 "; ".join(dict.fromkeys(p["notes"])), ", ".join(sorted(p["src"], key=lambda s: int(s.split(",")[0]) if s.split(",")[0].isdigit() else 999)),
                 "uncertain reading" if p["uncertain"] else ""])
write_sheet(ws, ["#", "Name", "Lineage", "Generation (Ali Shabel = 0)", "Branch (child of founder)", "Own clan (if married in)", "Other clan via FATHER", "Other clan via MOTHER", "Father", "Mother", "Spouse(s)",
                 "Named children", "Unnamed children (dots)", "Children status", "Children", "Line back to founder", "Notes / location", "Photo #", "Reading"],
            rows, [5, 34, 22, 12, 30, 22, 26, 26, 30, 30, 34, 9, 9, 24, 60, 60, 50, 10, 12])

ws2 = wb.create_sheet("Families")
rows = []
for fa, mo, kids, src, note, dots, noch in FAMILIES:
    rows.append([fa or "", mo or "", "; ".join(split(k)[0] for k in kids), len(kids), dots, "no children" if noch else "", note, src])
write_sheet(ws2, ["Father / husband", "Mother / wife", "Children (in notebook order)", "Named", "Dots", "", "Note", "Photo #"], rows, [32, 32, 70, 7, 7, 12, 60, 10])

ws3 = wb.create_sheet("Intermarriages")
rows = []
for fa, mo, kids, src, note, dots, noch in FAMILIES:
    if "INTERMARRIAGE" in note.upper() and fa and mo:
        rows.append([fa, lineage(fa), gen.get(fa, ""), mo, lineage(mo), gen.get(mo, ""), note, src])
write_sheet(ws3, ["Husband", "His lineage", "Gen", "Wife", "Her lineage", "Gen", "Note", "Photo #"], rows, [32, 22, 6, 32, 22, 6, 70, 10])

ws4 = wb.create_sheet("Dots - children unknown")
rows = [[n, lineage(n), gen.get(n, ""), "; ".join(P[n]["spouses"]), P[n]["dots"], len(P[n]["children"]), ", ".join(sorted(P[n]["src"]))]
        for n in order if P[n]["dots"]]
write_sheet(ws4, ["Person", "Lineage", "Gen", "Spouse(s)", "Unnamed children (dots)", "Named children", "Photo #"], rows, [34, 22, 6, 34, 10, 10, 10])

ws6 = wb.create_sheet("Other clans")
rows = []
for n, cl in sorted(CLAN.items(), key=lambda x: (x[1], x[0])):
    if n not in P: continue
    d = desc_of(n)
    rows.append([cl, n, "; ".join(P[n]["spouses"]), "; ".join(lineage(s) for s in P[n]["spouses"]), len(P[n]["children"]) , len(d), ", ".join(sorted(P[n]["src"]))])
write_sheet(ws6, ["Clan (as written in notebook)", "Married-in person", "Married to", "Spouse's lineage", "Children", "All descendants in notebook", "Photo #"], rows, [36, 30, 34, 26, 9, 12, 10])
ws5 = wb.create_sheet("Open questions")
qs = [
 ["Abibe Ali Shabel (Diima)", "Separate son of Ali Shabel living at Diima; his descendants are not yet recorded."],
 ["Ali Shabel's wife and parents", "Wife still unknown. Father per oral nasab: Abdulli Nuur Hille Kalaf Diini Haaji Umar Ziyad (recited by Gutama; needs cross-checking with other branches' elders)."],
 ["Abdalla Sharif's own nasab", "Name settled as Abdalla (given name) + Sharif (title), as people say it. His father and the chain above him are still unknown - ask the Sharifa elders."],
 ["Who is Abdalla (father of Ahmed Abdalla)?", "Ahmed Abdalla married Faxuma Ahmed Sharif. His father Abdalla has 6 children listed (photo 5) but no clan/parents."],
 ["Mumme Sharif's children", "Never listed as a family. Names like 'Sitti Ahmed Mumme Sharif', 'Aisha Habibe Mumme Sharif', 'Xaha Abrahim Mumme Sharif', 'Ammad Mumme Sharif', 'Abibe Mumme Sharif Harumte' imply sons Ahmed, Habibe, Abrahim, Ammad, Abibe."],

 ["Heydar (husband of Nafisa Abrahim Habibe)", "Header says 'Ammad Mumme Sharif' - is Heydar son of Ammad Mumme Sharif? Also Shukri Heydar married Ahmed Habibe Ahmed (Gutama's uncle) - same Shukri?"],
 ["Miski Aliyi Mummad Ahmed (wife of Mohammad Sheeka)", "Her father 'Aliyi Mummad Ahmed' not found in the notes."],
 ["Warra Abdisabur (Kamer Yusuf's family) - how related?", "Qaalluu, but the link from Abdisabur/Mumed-Nur up to Ali Shabel is not recorded. Known marriage links: Zeynaba Yusuf Mumed-Nur x Abibe Mumme Sharif; Fatuma Mude (Kamer's mother) <-> Yousuf Abdalla Adurman Sharif (daughter or wife?). Is 'Qalluu Abdusalamo' (photo 42) the same as Abdisabur?"],
 ["Warra Abdo Diini", "Boxed note: 'warra Abdo Diini kun qoratamuu qaba' - this family needs research. Xaha Abrahim Mumme Sharif married Faxuma Abdo Diini; Ussen Nalif? married daughter of Aisha Abdo Diini; Haji Aliyi Mammad Sharif & Zeynaba Din."],
 ["Qallu Diini origin", "Faxuma Usmane Ali Shabel married Diini (Dadar) -> Abrahim Diini. Is this the origin of the 'Qallu Diini' name?"],
 ["WEB FINDING - Abdi Qophee", "Oromia Writers Association biography: Abdii Qophee = Abdii Mohammad (Abdii Mohammed Yusuf); father Mohammad Yuusuf Abdusabuur, mother Faaxumaa Mohammadnur; born 1945 at Qalaadii near Soqaa, aanaa Dadar, East Hararghe; school in Dadar and Dire Dawa, Haile Selassie I University; Radio Harar presenter of 'Qophee Afaan Oromoo'; 500+ songs; died 1997; 5 children. https://oromiawritersassociation.org/the_authors/abdii-qophee/"],
 ["WEB FINDING - Qaalluu clan", "Afendi Muteki lists Qaalluu as one of the ten clans of the Ittu Oromo (with Baabbo, Alga, Waayye, Wacaale, Addaayyo, Arroojjii, Baaye, Gaamo, Gaadullaa). https://afendimutekiharar.blogspot.com/2013/03/the-oromo-of-harerghe-on-their-sub.html"],
 ["WEB - nothing found", "Ali Shabel, Sharif Abdalla, Qallu Diini, Biyyo Karaba, Sheik Usman Darasa, Dr. Abdurahman Abubakar Umare Jara, Abdusabur: no identifiable online record found (searched Nov 2026)."],
 ["Informants noted", "Jafar Mummad Ammade (Karamille) 'has information'; Mala (Laaftoo, Ganda Habib Borantoo, Oborra); Habib Borento Galmoo; Usman / Mukhtar Kaarruu?"],
 ["Abrahim Hassan family", "Noted 'Sharif (Abdalla miti)' - a Sharif family not from Sharif Abdalla. Married in twice (Kulsuma -> Mummad Adame; Halima -> Ahmed Ali Habibe)."],
]
write_sheet(ws5, ["Question", "Detail"], qs, [40, 120])
wb.save(OUT+"/Qaalluu_Sharifa_family_lookup.xlsx")

# ---------- Outline ----------
def outline(root, depth=0, seen=None, lines=None, maxd=12):
    seen = seen if seen is not None else set(); lines = lines if lines is not None else []
    p = P[root]
    sp = (" — fi " + " / ".join(p["spouses"])) if p["spouses"] else ""
    _n = [x for x in dict.fromkeys(p["notes"]) if x and not x.startswith("INTERMARRIAGE") and len(x) < 60 and x != "no children"]
    nt = (" [" + "; ".join(_n) + "]") if _n else ""
    inter = " ★" if any("INTERMARRIAGE" in x for x in p["notes"]) else ""
    if root in CLAN: inter += f" ⟨{CLAN[root]}⟩"
    else:
        _f, _m = side(root, "father"), side(root, "mother")
        if _f or _m: inter += " ⟨" + "; ".join(x for x in [("father: " + _f) if _f else "", ("mother: " + _m) if _m else ""] if x) + "⟩"
    tail = ""
    if p["dots"]: tail += f"  (+{p['dots']} children, names unknown)"
    if p["nochildren"]: tail += "  (no children)"
    lines.append("  " * depth + f"- {root}{sp}{nt}{inter}{tail}")
    if root in seen: lines[-1] += "  (see above)"; return lines
    seen.add(root)
    if root == "Ammo Ali Shabel":
        lines.append("  " * (depth+1) + "- (her children = the Sharifa clan: see section 2)"); return lines
    for c in p["children"]:
        outline(c, depth + 1, seen, lines, maxd)
    return lines

md = ["# Qaalluu and Sharifa family lookup — descendancy outline",
      "",
      "Compiled from Gutama's notebook (43 photos). Spelling follows the notebook. `fi` = married to. `★` = marriage between the two clans or between cousins. `?` = uncertain reading. `⟨…⟩` = other clan: for a married-in person their own clan, for a descendant the other clans they carry via father / via mother. Numbers in `(+N children, names unknown)` are the dotted bullets in the notebook.",
      "", "## 1. Ali Shabel (Abdullahi) — Qaalluu", ""]
SEEN=set(); md += outline(ALI_SHABEL, seen=SEEN)
md += ["", "## 2. Abdalla Sharif fi Ammo Ali Shabel — Sharifa", ""]
md += outline(SHARIF, seen=SEEN)
md += ["", "## 3. Families that married in (not descended from the founders)", ""]
for r in [r for _, r in ROOTS] + ["Abdalla (father of Ahmed Abdalla)"]:
    md += outline(r, seen=SEEN); md.append("")
open(OUT+"/family_outline.md", "w").write("\n".join(md))

import collections
cnt = collections.Counter()
for n in P:
    for c in set(side(n,"father").split("; ") + side(n,"mother").split("; ")) - {""}: cnt[c] += 1
print("descendants carrying other-clan blood:", cnt)
print("people:", len(P), "| descended from Ali Shabel:", len(qaalluu_all), "| Sharifa:", len(sharifa),
      "| with dots:", sum(1 for p in P.values() if p["dots"]), "| total dots:", sum(p["dots"] for p in P.values()))
# sanity: people with children but no parents and not founders (roots)
roots = [n for n, p in P.items() if not p["father"] and not p["mother"] and p["children"] and n not in (ALI_SHABEL,)]
print("roots:", len(roots))

# ---------- JSON export for the web lookup ----------
import json
ids = {n: i for i, n in enumerate(order, 1)}
out = []
for n in order:
    p = P[n]
    out.append({"id": ids[n], "n": n, "l": lineage(n), "g": gen.get(n), "f": p["father"], "m": p["mother"],
                "sp": p["spouses"], "ch": p["children"], "d": p["dots"], "nc": p["nochildren"],
                "cf": side(n, "father"), "cm": side(n, "mother"), "oc": CLAN.get(n, ""),
                "nt": [x for x in dict.fromkeys(p["notes"]) if x][:4], "r": root_generation(n)[0], "rg": root_generation(n)[1], "src": ", ".join(sorted(p["src"])),
                "unc": p["uncertain"]})
json.dump(out, open(OUT+"/people.json", "w"), ensure_ascii=False, separators=(",", ":"))
print("json people:", len(out), "bytes:", len(open(OUT+"/people.json").read()))
