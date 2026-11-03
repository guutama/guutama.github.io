# Build & review pipeline — Hidda Latinsa fi Hortee Qaalluu Shariifaa

- `families.py` — THE master data. Tuples `(father, mother, [children "Name|note"], photo, note, dots, nochildren)`.
- `clans.py` — married-in person → clan.
- `build.py` — graph, lineages, generations; writes people.json, xlsx, family_outline.md (into build/).
- `make_site.py` — runs build.py and writes ../index.html and ../data/*.csv. **Run this after every change**: `python3 build/make_site.py`, then commit & push `main`. GitHub Pages serves https://guutama.github.io; the Google Sheet `Qaallu_shariifa` (1rK7ovg8bSaShP-HG_aK1MdJEehMmr8AyrVYtgtNnUJQ) imports data/*.csv into its People / Families / Open questions tabs.
- `transcription.md` — raw reading of the 43 notebook photos (the evidence). `Uncertain_readings_to_check.md` — readings Gutama already corrected. `origins_research.md` — background on Qaalluu/Sharifa origins.

## Inbox
Submissions from the website land in the sheet tab `Skjemasvar 1` (columns: timestamp, name, contact, kind, person, details, father, mother, spouses, children, ganda/clan). Decisions are appended to tab `Merge log` (Date, From, Type, Person / lineage, Decision, Note). A submission is "new" if its timestamp is not yet in Merge log.

## Review rules
- Accept a change when it is consistent with the transcription and the tree (same generation for a marriage, parents exist or are given, names match a known spelling variant).
- Correct a spelling → edit the name everywhere it occurs in families.py (children lists reference parents by exact string).
- Add a person → add to the parent couple's children list, or a new tuple for a new couple. Unknown-gender children are fine. Place names go in the note, not the name.
- Add a lineage → new root tuples; add the founder to `ROOTS` in build.py and to the `roots` browse list in page_template.html.
- Contradiction with the notebook, or not enough detail → do NOT change data; log `NEEDS INFO` with what to ask, and add a row to the Open questions sheet in build.py if it is a real genealogical question.
- Never store phone numbers or emails in families.py or the page; they stay in the sheet only.
