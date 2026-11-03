# Hidda Latinsa fi Hortee Qaalluu Shariifaa — project guide for Claude Code

Family record (Oromo, East Hararghe; Qaalluu and Sharifa clans) maintained by Gutama Ibrahim.
Live site: https://guutama.github.io  (GitHub Pages, branch main, root).

## Layout
- `build/families.py` — THE master data. Every person, marriage and child comes from here.
- `build/clans.py` — married-in person → clan. `build/build.py` — graph/lineage logic. `build/page_template.html` — the website (bilingual strings in `S.om` / `S.en`).
- `build/transcription.md` — raw reading of the 43 notebook photos = the evidence. `build/Uncertain_readings_to_check.md` — corrections Gutama already gave.
- `build/README.md` — review rules for merging submissions. Read it before changing data.
- `index.html`, `data/*.csv` — GENERATED. Never edit by hand.

## Commands
- Rebuild site + CSVs: `python3 build/make_site.py`  (needs `pip install openpyxl`)
- Preview locally: `python3 -m http.server 8000` then open http://localhost:8000
- Publish: `git add -A && git commit -m "..." && git push` → live in ~1 min; the Google Sheet tabs People/Families/Open questions import data/*.csv automatically.

## Submissions loop
Website "Send" → Google Form → sheet `Qaallu_shariifa` (1rK7ovg8bSaShP-HG_aK1MdJEehMmr8AyrVYtgtNnUJQ), tab `Skjemasvar 1`. Decisions go to tab `Merge log` (col A = exact inbox timestamp). An hourly cloud task ("Qaalluu-Shariifaa: review new submissions") already does this; when working locally, "check submissions" = the same steps (needs the Google Sheets MCP connector, or a CSV export of the tab).

## Conventions
- Notebook notation: `fi`/`and` = spouse; a bare name after a woman = her child; "dots" = children exist but unnamed; Diima/Hadas/Wagur/Karra Mille/Ganda Bidu/Soqa are places.
- Spelling variants are the same person: Faxuma/Fatuma, Mohammad/Mummad/Mahammad, Abrahim/Ibrahim, Qamar/Kamer, Abibe/Habibe (for Abdi/Kadaro/Haloo's father).
- Never put phone numbers or emails into data files or the page.
- Commit messages: say which people/families changed.
