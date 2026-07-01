# Integration Template Generator

CIQ + EDP → filled AMOS/CMCLI integration templates, per scope of work.

This is the Streamlit version of the tool, built to run as a real hosted app instead
of a self-contained browser file — so templates can be updated by editing a file in
this repo, without anyone needing to rebuild anything.

## What's included in this minimal version

- **MCA → MMBB Pre-existing**
- **MCA → CRAN SA Rehome Trip 1**

More scopes (CRAN Trip-2, CRAN NSA, N2E, NSB, CENM) get added the same way —
drop the blank template into `templates/<SCOPE>/`, add one entry to `SCOPE_MAP`
in `app.py`, and (if it needs anything beyond the shared logic) a small generator
function following the same pattern as `generate_cran_trip1`.

## Folder structure

```
.
├── app.py                          # the Streamlit app
├── requirements.txt                # Python dependencies
└── templates/
    └── MCA/
        ├── MMBB_Pre-existing.txt
        └── CRAN_Trip1.txt
```

## How to deploy (Streamlit Community Cloud — free)

1. Push this folder's contents to your GitHub repo (root of the repo, not a
   subfolder — `app.py` should sit at the top level).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub.
3. Click **New app**, pick this repo, branch `main`, main file path `app.py`.
4. Click **Deploy**. First deploy takes a minute or two.
5. You'll get a URL like `https://your-app-name.streamlit.app` — that's the live tool.

## How template updates work once deployed

1. Edit the `.txt` file inside `templates/` directly on GitHub (or push a new
   version from your machine).
2. Streamlit Cloud detects the change and automatically redeploys — usually
   within a minute.
3. Next person who opens the app link is on the new version. No rebuild, no
   asking anyone to regenerate a file.

## Known limitations of this first pass

- Only the two scopes above are wired up — everything else needs porting
  from the browser-JS version the same way these two were.
- `xx5G_Cell_namexx` / `xxFDD_namexx` (CRAN NSA only, not used in Trip-1) are
  intentionally left unfilled — confirmed to be manual RF-judgment fields.
- Pre-checks PDF parsing is regex-based against one known PDF format and may
  not generalize to differently-formatted reports.
- Legacy `.xls` EDP files need the `xlrd` package (already in
  `requirements.txt`) — if an EDP still fails to load, try re-saving it as
  `.xlsx` first.
