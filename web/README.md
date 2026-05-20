# Web — Database (BD) page
Excel-faithful **BD** grid (English headers, data from row 6, virtual scroll).
## Run locally (Windows)
**Node is not on your PATH?** Use the scripts at repo root (they use Cursor’s Node or system Node):
```powershell
# From repo root in PowerShell:
.\run-bd-server.ps1
```
Or double-click **`run-bd-server.bat`** (opens **Microsoft Edge** automatically).
Open **http://127.0.0.1:5173/** if the browser did not open.
### Export data again (only if the `.xlsm` changed)
```powershell
.\export-bd-data.ps1
```
### If you install Node.js (recommended long-term)
Download **LTS** from https://nodejs.org/ — then `node` works everywhere:
```powershell
node tools/export-bd-sheet.mjs
node web/server.mjs
```
## Notes
- Data file: `public/data/bd-sheet.json` (~5 MB), generated once from `workbooks/base-de-donnees-complete-avec-liens.xlsm`.
- Formula cells are **read-only** (grey); manual cells are editable in the browser (not persisted until Laravel API exists).
- Rows with `A = _ADDBLUE` use light blue highlighting like Excel.
