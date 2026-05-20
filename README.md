# Vehicle Mass Platform
## View the Database (BD) page — **no Node.js, no admin**
1. Double-click **`run-bd-server.bat`** — it starts the server **and opens Microsoft Edge** automatically.  
   **or** in PowerShell:
   ```powershell
   cd C:\Users\TA55556\app\ai-app-optimization-for-mass-vehicle
   .\run-bd-server.ps1
   ```
2. If Edge did not open, go to **http://127.0.0.1:5173/** (or the port shown in the black window).
Uses only **PowerShell** (already on Windows). **Node.js is not used** to run the app.
### Still see the old page after an update?
1. **Stop** the server (Ctrl+C in the black window), run **`run-bd-server.bat`** again.
2. In the browser: **Ctrl+F5** (hard refresh) or Ctrl+Shift+R.
3. Check the top bar shows the current page name; use the **burger menu** (☰) for navigation.
Opening `index.html` directly from the folder (double-click) **does not work** — you must use the URL above while the server is running.
### If PowerShell blocks scripts
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```
(Allowed without admin — applies only to your user.)
Or use **`run-bd-server.bat`** only.
### If port 5173 is busy
```powershell
.\run-bd-server.ps1 -Port 8080
```
Then open **http://127.0.0.1:8080/**
---
## Data file
The grid loads **`web/public/data/bd-sheet.json`** (~5 MB).  
It is already in the repo after first export. If missing, copy it from a teammate or run export on a machine with Node (see below).
---
## Export from Excel (optional — needs Node once)
Only if the workbook changed and you have Node (or Cursor’s bundled node via script):
```powershell
.\export-bd-data.ps1
```
Installing [Node.js LTS](https://nodejs.org/) is optional and can be done **without admin** using the **zip** version: extract to e.g. `C:\Users\TA55556\node-portable\` and add that folder to your **user** PATH (no admin).
---
## Laravel API
See `api/README.md` — PHP stack when your environment allows it.
