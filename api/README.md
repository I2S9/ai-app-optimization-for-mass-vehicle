# API — Databricks + chargement progressif

Backend **Python FastAPI** pour :

1. **Charger par morceaux** (meta + cellules par fenêtre de lignes) — évite les 12 Mo JSON d’un coup  
2. **Sauvegarder** les sessions BD/Synthesis dans **Delta** (Unity Catalog)

## Démarrage local (sans Databricks)

Lit les JSON dans `web/public/data/` et expose la même API :

```powershell
cd api
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

Front : dans `web/index.html`, décommenter ou ajouter avant `main.js` :

```html
<script>window.__WGHT_API_BASE__ = 'http://127.0.0.1:8000';</script>
```

Ou lancer le serveur Node avec API intégrée (même origine, recommandé) :

```powershell
node web/server.mjs
```

## Databricks (prod)

1. Copier `.env.example` → `.env` et remplir `DATABRICKS_*`  
2. Exécuter `sql/001_schema.sql` dans un SQL Warehouse  
3. Ingérer les JSON exportés :

```powershell
cd api
python -m tools.ingest_json_to_delta --project-id default
```

4. `DATA_BACKEND=databricks` dans `.env`  
5. Déployer l’API (Databricks Apps, VM Azure, etc.)

Voir **`docs/databricks-setup.md`** pour le ticket IT et le plan complet.

## Endpoints

| Méthode | URL | Rôle |
|---------|-----|------|
| GET | `/api/v1/config` | `{ chunkedLoad: true }` |
| GET | `/api/v1/sheets/bd/meta` | Colonnes, headers, lastRow (sans cellules) |
| GET | `/api/v1/sheets/bd/cells?rowMin=2&rowMax=400` | Cellules d’une plage |
| GET | `/api/v1/sessions/{project_id}` | Snapshot complet (Databricks) |
| PUT | `/api/v1/sessions/{project_id}` | Save `{ revision, bd, syn }` |

Laravel (`api/README` historique) : **non utilisé** — Python uniquement.
