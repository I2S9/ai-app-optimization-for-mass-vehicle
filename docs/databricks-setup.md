# Configuration Databricks — WGHT Vehicle Mass

Objectif : **ne plus charger 12 Mo de JSON** au démarrage. Le navigateur reçoit d’abord les métadonnées (~50 Ko), puis les cellules **par blocs de lignes** depuis Delta via l’API.

---

## Architecture

```text
Navigateur                    API FastAPI                 Databricks
──────────                    ───────────                 ──────────
sheetDataApi.js  ──GET meta──►  /sheets/bd/meta     ──►  sheet_meta (Delta)
                 ──GET cells─►  /cells?rowMin..     ──►  sheet_cells (Delta)
                 ──PUT save──►  /sessions/{id}     ──►  workbook_sessions (Delta)
```

**Fallback** : si `/api/v1/config` est absent → chargement JSON classique (`bd-sheet.json`).

---

## Étape 1 — Ticket IT (template)

Copier-coller pour votre équipe infra :

> **Demande** : SQL Warehouse Databricks + Unity Catalog pour l’app WGHT Vehicle Mass  
>  
> - Catalog / schema : `main.vehicle_mass` (ou équivalent validé)  
> - SQL Warehouse avec **auto-stop** (ex. 10 min)  
> - Service principal ou PAT pour l’API (`DATABRICKS_TOKEN`)  
> - Droits : `CREATE TABLE`, `SELECT`, `INSERT`, `MERGE` sur le schema  
> - Hébergement API : **Databricks Apps** ou VM interne (port 8000)  
> - (Plus tard) **Entra ID** devant l’API + Static Web App pour le front  

---

## Étape 2 — Créer les tables Delta

Dans le **SQL Editor** du workspace, exécuter :

`api/sql/001_schema.sql`

Tables créées :

| Table | Usage |
|-------|--------|
| `workbook_sessions` | Snapshots complets BD+Syn à chaque Enregistrer |
| `sheet_meta` | JSON métadonnées (colonnes, headers, lastRow…) |
| `sheet_cells` | Cellules normalisées — requêtes `row BETWEEN …` |

---

## Étape 3 — Configurer l’API

```powershell
cd C:\Users\TA55556\app\ai-app-optimization-for-mass-vehicle\api
copy .env.example .env
# Éditer .env : DATABRICKS_HOST, DATABRICKS_HTTP_PATH, DATABRICKS_TOKEN
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Variables importantes :

| Variable | Exemple |
|----------|---------|
| `DATABRICKS_HOST` | `https://adb-1234567890123456.7.azuredatabricks.net` |
| `DATABRICKS_HTTP_PATH` | `/sql/1.0/warehouses/abc123def456` |
| `DATABRICKS_TOKEN` | PAT ou OAuth (secret — ne pas committer) |
| `DATA_BACKEND` | `databricks` en prod, `default` en local |

Test local sans Databricks :

```powershell
uvicorn app.main:app --reload --port 8000
curl http://127.0.0.1:8000/api/v1/health
```

---

## Étape 4 — Ingérer les JSON exportés

Une fois les tables créées et `.env` rempli :

```powershell
cd api
.\.venv\Scripts\Activate.ps1
python -m tools.ingest_json_to_delta --project-id default
```

Durée : quelques minutes (100k+ cellules BD). Ensuite les requêtes par fenêtre sont rapides.

---

## Étape 5 — Brancher le front

### Option A — API sur le même port (dev, recommandé)

```powershell
node web/server.mjs
```

Le serveur Node sert le front **et** `/api/v1/*` (chunks depuis JSON local).

### Option B — API Databricks sur port 8000

Dans `web/index.html` :

```html
<script>window.__WGHT_API_BASE__ = 'http://127.0.0.1:8000';</script>
```

Puis servir le front (`run-bd-server.bat` ou autre) et lancer l’API en parallèle.

---

## Comportement côté navigateur

1. **Boot Database** : meta + lignes 2–400 → grille visible en ~1–2 s  
2. **Arrière-plan** : autres blocs de lignes BD  
3. **Synthesis** : idem par blocs de 200 lignes  
4. **Enregistrer** (à venir) : `PUT /api/v1/sessions/default` avec `{ revision, bd, syn }`

Fichier client : `web/js/sheetDataApi.js`

---

## Prochaines étapes

- [ ] Auth Entra ID sur l’API  
- [ ] Save auto 60 s vers Delta (remplace IndexedDB seul)  
- [ ] Job Databricks nightly : snapshot → tables analytiques  
- [ ] Compression gzip sur les réponses API  

---

## Dépannage

| Symptôme | Cause probable |
|----------|----------------|
| Toujours 12 Mo au chargement | `/api/v1/config` inaccessible → fallback JSON |
| `503 Databricks` | Token expiré ou warehouse arrêté |
| Grille incomplète au scroll | Chunks pas finis — attendre indicateur ou recharger |
| `ingest` échoue | Schema non créé ou droits manquants |
