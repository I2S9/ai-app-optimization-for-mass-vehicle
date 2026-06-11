# Démo avec Supabase (recommandé)

Même architecture que Databricks : **meta + cellules par blocs** au chargement, **snapshot session** à chaque enregistrement (Bookmark Matrix + edits).

Les calculs **temps réel** restent dans le navigateur (`WorkbookSession`) — Supabase ne ralentit pas Synthesis.

---

## 1. Créer le projet Supabase

1. [supabase.com](https://supabase.com) → New project  
2. SQL Editor → coller et exécuter **`api/sql/supabase_001_schema.sql`**
3. SQL Editor → coller et exécuter **`api/sql/supabase_004_module_state.sql`** (Weight Tax, Waterline, …)

---

## 2. Configurer l’API

```powershell
cd api
copy .env.example .env
# Éditer SUPABASE_DB_URL (mot de passe DB du projet)
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

`.env` minimal :

```env
DATA_BACKEND=supabase
SUPABASE_DB_URL=postgresql://postgres.xxx:[PASSWORD]@....supabase.com:6543/postgres
```

---

## 3. Ingérer les JSON du projet

```powershell
cd api
.\.venv\Scripts\Activate.ps1
python tools/ingest_json_to_supabase.py --project-id default
```

(~2–5 min pour toutes les cellules BD)

---

## 4. Lancer API + front

**Terminal 1 — API**

```powershell
cd api
.\.venv\Scripts\Activate.ps1
uvicorn app.main:app --reload --port 8000
```

Vérifier : `http://127.0.0.1:8000/api/v1/config` → `"mode": "supabase"`

**Terminal 2 — Front**

Dans `web/index.html`, décommenter ou ajouter avant `main.js` :

```html
<script>window.__WGHT_API_BASE__ = 'http://127.0.0.1:8000';</script>
```

Puis :

```powershell
node web/server.mjs
```

---

## 5. Scénario démo

| Action | Résultat attendu |
|--------|------------------|
| Edit masse Database | Synthesis recalcule tout de suite |
| Bookmark Matrix : renommer / réordonner | Database + Synthesis alignés |
| F5 ou autre PC (même projet Supabase) | Session rechargée depuis le cloud |

---

## Plus tard : Databricks

Les tables Delta (`api/sql/001_schema.sql`) reprennent les mêmes rôles :

| Supabase | Databricks |
|----------|------------|
| `workbook_sessions` | `workbook_sessions` |
| `sheet_meta` | `sheet_meta` |
| `sheet_cells` | `sheet_cells` |

Changer `.env` : `DATA_BACKEND=databricks` + credentials, lancer `ingest_json_to_delta.py`.  
Le front ne change pas (même `/api/v1/*`).

---

## Dépannage

| Problème | Solution |
|----------|----------|
| `mode: local` | API pas joignable ou `SUPABASE_DB_URL` vide |
| 503 sur `/cells` | Tables non créées ou mauvais mot de passe URI |
| Save échoue | Vérifier ingest + `GET /sessions/default` |
| Calcul lent | Normal au 1er scroll ; puis cache navigateur |
