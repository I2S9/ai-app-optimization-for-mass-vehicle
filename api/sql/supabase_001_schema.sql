-- Supabase (PostgreSQL) — SQL Editor → New query → Run
-- Project: Settings → Database → Connection string (URI) → api/.env SUPABASE_DB_URL

-- Sessions (Bookmark Matrix structure + full BD/Syn snapshots on save)
CREATE TABLE IF NOT EXISTS public.workbook_sessions (
  project_id          TEXT PRIMARY KEY,
  revision            BIGINT NOT NULL DEFAULT 0,
  structure_revision  BIGINT NOT NULL DEFAULT 0,
  bd_snapshot         TEXT,
  syn_snapshot        TEXT,
  updated_at          TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_by          TEXT
);

-- Sheet metadata (columns, headers, lastRow — no cells)
CREATE TABLE IF NOT EXISTS public.sheet_meta (
  project_id  TEXT NOT NULL,
  sheet       TEXT NOT NULL,
  meta_json   JSONB NOT NULL,
  updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
  PRIMARY KEY (project_id, sheet)
);

-- Normalized cells (progressive load by row window)
CREATE TABLE IF NOT EXISTS public.sheet_cells (
  project_id   TEXT NOT NULL,
  sheet        TEXT NOT NULL,
  excel_row    INT NOT NULL,
  col          TEXT NOT NULL,
  v            TEXT,
  f            TEXT,
  user_edited  BOOLEAN DEFAULT false,
  PRIMARY KEY (project_id, sheet, excel_row, col)
);

CREATE INDEX IF NOT EXISTS sheet_cells_row_window
  ON public.sheet_cells (project_id, sheet, excel_row);

-- Precomputed display grid (transform output) so the browser skips the ~5s
-- in-browser transform. `pack` is the compressed (gz+b64) grid JSON; `fingerprint`
-- ties it to the raw cells/meta it was built from (stale packs are ignored).
CREATE TABLE IF NOT EXISTS public.sheet_grid_cache (
  project_id   TEXT NOT NULL,
  sheet        TEXT NOT NULL,
  fingerprint  TEXT NOT NULL,
  pack         TEXT NOT NULL,
  built_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
  PRIMARY KEY (project_id, sheet)
);

-- API FastAPI uses SUPABASE_DB_URL (postgres user) — RLS off for demo simplicity.
ALTER TABLE public.workbook_sessions DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.sheet_meta DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.sheet_cells DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.sheet_grid_cache DISABLE ROW LEVEL SECURITY;
