-- Supabase SQL Editor -> New query -> coller -> Run
-- Cree la table qui stocke la grille d'affichage pre-calculee (transform),
-- pour servir l'app sans le fichier statique /public/data/*-grid.json.

CREATE TABLE IF NOT EXISTS public.sheet_grid_cache (
  project_id   TEXT NOT NULL,
  sheet        TEXT NOT NULL,
  fingerprint  TEXT NOT NULL,
  pack         TEXT NOT NULL,
  built_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
  PRIMARY KEY (project_id, sheet)
);

ALTER TABLE public.sheet_grid_cache DISABLE ROW LEVEL SECURITY;
