-- Supabase SQL Editor → New query → Run
-- Persistance distante des modules (Weight Tax, Waterline, CDC Output, …)
-- Une ligne par (project_id, module_key), état JSON complet.

CREATE TABLE IF NOT EXISTS public.module_state (
  project_id   TEXT NOT NULL,
  module_key   TEXT NOT NULL,
  revision     BIGINT NOT NULL DEFAULT 0,
  state_json   JSONB NOT NULL DEFAULT '{}'::jsonb,
  updated_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_by   TEXT,
  PRIMARY KEY (project_id, module_key)
);

CREATE INDEX IF NOT EXISTS module_state_updated_at
  ON public.module_state (project_id, updated_at DESC);

ALTER TABLE public.module_state DISABLE ROW LEVEL SECURITY;
