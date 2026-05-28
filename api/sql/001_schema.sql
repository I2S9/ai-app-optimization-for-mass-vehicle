-- Databricks / Unity Catalog — run once in SQL warehouse (IT or notebook).
-- Replace catalog/schema if your naming differs.

-- Pick a catalog you have access to (prefer a *_lab catalog for dev).
-- Example:
--   USE CATALOG eng_lab;
--   CREATE SCHEMA IF NOT EXISTS vehicle_mass;
--
-- If you don't know which one: start with eng_lab (or your team's *_lab).
USE CATALOG eng_lab;
CREATE SCHEMA IF NOT EXISTS vehicle_mass;

-- Full session snapshots (v1 — simple save/load)
CREATE TABLE IF NOT EXISTS vehicle_mass.workbook_sessions (
  project_id   STRING  NOT NULL,
  revision     BIGINT  NOT NULL,
  bd_snapshot  STRING,
  syn_snapshot STRING,
  updated_at   TIMESTAMP NOT NULL,
  updated_by   STRING
)
USING DELTA
TBLPROPERTIES ('delta.feature.allowColumnDefaults' = 'supported');

-- Normalized cells (v2 — progressive load by row window)
CREATE TABLE IF NOT EXISTS vehicle_mass.sheet_cells (
  project_id   STRING  NOT NULL,
  sheet        STRING  NOT NULL COMMENT 'BD or SYNTHESIS',
  row          INT     NOT NULL,
  col          STRING  NOT NULL,
  v            STRING,
  f            STRING,
  user_edited  BOOLEAN
)
USING DELTA
PARTITIONED BY (project_id, sheet);

-- Sheet metadata without cells (small JSON blob per project/sheet)
CREATE TABLE IF NOT EXISTS vehicle_mass.sheet_meta (
  project_id STRING NOT NULL,
  sheet      STRING NOT NULL,
  meta_json  STRING NOT NULL,
  updated_at TIMESTAMP NOT NULL
)
USING DELTA;

-- Optional: seed project
-- INSERT INTO vehicle_mass.workbook_sessions VALUES (
--   'default', 0, NULL, NULL, current_timestamp(), 'bootstrap'
-- );
