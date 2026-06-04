-- Databricks / Unity Catalog — run once in SQL warehouse (IT or notebook).
--
-- PERMISSION_DENIED on CREATE SCHEMA?
--   → Do NOT use eng_lab unless IT gave you CREATE SCHEMA on that catalog.
--   → Run the discovery queries below, pick a catalog + schema you own, then
--     set api/.env:  DATABRICKS_CATALOG=<catalog>  DATABRICKS_SCHEMA=<schema>
--
-- ── Step 0: discover what you can use (SQL Editor) ──
-- SHOW CATALOGS;
-- SHOW SCHEMAS IN <your_catalog>;
-- If CREATE SCHEMA fails, ask IT for an existing schema or use one you already have.

-- ── Step 1: replace these two lines with YOUR catalog + schema ──
-- Examples IT often provisions:
--   USE CATALOG main;
--   CREATE SCHEMA IF NOT EXISTS vehicle_mass;
-- Or use an existing team schema (no CREATE needed):
--   USE CATALOG my_team_catalog;
--   USE SCHEMA existing_schema_name;

USE CATALOG main;
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
