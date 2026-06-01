/**
 * Ingest web/public/data/*.json into Databricks Delta (no Python required).
 * Usage: node tools/ingest-databricks.mjs [--project-id default]
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { DBSQLClient } from '@databricks/sql';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');

function loadEnv() {
  const envPath = path.join(root, 'api', '.env');
  if (!fs.existsSync(envPath)) {
    throw new Error('Missing api/.env — copy from .env.example and add DATABRICKS_TOKEN');
  }
  const env = {};
  for (const line of fs.readFileSync(envPath, 'utf8').split(/\r?\n/)) {
    const t = line.trim();
    if (!t || t.startsWith('#')) continue;
    const i = t.indexOf('=');
    if (i < 0) continue;
    env[t.slice(0, i).trim()] = t.slice(i + 1).trim();
  }
  return env;
}

function table(catalog, schema, name) {
  return `${catalog}.${schema}.${name}`;
}

async function runQuery(session, sql, params = {}) {
  const op = await session.executeStatement(sql, { namedParameters: params });
  await op.fetchAll();
  await op.close();
}

async function ingestSheet(session, { projectId, sheet, filePath, catalog, schema }) {
  if (!fs.existsSync(filePath)) {
    console.log(`Skip missing ${filePath}`);
    return null;
  }
  console.log(`Loading ${path.basename(filePath)} …`);
  const raw = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  const meta = { ...raw };
  delete meta.cells;
  const metaJson = JSON.stringify(meta);

  await runQuery(
    session,
    `
    MERGE INTO ${table(catalog, schema, 'sheet_meta')} t
    USING (SELECT :pid AS pid, :sh AS sh, :mj AS mj, current_timestamp() AS ts) s
    ON t.project_id = s.pid AND t.sheet = s.sh
    WHEN MATCHED THEN UPDATE SET meta_json = s.mj, updated_at = s.ts
    WHEN NOT MATCHED THEN INSERT (project_id, sheet, meta_json, updated_at)
      VALUES (s.pid, s.sh, s.mj, s.ts)
    `,
    { pid: projectId, sh: sheet, mj: metaJson }
  );

  await runQuery(
    session,
    `DELETE FROM ${table(catalog, schema, 'sheet_cells')} WHERE project_id = :pid AND sheet = :sh`,
    { pid: projectId, sh: sheet }
  );

  const cells = raw.cells || [];
  const BATCH = 5000;
  for (let i = 0; i < cells.length; i += BATCH) {
    const chunk = cells.slice(i, i + BATCH);
    const values = chunk
      .map((cell) => {
        const v = cell.v == null ? 'NULL' : `'${String(cell.v).replace(/'/g, "''")}'`;
        const f = cell.f ? `'${String(cell.f).replace(/'/g, "''")}'` : 'NULL';
        const ue = cell.userEdited ? 'true' : 'false';
        return `('${projectId}', '${sheet}', ${Number(cell.r)}, '${String(cell.c).replace(/'/g, "''")}', ${v}, ${f}, ${ue})`;
      })
      .join(',\n');
    await runQuery(
      session,
      `INSERT INTO ${table(catalog, schema, 'sheet_cells')}
       (project_id, sheet, row, col, v, f, user_edited) VALUES ${values}`
    );
    process.stdout.write(`  ${sheet}: ${Math.min(i + BATCH, cells.length)}/${cells.length}\r`);
  }
  console.log(`  ${sheet}: ${cells.length} cells`);
  return raw;
}

async function main() {
  const env = loadEnv();
  const host = (env.DATABRICKS_HOST || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
  const httpPath = env.DATABRICKS_HTTP_PATH;
  const token = env.DATABRICKS_TOKEN;
  const catalog = env.DATABRICKS_CATALOG || 'labcatalog';
  const schema = env.DATABRICKS_SCHEMA || 'ta55556';
  const projectId = process.argv.includes('--project-id')
    ? process.argv[process.argv.indexOf('--project-id') + 1]
    : 'default';

  if (!host || !httpPath || !token || token.includes('COLLE_ICI')) {
    throw new Error('Fill DATABRICKS_HOST, DATABRICKS_HTTP_PATH, DATABRICKS_TOKEN in api/.env');
  }

  const dataDir = path.join(root, 'web', 'public', 'data');
  const client = new DBSQLClient();
  await client.connect({ host, path: httpPath, token });
  const session = await client.openSession();

  try {
    const bdRaw = await ingestSheet(session, {
      projectId,
      sheet: 'BD',
      filePath: path.join(dataDir, 'bd-sheet.json'),
      catalog,
      schema,
    });
    const synRaw = await ingestSheet(session, {
      projectId,
      sheet: 'SYNTHESIS',
      filePath: path.join(dataDir, 'synthesis-sheet.json'),
      catalog,
      schema,
    });

    if (bdRaw && synRaw) {
      const bdJson = JSON.stringify(bdRaw).replace(/'/g, "''");
      const synJson = JSON.stringify(synRaw).replace(/'/g, "''");
      await runQuery(
        session,
        `
        MERGE INTO ${table(catalog, schema, 'workbook_sessions')} t
        USING (
          SELECT :pid AS pid, 0 AS rev, :bd AS bd, :syn AS syn,
                 current_timestamp() AS ts, 'ingest' AS by
        ) s
        ON t.project_id = s.pid
        WHEN MATCHED THEN UPDATE SET
          bd_snapshot = s.bd, syn_snapshot = s.syn, updated_at = s.ts, updated_by = s.by
        WHEN NOT MATCHED THEN INSERT
          (project_id, revision, bd_snapshot, syn_snapshot, updated_at, updated_by)
          VALUES (s.pid, s.rev, s.bd, s.syn, s.ts, s.by)
        `,
        { pid: projectId, bd: bdJson, syn: synJson }
      );
    }
    console.log('Done.');
  } finally {
    await session.close();
    await client.close();
  }
}

main().catch((e) => {
  console.error(e.message || e);
  process.exit(1);
});
