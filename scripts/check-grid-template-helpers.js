#!/usr/bin/env node
/**
 * Ensures Vue grid components expose every synStore helper referenced in inline templates.
 * Run: node scripts/check-grid-template-helpers.js
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.join(__dirname, '..');

const GRID_FILES = [
  {
    file: path.join(repoRoot, 'web', 'js', 'SynthesisGrid.js'),
    helpersObject: 'synGridTemplateHelpers',
  },
];

function extractTemplate(source) {
  const marker = 'template: `';
  const start = source.indexOf(marker);
  if (start < 0) return null;
  let i = start + marker.length;
  let escaped = false;
  for (; i < source.length; i++) {
    const ch = source[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (ch === '\\') {
      escaped = true;
      continue;
    }
    if (ch === '`') return source.slice(start + marker.length, i);
  }
  return null;
}

function extractObjectLiteralKeys(source, objectName) {
  const marker = `const ${objectName} = {`;
  const start = source.indexOf(marker);
  if (start < 0) return null;
  let i = start + marker.length;
  let depth = 1;
  let inString = null;
  let escaped = false;
  for (; i < source.length; i++) {
    const ch = source[i];
    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        escaped = true;
        continue;
      }
      if (ch === inString) inString = null;
      continue;
    }
    if (ch === "'" || ch === '"' || ch === '`') {
      inString = ch;
      continue;
    }
    if (ch === '{') depth += 1;
    if (ch === '}') {
      depth -= 1;
      if (depth === 0) {
        const block = source.slice(start + marker.length, i);
        const keys = new Set();
        const keyRe = /^\s*([A-Za-z_$][\w$]*)\s*(?::|,)/gm;
        let match;
        while ((match = keyRe.exec(block))) keys.add(match[1]);
        return keys;
      }
    }
  }
  return null;
}

function extractSynStoreImports(source) {
  const marker = "from './synStore.js";
  const importStart = source.lastIndexOf('import {', source.indexOf(marker));
  const importEnd = source.indexOf('}', importStart);
  if (importStart < 0 || importEnd < 0) return new Set();
  const block = source.slice(importStart, importEnd + 1);
  const names = new Set();
  const nameRe = /\b([A-Za-z_$][\w$]*)\b/g;
  let match;
  while ((match = nameRe.exec(block))) {
    const name = match[1];
    if (name !== 'import' && name !== 'from') names.add(name);
  }
  return names;
}

function extractSetupReturnKeys(source, helpersObject) {
  const anchor = `const ${helpersObject} = {`;
  const anchorIdx = source.indexOf(anchor);
  if (anchorIdx < 0) return new Set();
  const returnMarker = 'return {';
  const returnIdx = source.indexOf(returnMarker, anchorIdx);
  if (returnIdx < 0) return new Set();
  let i = returnIdx + returnMarker.length;
  let depth = 1;
  let inString = null;
  let escaped = false;
  const keys = new Set();
  let blockStart = i;
  for (; i < source.length; i++) {
    const ch = source[i];
    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        escaped = true;
        continue;
      }
      if (ch === inString) inString = null;
      continue;
    }
    if (ch === "'" || ch === '"' || ch === '`') {
      inString = ch;
      continue;
    }
    if (ch === '{') depth += 1;
    if (ch === '}') {
      depth -= 1;
      if (depth === 0) {
        const block = source.slice(blockStart, i);
        const keyRe = /^\s*(?:\.\.\.([A-Za-z_$][\w$]*)|([A-Za-z_$][\w$]*)\s*(?::|,))/gm;
        let match;
        while ((match = keyRe.exec(block))) {
          if (match[1]) keys.add(`...${match[1]}`);
          if (match[2]) keys.add(match[2]);
        }
        break;
      }
    }
  }
  return keys;
}

function resolveExposedKeys(source, helpersObject, helperKeys) {
  const exposed = new Set();
  for (const key of extractSetupReturnKeys(source, helpersObject)) {
    if (key === `...${helpersObject}`) {
      for (const helperKey of helperKeys) exposed.add(helperKey);
      continue;
    }
    exposed.add(key);
  }
  return exposed;
}

function extractSynTemplateCalls(template) {
  const names = new Set();
  const callRe = /\b([A-Za-z_$][\w$]*)\s*\(/g;
  let match;
  while ((match = callRe.exec(template))) names.add(match[1]);
  return names;
}

let failed = false;

for (const { file, helpersObject } of GRID_FILES) {
  const rel = path.relative(repoRoot, file);
  const source = fs.readFileSync(file, 'utf8');
  const template = extractTemplate(source);
  if (!template) {
    console.error(`${rel}: template string not found`);
    failed = true;
    continue;
  }

  const helperKeys = extractObjectLiteralKeys(source, helpersObject);
  if (!helperKeys || helperKeys.size === 0) {
    console.error(`${rel}: ${helpersObject} object not found`);
    failed = true;
    continue;
  }

  if (!source.includes(`...${helpersObject}`)) {
    console.error(
      `${rel}: setup() return must spread ...${helpersObject} so template helpers stay in sync`
    );
    failed = true;
  }

  const synStoreImports = extractSynStoreImports(source);
  const templateCalls = extractSynTemplateCalls(template);
  const synStoreCallsInTemplate = [...templateCalls].filter((name) =>
    synStoreImports.has(name)
  );
  const missingFromHelpers = synStoreCallsInTemplate
    .filter((name) => !helperKeys.has(name))
    .sort();
  const exposed = resolveExposedKeys(source, helpersObject, helperKeys);
  const missingFromReturn = synStoreCallsInTemplate
    .filter((name) => !exposed.has(name))
    .sort();

  if (missingFromHelpers.length) {
    failed = true;
    console.error(`${rel}: synStore helpers used in template but missing from ${helpersObject}:`);
    for (const name of missingFromHelpers) console.error(`  - ${name}`);
  }
  if (missingFromReturn.length) {
    failed = true;
    console.error(`${rel}: synStore helpers used in template but not exposed from setup():`);
    for (const name of missingFromReturn) console.error(`  - ${name}`);
  }
  if (!missingFromHelpers.length && !missingFromReturn.length) {
    console.log(
      `${rel}: OK (${synStoreCallsInTemplate.length} synStore template helper call(s) checked)`
    );
  }
}

process.exit(failed ? 1 : 0);
