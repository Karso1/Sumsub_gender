#!/usr/bin/env node

/**
 * Safely copy user-supplied gender values from an Excel workbook into Sumsub.
 * Dry-run is the default. Use --apply to save updates.
 *
 * The script launches a temporary Chrome context and never reads or stores
 * the user's existing Chrome credentials. Sign in to Sumsub manually.
 */

import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import readline from 'node:readline/promises';
import { stdin, stdout } from 'node:process';
import { fileURLToPath } from 'node:url';
import zlib from 'node:zlib';

const require = createRequire(import.meta.url);
const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const packageRoot = path.resolve(scriptDir, '..');
const localRuntime = path.join(packageRoot, 'runtime');
const depsRoot = fs.existsSync(path.join(localRuntime, 'node_modules/playwright'))
  ? localRuntime
  : packageRoot;
const bundledNodeModules = path.join(depsRoot, 'node_modules');
const playwrightEntry = path.join(bundledNodeModules, 'playwright');
const { chromium } = require(playwrightEntry);

const DEFAULT_INPUT = path.join(packageRoot, 'input', 'workbook.xlsx');
const CLIENT_ID = process.env.SUMSUB_CLIENT_ID;
if (!CLIENT_ID) throw new Error('Set SUMSUB_CLIENT_ID before running.');
const SEARCH_BASE = `https://cockpit.sumsub.com/checkus/applicants/individual?clientId=${encodeURIComponent(CLIENT_ID)}&filter=%7B%22date%22%3A%7B%22m%22%3A%22reviewStarted%22%7D%7D&limit=25&offset=0&search=`;
const ALIASES = new Map([
  ['male', 'Male'], ['m', 'Male'], ['男', 'Male'],
  ['female', 'Female'], ['f', 'Female'], ['女', 'Female'],
  ['x', 'X'], ['non-binary', 'X'], ['nonbinary', 'X'],
]);

function parseArgs(argv) {
  const args = { input: DEFAULT_INPUT, apply: false, limit: Infinity, results: path.resolve('sumsub_gender_results.csv'), checkpoint: path.resolve('sumsub_gender_progress.jsonl') };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--apply') args.apply = true;
    else if (a === '--input') args.input = path.resolve(argv[++i] ?? '');
    else if (a === '--limit') args.limit = Number(argv[++i]);
    else if (a === '--results') args.results = path.resolve(argv[++i] ?? '');
    else if (a === '--checkpoint') args.checkpoint = path.resolve(argv[++i] ?? '');
    else if (a === '--help' || a === '-h') args.help = true;
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!Number.isInteger(args.limit) && args.limit !== Infinity) throw new Error('--limit must be a positive integer');
  if (args.limit < 1) throw new Error('--limit must be a positive integer');
  return args;
}

function loadRows(input) {
  const zip = fs.readFileSync(input);
  const entries = readZipEntries(zip);
  const sharedXml = entries.get('xl/sharedStrings.xml') ?? '';
  const sharedStrings = [...sharedXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m =>
    [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t => decodeXml(t[1])).join(''));
  const sheetPath = [...entries.keys()].filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)).sort()[0];
  if (!sheetPath) throw new Error('No worksheet XML was found in the Excel file.');
  const xml = entries.get(sheetPath);
  const xmlRows = [...xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)];
  if (!xmlRows.length) throw new Error('The first worksheet has no rows.');
  const rowValues = xmlRows.map(rowMatch => {
    const cells = [];
    for (const c of rowMatch[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = c[1].match(/\br="([A-Z]+)\d+"/);
      if (!ref) continue;
      const col = [...ref[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      const kind = c[1].match(/\bt="([^"]+)"/)?.[1] ?? '';
      const value = c[2] ?? '';
      const raw = value.match(/<v\b[^>]*>([\s\S]*?)<\/v>/)?.[1] ?? '';
      let cellValue = '';
      if (kind === 's') cellValue = sharedStrings[Number(raw)] ?? '';
      else if (kind === 'inlineStr') cellValue = [...value.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t => decodeXml(t[1])).join('');
      else cellValue = decodeXml(raw);
      cells[col] = cellValue;
    }
    return cells;
  });
  const header = rowValues[0] ?? [];
  const norm = header.map(v => String(v ?? '').trim().toLowerCase());
  const uid_i = norm.findIndex(v => ['用户id', '用户 id', 'uid', 'user id', 'external id'].includes(v));
  const gender_i = norm.findIndex(v => ['gender', '性别', '真实性别', '真实 性别'].includes(v));
  const rows = [];
  if (uid_i >= 0 && gender_i >= 0) {
    for (let i = 1; i < rowValues.length; i++) {
      const row = rowValues[i];
      const uid = String(row[uid_i] ?? '').trim();
      const gender = String(row[gender_i] ?? '').trim();
      if (uid) rows.push({ uid, gender, sheet: path.basename(sheetPath), row: i + 1 });
    }
  }
  if (!rows.length) throw new Error('Could not find a sheet with both 用户ID and gender columns.');
  return rows;
}

function readZipEntries(buffer) {
  let eocd = -1;
  const start = Math.max(0, buffer.length - 0xffff - 22);
  for (let i = buffer.length - 22; i >= start; i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Excel file is not a supported ZIP-based .xlsx workbook.');
  const count = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const entries = new Map();
  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) throw new Error('Invalid Excel ZIP directory.');
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.subarray(offset + 46, offset + 46 + nameLength).toString('utf8');
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
    let content;
    if (method === 0) content = compressed;
    else if (method === 8) content = zlib.inflateRawSync(compressed);
    else throw new Error(`Unsupported compression method ${method} in ${name}.`);
    entries.set(name, content.toString('utf8'));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

function decodeXml(value) {
  return String(value ?? '')
    .replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'").replaceAll('&amp;', '&')
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}

function normalizeGender(value) {
  return ALIASES.get(String(value ?? '').trim().toLowerCase()) ?? null;
}

function csvCell(value) {
  const s = String(value ?? '');
  return `"${s.replaceAll('"', '""')}"`;
}

function appendResult(file, row) {
  const header = 'timestamp,uid,source_gender,result,detail\n';
  if (!fs.existsSync(file)) fs.writeFileSync(file, header, { mode: 0o600 });
  fs.appendFileSync(file, [new Date().toISOString(), row.uid, row.gender, row.result, row.detail].map(csvCell).join(',') + '\n', { mode: 0o600 });
  fs.chmodSync(file, 0o600);
}

function appendCheckpoint(file, row) {
  fs.appendFileSync(file, JSON.stringify({ timestamp: new Date().toISOString(), uid: row.uid, result: row.result }) + '\n', { mode: 0o600 });
  fs.chmodSync(file, 0o600);
}

function readCompleted(file) {
  if (!fs.existsSync(file)) return new Set();
  return new Set(fs.readFileSync(file, 'utf8').split(/\r?\n/).filter(Boolean).flatMap(line => {
    try { const r = JSON.parse(line); return r.result === 'saved' || r.result === 'already-set' ? [String(r.uid)] : []; }
    catch { return []; }
  }));
}

async function checkedState(locator) {
  const aria = await locator.getAttribute('aria-checked', { timeout: 1500 }).catch(() => null);
  if (aria !== null) return aria === 'true';
  const checked = await locator.getAttribute('checked', { timeout: 1500 }).catch(() => null);
  return checked !== null;
}

async function currentGender(page) {
  const choices = ['Male', 'Female', 'X'];
  let selected = [];
  for (const label of choices) {
    const radio = page.getByRole('radio', { name: label, exact: true });
    if (await radio.count() && await checkedState(radio.first())) selected.push(label);
  }
  if (selected.length > 1) throw new Error(`More than one gender radio selected: ${selected.join(', ')}`);
  return selected[0] ?? null;
}

async function findApplicant(page, uid) {
  await page.goto(SEARCH_BASE + encodeURIComponent(uid), { waitUntil: 'domcontentloaded' });
  const search = page.getByRole('textbox', { name: /table search/i });
  await search.waitFor({ state: 'visible', timeout: 30000 });
  // The search URL already contains this UID. Wait for its result instead of
  // submitting the field again; Sumsub updates this SPA route asynchronously.
  const links = page.locator('a[href*="/checkus/applicant/"]');
  await links.first().waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});
  const hrefs = new Set();
  const linkCount = await links.count();
  for (let i = 0; i < linkCount; i++) {
    const href = await links.nth(i).getAttribute('href');
    if (href) hrefs.add(href);
  }
  if (hrefs.size === 0) return { status: 'not-found', detail: 'No applicant row matched this UID.' };
  if (hrefs.size !== 1) return { status: 'ambiguous', detail: `Search returned ${hrefs.size} distinct applicant records.` };
  const href = [...hrefs][0];
  // Navigate to the exact applicant URL directly. Clicking the Sumsub row can
  // be handled through client-side routing or a new tab, so waitForURL on the
  // search page is unreliable and was causing a 30-second timeout per UID.
  await page.goto(new URL(href, page.url()).toString(), { waitUntil: 'domcontentloaded' });
  const genderRadio = page.getByRole('radio', { name: 'Male', exact: true });
  await genderRadio.waitFor({ state: 'visible', timeout: 15000 });
  const external = page.getByRole('button', { name: /copy external id/i });
  if (!(await external.count())) return { status: 'unverified', detail: 'Applicant opened, but external ID control was not present.' };
  return { status: 'opened' };
}

async function processOne(page, row, apply) {
  const expected = normalizeGender(row.gender);
  if (!expected) return { result: 'invalid-source-gender', detail: `Unrecognized gender value: ${row.gender}` };
  const found = await findApplicant(page, row.uid);
  if (found.status !== 'opened') return { result: found.status, detail: found.detail };
  const existing = await currentGender(page);
  if (existing) {
    if (existing === expected) return { result: 'already-set', detail: existing };
    return { result: 'existing-conflict', detail: `Sumsub already has ${existing}; source says ${expected}; left unchanged.` };
  }
  if (!apply) return { result: 'would-set', detail: `Blank in Sumsub; source value is ${expected}.` };
  const target = page.getByRole('radio', { name: expected, exact: true });
  await target.click();
  if (!(await checkedState(target))) throw new Error(`Failed to select ${expected}`);
  const save = page.getByRole('button', { name: 'Save', exact: true });
  if (await save.count() !== 1 || !(await save.isEnabled())) throw new Error('Save button was not uniquely enabled.');
  await save.click();
  await page.waitForTimeout(400);
  await page.reload({ waitUntil: 'domcontentloaded' });
  await page.getByRole('radio', { name: expected, exact: true }).waitFor({ state: 'visible', timeout: 30000 });
  const after = await currentGender(page);
  if (after !== expected) throw new Error(`Post-save verification failed: expected ${expected}, found ${after ?? 'blank'}.`);
  return { result: 'saved', detail: after };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
  console.log(`Usage: node scripts/sumsub_gender_from_excel.mjs [options]\n\nOptions:\n  --input PATH       Excel input (default: the supplied workbook)\n  --apply            Save changes (default is preview only)\n  --limit N          Process at most N rows\n  --results PATH     CSV audit log\n  --checkpoint PATH  JSONL resume file\n`);
    return;
  }
  if (!fs.existsSync(args.input)) throw new Error(`Input workbook not found: ${args.input}`);
  const rows = loadRows(args.input).slice(0, args.limit);
  const done = readCompleted(args.checkpoint);
  const todo = rows.filter(r => !done.has(r.uid));
  const invalid = rows.filter(r => !normalizeGender(r.gender));
  console.log(`Workbook rows: ${rows.length}; already completed: ${rows.length - todo.length}; queued: ${todo.length}; invalid gender values in this batch: ${invalid.length}.`);
  console.log(args.apply ? 'APPLY MODE: blank Sumsub values will be saved from the workbook.' : 'PREVIEW MODE: no changes will be saved.');
  console.log('A temporary Chrome window will open. Sign in to the intended Sumsub production account manually. The browser session is discarded when the run ends.');

  const browser = await chromium.launch({ channel: 'chrome', headless: false, args: ['--disable-session-crashed-bubble'] });
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  try {
    const loginPage = await context.newPage();
    await loginPage.goto(SEARCH_BASE, { waitUntil: 'domcontentloaded' });
    const rl = readline.createInterface({ input: stdin, output: stdout });
    await rl.question(`After signing in and confirming the Sumsub client is ${CLIENT_ID}, press Enter here to start: `);
    rl.close();
    await loginPage.getByRole('textbox', { name: /table search/i }).waitFor({ state: 'visible', timeout: 30000 });
    if (!(await loginPage.getByText('${CLIENT_ID}', { exact: true }).count())) {
      throw new Error(`Could not verify the Sumsub client ID ${CLIENT_ID} on the page. Stopping without processing rows.`);
    }

    console.log('已登录，开始在当前 Chrome 页面处理；全程复用同一个页面。');
    let saved = 0, skipped = 0, failed = 0;
    for (let i = 0; i < todo.length; i++) {
      const row = todo[i];
      let result;
      const page = loginPage;
      try { result = await processOne(page, row, args.apply); }
      catch (error) { result = { result: 'error', detail: String(error?.message ?? error) }; }
      appendResult(args.results, { ...row, ...result });
      if (result.result === 'saved' || result.result === 'already-set') appendCheckpoint(args.checkpoint, { uid: row.uid, result: result.result });
      if (result.result === 'saved') saved++;
      else if (result.result === 'error') failed++;
      else skipped++;
      console.log(`[${i + 1}/${todo.length}] ${row.uid}: ${result.result} — ${result.detail}`);
      if (i < todo.length - 1) await loginPage.waitForTimeout(150);
    }
    console.log(`Finished this run. Saved: ${saved}; skipped/unchanged: ${skipped}; errors: ${failed}.`);
    console.log(`Audit CSV: ${args.results}`);
    console.log(`Resume checkpoint: ${args.checkpoint}`);
    await loginPage.waitForTimeout(500);
  } finally {
    await context.close().catch(() => {});
    await browser.close();
  }
}

main().catch(error => {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
});
