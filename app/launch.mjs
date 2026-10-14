import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { spawn } from 'node:child_process';
import readline from 'node:readline/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const input = path.join(root, 'input');
const preview = process.argv.includes('--preview');
try {
  fs.mkdirSync(input, { recursive: true });
  const files = fs.readdirSync(input, { withFileTypes: true })
    .filter(e => e.isFile() && !e.name.startsWith('~$') && /\.xlsx$/i.test(e.name))
    .map(e => e.name).sort((a, b) => a.localeCompare(b, 'zh-CN'));
  if (!files.length) throw new Error('input 文件夹里没有 .xlsx 表格。请将 Excel 保存为 .xlsx 后放入 input；文件名不限。');
  let selected = files[0];
  if (files.length > 1) {
    console.log('input 中有多个表格，请选择本次使用的文件：');
    files.forEach((name, i) => console.log(`${i + 1}. ${name}`));
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    try {
      while (true) {
        const answer = (await rl.question('输入文件序号（输入 q 取消）：')).trim();
        if (answer.toLowerCase() === 'q') process.exit(0);
        const n = Number(answer);
        if (Number.isInteger(n) && n >= 1 && n <= files.length) { selected = files[n - 1]; break; }
        console.log('序号无效，请重新输入。');
      }
    } finally { rl.close(); }
  }
  const workbook = path.join(input, selected);
  // Changed workbooks get independent resume records, even when the name is reused.
  const digest = createHash('sha256').update(fs.readFileSync(workbook)).digest('hex').slice(0, 16);
  const output = path.join(root, 'output', `${path.parse(selected).name}-${digest}`);
  fs.mkdirSync(output, { recursive: true });
  console.log(`本次表格：${selected}`);
  console.log(preview ? '预览模式：只检查前 10 条，不保存。' : '正式模式：按表格填写空缺性别并保存。');
  console.log(`结果目录：${output}`);
  const args = [path.join(root, 'app', 'sumsub_gender_from_excel.mjs'), '--input', workbook,
    '--results', path.join(output, preview ? 'preview.csv' : 'results.csv'),
    '--checkpoint', path.join(output, preview ? 'preview-progress.jsonl' : 'progress.jsonl'),
    ...(preview ? ['--limit', '10'] : ['--apply'])];
  const child = spawn(process.execPath, args, { stdio: 'inherit', cwd: root });
  child.on('error', error => { console.error(error.message); process.exitCode = 1; });
  child.on('exit', code => { process.exitCode = code ?? 1; });
} catch (error) { console.error(error.message); process.exitCode = 1; }
