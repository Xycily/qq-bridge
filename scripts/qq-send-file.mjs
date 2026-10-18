#!/usr/bin/env node
// qq-send-file.mjs —— 通过 OneBot 把本地文件发到 QQ（私聊或群）。
//
// 用法：
//   node scripts/qq-send-file.mjs --to private:10001 --file "qq-files/outbox/a.docx"
//   node scripts/qq-send-file.mjs --to group:123456 --file "D:\x.pdf" --name 报告.pdf
//
// 说明：
// - 只允许发送「工作区」内的文件（工作区 = 本仓库的上一级目录，可用环境变量
//   QQ_FILES_WORKSPACE 覆盖），防止被诱导把系统文件发出去。确实要发外部文件时
//   加 --allow-outside。
// - 发送用 OneBot 的 send_private_msg / send_group_msg，消息段 type=file，
//   这是 SnowLuma 支持的扩展段（实测可用：日志 highway upload done）。
// - 配置取自 qq-bridge/config.json 的 snowluma.httpUrl / accessToken，无需额外填写。
import { readFileSync, existsSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(__dirname, '..');
// 工作区默认是仓库的上一级目录（qq-files/ 就在那里），可用环境变量覆盖。
const WORKSPACE = process.env.QQ_FILES_WORKSPACE
  ? path.resolve(process.env.QQ_FILES_WORKSPACE)
  : path.resolve(REPO_ROOT, '..');
const CONFIG = path.join(REPO_ROOT, 'config.json');

function usage(code = 1) {
  console.error('用法: node scripts/qq-send-file.mjs --to <private:QQ|group:群号> --file <路径> [--name <显示名>]');
  process.exit(code);
}

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--to') out.to = argv[++i];
    else if (a === '--file') out.file = argv[++i];
    else if (a === '--name') out.name = argv[++i];
    else if (a === '--allow-outside') out.allowOutside = true;
    else if (a === '-h' || a === '--help') usage(0);
  }
  return out;
}

const args = parseArgs(process.argv.slice(2));
if (!args.to || !args.file) usage();

const m = args.to.match(/^(private|group)[:：](.+)$/i);
if (!m) {
  console.error(`--to 格式不对: ${args.to}（应为 private:QQ号 或 group:群号）`);
  process.exit(1);
}
const kind = m[1].toLowerCase();
const id = Number(m[2]);
if (!Number.isInteger(id) || id <= 0) {
  console.error(`--to 里的号码不合法: ${m[2]}`);
  process.exit(1);
}

const abs = path.resolve(args.file);
if (!existsSync(abs)) {
  console.error(`文件不存在: ${abs}`);
  process.exit(1);
}
const st = statSync(abs);
if (!st.isFile()) {
  console.error(`不是普通文件: ${abs}`);
  process.exit(1);
}

// 工作区边界检查
const rel = path.relative(WORKSPACE, abs);
const inside = rel && !rel.startsWith('..') && !path.isAbsolute(rel);
if (!inside && !args.allowOutside) {
  console.error(`拒绝发送工作区之外的文件: ${abs}`);
  console.error(`工作区是: ${WORKSPACE}`);
  console.error('（确实要发就加 --allow-outside，但请先确认这个文件不敏感）');
  process.exit(1);
}

if (!existsSync(CONFIG)) {
  console.error(`找不到桥接配置: ${CONFIG}`);
  process.exit(1);
}
const cfg = JSON.parse(readFileSync(CONFIG, 'utf8'));
const base = String(cfg?.snowluma?.httpUrl ?? '').replace(/\/+$/, '');
const token = String(cfg?.snowluma?.accessToken ?? '');
if (!base) {
  console.error('config.json 里没有 snowluma.httpUrl');
  process.exit(1);
}

const action = kind === 'private' ? 'send_private_msg' : 'send_group_msg';
const displayName = args.name || path.basename(abs);
const body = kind === 'private'
  ? { user_id: id, message: [{ type: 'file', data: { file: abs, name: displayName } }] }
  : { group_id: id, message: [{ type: 'file', data: { file: abs, name: displayName } }] };

console.log(`发送 ${abs}`);
console.log(`  -> ${kind === 'private' ? '私聊' : '群'} ${id}  文件名: ${displayName}  大小: ${st.size} 字节`);

const ctrl = new AbortController();
const timer = setTimeout(() => ctrl.abort(), 120000);
try {
  const res = await fetch(`${base}/${action}`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      ...(token ? { authorization: `Bearer ${token}` } : {}),
    },
    body: JSON.stringify(body),
    signal: ctrl.signal,
  });
  const text = await res.text();
  let json = {};
  try { json = JSON.parse(text); } catch {}
  if (!res.ok || json.status !== 'ok' || json.retcode !== 0) {
    console.error(`发送失败 HTTP ${res.status}: ${json.wording || text.slice(0, 300)}`);
    process.exitCode = 1;
  } else {
    console.log(`发送成功 message_id=${json?.data?.message_id ?? '(无)'}`);
  }
} catch (e) {
  console.error('发送出错:', e.message);
  process.exitCode = 1;
} finally {
  clearTimeout(timer);
}
