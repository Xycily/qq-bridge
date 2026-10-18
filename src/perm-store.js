// perm-store.js —— QQ 侧的"敏感文件读取授权"存储。
//
// 场景：QQ 会话里的模型只有 workspace_read_file 这一个只读入口。默认它可以读工作区里
// 任何文本文件，但**敏感文件**（config.json、含 token/secret/密码 的文件、SnowLuma 的
// 配置目录等）需要管理员在 QQ 上明确同意。
//
// 流程：
//   1. 模型尝试读敏感文件 → 工具拒绝，并提示先调 request_permission
//   2. request_permission → 生成 4 位授权码，给管理员（ownerQQ）发 QQ 私聊，写明要读什么
//   3. 管理员在 QQ 回复「同意 AB12」或「拒绝 AB12」
//   4. 桥接收到这条回复 → 更新记录 → 模型重试 workspace_read_file 即可通过
//
// 存储：state/perm-requests.json（桥接与 MCP server 两个进程共用同一个文件）。
// 并发：写入用「读-改-写 + 原子 rename」，读方容忍解析失败（返回空表）。
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const STATE_DIR = path.join(ROOT, 'state');
const FILE = path.join(STATE_DIR, 'perm-requests.json');

const TTL_MS = 10 * 60 * 1000; // 授权码 10 分钟有效
const MAX_RECORDS = 50;

// 需要授权才能读的路径特征（大小写不敏感）。命中任一即视为敏感。
const SENSITIVE_PATTERNS = [
  /(^|[\\/])config\.json$/i,
  /(^|[\\/])config\.example\.json$/i,
  /(^|[\\/])\.env(\.[^\\/]*)?$/i,
  /\.credentials[^\\/]*$/i,
  /token/i,
  /secret/i,
  /password/i,
  /passwd/i,
  /cookie/i,
  /(^|[\\/])snowluma[\\/].*[\\/]config[\\/]/i,
  /(^|[\\/])state[\\/](console-token|sessions\.json)$/i
];

export function isSensitivePath(p) {
  const s = String(p ?? '');
  return SENSITIVE_PATTERNS.some((re) => re.test(s));
}

function readAll() {
  try {
    const text = fs.readFileSync(FILE, 'utf8');
    const parsed = JSON.parse(text);
    return Array.isArray(parsed?.requests) ? parsed.requests : [];
  } catch {
    return [];
  }
}

function writeAll(list) {
  try {
    fs.mkdirSync(STATE_DIR, { recursive: true });
    const trimmed = list.slice(-MAX_RECORDS);
    const payload = JSON.stringify({ version: 1, requests: trimmed }, null, 2);
    const tmp = `${FILE}.tmp`;
    try {
      fs.writeFileSync(tmp, payload, 'utf8');
      fs.renameSync(tmp, FILE);
    } catch {
      // Windows 上 rename 覆盖偶发 EPERM（文件被占用/权限），退化为直接覆盖写
      try { fs.unlinkSync(tmp); } catch {}
      fs.writeFileSync(FILE, payload, 'utf8');
    }
  } catch (error) {
    // 只写 stderr：stdio MCP 服务器的日志由调用方（DSH）收集，不污染 JSON-RPC 通道
    try { console.error('[perm-store] 写入失败:', error?.message ?? error); } catch {}
  }
}

function prune(list) {
  const now = Date.now();
  return list.filter((r) => {
    if (!r) return false;
    if (r.status === 'pending') return Number(r.expiresAt ?? 0) > now;
    // 已决的保留 10 分钟供模型重试读取
    return Number(r.decidedAt ?? 0) + TTL_MS > now;
  });
}

function newCode(existing) {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 去掉易混淆的 I/O/0/1
  for (let attempt = 0; attempt < 50; attempt++) {
    let code = '';
    for (let i = 0; i < 4; i++) code += alphabet[crypto.randomInt(alphabet.length)];
    if (!existing.some((r) => r.code === code && r.status === 'pending')) return code;
  }
  return String(Date.now()).slice(-4);
}

/** 新建一条授权申请，返回记录（含 code）。 */
export function createPermissionRequest({ path: targetPath, reason = '', requester = '' } = {}) {
  const list = prune(readAll());
  const code = newCode(list);
  const now = Date.now();
  const record = {
    code,
    path: String(targetPath ?? ''),
    reason: String(reason ?? ''),
    requester: String(requester ?? ''),
    status: 'pending',
    createdAt: now,
    expiresAt: now + TTL_MS,
    decidedAt: 0
  };
  list.push(record);
  writeAll(list);
  return record;
}

/** 指定路径是否已被明确批准（且未过期）。 */
export function isPathApproved(targetPath) {
  const abs = String(targetPath ?? '').toLowerCase();
  const now = Date.now();
  return prune(readAll()).some(
    (r) => r.status === 'approved' && Number(r.decidedAt) + TTL_MS > now && String(r.path ?? '').toLowerCase() === abs
  );
}

/** 按授权码批准/拒绝；返回更新后的记录或 null。 */
export function decideRequest(code, approved) {
  const wanted = String(code ?? '').trim().toUpperCase();
  if (!wanted) return null;
  const list = prune(readAll());
  const rec = list.find((r) => String(r.code).toUpperCase() === wanted && r.status === 'pending');
  if (!rec) return null;
  rec.status = approved ? 'approved' : 'denied';
  rec.decidedAt = Date.now();
  writeAll(list);
  return rec;
}

export function listRequests() {
  return prune(readAll());
}

export const PERM_FILE = FILE;
export const PERM_TTL_MS = TTL_MS;
