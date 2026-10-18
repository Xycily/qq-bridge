# 文件收发（本仓库新增能力）

> 本文档描述本仓库相对上游 `qq-bridge` 新增的一套「QQ 文件收发 + Office 读写」能力。
> 上游默认只处理文本、图片、表情；文件只会显示成 `[文件 xxx]` 一个占位符。

## 一句话概括

**你发文件给她 → 她存到磁盘、能读内容、能改、能改完发回给你。**

```
QQ 收到文件 ──► <工作区>/qq-files/inbox/    （桥接自动落盘）
                     │
                     ├─ doc_read / workspace_read_file   读内容
                     ├─ doc_edit / doc_create            改 / 新建
                     └─ qq_send_file                     发回 QQ
```

## 目录约定

| 目录 | 用途 |
|---|---|
| `<工作区>/qq-files/inbox/` | 你发来的文件，桥接收到的瞬间就落盘在这里 |
| `<工作区>/qq-files/outbox/` | 她的产出放这里（`doc_create` / `doc_edit` 只允许写这片） |
| `<工作区>/qq-files/work/` | 中转/临时文件 |

「工作区」默认是本仓库的上一级目录；`scripts/qq-send-file.mjs` 可用
`QQ_FILES_WORKSPACE` 环境变量覆盖。

## 收文件

桥接在解析 QQ 消息段时，遇到 `file` 段会：

1. 用 `get_private_file_url` / `get_group_file_url`（按会话类型）换取下载地址；
   这两个动作是 SnowLuma 明确要求的，`get_file` 只解析图片/语音缓存。
2. 下载后写入 `qq-files/inbox/<ISO 时间戳>-<原文件名>`。
3. 在给模型的上下文里写成 `[文件 xxx（已保存到 <绝对路径>；需要看内容可用 workspace_read_file 读取）]`。

下载失败不会中断聊天，只写一行日志：
`[file] 未能保存「xxx」 url=… file_id=…`（不打印完整 URL，避免带出可访问链接）。

排查用的独立日志在 `state/file-debug.log`（绕开日志脱敏）。

## 读文件

| 工具 | 范围 | 说明 |
|---|---|---|
| `workspace_read_file(path, maxChars?)` | 整个工作区 | 纯文本；≤200KB；含 NUL 字节判为二进制拒绝 |
| `workspace_list_dir(path?)` | 整个工作区 | 默认列 `qq-files/inbox` |
| `doc_read(path, maxChars?)` | 整个工作区 | docx / xlsx / pptx / pdf / 文本 |

**敏感文件需要授权**：命中 `config.json`、`.env*`、`*token*`、`*secret*`、
`*password*`、`snowluma/**/config/**` 等特征的路径，会被拒绝并提示先申请授权。

## 写文件

| 工具 | 范围 | 支持 |
|---|---|---|
| `workspace_write_file(path, content, overwrite?)` | 仅 `qq-files/` | 纯文本，≤200KB，默认不覆盖 |
| `doc_edit(path, ops)` | 仅 `qq-files/` | docx `replace`/`append`；xlsx `set`/`append_row`；文本 `overwrite` |
| `doc_create(path, content, overwrite?)` | 仅 `qq-files/` | 新建 docx / xlsx / txt / md / csv |

`doc_create` 的文本按行解析：`# 标题`、`## 小标题`、`- 列表项`、其余为正文；
`.xlsx` 每行一行数据（制表符优先，其次逗号），首行自动加粗。

## 发文件

**给模型用**：`qq_send_file(key, path, name?)`
—— `key` 形如 `private:10001` / `group:123456`，必须在 `config.json` 的 allow 白名单内；
文件必须在 `qq-files/` 下，≤20MB。

**给自己用（命令行）**：

```bash
node scripts/qq-send-file.mjs --to private:10001 --file "qq-files/outbox/报告.docx"
node scripts/qq-send-file.mjs --to group:123456 --file "D:/x.pdf" --name 报告.pdf --allow-outside
```

只允许发工作区内的文件，`--allow-outside` 才放行外部路径。

## 敏感文件授权流程

```
模型读敏感文件被拒
   ↓
调用 request_permission(path, reason)
   ↓
管理员 QQ 私聊收到：【授权申请 AB12】想读取 … 原因 …
   ↓
管理员回复「同意 AB12」或「拒绝 AB12」
   ↓
桥接拦截该回复并更新状态（不进 agent 上下文，群友伪造无效）
   ↓
模型重试读取 → 通过（10 分钟内有效）
```

授权记录存在 `state/perm-requests.json`，每条按**具体路径**授权，不按目录。

## 格式支持矩阵

| 格式 | 读 | 写 | 备注 |
|---|---|---|---|
| `.txt` `.md` `.csv` `.json` | ✅ | ✅ | |
| `.docx` | ✅ | ✅ replace / append | 依赖 python-docx |
| `.xlsx` | ✅ | ✅ set / append_row | 依赖 openpyxl；读的是缓存值 |
| `.pptx` | ✅ | ❌ | 只读 |
| `.pdf` | ⚠️ 仅文本层 | ❌ | 扫描件抽不出文字；不支持就地修改 |
| `.doc` `.xls` `.ppt` | ❌ | ❌ | 旧二进制格式，需 LibreOffice/antiword，本环境未安装 |

**PDF 的常见用法**：`doc_read` 抽出文本 → `doc_create` 生成新的 `.docx` → `qq_send_file` 发回。
版式（分栏、表格线、图片）不会保留，只保留文字顺序。

## 安全边界一览

| 能力 | 边界 | 为什么 |
|---|---|---|
| 读 | 整个工作区 | 敏感文件仍需 QQ 授权 |
| 写 / 新建 | 仅 `qq-files/` | 否则一句「把 config.json 改一下」就能改掉机器人配置 |
| 发送 | 仅 `qq-files/` + 白名单目标 | 防止把系统文件或令牌发进聊天 |
| 出站文本 | `SENSITIVE_RE` 过滤 | 路径与凭据形态的内容会被 `***` 或拦截，授权也绕不过 |

> 出站过滤和读取授权是**两层独立**的机制：授权只解决「能不能读」，
> 「能不能发到 QQ」由出站过滤决定。

## 依赖

| 依赖 | 用途 | 来源 |
|---|---|---|
| python-docx / openpyxl / python-pptx | docx / xlsx / pptx | DSH 自带 Python（`load_workspace_dependencies` 返回的路径） |
| Python 解释器 | 执行 `scripts/office-tool.py` | 优先 `QQ_OFFICE_PYTHON` / `DSH_PYTHON` / `$DSH_HOME/dsh-runtimes/...`，回退 `python` |

PDF 抽取用纯标准库（zlib + 正则），不依赖任何第三方包。

## 相关源码

| 文件 | 作用 |
|---|---|
| `src/bridge.js` | 收文件落盘（`saveIncomingFile`）、授权回复拦截、启动标记 |
| `src/mcp-snowluma-safe.js` | 全部 MCP 工具（读/写/发/授权/Office） |
| `src/perm-store.js` | 授权记录存取与敏感路径判定 |
| `scripts/office-tool.py` | docx/xlsx/pptx/pdf 的读写与新建 |
| `scripts/qq-send-file.mjs` | 命令行发文件 |
