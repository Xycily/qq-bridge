#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""office-tool.py —— QQ 桥接的文档读写助手（由 MCP server 以子进程调用）。

用法：
  python office-tool.py read <path> [max_chars]
  python office-tool.py edit <path> <ops_json>

输出：单行 JSON（UTF-8），便于 Node 侧解析。

支持：
  .docx  读：段落 + 表格；写：replace / append
  .xlsx  读：各表单元格；写：set / append_row
  .pptx  读：各页文本（暂不支持写）
  .pdf   读：内置轻量文本抽取（仅未加密、文本型 PDF）
  文本类 读：原样；写：overwrite
"""
import json
import os
import re
import sys
import zlib


def emit(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False))
    sys.stdout.flush()


def fail(msg):
    emit({"ok": False, "error": str(msg)})
    sys.exit(0)


# ── docx ────────────────────────────────────────────────────────────────────
def read_docx(path, max_chars):
    import docx  # python-docx

    d = docx.Document(path)
    out = []
    for p in d.paragraphs:
        if p.text.strip():
            out.append(p.text)
    for ti, table in enumerate(d.tables, 1):
        out.append(f"[表格 {ti}]")
        for row in table.rows:
            out.append(" | ".join(c.text.strip() for c in row.cells))
    text = "\n".join(out)
    return {
        "ok": True,
        "kind": "docx",
        "text": text[:max_chars],
        "truncated": len(text) > max_chars,
        "meta": {"paragraphs": len(d.paragraphs), "tables": len(d.tables)},
    }


def _replace_in_paragraph(p, old, new):
    """优先在单个 run 内替换（保留格式）；跨 run 时重建整段。"""
    for r in p.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
            return 1
    if old not in p.text:
        return 0
    merged = p.text.replace(old, new)
    if p.runs:
        for r in list(p.runs):
            r.text = ""
        p.runs[0].text = merged
    else:
        p.add_run(merged)
    return 1


def edit_docx(path, ops):
    import docx

    d = docx.Document(path)
    applied = 0
    for op in ops:
        kind = (op or {}).get("op")
        if kind == "replace":
            old = str(op.get("old") or "")
            new = str(op.get("new") or "")
            if not old:
                continue
            for p in d.paragraphs:
                applied += _replace_in_paragraph(p, old, new)
            for table in d.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            applied += _replace_in_paragraph(p, old, new)
        elif kind == "append":
            text = str(op.get("text") or "")
            if text:
                d.add_paragraph(text)
                applied += 1
        else:
            raise ValueError(f"不支持的 docx 操作: {kind}")
    d.save(path)
    return {"ok": True, "kind": "docx", "applied": applied}


# ── xlsx ────────────────────────────────────────────────────────────────────
def read_xlsx(path, max_chars):
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    names = list(wb.sheetnames)
    out = []
    for ws in wb.worksheets:
        out.append(f"[工作表 {ws.title}]")
        for row in ws.iter_rows(values_only=True):
            if row is None:
                continue
            cells = ["" if v is None else str(v) for v in row]
            if any(c.strip() for c in cells):
                out.append("\t".join(cells).rstrip())
    wb.close()
    text = "\n".join(out)
    return {
        "ok": True,
        "kind": "xlsx",
        "text": text[:max_chars],
        "truncated": len(text) > max_chars,
        "meta": {"sheets": names},
    }


def edit_xlsx(path, ops):
    import openpyxl

    wb = openpyxl.load_workbook(path)
    applied = 0
    for op in ops:
        kind = (op or {}).get("op")
        sheet = str(op.get("sheet") or wb.sheetnames[0])
        if kind == "set":
            ws = wb[sheet]
            ws[str(op["cell"])] = op.get("value")
            applied += 1
        elif kind == "append_row":
            ws = wb[sheet]
            ws.append(list(op.get("values") or []))
            applied += 1
        else:
            raise ValueError(f"不支持的 xlsx 操作: {kind}")
    wb.save(path)
    return {"ok": True, "kind": "xlsx", "applied": applied}


# ── pptx ────────────────────────────────────────────────────────────────────
def read_pptx(path, max_chars):
    from pptx import Presentation

    prs = Presentation(path)
    out = []
    for i, slide in enumerate(prs.slides, 1):
        out.append(f"[第 {i} 页]")
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                out.append(shape.text.strip())
    text = "\n".join(out)
    return {
        "ok": True,
        "kind": "pptx",
        "text": text[:max_chars],
        "truncated": len(text) > max_chars,
        "meta": {"slides": len(prs.slides)},
    }


# ── pdf（轻量文本抽取，无第三方依赖）─────────────────────────────────────
PDF_STRING_RE = re.compile(rb"\((?:\\.|[^\\()])*\)")


def _pdf_unescape(raw):
    s = raw[1:-1]
    out = bytearray()
    i = 0
    while i < len(s):
        c = s[i]
        if c == 0x5C and i + 1 < len(s):  # 反斜杠转义
            n = s[i + 1]
            mapping = {0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12}
            if n in mapping:
                out.append(mapping[n])
                i += 2
                continue
            if 0x30 <= n <= 0x37:
                oct_digits = s[i + 1:i + 4]
                try:
                    out.append(int(oct_digits, 8) & 0xFF)
                except Exception:
                    pass
                i += 1 + len(oct_digits)
                continue
            out.append(n)
            i += 2
            continue
        out.append(c)
        i += 1
    try:
        return out.decode("utf-8")
    except Exception:
        return out.decode("latin-1", "ignore")


def read_pdf(path, max_chars):
    with open(path, "rb") as f:
        data = f.read()
    if b"/Encrypt" in data:
        return {"ok": False, "error": "PDF 已加密，无法抽取文本"}
    texts = []
    for m in re.finditer(rb"stream\r?\n", data):
        start = m.end()
        end = data.find(b"endstream", start)
        if end < 0:
            continue
        chunk = data[start:end]
        raw = None
        for candidate in (chunk, chunk.rstrip(b"\r\n")):
            try:
                raw = zlib.decompress(candidate)
                break
            except Exception:
                raw = None
        if not raw:
            continue
        parts = [_pdf_unescape(sm.group(0)) for sm in PDF_STRING_RE.finditer(raw)]
        line = "".join(parts)
        if line.strip():
            texts.append(line)
    text = "\n".join(texts)
    if not text.strip():
        return {"ok": False, "error": "未能从该 PDF 抽取到文本（可能是扫描件/图片型 PDF）"}
    return {
        "ok": True,
        "kind": "pdf",
        "text": text[:max_chars],
        "truncated": len(text) > max_chars,
        "meta": {"note": "轻量抽取，复杂排版可能丢失结构"},
    }


# ── 入口 ────────────────────────────────────────────────────────────────────
# ── 新建文档（内容从 stdin 读，避免命令行长度限制）──────────────────────
def _read_stdin_text():
    data = sys.stdin.buffer.read()
    for enc in ("utf-8", "utf-8-sig", "gb18030"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "ignore")


def create_docx(path, text):
    """按行生成 docx：# 标题 / ## 小标题 / 其它为正文 / 空行分段。"""
    import docx

    d = docx.Document()
    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("## "):
            d.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            d.add_heading(line[2:].strip(), level=1)
        elif line.startswith("- "):
            d.add_paragraph(line[2:].strip(), style="List Bullet")
        else:
            d.add_paragraph(line)
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    d.save(path)
    return {"ok": True, "kind": "docx", "path": path, "paragraphs": len(d.paragraphs)}


def create_xlsx(path, text):
    """按行生成 xlsx：制表符优先，其次逗号；第一行当表头加粗。"""
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    rows = 0
    for raw in text.splitlines():
        if not raw.strip():
            continue
        cells = raw.split("\t") if "\t" in raw else [c.strip() for c in raw.split(",")]
        ws.append([_coerce(v) for v in cells])
        rows += 1
    for cell in ws[1]:
        cell.font = Font(bold=True)
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb.save(path)
    return {"ok": True, "kind": "xlsx", "path": path, "rows": rows}


def _coerce(v):
    s = str(v)
    try:
        if re.fullmatch(r"-?\d+", s):
            return int(s)
        if re.fullmatch(r"-?\d+\.\d+", s):
            return float(s)
    except Exception:
        pass
    return s


def create_file(path, text):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return create_docx(path, text)
    if ext == ".xlsx":
        return create_xlsx(path, text)
    if ext in (".txt", ".md", ".csv", ".json"):
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return {"ok": True, "kind": "text", "path": path, "bytes": len(text.encode("utf-8"))}
    return {"ok": False, "error": f"不支持新建 {ext}；支持 .docx / .xlsx / .txt / .md / .csv"}


# ── 图片（交给多模态模型看）──────────────────────────────────────────────
def read_image_for_model(path, max_edge=1024):
    """把图片读成 base64，供 MCP 以 image 内容块返回给模型。

    超过 max_edge 会等比缩小并重编码为 JPEG，控制 token 消耗；
    带透明通道的 PNG 仍转 JPEG（模型不需要 alpha）。
    """
    try:
        from PIL import Image
    except Exception as e:
        return {"ok": False, "error": f"缺少 Pillow：{e}"}
    try:
        with Image.open(path) as im:
            w, h = im.size
            im = im.convert("RGB")
            scale = min(1.0, float(max_edge) / max(w, h))
            if scale < 1.0:
                im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
            import io as _io
            buf = _io.BytesIO()
            im.save(buf, format="JPEG", quality=82, optimize=True)
            data = buf.getvalue()
        import base64 as _b64
        return {
            "ok": True,
            "mimeType": "image/jpeg",
            "data": _b64.b64encode(data).decode("ascii"),
            "width": im.size[0],
            "height": im.size[1],
            "srcWidth": w,
            "srcHeight": h,
            "bytes": len(data),
        }
    except Exception as e:
        return {"ok": False, "error": f"读取图片失败：{e}"}


# ── doc/xls/ppt 旧格式转换（调 Word/Excel/PowerPoint COM）────────────────
def _convert_with_com(path, ext, workdir):
    """用 PowerShell + COM 把旧格式转成新格式，返回转换后的路径（失败返回 None）。"""
    import subprocess

    target_ext = {".doc": ".docx", ".xls": ".xlsx", ".ppt": ".pptx"}.get(ext)
    if not target_ext:
        return None
    os.makedirs(workdir, exist_ok=True)
    out = os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + target_ext)
    if os.path.exists(out):
        try:
            os.remove(out)
        except Exception:
            pass
    ps1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "convert-legacy.ps1")
    if not os.path.isfile(ps1):
        return None
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", ps1, "-InPath", path, "-OutPath", out],
            capture_output=True, text=True, timeout=180,
        )
        if r.returncode == 0 and os.path.isfile(out) and os.path.getsize(out) > 0:
            return out
        return None
    except Exception:
        return None


def convert_legacy(path):
    """旧格式入口：转成新格式，并回传转换后的路径。"""
    ext = os.path.splitext(path)[1].lower()
    workdir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(path))), "work", "_converted")
    converted = _convert_with_com(path, ext, workdir)
    if converted:
        return {"ok": True, "kind": "conversion", "converted": converted, "from": path}
    return {"ok": False, "error": f"无法转换 {ext}：本机没有可用的 Word/Excel/PowerPoint COM，或转换失败"}


def main():
    if len(sys.argv) < 3:
        fail("用法: office-tool.py read <path> [max_chars] | edit <path> <ops_json>")
    cmd = sys.argv[1].lower()
    path = sys.argv[2]
    # create 用于新建文件，其余命令要求目标已存在
    if cmd != "create" and not os.path.isfile(path):
        fail(f"文件不存在: {path}")
    ext = os.path.splitext(path)[1].lower()

    if cmd == "read":
        max_chars = int(sys.argv[3]) if len(sys.argv) > 3 else 20000
        # 旧格式（.doc/.xls/.ppt）先转成新格式再读
        if ext in (".doc", ".xls", ".ppt"):
            conv = _convert_with_com(path, ext, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(path))), "work", "_converted"))
            if not conv:
                return fail(f"无法读取 {ext}：需要本机安装 Word/Excel/PowerPoint（或 WPS）用于转换；也可以先手动另存为 {ext}x 再发")
            path = conv
            ext = os.path.splitext(path)[1].lower()
        if ext == ".docx":
            return emit(read_docx(path, max_chars))
        if ext in (".xlsx", ".xlsm"):
            return emit(read_xlsx(path, max_chars))
        if ext == ".pptx":
            return emit(read_pptx(path, max_chars))
        if ext == ".pdf":
            return emit(read_pdf(path, max_chars))
        try:
            with open(path, "r", encoding="utf-8") as f:
                text = f.read()
        except UnicodeDecodeError:
            return fail("不是 UTF-8 文本文件")
        return emit({"ok": True, "kind": "text", "text": text[:max_chars], "truncated": len(text) > max_chars})

    if cmd == "edit":
        if len(sys.argv) < 4:
            fail("edit 需要 ops_json 参数")
        try:
            ops = json.loads(sys.argv[3])
        except Exception as e:
            return fail(f"ops_json 解析失败: {e}")
        if isinstance(ops, dict):
            ops = [ops]
        if not isinstance(ops, list):
            return fail("ops 必须是数组或对象")
        if ext == ".docx":
            return emit(edit_docx(path, ops))
        if ext in (".xlsx", ".xlsm"):
            return emit(edit_xlsx(path, ops))
        if ext == ".pptx":
            return fail("暂不支持修改 pptx（只读）")
        if ext == ".pdf":
            return fail("PDF 不支持就地修改；可读出文本后另存为 docx/txt")
        for op in ops:
            if (op or {}).get("op") == "overwrite":
                with open(path, "w", encoding="utf-8") as f:
                    f.write(str(op.get("content") or ""))
                return emit({"ok": True, "kind": "text", "applied": 1})
        return fail("文本文件只支持 {op: overwrite, content: ...}")

    if cmd == "create":
        # python office-tool.py create <out_path>  （内容从 stdin 读）
        text = _read_stdin_text()
        if not text.strip():
            return fail("stdin 里没有内容")
        return emit(create_file(path, text))

    if cmd == "convert":
        return emit(convert_legacy(path))

    if cmd == "image":
        # python office-tool.py image <path> [max_edge]
        max_edge = int(sys.argv[3]) if len(sys.argv) > 3 else 1024
        return emit(read_image_for_model(path, max_edge))

    fail(f"未知命令: {cmd}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        fail(exc)
