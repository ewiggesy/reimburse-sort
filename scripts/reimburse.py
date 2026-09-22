# -*- coding: utf-8 -*-
"""发票扫描、分类、统计、整理、打印工具。

用法:
  python reimburse.py scan <文件夹> [--json out.json]
      扫描文件夹内所有发票/票据文件，输出分类与金额明细（CSV 风格打印，可选存 JSON）。
  python reimburse.py organize <文件夹> [--out 输出目录] [--move]
      按类别复制（默认）或移动文件到 输出目录/<类别>/，并生成 报销统计.md 与 报销明细.csv。
  python reimburse.py print <文件或目录> [--printer 打印机名]
      调用 Windows 默认打印（PDF 用默认关联程序静默打印，图片用 Shell 打印动词）。

类别: 交通 / 会议 / 图书 / 顺丰快递 / 其他
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys

CATEGORIES = ["交通", "会议", "图书", "顺丰快递", "其他"]

AMOUNT_PATTERNS = [
    r"价税合计[（(]大写[)）].*?[（(]小写[)）]\s*[¥￥]?\s*([0-9,]+\.\d{2})",
    r"[（(]小写[)）]\s*[¥￥]?\s*([0-9,]+\.\d{2})",
    r"价税合计[^¥￥\d\n]{0,20}[¥￥]\s*([0-9,]+\.\d{2})",
    r"金\s*额\s*合\s*计[^。\n]{0,40}?[¥￥]\s*([0-9,]+\.\d{2})",
    r"合\s*计\s*[¥￥]\s*([0-9,]+\.\d{2})",
]
FALLBACK_PATTERN = r"[¥￥]\s*([0-9,]+\.\d{2})"
INVOICE_NO_PATTERN = r"(?:发票|票据)号码[::]?\s*([0-9]{8,})"
DATE_PATTERN = r"开票日期[::]?\s*(\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2})"

# (类别, 关键词列表) —— 按优先级匹配：文件夹名 -> 文件名 -> 内容
CATEGORY_KEYWORDS = {
    "交通": ["交通", "滴滴", "出租", "打车", "高铁", "火车", "动车", "机票"],
    "顺丰快递": ["顺丰"],
    "图书": ["图书", "书店", "新华", "京东图书", "digital"],
    "会议": ["会议", "会务", "注册费", "答辩", "会费", "研讨会", "会展", "学术"],
}


def extract_text(pdf_path):
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def parse_amount(text):
    for pat in AMOUNT_PATTERNS:
        m = re.findall(pat, text, re.S)
        if m:
            return float(m[-1].replace(",", ""))
    m = re.findall(FALLBACK_PATTERN, text)
    return float(m[-1].replace(",", "")) if m else None


def parse_invoice_no(text):
    m = re.findall(INVOICE_NO_PATTERN, text)
    return m[0] if m else None


def parse_date(text):
    m = re.findall(DATE_PATTERN, text)
    return m[0] if m else None


def filename_amount(name):
    """从文件名提取金额，如 20260823120440_2600.00_浙江大学.pdf、付款记录2000.jpg"""
    m = re.findall(r"[_-]([0-9]+(?:\.[0-9]{1,2})?)(?=[_.\s]|元|$)", name)
    if not m:
        return None
    try:
        return float(m[-1])
    except ValueError:
        return None


def classify(folder_hint, filename, text):
    """文件夹名 -> 文件名 -> 内容关键词。注意：会议文件夹内的火车票/机票跟随会议计入
    "会议"总额（用户口径），不单独归入交通；交通类只放市内交通（滴滴等）和独立交通文件夹。"""
    for hay in (folder_hint or "", filename or "", (text or "")[:600]):
        for cat, kws in CATEGORY_KEYWORDS.items():
            if any(k in hay for k in kws):
                return cat
    return "其他"


def meeting_expense_type(folder_hint, filename, text):
    """会议类发票细分：会务费 / 住宿费 / 其他费用（对应报销单栏目）。"""
    hay = " ".join([folder_hint or "", filename or "", (text or "")[:1500]])
    if "住宿" in hay:
        return "住宿费"
    if any(k in hay for k in ("会务", "注册", "会费", "会议费")):
        return "会务费"
    return "其他费用"


MANUAL_CSV = "纸质发票.csv"


def load_manual_entries(root):
    """读取 <文件夹>/纸质发票.csv：列 = 类别,分组,费用类型,金额,备注（分组/费用类型/备注可留空）。
    用于登记没有电子版的纸质发票，金额并入统计。"""
    path = os.path.join(root, MANUAL_CSV)
    entries = []
    if not os.path.exists(path):
        return entries
    import csv
    with open(path, encoding="utf-8-sig", newline="") as f:
        for i, row in enumerate(csv.reader(f), 1):
            row = [c.strip() for c in row]
            if not row or not any(row) or row[0].startswith("#") or row[0] == "类别":
                continue
            cat = row[0] if row[0] in CATEGORIES else "其他"
            group = row[1] if len(row) > 1 and row[1] else "(纸质发票)"
            etype = row[2] if len(row) > 2 else ""
            try:
                amount = float(row[3]) if len(row) > 3 and row[3] else None
            except ValueError:
                amount = None
            note = row[4] if len(row) > 4 else ""
            if amount is None:
                print(f"警告: {MANUAL_CSV} 第{i}行金额无效，已跳过: {row}")
                continue
            entries.append({"path": path, "rel": f"{MANUAL_CSV}#{i}", "filename": f"{MANUAL_CSV} 第{i}行",
                            "ext": ".csv", "folder_hint": "", "category": cat,
                            "expense_type": etype if cat == "会议" else "", "group": group,
                            "amount": amount, "invoice_no": None, "date": None,
                            "kind": "manual", "note": (note + " " if note else "") + "纸质发票（手工登记）"})
    return entries


def scan_folder(root):
    """返回明细列表: dict(path, rel, category, amount, invoice_no, date, kind)"""
    import pdfplumber  # noqa: F401  (提前确认依赖可用)
    records = []
    seen_invoice_nos = {}
    for dirpath, _dirs, files in os.walk(root):
        for fn in sorted(files):
            if dirpath == root and fn == MANUAL_CSV:
                continue  # 数据文件，不当普通文件处理
            path = os.path.join(dirpath, fn)
            rel = os.path.relpath(path, root)
            if fn.lower().endswith((".bat", ".cmd")) or "分类整理" in rel.split(os.sep):
                continue  # 批处理与历史输出目录（含其子目录），跳过，避免重复扫描/递归
            ext = os.path.splitext(fn)[1].lower()
            folder_hint = os.path.basename(dirpath) if os.path.abspath(dirpath) != os.path.abspath(root) else ""
            rec = {"path": path, "rel": rel, "filename": fn, "ext": ext,
                   "folder_hint": folder_hint, "category": "其他",
                   "expense_type": "", "group": folder_hint or "(根目录)",
                   "amount": None, "invoice_no": None, "date": None,
                   "kind": "unknown", "note": ""}
            if ext == ".pdf":
                text = None
                try:
                    text = extract_text(path)
                    rec["amount"] = parse_amount(text)
                    rec["invoice_no"] = parse_invoice_no(text)
                    rec["date"] = parse_date(text)
                    rec["kind"] = "invoice" if rec["amount"] is not None else "document"
                    if rec["kind"] == "document":
                        rec["note"] = "非金额票据（通知/说明等）"
                    rec["category"] = classify(folder_hint, fn, text)
                except Exception as e:
                    rec["note"] = f"PDF解析失败: {e}"
                    rec["category"] = classify(folder_hint, fn, None)
                if rec["category"] == "会议" and rec["kind"] == "invoice":
                    rec["expense_type"] = meeting_expense_type(folder_hint, fn, text)
            elif ext == ".ofd":
                rec["kind"] = "ofd"
                rec["note"] = "OFD格式，未解析金额"
                rec["category"] = classify(folder_hint, fn, None)
            elif ext in (".jpg", ".jpeg", ".png", ".bmp"):
                rec["kind"] = "image"
                rec["amount"] = filename_amount(fn)
                rec["note"] = "图片附件（付款凭证等）"
                rec["category"] = classify(folder_hint, fn, None)
            elif ext == ".zip":
                rec["kind"] = "zip"
                rec["note"] = "压缩包"
                rec["category"] = classify(folder_hint, fn, None)
            else:
                rec["note"] = "未识别类型"
            records.append(rec)
    # 去重：相同发票号码只保留第一个（避免 pdf/zip/ofd 重复计费）
    final = []
    for rec in records:
        no = rec["invoice_no"]
        if no and rec["kind"] == "invoice":
            if no in seen_invoice_nos:
                rec["note"] = (rec["note"] + " " if rec["note"] else "") + f"与 {seen_invoice_nos[no]} 发票号码重复，不计入合计"
                rec["kind"] = "duplicate"
            else:
                seen_invoice_nos[no] = rec["rel"]
        final.append(rec)
    final.extend(load_manual_entries(root))
    return final


def _stats(records):
    """返回 (类别合计dict, 分组合计dict)。分组 = (类别, 子文件夹)；会议类再细分费用类型。"""
    totals = {c: 0.0 for c in CATEGORIES}
    counts = {c: 0 for c in CATEGORIES}
    groups = {}
    for r in records:
        if r["kind"] not in ("invoice", "manual") or r["amount"] is None:
            continue
        c = r["category"]
        totals[c] += r["amount"]
        counts[c] += 1
        key = (c, r["group"])
        g = groups.setdefault(key, {"count": 0, "amount": 0.0,
                                    "会务费": 0.0, "住宿费": 0.0, "其他费用": 0.0})
        g["count"] += 1
        g["amount"] += r["amount"]
        if r["expense_type"]:
            g[r["expense_type"]] += r["amount"]
    return totals, counts, groups


def _print_stats(totals, counts, groups):
    print(f"{'类别':<8}{'张数':>6}{'金额合计':>14}")
    for c in CATEGORIES:
        print(f"{c:<10}{counts[c]:>6}{totals[c]:>14.2f}")
    print(f"{'总计':<10}{sum(counts.values()):>6}{sum(totals.values()):>14.2f}")
    print()
    print(f"{'类别':<8}{'分组（子文件夹）':<28}{'张数':>6}{'金额合计':>14}")
    for (c, gname), g in groups.items():
        print(f"{c:<10}{gname:<30}{g['count']:>6}{g['amount']:>14.2f}")
    meet = [(k, g) for k, g in groups.items() if k[0] == "会议"]
    if meet:
        print()
        print("会议费用细分（对应报销单栏目）:")
        print(f"{'会议':<28}{'会务费':>12}{'住宿费':>12}{'其他费用':>12}{'小计':>12}")
        for (c, gname), g in meet:
            print(f"{gname:<30}{g['会务费']:>12.2f}{g['住宿费']:>12.2f}{g['其他费用']:>12.2f}{g['amount']:>12.2f}")


def cmd_scan(a):
    records = scan_folder(a.folder)
    totals, counts, groups = _stats(records)
    _print_stats(totals, counts, groups)
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        print(f"\n明细已保存: {a.json}")


def cmd_organize(a):
    records = scan_folder(a.folder)
    out = a.out or os.path.join(a.folder, "分类整理")
    os.makedirs(out, exist_ok=True)
    lines = ["| 原路径 | 类别 | 费用类型 | 金额 | 发票号码 | 日期 | 备注 |",
             "|---|---|---|---|---|---|---|"]
    totals, counts, groups = _stats(records)
    for r in records:
        if r["kind"] == "manual":
            amt = f"{r['amount']:.2f}"
            lines.append(f"| {r['rel']} | {r['category']} | {r['expense_type']} | {amt} |  |  | {r['note']} |")
            continue  # 纸质登记不复制文件
        # 输出结构: <类别>/<原子文件夹>/<文件>；根目录文件直接放 <类别>/ 下
        dest_dir = os.path.join(out, r["category"], r["folder_hint"]) if r["folder_hint"] \
            else os.path.join(out, r["category"])
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, r["filename"])
        if os.path.abspath(dest) == os.path.abspath(r["path"]):
            continue
        if os.path.exists(dest):
            base, ext = os.path.splitext(r["filename"])
            i = 1
            while os.path.exists(dest):
                dest = os.path.join(dest_dir, f"{base}_{i}{ext}")
                i += 1
        if a.move:
            shutil.move(r["path"], dest)
        else:
            shutil.copy2(r["path"], dest)
        amt = f"{r['amount']:.2f}" if r["amount"] is not None else ""
        lines.append(f"| {r['rel']} | {r['category']} | {r['expense_type']} | {amt} | "
                     f"{r['invoice_no'] or ''} | {r['date'] or ''} | {r['note']} |")
    # 写报告
    md = ["# 报销统计", "", f"来源文件夹: {os.path.abspath(a.folder)}", "",
          "## 类别合计", "", "| 类别 | 张数 | 金额合计（元） |", "|---|---|---|"]
    for c in CATEGORIES:
        md.append(f"| {c} | {counts[c]} | {totals[c]:.2f} |")
    md.append(f"| **总计** | **{sum(counts.values())}** | **{sum(totals.values()):.2f}** |")
    md += ["", "## 分组合计（不合并，每个子文件夹单独一行）", "",
           "| 类别 | 分组（子文件夹） | 张数 | 金额合计（元） |", "|---|---|---|---|"]
    for (c, gname), g in groups.items():
        md.append(f"| {c} | {gname} | {g['count']} | {g['amount']:.2f} |")
    meet = [(k, g) for k, g in groups.items() if k[0] == "会议"]
    if meet:
        md += ["", "## 会议费用细分（对应报销单栏目：会务费 / 住宿费 / 其他费用）", "",
               "| 会议 | 会务费 | 住宿费 | 其他费用 | 小计 |", "|---|---|---|---|---|"]
        for (c, gname), g in meet:
            md.append(f"| {gname} | {g['会务费']:.2f} | {g['住宿费']:.2f} | "
                      f"{g['其他费用']:.2f} | {g['amount']:.2f} |")
    md += ["", "## 文件明细", ""] + lines + [""]
    with open(os.path.join(out, "报销统计.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(md))
    import csv
    with open(os.path.join(out, "报销明细.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["原路径", "类别", "分组", "费用类型", "金额", "发票号码", "日期", "备注"])
        for r in records:
            w.writerow([r["rel"], r["category"], r["group"], r["expense_type"],
                        r["amount"] if r["amount"] is not None else "",
                        r["invoice_no"] or "", r["date"] or "", r["note"]])
    print(f"整理完成: {os.path.abspath(out)}")
    print()
    _print_stats(totals, counts, groups)


def _gdi_print(images, job_name="发票打印", landscape=False):
    """用 ctypes 调 winspool/gdi32 直接把位图发给默认打印机。不依赖任何阅读器，
    无对话框。images: PIL.Image 列表，一页一张。landscape=True 时把纸张设为横向
    （内容需调用方先旋转好，这样横着拿纸字是正的）。"""
    import ctypes
    from ctypes import wintypes

    gdi = ctypes.windll.gdi32
    spool = ctypes.WinDLL("winspool.drv")  # windll.winspool.drv 的名称解析有坑，必须显式加载

    class DOCINFOW(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.INT), ("lpszDocName", wintypes.LPCWSTR),
                    ("lpszOutput", wintypes.LPCWSTR), ("lpszDatatype", wintypes.LPCWSTR),
                    ("fwType", wintypes.DWORD)]

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", ctypes.c_long),
                    ("biHeight", ctypes.c_long), ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", ctypes.c_long),
                    ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD)]

    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 1)]

    gdi.CreateDCW.restype = wintypes.HDC
    gdi.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPVOID]
    gdi.StartDocW.argtypes = [wintypes.HDC, ctypes.POINTER(DOCINFOW)]
    gdi.StartPage.argtypes = [wintypes.HDC]
    gdi.EndPage.argtypes = [wintypes.HDC]
    gdi.EndDoc.argtypes = [wintypes.HDC]
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    gdi.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
    gdi.StretchDIBits.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                  ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                  ctypes.c_int, wintypes.LPCVOID,
                                  ctypes.POINTER(BITMAPINFO), wintypes.UINT, wintypes.DWORD]
    spool.GetDefaultPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]

    size = wintypes.DWORD(260)
    buf = ctypes.create_unicode_buffer(260)
    if not spool.GetDefaultPrinterW(buf, ctypes.byref(size)):
        raise RuntimeError("系统没有默认打印机，请先在 Windows 设置里添加/选定打印机")
    printer = buf.value
    devmode = None
    hprinter = None
    if landscape:
        # 取打印机默认 DEVMODE，改 dmOrientation=2(横向)，让 DC 按横向纸张工作
        spool.OpenPrinterW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(wintypes.HANDLE), wintypes.LPVOID]
        spool.ClosePrinter.argtypes = [wintypes.HANDLE]
        spool.DocumentPropertiesW.argtypes = [wintypes.HWND, wintypes.HANDLE, wintypes.LPCWSTR,
                                              wintypes.LPVOID, wintypes.LPVOID, wintypes.DWORD]
        if not spool.OpenPrinterW(printer, ctypes.byref(hprinter := wintypes.HANDLE()), None):
            raise RuntimeError(f"无法打开打印机: {printer}")
        dm_size = spool.DocumentPropertiesW(None, hprinter, printer, None, None, 0)
        if dm_size > 0:
            dm_buf = ctypes.create_string_buffer(dm_size)
            # DM_OUT_BUFFER=2
            if spool.DocumentPropertiesW(None, hprinter, printer, dm_buf, None, 2) == 1:
                fields = int.from_bytes(dm_buf.raw[72:76], "little")
                dm_buf.raw = dm_buf.raw[:72] + (fields | 0x00000001).to_bytes(4, "little") \
                    + (2).to_bytes(2, "little") + dm_buf.raw[78:]  # dmFields|=DM_ORIENTATION; dmOrientation=横向
                devmode = dm_buf
        spool.ClosePrinter(hprinter)
    hdc = gdi.CreateDCW("WINSPOOL", printer, None, devmode)
    if not hdc:
        raise RuntimeError(f"无法连接打印机: {printer}")
    doc = DOCINFOW(ctypes.sizeof(DOCINFOW), job_name, None, None, 0)
    if gdi.StartDocW(hdc, ctypes.byref(doc)) <= 0:
        gdi.DeleteDC(hdc)
        raise RuntimeError("StartDoc 失败，打印机可能脱机或未响应")
    try:
        page_w = gdi.GetDeviceCaps(hdc, 8)   # HORZRES
        page_h = gdi.GetDeviceCaps(hdc, 10)  # VERTRES
        margin = int(min(page_w, page_h) * 0.04)
        max_w, max_h = page_w - 2 * margin, page_h - 2 * margin
        for img in images:
            iw, ih = img.size
            scale = min(max_w / iw, max_h / ih)
            dw, dh = max(1, int(iw * scale)), max(1, int(ih * scale))
            x, y = (page_w - dw) // 2, (page_h - dh) // 2
            # BMP 每行必须 4 字节对齐，PIL 输出是紧凑排列，需手动补零，否则越界读崩溃
            raw = img.convert("RGB").tobytes("raw", "BGR")
            row_size = (iw * 3 + 3) & ~3
            pad = row_size - iw * 3
            if pad:
                data = b"".join(raw[i * iw * 3:(i + 1) * iw * 3] + b"\x00" * pad
                                for i in range(ih))
            else:
                data = raw
            bmi = BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.bmiHeader.biWidth = iw
            bmi.bmiHeader.biHeight = -ih  # 负值=自上而下
            bmi.bmiHeader.biPlanes = 1
            bmi.bmiHeader.biBitCount = 24
            bmi.bmiHeader.biCompression = 0
            bmi.bmiHeader.biSizeImage = len(data)
            pixels = ctypes.create_string_buffer(data, len(data))
            if gdi.StartPage(hdc) <= 0:
                raise RuntimeError("StartPage 失败")
            gdi.StretchDIBits(hdc, x, y, dw, dh, 0, 0, iw, ih, pixels,
                              ctypes.byref(bmi), 0, 0x00CC0020)  # SRCCOPY
            if gdi.EndPage(hdc) <= 0:
                raise RuntimeError("EndPage 失败")
    finally:
        gdi.EndDoc(hdc)
        gdi.DeleteDC(hdc)
    return printer


def _load_as_pages(path, dpi=300):
    """PDF 用 pypdfium2 渲染成 PIL 图列表（默认 300 DPI，文字清晰无锯齿）；
    JPG/PNG 原图直接打印（本身是位图，提 DPI 无意义）。"""
    from PIL import Image
    if path.lower().endswith(".pdf"):
        import pypdfium2 as pdfium
        pages = []
        pdf = pdfium.PdfDocument(path)
        for i in range(len(pdf)):
            page = pdf[i]
            pages.append(page.render(scale=dpi / 72.0).to_pil())
            page.close()
        pdf.close()
        return pages
    return [Image.open(path)]


def _default_printer():
    import ctypes
    from ctypes import wintypes
    spool = ctypes.WinDLL("winspool.drv")
    spool.GetDefaultPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    size = wintypes.DWORD(260)
    buf = ctypes.create_unicode_buffer(260)
    if not spool.GetDefaultPrinterW(buf, ctypes.byref(size)):
        return ""
    return buf.value


def _find_ghostscript():
    import glob as _glob
    hits = sorted(_glob.glob(r"C:\Program Files\gs\gs*\bin\gswin64c.exe"), reverse=True)
    return hits[0] if hits and os.path.exists(hits[0]) else None


def cmd_print(a):
    """打印：PDF 优先 Ghostscript/SumatraPDF 矢量直打默认打印机；无阅读器时渲染位图走 GDI；
    图片走 GDI。全部静默无对话框，逐份间隔 2 秒防队列拥堵。"""
    import time
    gs = _find_ghostscript()
    sumatra = None
    for cand in (r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
                 r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
                 os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe")):
        if os.path.exists(cand):
            sumatra = cand
            break
    wps = None
    try:  # WPS: wps.exe /pdf /p 可直接打印 PDF（注册表 KWPS.PDF.9 的 print 动词）
        import glob as _glob
        hits = sorted(_glob.glob(os.path.expandvars(
            r"%LOCALAPPDATA%\Kingsoft\WPS Office\*\office6\wps.exe")), reverse=True)
        wps = hits[0] if hits and os.path.exists(hits[0]) else None
    except Exception:
        pass
    targets = []
    if os.path.isdir(a.target):
        for f in sorted(os.listdir(a.target)):
            if f.lower().endswith((".pdf", ".jpg", ".jpeg", ".png")):
                targets.append(os.path.join(a.target, f))
    else:
        targets = [a.target]
    printed_pages = 0
    printer = _default_printer()
    for t in targets:
        if a.landscape:
            # 横版需求：只对"发票"旋转90°横打；会议通知等非发票文档、付款记录图片仍竖版。
            # 发票判定 = 能提取到金额（与统计口径一致）；旋转方向若反了把 -90 改成 90。
            try:
                pages = _load_as_pages(t, a.dpi)
                is_invoice = False
                if t.lower().endswith(".pdf"):
                    try:
                        is_invoice = parse_amount(extract_text(t)) is not None
                    except Exception:
                        is_invoice = False
                if is_invoice:
                    pages = [p.rotate(-90, expand=True) for p in pages]
                printer = _gdi_print(pages, os.path.basename(t)) or printer
                tag = "发票横版" if is_invoice else "竖版"
                print(f"已发送打印({tag}): {os.path.basename(t)}（{len(pages)} 页 -> {printer}）")
                printed_pages += len(pages)
                time.sleep(2)
                continue
            except Exception as e:
                print(f"打印失败: {e}")
                continue
        # PDF 优先矢量通路：SumatraPDF 静默直打；其次 Ghostscript mswinpr2
        if t.lower().endswith(".pdf") and sumatra:
            try:
                cmd = [sumatra, "-print-to-default", "-silent"]
                if a.landscape:
                    cmd += ["-print-settings", "landscape,fit"]  # 横版+铺满纸张
                cmd.append(t)
                subprocess.run(cmd, check=True)
                print(f"已发送打印(SumatraPDF矢量{'横版' if a.landscape else ''}): {os.path.basename(t)}")
                printed_pages += 1
                time.sleep(2)
                continue
            except Exception as e:
                print(f"SumatraPDF 打印失败，尝试下一通路: {e}")
        if t.lower().endswith(".pdf") and gs and printer:
            try:
                subprocess.run([gs, "-dBATCH", "-dNOPAUSE", "-dQUIET", "-dNOPROMPT",
                                "-dNumCopies=1", "-sDEVICE=mswinpr2",
                                f"-sOutputFile=%printer%{printer}", t], check=True)
                print(f"已发送打印(Ghostscript矢量): {os.path.basename(t)} -> {printer}")
                printed_pages += 1
                time.sleep(2)
                continue
            except Exception as e:
                print(f"Ghostscript 打印失败，改 GDI 位图打印: {e}")
        try:
            pages = _load_as_pages(t, a.dpi)
            printer = _gdi_print(pages, os.path.basename(t)) or printer
            print(f"已发送打印: {os.path.basename(t)}（{len(pages)} 页 -> {printer}）")
            printed_pages += len(pages)
            time.sleep(2)
        except Exception as e:
            # 回退：SumatraPDF / WPS / 系统 print 动词
            try:
                if t.lower().endswith(".pdf") and sumatra:
                    subprocess.run([sumatra, "-print-to-default", "-silent", t], check=True)
                elif t.lower().endswith(".pdf") and wps:
                    subprocess.run([wps, "/prometheus", "/pdf", "/n", "/p", t], check=True)
                else:
                    os.startfile(t, "print")  # noqa: SH601 - Windows 专用
                print(f"已发送打印（回退通路）: {os.path.basename(t)}")
                printed_pages += 1
                time.sleep(2)
            except Exception as e2:
                print(f"打印失败 {t}: {e}；回退亦失败: {e2}")
    print(f"\n共发送 {printed_pages} 页打印任务到 {printer or '默认打印机'}，请检查打印机输出。")


def cmd_add(a):
    """向 <文件夹>/纸质发票.csv 追加登记。不带 --amount 时进入连续交互录入，逐张输入，空行结束。"""
    import csv
    path = os.path.join(a.folder, MANUAL_CSV)
    exists = os.path.exists(path)
    cat = a.category if a.category in CATEGORIES else "其他"
    with open(path, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["类别", "分组", "费用类型", "金额", "备注"])
        if a.amount is not None:
            w.writerow([cat, a.group or "", a.etype or "", f"{a.amount:.2f}", a.note or ""])
            print(f"已登记: {cat} | {a.group or '-'} | {a.etype or '-'} | {a.amount:.2f} -> {path}")
            return
        # 连续录入模式
        print("纸质发票连续录入（直接回车结束某项，金额为空则结束录入）")
        print("类别可选: " + " / ".join(CATEGORIES))
        n = 0
        while True:
            amount_s = input("金额: ").strip()
            if not amount_s:
                break
            try:
                amount = float(amount_s)
            except ValueError:
                print("  金额无效，重试")
                continue
            c = input("类别(默认其他): ").strip() or "其他"
            if c not in CATEGORIES:
                c = "其他"
            g = input("分组(如 8月-杭州会议，可空): ").strip()
            e = input("费用类型(会务费/住宿费/其他费用，可空): ").strip() if c == "会议" else ""
            note = input("备注(可空): ").strip()
            w.writerow([c, g, e, f"{amount:.2f}", note])
            f.flush()
            n += 1
            print(f"  已登记第{n}张: {c} {amount:.2f}")
    print(f"共登记 {n if a.amount is None else 1} 张 -> {path}")


def cmd_total(a):
    """只输金额算总和：命令行传金额，或逐行输入、空行结束。输错可输负数冲减。"""
    amounts = list(a.amounts) if a.amounts else []
    if not amounts:
        print("逐个输入金额，直接回车结束。输错了输个负数冲减即可。")
        while True:
            s = input("金额: ").strip()
            if not s:
                break
            try:
                amounts.append(float(s))
            except ValueError:
                print("  无效数字，重试")
    total = sum(amounts)
    print(f"共 {len(amounts)} 笔，总金额: {total:.2f} 元")


def main():
    p = argparse.ArgumentParser(description="发票分类统计整理工具")
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("scan", help="扫描并统计")
    s.add_argument("folder")
    s.add_argument("--json", help="明细保存为 JSON")
    o = sub.add_parser("organize", help="按类别整理文件并生成统计")
    o.add_argument("folder")
    o.add_argument("--out", help="输出目录（默认 <文件夹>/分类整理）")
    o.add_argument("--move", action="store_true", help="移动而非复制")
    ad = sub.add_parser("add", help="登记纸质发票（无电子版）；不带 --amount 进入连续录入")
    ad.add_argument("folder")
    ad.add_argument("--category", default="其他", help="类别: 交通/会议/图书/顺丰快递/其他")
    ad.add_argument("--amount", type=float, help="金额（不带则连续交互录入）")
    ad.add_argument("--group", default="", help="分组名，如 8月-杭州会议")
    ad.add_argument("--etype", default="", help="费用类型（会议类）: 会务费/住宿费/其他费用")
    ad.add_argument("--note", default="", help="备注")
    pr = sub.add_parser("print", help="打印文件或目录")
    pr.add_argument("target")
    pr.add_argument("--dpi", type=int, default=300, help="PDF 渲染精度（默认 300，最高 600）")
    pr.add_argument("--landscape", action="store_true", help="横版打印并铺满纸张（发票常用）")
    pr.add_argument("--printer", help="打印机名（暂用系统默认）")
    t = sub.add_parser("total", help="只输金额算总金额（可带金额参数，否则逐行输入）")
    t.add_argument("amounts", nargs="*", type=float, help="金额列表，如: total 120 45.5 200")
    a = p.parse_args()
    {"scan": cmd_scan, "organize": cmd_organize, "add": cmd_add, "print": cmd_print,
     "total": cmd_total}[a.cmd](a)


if __name__ == "__main__":
    main()
