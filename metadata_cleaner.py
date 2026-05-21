#!/usr/bin/env python3
"""
Metadata Cleaner - Cross-platform tool to remove metadata from Office files & PDFs.
Works on Windows, macOS, and Linux.
"""

import os
import sys
import shutil
import zipfile
try:
    import defusedxml.ElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET  # fallback, Python 3.8+ disables external entities by default
from xml.etree.ElementTree import register_namespace
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from threading import Thread

# Optional drag-and-drop support
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES

    HAS_DND = True
except ImportError:
    HAS_DND = False
    TkinterDnD = None  # type: ignore[assignment]
    DND_FILES = None

# ============================================================
# Configuration
# ============================================================

SUPPORTED_EXTENSIONS = {
    ".docx", ".xlsx", ".pptx",     # Microsoft Office
    ".wps",  ".et",   ".dps",      # WPS Office (new XML-based format)
    ".pdf",
}
OFFICE_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".wps", ".et", ".dps"}

# XML namespaces in Office Open XML properties
NS_CORE = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

NS_APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

# Try importing pypdf (modern) then PyPDF2 (legacy) for PDF support
try:
    from pypdf import PdfReader, PdfWriter

    HAS_PDF_SUPPORT = True
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter  # type: ignore[no-redef]

        HAS_PDF_SUPPORT = True
    except ImportError:
        HAS_PDF_SUPPORT = False


# ============================================================
# Office file metadata cleaning
# ============================================================

# Fields in core.xml — namespace-qualified tag, action, and optional new value
_CORE_FIELDS = [
    # (full Clark-notation tag, action, new_value_or_None)
    ("{http://purl.org/dc/elements/1.1/}creator", "clear", None),
    ("{http://purl.org/dc/elements/1.1/}description", "clear", None),
    ("{http://purl.org/dc/elements/1.1/}subject", "clear", None),
    ("{http://purl.org/dc/elements/1.1/}title", "clear", None),
    ("{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy", "clear", None),
    ("{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}version", "clear", None),
    ("{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}revision", "set", "1"),
    ("{http://purl.org/dc/terms/}created", "clear", None),
    ("{http://purl.org/dc/terms/}modified", "clear", None),
]


def _clean_core_xml(xml_bytes: bytes) -> bytes:
    """Parse core.xml, strip/rewrite metadata fields, return cleaned XML bytes.

    Elements are cleared (not removed) to preserve namespace declarations
    that Word requires.  The XML declaration is normalised to double-quoted
    attributes with standalone='yes' to match Office conventions.
    """
    for prefix, uri in NS_CORE.items():
        register_namespace(prefix, uri)

    root = ET.fromstring(xml_bytes)

    for tag, action, value in _CORE_FIELDS:
        for elem in root.findall(tag):
            if action == "clear":
                elem.text = ""
            elif action == "set":
                elem.text = value

    raw = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    # Normalise the XML declaration to Office-compatible format
    raw = raw.replace(
        b"<?xml version='1.0' encoding='UTF-8'?>",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    )
    return raw


# Fields in app.xml — the metadata we strip from extended properties
_APP_FIELDS_TO_CLEAR = [
    "Company",
    "Manager",
    "TotalTime",
]


def _clean_app_xml(xml_bytes: bytes) -> bytes:
    """Parse app.xml, strip company/manager metadata, return cleaned XML bytes."""
    register_namespace("", NS_APP)
    root = ET.fromstring(xml_bytes)

    for field in _APP_FIELDS_TO_CLEAR:
        for elem in root.findall(f"{{{NS_APP}}}{field}"):
            elem.text = ""

    raw = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    raw = raw.replace(
        b"<?xml version='1.0' encoding='UTF-8'?>",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    )
    return raw


def clean_office_file(filepath: str) -> tuple:
    """
    Remove metadata from a .docx / .xlsx / .pptx / .wps / .et / .dps file.
    Streams ZIP entries to avoid loading large files into memory.
    Uses atomic temp-file replacement to prevent data loss on interruption.
    Returns (success: bool, error_message: str|None).
    """
    tmp_path = filepath + ".tmp"

    try:
        with zipfile.ZipFile(filepath, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename

                # Clean core.xml metadata
                if name == "docProps/core.xml":
                    cleaned = _clean_core_xml(zin.read(name))
                    zout.writestr(item, cleaned)
                # Clean app.xml metadata (company, manager)
                elif name == "docProps/app.xml":
                    cleaned = _clean_app_xml(zin.read(name))
                    zout.writestr(item, cleaned)
                # Drop custom.xml entirely — can contain user paths / identifiers
                elif name == "docProps/custom.xml":
                    continue
                # Stream everything else as-is (no memory accumulation)
                else:
                    zout.writestr(item, zin.read(name))

        # Atomic replace — original stays intact until tmp is fully written
        os.replace(tmp_path, filepath)
        return True, None

    except PermissionError:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False, "文件正被其他程序(如 Office/WPS)占用，请先关闭该文件再试。"
    except zipfile.BadZipFile:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False, "文件已损坏或不是有效的 Office/WPS 文档（旧版二进制格式不支持）。"
    except Exception as exc:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False, str(exc)


# ============================================================
# PDF metadata cleaning
# ============================================================


def clean_pdf_file(filepath: str) -> tuple:
    """
    Remove metadata from a PDF file.
    Returns (success: bool, error_message: str|None).
    """
    if not HAS_PDF_SUPPORT:
        return False, "PDF 组件未安装，请运行 setup.bat (Windows) 或 setup.sh (macOS) 安装依赖。"

    tmp_path = filepath + ".tmp"

    try:
        reader = PdfReader(filepath)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        # Explicitly overwrite metadata to clear XMP /Info remnants
        writer.add_metadata({})

        with open(tmp_path, "wb") as f:
            writer.write(f)

        # Strip the /Producer entry that pypdf/PyPDF2 injects automatically
        _strip_pdf_producer(tmp_path)

        os.replace(tmp_path, filepath)
        return True, None

    except PermissionError:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False, "文件正被其他程序占用，请先关闭该 PDF 再试。"
    except Exception as exc:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False, str(exc)


def _strip_pdf_producer(path: str) -> None:
    """Remove /Producer entry from a PDF's Info dictionary at byte level."""
    import re

    with open(path, "rb") as f:
        content = f.read()

    # Match /Producer (value) — string literal
    content = re.sub(rb"/Producer\s*\([^)]*\)", b"", content)
    # Match /Producer <hex> — hex string alternative
    content = re.sub(rb"/Producer\s*<[0-9A-Fa-f]*>", b"", content)

    with open(path, "wb") as f:
        f.write(content)


# ============================================================
# Unified entry point
# ============================================================


def clean_file(filepath: str) -> tuple:
    """Clean a single file. Returns (success, error_message)."""
    ext = Path(filepath).suffix.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        return False, f"不支持的文件类型: {ext}"

    if ext == ".pdf":
        return clean_pdf_file(filepath)
    else:
        return clean_office_file(filepath)


# ============================================================
# Document content warnings (comments / tracked changes)
# ============================================================


def scan_document_warnings(filepath: str) -> list[str]:
    """
    Scan an Office file for comments and tracked changes that the tool
    cannot strip from raw XML.  Returns a human-readable list of issues.
    """
    ext = Path(filepath).suffix.lower()
    if ext not in OFFICE_EXTENSIONS:
        return []

    warnings: list[str] = []

    try:
        with zipfile.ZipFile(filepath, "r") as z:
            names = z.namelist()
            name_set = frozenset(names)

            # -- Word comments: file must exist AND contain actual <w:comment elements --
            if "word/comments.xml" in name_set:
                raw = z.read("word/comments.xml")
                if b"<w:comment " in raw or b"<w:comment>" in raw or b"<w:comment/" in raw:
                    warnings.append("批注")

            # -- Excel comments --------------------------------------------------
            for n in names:
                if n.startswith("xl/comments") and n.endswith(".xml"):
                    raw = z.read(n)
                    if raw.strip():
                        if "批注" not in warnings:
                            warnings.append("批注")
                    break

            # -- PowerPoint comments --------------------------------------------
            for n in names:
                if n.startswith("ppt/comments/") and n.endswith(".xml"):
                    raw = z.read(n)
                    if raw.strip():
                        if "批注" not in warnings:
                            warnings.append("批注")
                    break

            # -- Word tracked changes: exact XML tag (not substring) ------------
            doc_xml_path = "word/document.xml"
            if doc_xml_path in name_set:
                reader = z.open(doc_xml_path, "r")
                chunk = reader.read(65536)
                reader.close()
                if b"<w:ins " in chunk or b"<w:del " in chunk or \
                   b"<w:ins>" in chunk or b"<w:del>" in chunk:
                    warnings.append("修订记录")

    except Exception:
        pass  # scanning is best-effort; don't block the user over it

    return warnings


# ============================================================
# GUI
# ============================================================


class MetadataCleanerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("元数据清除工具")
        self.root.geometry("620x480")
        self.root.minsize(500, 380)

        self.files: list[str] = []
        self._cleaning = False  # guard against closing during processing
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # -- UI construction --------------------------------------------------

    def _build_ui(self):
        main = ttk.Frame(self.root, padding="12")
        main.pack(fill=tk.BOTH, expand=True)

        # Header
        ttk.Label(main, text="元数据清除工具", font=("", 16, "bold")).pack(pady=(0, 2))
        ttk.Label(
            main,
            text="支持 Word / Excel / PPT (.docx/.xlsx/.pptx)  |  WPS (.wps/.et/.dps)  |  PDF",
            font=("", 10),
        ).pack(pady=(0, 12))

        # Disclaimer: what this tool does NOT clean
        ttk.Label(
            main,
            text="注：本工具清除文件属性元数据，不清除文档正文中的批注与修订记录。",
            font=("", 8),
        ).pack(pady=(0, 6))

        # Warning if PyPDF2 is missing
        if not HAS_PDF_SUPPORT:
            warn_frame = ttk.Frame(main)
            warn_frame.pack(fill=tk.X, pady=(0, 10))
            ttk.Label(
                warn_frame,
                text="[!] PDF 功能需要 PyPDF2 库，请先运行 setup.bat (Windows) 或 setup.sh (macOS) 安装依赖",
                foreground="#c0392b",
                font=("", 9),
            ).pack()

        # File list
        list_frame = ttk.LabelFrame(main, text="文件列表", padding="6")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        inner = ttk.Frame(list_frame)
        inner.pack(fill=tk.BOTH, expand=True)

        sb = ttk.Scrollbar(inner)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(
            inner,
            selectmode=tk.EXTENDED,
            yscrollcommand=sb.set,
            font=("", 11),
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=self.listbox.yview)

        # Drag-and-drop registration
        if HAS_DND:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)
            # Also register the whole window for drops
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)

        # Buttons
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_add = ttk.Button(btn_frame, text="添加文件", command=self._add_files)
        self.btn_add.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_remove = ttk.Button(btn_frame, text="移除选中", command=self._remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_clear = ttk.Button(btn_frame, text="清空列表", command=self._clear_files)
        self.btn_clear.pack(side=tk.LEFT)

        self.btn_clean = ttk.Button(
            btn_frame, text="清除元数据", command=self._start_cleaning
        )
        self.btn_clean.pack(side=tk.RIGHT)

        # Progress bar
        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 4))

        # Status
        ready_text = "就绪 — 请添加文件"
        if not HAS_PDF_SUPPORT:
            ready_text += "  (PDF 功能需安装 pypdf)"
        self.status_var = tk.StringVar(value=ready_text)
        ttk.Label(main, textvariable=self.status_var, font=("", 9)).pack(anchor=tk.W)

    # -- Button callbacks -------------------------------------------------

    def _add_files(self):
        types = [
            ("支持的文件", "*.docx *.xlsx *.pptx *.wps *.et *.dps *.pdf"),
            ("Word / WPS 文档", "*.docx *.wps"),
            ("Excel / WPS 表格", "*.xlsx *.et"),
            ("PowerPoint / WPS 演示", "*.pptx *.dps"),
            ("PDF 文件", "*.pdf"),
        ]
        paths = filedialog.askopenfilenames(title="选择文件", filetypes=types)

        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, os.path.basename(p))

        self._refresh_status()

    def _remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.files[idx]
        self._refresh_status()

    def _clear_files(self):
        self.listbox.delete(0, tk.END)
        self.files.clear()
        self._refresh_status()

    def _on_drop(self, event):
        """Handle file drop from OS file manager (requires tkinterdnd2)."""
        raw = event.data
        # Parse file paths from tkinterdnd2 format:
        #   Windows: {C:/path/file.docx} {C:/path/file.pdf}
        #   macOS:   /path/file.docx /path/file.pdf   (space-separated, brace-wrapped)
        paths: list[str] = []
        brace_depth = 0
        current = ""
        for ch in raw:
            if ch == "{":
                brace_depth += 1
                continue
            elif ch == "}":
                brace_depth -= 1
                if brace_depth == 0 and current:
                    paths.append(current)
                    current = ""
                continue
            elif ch == " " and brace_depth == 0:
                if current:
                    paths.append(current)
                    current = ""
                continue
            current += ch
        if current:
            paths.append(current)

        added = 0
        for p in paths:
            p = p.strip()
            if p and os.path.isfile(p) and p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, os.path.basename(p))
                added += 1

        if added:
            self._refresh_status()

    def _refresh_status(self):
        n = len(self.files)
        self.status_var.set(f"已选择 {n} 个文件" if n else "就绪 — 请添加文件")

    # -- Cleaning flow ----------------------------------------------------

    def _start_cleaning(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加要处理的文件。")
            return

        # Scan Office files for comments / tracked changes
        warnings_map: dict[str, list[str]] = {}
        for fp in self.files:
            issues = scan_document_warnings(fp)
            if issues:
                warnings_map[os.path.basename(fp)] = issues

        if warnings_map:
            lines = ["以下文件中检测到：\n"]
            for fname, issues in warnings_map.items():
                lines.append(f"    {fname}  —  {'、'.join(issues)}")
            lines.append("\n批注和修订记录嵌在文档正文中，本工具无法清除。")
            lines.append("请在 Office / WPS 中手动操作：")
            lines.append("  1) 审阅 → 接受所有修订")
            lines.append("  2) 审阅 → 删除文档中的所有批注")
            lines.append("  3) 保存后重新添加到本工具\n")
            lines.append("是否仍然继续清除？（跳过批注和修订）")

            if not messagebox.askokcancel("警告：发现批注/修订记录", "\n".join(lines)):
                return

        ok = messagebox.askokcancel(
            "确认清除",
            f"即将清除 {len(self.files)} 个文件的元数据。\n"
            "确认继续？",
        )
        if not ok:
            return

        self._cleaning = True
        self._toggle_buttons(tk.DISABLED)
        self.progress["value"] = 0
        self.status_var.set("正在处理...")

        t = Thread(target=self._clean_all, daemon=True)
        t.start()

    def _clean_all(self):
        total = len(self.files)
        results = []

        for i, fp in enumerate(self.files):
            fname = os.path.basename(fp)
            self._ui_update(i, total, f"处理中: {fname}")
            ok, err = clean_file(fp)
            results.append((fname, ok, err))
            self._ui_update(i + 1, total, f"完成: {fname}")

        self.root.after(0, lambda: self._show_results(results))

    def _ui_update(self, cur, total, msg):
        self.root.after(0, lambda: self._do_update(cur, total, msg))

    def _do_update(self, cur, total, msg):
        self.progress["value"] = (cur / total) * 100 if total else 0
        self.status_var.set(msg)

    def _show_results(self, results):
        self._cleaning = False
        self._toggle_buttons(tk.NORMAL)
        self.progress["value"] = 100

        ok_count = sum(1 for _, s, _ in results if s)
        bad_count = len(results) - ok_count

        if bad_count == 0:
            self.status_var.set(f"全部成功 — 已清除 {ok_count} 个文件的元数据")
            messagebox.showinfo("完成", f"成功清除 {ok_count} 个文件的元数据！")
            self._clear_files()
        else:
            self.status_var.set(f"完成：{ok_count} 成功, {bad_count} 失败")
            details = "\n".join(
                f"  {name}: {err}" for name, ok, err in results if not ok
            )
            messagebox.showwarning(
                "完成（有错误）",
                f"成功: {ok_count}  失败: {bad_count}\n\n失败详情:\n{details}",
            )

    def _toggle_buttons(self, state):
        for b in (self.btn_add, self.btn_remove, self.btn_clear, self.btn_clean):
            b.config(state=state)

    def _on_close(self):
        if self._cleaning:
            messagebox.showwarning("提示", "正在处理中，请等待完成后再关闭窗口。")
        else:
            self.root.destroy()


# ============================================================
# CLI mode (drag files onto script icon)
# ============================================================


def cli_mode(paths: list[str]):
    """Process files from command-line arguments."""
    print("元数据清除工具")
    print("=" * 40)
    print()

    ok = bad = 0

    for p in paths:
        if not os.path.exists(p):
            print(f"[跳过] 文件不存在: {p}")
            bad += 1
            continue

        ext = Path(p).suffix.lower()
        if ext not in SUPPORTED_EXTENSIONS:
            print(f"[跳过] 不支持的类型 ({ext}): {os.path.basename(p)}")
            bad += 1
            continue

        print(f"处理: {os.path.basename(p)} ... ", end="", flush=True)
        success, err = clean_file(p)
        if success:
            print("成功")
            ok += 1
        else:
            print(f"失败 — {err}")
            bad += 1

    print(f"\n完成: {ok} 成功, {bad} 失败")

    if os.name == "nt":
        input("\n按 Enter 键退出...")


# ============================================================
# Entry point
# ============================================================


def main():
    if len(sys.argv) > 1:
        cli_mode(sys.argv[1:])
    else:
        if HAS_DND:
            root = TkinterDnD.Tk()
        else:
            root = tk.Tk()
        MetadataCleanerApp(root)
        root.mainloop()


if __name__ == "__main__":
    main()
