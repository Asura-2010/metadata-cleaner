#!/usr/bin/env python3
"""
Metadata Cleaner - Cross-platform tool to remove metadata from Office files & PDFs.
Works on Windows, macOS, and Linux.
"""

__version__ = "1.5.0"

import os
import re
import sys
import json
import time
import random
import shutil
import tempfile
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from threading import Thread

from randomizer import (
    generate_unique_identities,
    randomize_times,
    random_template,
    random_total_time,
    random_rsid_mapping,
    format_metadata_report,
    ReportPopup,
)

# Optional drag-and-drop support
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES

    HAS_DND = True
except ImportError:
    HAS_DND = False
    TkinterDnD = None  # type: ignore[assignment]
    DND_FILES = None

# Config file for remembering window size and layout
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".metadata-cleaner-config.json")

# ============================================================
# Configuration
# ============================================================

SUPPORTED_EXTENSIONS = {
    ".docx", ".xlsx", ".pptx",     # Microsoft Office
    ".wps",  ".et",   ".dps",      # WPS Office (new XML-based format)
    ".pdf",
    ".png", ".jpg", ".jpeg",       # Standalone images
    ".gif", ".bmp", ".tiff", ".tif", ".webp",
    ".heic",
}
OFFICE_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".wps", ".et", ".dps"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".heic"}

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


# Try importing Pillow for image metadata stripping (optional)
try:
    from PIL import Image

    HAS_IMAGE_SUPPORT = True
except ImportError:
    HAS_IMAGE_SUPPORT = False

# Register HEIC/HEIF support with Pillow (optional)
try:
    import pillow_heif

    pillow_heif.register_heif_opener()
    HAS_HEIF_SUPPORT = True
except ImportError:
    HAS_HEIF_SUPPORT = False


# ============================================================
# Image metadata cleaning (embedded in Office files)
# ============================================================


def _strip_image_metadata(data: bytes) -> bytes:
    """Remove EXIF, PNG text chunks, and other metadata from image bytes.
    Returns cleaned bytes, or original if Pillow is unavailable or fails."""
    if not HAS_IMAGE_SUPPORT:
        return data

    try:
        import io

        img = Image.open(io.BytesIO(data))
        fmt = img.format
        if fmt not in ("PNG", "JPEG", "TIFF", "BMP", "GIF", "HEIF"):
            return data

        out = io.BytesIO()
        # Re-save without metadata: strip EXIF, PNG text chunks, etc.
        if fmt == "JPEG":
            try:
                img.save(out, format="JPEG", quality="keep")
            except ValueError:
                # quality="keep" requires quantization tables; degrade gracefully
                img.save(out, format="JPEG", quality=95)
        elif fmt == "PNG":
            img.save(out, format="PNG")
        elif fmt == "TIFF":
            # Re-save without extra metadata. Do NOT rebuild from pixel data —
            # that changes the TIFF structure and PowerPoint flags the file as
            # corrupted. Structural IFD tags (resolution, strip layout, etc.)
            # are necessary for integrity and are filtered from the metadata viewer.
            img.save(out, format="TIFF")
        elif fmt == "HEIF":
            # Pillow's generic save preserves EXIF for HEIF; use pillow-heif
            # directly after stripping metadata from the Pillow image.
            try:
                from pillow_heif import from_pillow as heif_from_pillow
                for key in list(img.info.keys()):
                    try:
                        del img.info[key]
                    except (KeyError, TypeError):
                        pass
                heif_img = heif_from_pillow(img)
                heif_img.info.clear()
                heif_img.save(out, quality=80)
            except ImportError:
                img.save(out, format="HEIF", quality=95)
        else:
            img.save(out, format=fmt)
        return out.getvalue()
    except Exception:
        return data  # best-effort


# ============================================================
# Office file metadata cleaning
# ============================================================


def _xml_escape(val: bytes) -> bytes:
    """Escape &, <, > in bytes for safe XML text content insertion."""
    val = val.replace(b"&", b"&amp;")
    val = val.replace(b"<", b"&lt;")
    val = val.replace(b">", b"&gt;")
    return val


def _maybe_empty(val: bytes, probability: float = 0.15) -> bytes:
    """Randomly return empty instead of the value, for natural-looking metadata.

    Real files don't always have every field populated — ~15% of fields
    are blank in practice.  This avoids the "everything filled" pattern
    that looks obviously machine-generated.
    """
    return b"" if random.random() < probability else val


def _clean_core_xml(xml_bytes: bytes, randomize: bool = False,
                    random_state: dict | None = None) -> bytes:
    """Remove identity-related metadata from core.xml.

    When randomize=False (default): only removes author/lastModifiedBy elements
    entirely (no empty tags left behind). All other fields (dates, revision,
    title, subject, etc.) are preserved as-is — they don't expose the operator.
    When randomize=True: creator / lastModifiedBy filled with random name,
    dates filled with random timestamps, revision set to random number,
    other string fields cleared.
    """
    data = xml_bytes

    if randomize and random_state:
        author = _xml_escape(random_state["author_name"].encode())
        created, last_printed, modified = random_state["times"]
        revision = random_state["revision"].encode()

        # Helper: replacement function that inserts value between groups
        def _ins(val: bytes):
            return lambda m: m.group(1) + val + m.group(2)

        # Creator / lastModifiedBy → random name (sometimes left empty)
        for prefix, local in [("dc", "creator"), ("cp", "lastModifiedBy")]:
            tag_pat = prefix.encode() + b":" + local.encode()
            name_or_empty = _maybe_empty(author)
            data = re.sub(
                rb"(<" + tag_pat + rb"(?:\s[^>]*)?>)[^<]*(</" + tag_pat + rb">)",
                _ins(name_or_empty), data,
            )

        # Date fields → random timestamps (rarely left empty — 8%)
        for (prefix, local), val in [
            (("dcterms", "created"), created),
            (("dcterms", "modified"), modified),
            (("cp", "lastPrinted"), last_printed),
        ]:
            tag_pat = prefix.encode() + b":" + local.encode()
            val_bytes = _maybe_empty(val.encode(), 0.08)
            data = re.sub(
                rb"(<" + tag_pat + rb"(?:\s[^>]*)?>)[^<]*(</" + tag_pat + rb">)",
                _ins(val_bytes), data,
            )
            data = re.sub(
                rb"<" + tag_pat + rb"(?:\s[^>]*)?/>",
                b"<" + tag_pat + b">" + val_bytes + b"</" + tag_pat + b">",
                data,
            )

        # Revision → random number (sometimes 0 — 10%)
        rev_or_empty = _maybe_empty(revision, 0.10)
        data = re.sub(
            rb"<cp:revision(?:\s[^>]*)?>[^<]*</cp:revision>",
            b"<cp:revision>" + rev_or_empty + b"</cp:revision>",
            data,
        )
        data = re.sub(
            rb"<cp:revision(?:\s[^>]*)?/>",
            b"<cp:revision>" + rev_or_empty + b"</cp:revision>",
            data,
        )

        # Remaining string fields: clear text
        for prefix, local in [
            ("dc", "description"), ("dc", "subject"), ("dc", "title"),
            ("cp", "version"), ("cp", "keywords"), ("cp", "category"),
        ]:
            tag_pat = prefix.encode() + b":" + local.encode()
            data = re.sub(
                rb"(<" + tag_pat + rb"(?:\s[^>]*)?>)[^<]*(</" + tag_pat + rb">)",
                rb"\1\2",
                data,
            )
    else:
        # Only remove identity-leaking fields — delete the entire element,
        # not just clear content (empty tags are a forensic signal).
        # Preserve everything else (dates, revision, title, subject, etc.)
        # since they don't expose who operated the tool.
        for prefix, local in [("dc", "creator"), ("cp", "lastModifiedBy")]:
            tag_pat = prefix.encode() + b":" + local.encode()
            data = re.sub(
                rb"<" + tag_pat + rb"(?:\s[^>]*)?>"
                rb"[^<]*"
                rb"</" + tag_pat + rb">",
                b"",
                data,
            )
            data = re.sub(
                rb"<" + tag_pat + rb"(?:\s[^>]*)?/>",
                b"",
                data,
            )

    return data


# Application/AppVersion — used in randomize mode to set realistic Office values.
# Office 2016/2019/365 = 16.0, Office 2013 = 15.0, Office 2010 = 14.0
_OFFICE_VERSIONS = ["16.0", "16.0", "16.0",  # most common — just major version
                    "16.0.18526.2016", "16.0.17928.2024",
                    "16.0.17328.2012", "16.0.16924.2008",
                    "15.0.5631.1000", "15.0.5589.1000",
                    "14.0.7268.5000"]

_CUSTOM_PROP_NAMES = [
    "DocumentId", "Version", "Status", "Category", "Manager",
    "CheckedBy", "Department", "ProjectCode", "RefNo", "BatchId",
    "ClientId", "Priority", "Source", "Language", "Region",
]

_CUSTOM_PROP_VALUES = [
    "v1.0", "v2.1", "Draft", "Final", "Internal",
    "A", "B", "C", "1", "2", "3",
    str(time.localtime().tm_year), str(time.localtime().tm_year - 1),
    "Q1", "Q2", "Q3", "Q4",
]


def _unique_empty_custom_xml() -> bytes:
    """Return a custom.xml with random properties per file.

    An empty skeleton with only xmlns declarations is a forensic fingerprint —
    all batch-processed files would hash identically.  Adding 0-3 random
    custom properties makes each file unique and looks like normal Office
    usage where applications set their own custom metadata.
    """
    lines = [
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"'
        b' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    ]

    # Add 0-3 random custom properties
    pid = 2  # Office starts property IDs at 2
    for _ in range(random.randint(0, 3)):
        name = random.choice(_CUSTOM_PROP_NAMES)
        value = random.choice(_CUSTOM_PROP_VALUES)
        # Mix value types for realism
        vtype = random.choice(["vt:lpwstr", "vt:i4", "vt:bool"])
        if vtype == "vt:i4":
            val_xml = f"<vt:i4>{random.randint(1, 999)}</vt:i4>"
        elif vtype == "vt:bool":
            val_xml = f"<vt:bool>{random.choice(['true', 'false'])}</vt:bool>"
        else:
            val_xml = f"<vt:lpwstr>{value}</vt:lpwstr>"
        lines.append(
            f'  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" '
            f'pid="{pid}" name="{name}">{val_xml}</property>'.encode()
        )
        pid += 1

    lines.append(b'</Properties>')
    return b'\n'.join(lines)


# ============================================================
# Anti-forensics: filesystem timestamp scattering
# ============================================================

if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    class _FILETIME(ctypes.Structure):
        _fields_ = [("dwLowDateTime", wintypes.DWORD),
                    ("dwHighDateTime", wintypes.DWORD)]

    def _scatter_file_timestamps(filepath: str) -> None:
        """Windows: scatter ctime/atime/mtime and sync NTFS $SIA/$FNA.

        Uses SetFileTime to set all three timestamps, then renames the file
        (rename-then-back) to force Windows to copy $SIA → $FNA, eliminating
        the timestomping signature that forensic tools detect.
        """
        scattered = time.time() - random.randint(300, 259200)
        ft = _unix_to_filetime(scattered)

        handle = ctypes.windll.kernel32.CreateFileW(
            filepath, 0x40000000, 0, None, 3, 0x80, None
        )
        if handle == -1:
            return
        try:
            ctypes.windll.kernel32.SetFileTime(
                handle, ctypes.byref(ft), ctypes.byref(ft), ctypes.byref(ft)
            )
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)

        # Rename trick: sync $SIA → $FNA so forensic tools see consistent timestamps
        tmp_name = filepath + ".ts_sync"
        try:
            os.rename(filepath, tmp_name)
            os.rename(tmp_name, filepath)
        except OSError:
            pass

    def _unix_to_filetime(unix_ts: float):
        """Convert Unix timestamp to Windows FILETIME (100ns intervals since 1601)."""
        ticks = int((unix_ts + 11644473600) * 10000000)
        return _FILETIME(ticks & 0xFFFFFFFF, ticks >> 32)

else:
    def _scatter_file_timestamps(filepath: str) -> None:
        """macOS/Linux: scatter only mtime/atime (ctime is kernel-controlled).

        ctime will cluster at processing time — see README for mitigation.
        """
        scattered = time.time() - random.randint(300, 259200)
        try:
            os.utime(filepath, (scattered, scattered))
        except OSError:
            pass


def _clean_app_xml(xml_bytes: bytes, randomize: bool = False,
                   random_state: dict | None = None) -> bytes:
    """Remove identity-related metadata from app.xml.

    When randomize=False (default): only removes Company/Manager elements
    entirely. All other fields (Template, TotalTime, Application, AppVersion,
    HeadingPairs, TitlesOfParts) are preserved as-is.
    When randomize=True: Company / Manager filled with random company name,
    Template filled with random path, TotalTime set to random value.
    """
    data = xml_bytes

    def _ins(val: bytes):
        return lambda m: m.group(1) + val + m.group(2)

    if randomize and random_state:
        # Company and Manager → random company name (sometimes left empty — 15%)
        company_bytes = _maybe_empty(_xml_escape(random_state["company_name"].encode()))
        for local in ["Company", "Manager"]:
            data = re.sub(
                rb"(<" + local.encode() + rb"(?:\s[^>]*)?>)[^<]*(</" + local.encode() + rb">)",
                _ins(company_bytes), data,
            )

        # Template → random template path (sometimes left empty — 15%)
        tmpl_bytes = _maybe_empty(_xml_escape(random_state["template"].encode()))
        data = re.sub(
            rb"(<Template(?:\s[^>]*)?>)[^<]*(</Template>)",
            _ins(tmpl_bytes), data,
        )

        # TotalTime → random value (sometimes 0 — 15%)
        tt_bytes = _maybe_empty(random_state["total_time"].encode())
        data = re.sub(
            rb"(<TotalTime(?:\s[^>]*)?>)[^<]*(</TotalTime>)",
            _ins(tt_bytes), data,
        )

        # Application / AppVersion → realistic values (randomize mode only)
        file_ext = random_state.get("ext", "")
        if file_ext in (".wps", ".et", ".dps"):
            app_val = random.choice(["WPS Office", "WPS"]).encode()
            ver_val = random.choice(["12.1.0.25865", "12.1.0.23125", "11.8.2.12013",
                                      "11.6.0.10872", "11.2.0.10382"]).encode()
        else:
            app_val = random.choice(["Microsoft Office Word", "Microsoft Office Excel",
                                      "Microsoft Office PowerPoint"]).encode()
            ver_val = random.choice(_OFFICE_VERSIONS).encode()
        data = re.sub(
            rb"(<Application(?:\s[^>]*)?>)[^<]*(</Application>)",
            _ins(app_val), data,
        )
        data = re.sub(
            rb"(<AppVersion(?:\s[^>]*)?>)[^<]*(</AppVersion>)",
            _ins(ver_val), data,
        )

        # HeadingPairs / TitlesOfParts → remove in randomize mode (leaks structure)
        data = re.sub(rb"<HeadingPairs>.*?</HeadingPairs>\s*", b"", data, flags=re.DOTALL)
        data = re.sub(rb"<TitlesOfParts>.*?</TitlesOfParts>\s*", b"", data, flags=re.DOTALL)
    else:
        # Only remove identity-leaking fields — delete the entire element.
        # Preserve everything else (Template, TotalTime, Application, etc.)
        for local in ["Company", "Manager"]:
            data = re.sub(
                rb"<" + local.encode() + rb"(?:\s[^>]*)?>"
                rb"[^<]*"
                rb"</" + local.encode() + rb">",
                b"",
                data,
            )
            data = re.sub(
                rb"<" + local.encode() + rb"(?:\s[^>]*)?/>",
                b"",
                data,
            )

    return data


def _clean_document_xml(xml_bytes: bytes, randomize: bool = False) -> bytes:
    """Clear descr attributes in document XML (cNvPr / docPr elements).
    Office stores source file paths of pasted/dragged images in descr attrs.
    Word uses wp:docPr / wp:cNvPr, PowerPoint uses p:cNvPr, Excel uses xdr:cNvPr.

    When randomize=True, also replaces RSID values with random hex strings.
    """
    data = xml_bytes

    if randomize:
        # Collect all unique RSID values and build a per-document mapping
        rsid_pattern = rb'w:rsid\w+="([0-9A-Fa-f]+)"'
        unique_rsids = set(re.findall(rsid_pattern, data))
        if unique_rsids:
            mapping = random_rsid_mapping(unique_rsids)

            def _replace_rsid(m):
                return m.group(1) + mapping.get(m.group(2), m.group(2)) + m.group(3)

            data = re.sub(
                rb'(w:rsid\w+=")([0-9A-Fa-f]+)(")',
                _replace_rsid,
                data,
            )

    # Clear image source paths from descr attrs (both modes)
    data = re.sub(
        rb'(<(?:wp:docPr|wp:cNvPr|pic:cNvPr|p:cNvPr|xdr:cNvPr)[^>]*?)\s+descr="[^"]*"',
        rb"\1", data)
    return data


def clean_office_file(filepath: str, randomize: bool = False,
                      identity: dict | None = None) -> tuple:
    """
    Remove metadata from a .docx / .xlsx / .pptx / .wps / .et / .dps file.
    Streams ZIP entries to avoid loading large files into memory.
    Uses atomic temp-file replacement to prevent data loss on interruption.
    Returns (success: bool, error_message: str|None).

    When randomize=True, metadata is replaced with realistic random values
    instead of being cleared.  identity must be a dict with author_name
    and company_name keys (from generate_unique_identities).
    """
    # Build random_state once per file so sub-functions share the same values
    ext = Path(filepath).suffix.lower()
    if randomize and identity:
        random_state = {
            "author_name": identity["author_name"],
            "company_name": identity["company_name"],
            "times": randomize_times(os.path.getmtime(filepath)),
            "template": random_template(identity["author_name"], ext),
            "total_time": random_total_time(),
            "revision": str(random.randint(1, 50)),
            "ext": ext,
        }
    else:
        random_state = None

    fd, tmp_path = tempfile.mkstemp(dir=os.path.dirname(filepath),
                                     prefix=".~tmp-", suffix=".tmp")
    os.close(fd)

    try:
        with zipfile.ZipFile(filepath, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename

                # Clean core.xml metadata
                if name == "docProps/core.xml":
                    cleaned = _clean_core_xml(zin.read(name), randomize, random_state)
                    zout.writestr(item, cleaned)
                # Clean app.xml metadata (company, manager, template)
                elif name == "docProps/app.xml":
                    cleaned = _clean_app_xml(zin.read(name), randomize, random_state)
                    zout.writestr(item, cleaned)
                # Clean custom.xml — replace with empty skeleton
                elif name == "docProps/custom.xml":
                    zout.writestr(item, _unique_empty_custom_xml())
                # Clean image source paths + optionally RSIDs
                elif (name == "word/document.xml"
                      or name.startswith("ppt/slides/slide")
                      or name.startswith("xl/worksheets/sheet")):
                    cleaned = _clean_document_xml(zin.read(name), randomize)
                    zout.writestr(item, cleaned)
                # Strip EXIF/metadata from embedded images (screenshots, photos)
                elif name.lower().endswith((".png", ".jpg", ".jpeg", ".gif",
                                                ".bmp", ".tiff", ".tif", ".webp",
                                                ".heic")):
                    cleaned = _strip_image_metadata(zin.read(name))
                    zout.writestr(item, cleaned)
                # Strip metadata from document thumbnail (preview of first slide)
                elif name.startswith("docProps/thumbnail."):
                    cleaned = _strip_image_metadata(zin.read(name))
                    zout.writestr(item, cleaned)
                # Clean WPS/Kingsoft custom tags that carry device identifiers
                elif "/tags/" in name and name.endswith(".xml"):
                    cleaned = re.sub(
                        rb'<p:tag\s[^>]*?\bname="(?:KSO_|COMMONDATA)[^"]*"[^>]*/>',
                        b"",
                        zin.read(name),
                    )
                    zout.writestr(item, cleaned)
                # Clean settings.xml — remove w:docId (GUID that links copies)
                elif name == "word/settings.xml":
                    cleaned = zin.read(name)
                    cleaned = re.sub(
                        rb'<w:docId\s[^>]*/>',
                        b'',
                        cleaned,
                    )
                    zout.writestr(item, cleaned)
                # Stream everything else as-is (no memory accumulation)
                else:
                    with zin.open(item) as f_in, \
                         zout.open(item, "w") as f_out:
                        shutil.copyfileobj(f_in, f_out)

        # Preserve original file permissions before atomic replace
        try:
            shutil.copymode(filepath, tmp_path)
        except OSError:
            pass
        os.replace(tmp_path, filepath)

        # Scatter filesystem timestamps within last 3 days.
        # Windows: all three (ctime/atime/mtime) via SetFileTime + rename sync.
        # macOS/Linux: mtime/atime only (ctime is kernel-controlled).
        _scatter_file_timestamps(filepath)

        return True, None

    except Exception as exc:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        if isinstance(exc, PermissionError):
            return False, "文件正被其他程序(如 Office/WPS)占用，请先关闭该文件再试。"
        if isinstance(exc, zipfile.BadZipFile):
            return False, "文件是旧版二进制格式（.doc/.xls/.ppt），请用 Office / WPS 打开后「另存为」新版格式（.docx/.xlsx/.pptx）再试。"
        return False, str(exc)


# ============================================================
# PDF raw-byte post-processing (single read/write)
# ============================================================

_XMP_ATTRS = (
    "xmp:CreateDate", "xmp:ModifyDate", "xmp:MetadataDate",
    "xmp:CreatorTool", "xmpMM:DocumentID", "xmpMM:InstanceID",
    "xmpMM:OriginalDocumentID", "dc:format", "photoshop:ColorMode",
    "photoshop:ICCProfile", "stEvt:softwareAgent", "stEvt:when",
    "stEvt:action",
)

def _strip_pdf_raw_bytes(path: str) -> None:
    """Single-pass raw-byte cleaning of a pypdf-written PDF.
    - /Alt image source paths → equal-width space padding
    - /URI links → equal-width space padding
    - Per-image XMP attributes → empty values
    - Document-level XMP streams → blanked with spaces"""
    import re

    with open(path, "rb") as f:
        content = f.read()

    # --- /Alt image source paths ---
    def _pad_alt(m: re.Match) -> bytes:
        prefix, inner, suffix = m.group(1), m.group(2), m.group(3)
        return prefix + b" " * len(inner) + suffix

    content = re.sub(rb"(/Alt\s?\()((?:[^\\]|\\.)*?)(\))", _pad_alt, content)

    # --- /URI links ---
    def _pad_uri(m: re.Match) -> bytes:
        prefix, inner, suffix = m.group(1), m.group(2), m.group(3)
        return prefix + b" " * len(inner) + suffix

    content = re.sub(rb"(/URI\s?\()((?:[^\\]|\\.)*?)(\))", _pad_uri, content)

    # --- Per-image XMP attributes ---
    # Use equal-width space padding to preserve stream /Length and XREF offsets.
    # Clearing values to "" would change byte count and corrupt the PDF.
    def _pad_xmp_attr(m: re.Match) -> bytes:
        prefix = m.group(1)   # e.g. xmp:CreatorTool="
        value = m.group(2)    # the value between quotes
        suffix = m.group(3)   # closing "
        return prefix + b" " * len(value) + suffix

    for attr in _XMP_ATTRS:
        content = re.sub(
            rb'(' + attr.encode() + rb'\s*=\s*")([^"]*)(")',
            _pad_xmp_attr, content)

    # --- Document-level XMP streams ---
    # Only match complete XMP packets (<?xpacket begin= ... <?xpacket end=).
    # Avoid matching standalone <?xpacket end="w"?> inside ICC colour profiles.
    # Use a bounded repeat instead of .*? with re.DOTALL to prevent catastrophic
    # matches across large swaths of binary data.
    def _pad_xmp(m: re.Match) -> bytes:
        return b" " * len(m.group())

    content = re.sub(
        rb'<\?xpacket begin=[\x00-\xff]{0,50000}?<\?xpacket end=',
        _pad_xmp, content, flags=re.DOTALL)
    # Also match any standalone x:xmpmeta blocks not wrapped in xpacket
    content = re.sub(
        rb'<x:xmpmeta[\x00-\xff]{0,50000}?</x:xmpmeta>',
        _pad_xmp, content, flags=re.DOTALL)

    with open(path, "wb") as f:
        f.write(content)


def clean_pdf_file(filepath: str, randomize: bool = False,
                   identity: dict | None = None) -> tuple:
    """
    Remove metadata from a PDF file — info dict, XMP (doc + image level),
    /Alt image source paths, /URI links, and Producer.
    Returns (success: bool, error_message: str|None).
    """
    if not HAS_PDF_SUPPORT:
        return False, "PDF 组件未安装，请运行 setup.bat (Windows) 或 setup.sh (macOS) 安装依赖。"

    fd, tmp_path = tempfile.mkstemp(dir=os.path.dirname(filepath),
                                     prefix=".~tmp-", suffix=".tmp")
    os.close(fd)

    try:
        with open(filepath, "rb") as fin:
            reader = PdfReader(fin)
            writer = PdfWriter()

            if hasattr(writer, "clone_document_from_reader"):
                writer.clone_document_from_reader(reader)
            else:
                for page in reader.pages:
                    writer.add_page(page)

            # Clear all info-dict fields — remove keys entirely from the
            # internal dict so pypdf doesn't write space-padded stubs.
            # Using add_metadata(" ") leaves "( )" in the file, which is
            # a detectable tool fingerprint.
            _fields_to_clear = [
                "/Producer", "/Creator", "/CreationDate", "/ModDate",
                "/Title", "/Author", "/Subject", "/Keywords",
                "/Comments", "/Company", "/SourceModified",
            ]
            info_obj = writer._info.get_object() if hasattr(writer._info, 'get_object') else writer._info
            for key in _fields_to_clear:
                info_obj.pop(key, None)

            with open(tmp_path, "wb") as fout:
                writer.write(fout)

        # Single-pass raw-byte post-processing
        _strip_pdf_raw_bytes(tmp_path)

        try:
            shutil.copymode(filepath, tmp_path)
        except OSError:
            pass
        os.replace(tmp_path, filepath)
        return True, None

    except Exception as exc:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        if isinstance(exc, PermissionError):
            return False, "文件正被其他程序占用，请先关闭该 PDF 再试。"
        return False, str(exc)



# ============================================================
# Unified entry point
# ============================================================


def clean_image_file(filepath: str, randomize: bool = False,
                     identity: dict | None = None) -> tuple:
    """
    Remove metadata from a standalone image file (PNG / JPEG / GIF / etc).
    Uses atomic temp-file replacement. Returns (success, error_message).
    """
    fd, tmp_path = tempfile.mkstemp(dir=os.path.dirname(filepath),
                                     prefix=".~tmp-", suffix=".tmp")
    os.close(fd)

    try:
        with open(filepath, "rb") as f:
            original = f.read()

        cleaned = _strip_image_metadata(original)

        with open(tmp_path, "wb") as f:
            f.write(cleaned)

        try:
            shutil.copymode(filepath, tmp_path)
        except OSError:
            pass
        os.replace(tmp_path, filepath)
        return True, None

    except Exception as exc:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        if isinstance(exc, PermissionError):
            return False, "文件正被其他程序占用，请先关闭再试。"
        return False, str(exc)


def clean_file(filepath: str, randomize: bool = False,
               identity: dict | None = None) -> tuple:
    """Clean a single file. Returns (success, error_message)."""
    ext = Path(filepath).suffix.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        if ext in {".doc", ".xls", ".ppt"}:
            new_ext = ".docx" if ext == ".doc" else (".xlsx" if ext == ".xls" else ".pptx")
            return False, f"旧版二进制格式 {ext} 不支持，请用 Office / WPS 打开后「另存为」{new_ext} 格式再试。"
        if ext in {".wps1", ".wpt"}:
            return False, f"旧版 WPS 格式 {ext} 不支持，请用 WPS 打开后「另存为」新版格式（.wps/.et/.dps）再试。"
        return False, f"不支持的文件类型: {ext}"

    if ext == ".pdf":
        return clean_pdf_file(filepath, randomize, identity)
    elif ext in IMAGE_EXTENSIONS:
        return clean_image_file(filepath, randomize, identity)
    else:
        return clean_office_file(filepath, randomize, identity)


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
                with z.open(doc_xml_path, "r") as fh:
                    for chunk in iter(lambda: fh.read(65536), b""):
                        if b"<w:ins " in chunk or b"<w:del " in chunk or \
                           b"<w:ins>" in chunk or b"<w:del>" in chunk:
                            warnings.append("修订记录")
                            break

    except Exception:
        pass  # scanning is best-effort; don't block the user over it

    return warnings


# ============================================================
# Metadata reading (view original metadata before cleaning)
# ============================================================


def _read_office_metadata(filepath: str) -> dict:
    """Extract metadata from an Office file's XML parts. Returns a dict of
    {category: {chinese_label: value}} for display."""
    result: dict[str, dict[str, str]] = {}

    try:
        with zipfile.ZipFile(filepath, "r") as z:
            names = frozenset(z.namelist())

            # --- core.xml -------------------------------------------------
            if "docProps/core.xml" in names:
                raw = z.read("docProps/core.xml").decode("utf-8", errors="replace")
                fields = {
                    "dc:title": "标题",
                    "dc:subject": "主题",
                    "dc:creator": "作者",
                    "cp:keywords": "关键词",
                    "dc:description": "描述",
                    "cp:lastModifiedBy": "最后修改者",
                    "cp:revision": "修订号",
                    "cp:version": "版本",
                    "cp:category": "类别",
                    "cp:contentStatus": "内容状态",
                    "dcterms:created": "创建时间",
                    "dcterms:modified": "修改时间",
                    "cp:lastPrinted": "最后打印",
                }
                core = {}
                for tag, label in fields.items():
                    # Match both <tag>text</tag> and <tag />
                    m = re.search(
                        r"<" + tag + r"(?:\s[^>]*)?>([^<]*)</" + tag + r">",
                        raw,
                    )
                    if m:
                        val = m.group(1).strip()
                        if val:
                            core[label] = val
                if core:
                    result["文档属性 (core.xml)"] = core

            # --- app.xml --------------------------------------------------
            if "docProps/app.xml" in names:
                raw = z.read("docProps/app.xml").decode("utf-8", errors="replace")
                fields = {
                    "Application": "创建程序",
                    "AppVersion": "程序版本",
                    "Template": "模板",
                    "TotalTime": "编辑时长(分钟)",
                    "Pages": "页数",
                    "Words": "字数",
                    "Characters": "字符数",
                    "Lines": "行数",
                    "Paragraphs": "段落数",
                    "Company": "公司",
                    "Manager": "管理者",
                    "ScaleCrop": "缩放裁剪",
                    "LinksUpToDate": "链接最新",
                    "SharedDoc": "共享文档",
                    "HyperlinksChanged": "超链接已更改",
                }
                app = {}
                for tag, label in fields.items():
                    m = re.search(
                        r"<" + tag + r"(?:\s[^>]*)?>([^<]*)</" + tag + r">",
                        raw,
                    )
                    if m:
                        val = m.group(1).strip()
                        if val:
                            app[label] = val
                if app:
                    result["扩展属性 (app.xml)"] = app

            # --- custom.xml -----------------------------------------------
            if "docProps/custom.xml" in names:
                raw = z.read("docProps/custom.xml").decode("utf-8", errors="replace")
                # Extract custom property names and values
                props = {}
                for m in re.finditer(
                    r'name="([^"]+)".*?<vt:lpstr>([^<]*)</vt:lpstr>', raw
                ):
                    props[m.group(1)] = m.group(2)
                for m in re.finditer(
                    r'name="([^"]+)".*?<vt:lpwstr>([^<]*)</vt:lpwstr>', raw
                ):
                    props[m.group(1)] = m.group(2)
                for m in re.finditer(
                    r'name="([^"]+)".*?<vt:i4>([^<]*)</vt:i4>', raw
                ):
                    props[m.group(1)] = m.group(2)
                for m in re.finditer(
                    r'name="([^"]+)".*?<vt:bool>([^<]*)</vt:bool>', raw
                ):
                    props[m.group(1)] = m.group(2)
                for m in re.finditer(
                    r'name="([^"]+)".*?<vt:r8>([^<]*)</vt:r8>', raw
                ):
                    props[m.group(1)] = m.group(2)
                for m in re.finditer(
                    r'name="([^"]+)".*?<vt:filetime>([^<]*)</vt:filetime>', raw
                ):
                    props[m.group(1)] = m.group(2)
                if props:
                    result["自定义属性 (custom.xml)"] = props

            # --- embedded image metadata ----------------------------------
            image_meta: dict[str, str] = {}
            for name in z.namelist():
                if name.lower().endswith((".png", ".jpg", ".jpeg", ".gif",
                                           ".bmp", ".tiff", ".tif", ".webp",
                                           ".heic")):
                    try:
                        raw = z.read(name)
                        info = _read_image_metadata(raw)
                        if info:
                            # Use filename + metadata as key
                            short_name = name.rsplit("/", 1)[-1]
                            for key, val in info.items():
                                image_meta[f"{short_name} → {key}"] = val
                    except Exception:
                        pass
            if image_meta:
                result["嵌入图片元数据"] = image_meta

            # --- image source paths in document/slide/sheet XML -------------
            descr_paths: dict[str, str] = {}
            for name in z.namelist():
                if (name == "word/document.xml"
                        or name.startswith("ppt/slides/slide")
                        or name.startswith("xl/worksheets/sheet")):
                    raw = z.read(name).decode("utf-8", errors="replace")
                    for m in re.finditer(r'descr="([^"]+)"', raw):
                        path = m.group(1)
                        # Only collect paths that look like file paths
                        if "\\" in path or "/" in path:
                            # Use parent element name as key
                            ctx = raw[max(0, m.start() - 200):m.start()]
                            name_m = re.search(r'name="([^"]+)"', ctx)
                            img_name = name_m.group(1) if name_m else "?"
                            descr_paths[img_name] = path
            if descr_paths:
                result["图片来源路径 (document.xml)"] = descr_paths

    except Exception as exc:
        result["错误"] = {"读取失败": str(exc)}

    return result


def _read_image_metadata(data: bytes) -> dict[str, str]:
    """Extract metadata from raw image bytes. Returns {label: value}."""
    if not HAS_IMAGE_SUPPORT:
        return {}
    try:
        import io

        img = Image.open(io.BytesIO(data))
        info: dict[str, str] = {}

        if img.format == "PNG":
            for key, val in (img.text or {}).items():
                info[key] = val

        # EXIF may be present in PNG (eXIf chunk), JPEG, TIFF, and HEIF
        if img.format in ("PNG", "JPEG", "TIFF", "HEIF"):
            exif = img.getexif()
            if exif:
                # TIFF/EXIF tags that describe image encoding, not user metadata
                _STRUCTURAL_TAGS = {
                    254, 256, 257, 258, 259, 262, 266, 273, 274,
                    277, 278, 279, 282, 283, 284, 296, 317, 320,
                    322, 323, 324, 325, 330, 338, 339, 513, 514,
                    530, 531, 532, 33421, 33422,
                }
                for tag_id, val in exif.items():
                    from PIL.ExifTags import TAGS
                    if tag_id in _STRUCTURAL_TAGS:
                        continue
                    tag_name = TAGS.get(tag_id, f"Tag{tag_id}")
                    if val and tag_name not in ("ExifOffset", "MakerNote"):
                        info[tag_name] = str(val).strip()

        if img.format == "GIF":
            if hasattr(img, "info"):
                for key in ("comment", "duration"):
                    val = img.info.get(key)
                    if val:
                        info[key] = str(val)

        return info
    except Exception:
        return {}


def _decode_utf16be(raw_bytes: bytes) -> str:
    """Decode a UTF-16BE byte string (without BOM).  Handles malformed input."""
    try:
        return raw_bytes.decode("utf-16-be")
    except (UnicodeDecodeError, UnicodeError):
        return raw_bytes.decode("utf-16-be", errors="replace")


def _scan_pdf_alt_text(content: bytes) -> list[str]:
    """Scan raw PDF bytes for /Alt entries with UTF-16BE encoded paths.
    Searches both uncompressed objects and ObjStm compressed streams."""
    paths: list[str] = []
    # Match /Alt(..) — both plain ASCII and UTF-16BE BOM variants
    for m in re.finditer(rb"/Alt\((.*?)\)", content):
        inner = m.group(1)
        if not inner:
            continue
        # UTF-16BE with BOM
        if inner.startswith(b"\xfe\xff"):
            text = _decode_utf16be(inner[2:])
            text = text.replace("\x00", "").strip()
            if text and text not in paths:
                paths.append(text)
        # Plain ASCII (no BOM, detectable printable text)
        else:
            try:
                text = inner.decode("ascii")
                # Heuristic: only keep if looks like a path or URL
                if text and ("/" in text or "\\" in text or ":" in text) and text not in paths:
                    paths.append(text)
            except UnicodeDecodeError:
                pass
    return paths


def _scan_pdf_uri_links(content: bytes) -> list[str]:
    """Scan raw PDF bytes for /URI link annotations."""
    uris: list[str] = []
    # Match PDF URI actions: /URI(http://...) or /URI(https://...)
    for m in re.finditer(rb"/URI\s*\((.*?)\)", content):
        inner = m.group(1)
        try:
            text = inner.decode("ascii").strip()
            if text and text not in uris:
                uris.append(text)
        except UnicodeDecodeError:
            try:
                text = inner.decode("utf-8", errors="replace").strip()
                if text and text not in uris:
                    uris.append(text)
            except UnicodeDecodeError:
                pass
    return uris


def _scan_pdf_xmp_metadata(content: bytes) -> dict[str, dict[str, str]]:
    """Extract per-image XMP metadata from PDF content.
    Returns a summary: {field_label: {unique_value: count}}."""
    import re as _re

    result: dict[str, dict[str, str]] = {}

    # XMP fields worth reporting (label -> attribute name)
    xmp_fields = [
        ("创建工具", "xmp:CreatorTool"),
        ("软件代理", "stEvt:softwareAgent"),
        ("创建时间", "xmp:CreateDate"),
        ("修改时间", "xmp:ModifyDate"),
        ("元数据日期", "xmp:MetadataDate"),
        ("标题", "dc:title"),
        ("描述", "dc:description"),
        ("作者", "dc:creator"),
        ("主题", "dc:subject"),
        ("权利", "dc:rights"),
        ("来源", "photoshop:Source"),
        ("评级", "xmp:Rating"),
    ]

    for label, attr in xmp_fields:
        # Match attribute="value" or attribute='value'
        pattern = (attr.encode() + rb'\s*=\s*"([^"]*)"').replace(rb'\\', rb'\\\\')
        pattern_alt = (attr.encode() + rb"\s*=\s*'([^']*)'")

        values: dict[str, int] = {}
        for m in _re.finditer(pattern, content):
            val = m.group(1).decode("utf-8", errors="replace").strip()
            if val:
                values[val] = values.get(val, 0) + 1
        for m in _re.finditer(pattern_alt, content):
            val = m.group(1).decode("utf-8", errors="replace").strip()
            if val:
                values[val] = values.get(val, 0) + 1

        if values:
            display: dict[str, str] = {}
            for val, count in values.items():
                display[val] = f"{count} 张图片" if count > 1 else "1 张图片"
            result[label] = display

    return result


def _scan_pdf_hidden_text(raw: bytes) -> dict[str, list[str]]:
    """Scan raw PDF bytes for /Alt image paths and /URI links.
    Only scans uncompressed content — fast enough for 64MB+ files."""
    result: dict[str, list[str]] = {}

    alt_paths = _scan_pdf_alt_text(raw)
    uris = _scan_pdf_uri_links(raw)

    if alt_paths:
        result["图片源路径"] = alt_paths
    if uris:
        result["链接 URL"] = uris

    return result


def _extract_structure_tree_alts(reader) -> list[str]:
    """Traverse the PDF structure tree via pypdf and extract all /Alt text.
    This catches entries inside compressed objects that raw-byte scanning misses."""
    alts: list[str] = []

    def _walk(obj, depth: int = 0) -> None:
        if depth > 8:
            return
        if hasattr(obj, "get_object"):
            obj = obj.get_object()
        if not hasattr(obj, "keys"):
            return
        if "/Alt" in obj:
            val = obj["/Alt"]
            if hasattr(val, "get_object"):
                val = val.get_object()
            text = str(val).strip()
            if text:
                alts.append(text)
        if "/K" in obj:
            kids = obj["/K"]
            if hasattr(kids, "get_object"):
                kids = kids.get_object()
            items = kids if isinstance(kids, list) else [kids]
            for kid in items:
                _walk(kid, depth + 1)

    try:
        root = reader.trailer["/Root"]
        if "/StructTreeRoot" in root:
            struct = root["/StructTreeRoot"].get_object()
            _walk(struct)
    except Exception:
        pass

    # Deduplicate while preserving order
    seen: set[str] = set()
    unique: list[str] = []
    for text in alts:
        if text not in seen:
            seen.add(text)
            unique.append(text)

    # Cap at 40 entries to keep the metadata window manageable
    return unique[:40]


def _read_pdf_metadata(filepath: str) -> dict:
    """Extract metadata from a PDF file — standard info dict, XMP, deep scan,
    page count.  Reads the file once and scans in-memory for best performance."""
    result: dict[str, dict[str, str]] = {}
    if not HAS_PDF_SUPPORT:
        return {"错误": {"PDF 支持": "未安装 pypdf/PyPDF2 库"}}

    import io

    # Read file once into memory, then pass BytesIO to pypdf
    try:
        with open(filepath, "rb") as f:
            raw = f.read()
    except OSError as exc:
        result["错误"] = {"读取失败": str(exc)}
        return result

    # Standard metadata via pypdf (uses BytesIO to avoid re-reading)
    try:
        reader = PdfReader(io.BytesIO(raw))
        info = reader.metadata
        pdf: dict[str, str] = {}

        # --- PDF info dict fields ---
        if info:
            info_fields = {
                "/Title": "标题",
                "/Author": "作者",
                "/Subject": "主题",
                "/Keywords": "关键词",
                "/Creator": "创建程序",
                "/Producer": "生成工具",
                "/CreationDate": "创建时间",
                "/ModDate": "修改时间",
                "/Comments": "注释",
                "/Company": "公司",
                "/SourceModified": "源修改时间",
            }
            for key, label in info_fields.items():
                val = getattr(info, key[1:].lower(), None) or info.get(key, None)
                if val and str(val).strip():
                    val_str = str(val).strip()
                    if val_str.startswith("D:"):
                        val_str = val_str[2:].replace("'", "")
                    pdf[label] = val_str

        # --- Page count ---
        try:
            pdf["页数"] = str(len(reader.pages))
        except Exception:
            pass

        # --- Document-level XMP metadata ---
        try:
            xmp = reader.xmp_metadata
            if xmp:
                xmp_fields = {
                    "xmp_creator_tool": "XMP-创建工具",
                    "xmp_create_date": "XMP-创建时间",
                    "xmp_modify_date": "XMP-修改时间",
                    "xmp_metadata_date": "XMP-元数据日期",
                    "xmpmm_document_id": "XMP-文档ID",
                    "dc_title": "XMP-标题",
                    "dc_description": "XMP-描述",
                    "dc_format": "XMP-格式",
                }
                for attr, label in xmp_fields.items():
                    val = getattr(xmp, attr, None)
                    if val:
                        val_str = str(val)
                        # Unwrap XMP lang-alt dicts
                        if isinstance(val, dict):
                            val_str = str(val.get("x-default", list(val.values())[0] if val else ""))
                        if val_str.strip():
                            pdf[label] = val_str.strip()
        except Exception:
            pass  # XMP reading can fail on malformed metadata

        # --- Structure tree /Alt image source paths ---
        struct_alts = _extract_structure_tree_alts(reader)
        if struct_alts:
            indexed: dict[str, str] = {}
            for i, alt in enumerate(struct_alts):
                key = f"[{i+1}]" if len(struct_alts) > 1 else ""
                indexed[key] = str(alt)
            result["图片源路径"] = indexed

        if pdf:
            result["PDF 属性"] = pdf

    except Exception as exc:
        result["错误"] = {"读取失败": str(exc)}

    # Deep scan for /URI links and raw /Alt (uncompressed fallback)
    deep = _scan_pdf_hidden_text(raw)
    for category, items in deep.items():
        # Skip "图片源路径" from raw scan if we already have struct-tree results
        if category == "图片源路径" and "图片源路径" in result:
            continue
        if items:
            indexed = {}
            for i, item in enumerate(items):
                key = f"[{i+1}]" if len(items) > 1 else ""
                indexed[key] = str(item)
            result[category] = indexed

    # XMP per-image metadata (scan raw bytes — XMP is typically uncompressed)
    xmp = _scan_pdf_xmp_metadata(raw)
    for label, values in xmp.items():
        if values:
            indexed: dict[str, str] = {}
            if len(values) == 1:
                val, count = next(iter(values.items()))
                indexed[""] = f"{val}（{count}）"
            else:
                for i, (val, count) in enumerate(values.items()):
                    indexed[f"[{i+1}]"] = f"{val}（{count}）"
            result[f"图片XMP-{label}"] = indexed

    return result


def _read_image_file_metadata(filepath: str) -> dict:
    """Extract metadata from a standalone image file."""
    result: dict[str, dict[str, str]] = {}
    try:
        with open(filepath, "rb") as f:
            data = f.read()
        info = _read_image_metadata(data)
        if info:
            result["图片元数据"] = info
        else:
            result["图片元数据"] = {"（无）": "此图片不含元数据"}
    except Exception as exc:
        result["错误"] = {"读取失败": str(exc)}
    return result


def read_metadata(filepath: str) -> dict:
    """Read metadata from an Office file, PDF, or standalone image.
    Returns {category: {label: value}} or {"错误": {reason: detail}}.
    """
    ext = Path(filepath).suffix.lower()
    if ext == ".pdf":
        return _read_pdf_metadata(filepath)
    elif ext in OFFICE_EXTENSIONS:
        return _read_office_metadata(filepath)
    elif ext in IMAGE_EXTENSIONS:
        return _read_image_file_metadata(filepath)
    if ext in {".doc", ".xls", ".ppt"}:
        new_ext = ".docx" if ext == ".doc" else (".xlsx" if ext == ".xls" else ".pptx")
        return {"错误": {"不支持": f"旧版格式 {ext}，请用 Office/WPS「另存为」{new_ext} 后重试"}}
    if ext in {".wps1", ".wpt"}:
        return {"错误": {"不支持": f"旧版 WPS 格式 {ext}，请用 WPS「另存为」新版格式后重试"}}
    return {"错误": {"不支持": f"文件类型 {ext} 不支持"}}


# ============================================================
# GUI
# ============================================================


class MetadataCleanerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"元数据清除工具 v{__version__}")
        self.root.geometry("1200x800")
        self.root.minsize(960, 600)

        # Set window icon (Windows taskbar and thumbnail)
        if sys.platform == "win32":
            # In PyInstaller onefile mode, icon is bundled via --add-data
            if getattr(sys, "_MEIPASS", None):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            else:
                icon_path = "icon.ico"
            if os.path.exists(icon_path):
                try:
                    # Use PhotoImage as fallback if iconbitmap fails
                    from PIL import Image
                    img = Image.open(icon_path)
                    photo = tk.PhotoImage(file=icon_path)
                    self.root.iconphoto(True, photo)
                except Exception:
                    try:
                        self.root.iconbitmap(icon_path)
                    except Exception:
                        pass  # icon loading failed, continue without icon

        self.files: list[str] = []
        self._cleaning = False  # guard against closing during processing
        self._build_ui()
        self._load_config()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # -- UI construction --------------------------------------------------

    def _build_ui(self):
        main = ttk.Frame(self.root, padding="12")
        main.pack(fill=tk.BOTH, expand=True)

        # Header row: title on left, about button on right
        header = ttk.Frame(main)
        header.pack(fill=tk.X, pady=(0, 2))

        ttk.Label(header, text="元数据清除工具", font=("", 16, "bold")).pack(side=tk.LEFT)

        self.btn_about = ttk.Button(header, text="说明", command=self._show_about, width=6)
        self.btn_about.pack(side=tk.RIGHT)

        ttk.Label(
            main,
            text="支持 Word / Excel / PPT (.docx/.xlsx/.pptx)  |  WPS (.wps/.et/.dps)  |  PDF  |  图片 (.jpg/.png/.gif/.bmp/.tiff/.webp/.heic)",
            font=("", 10),
        ).pack(pady=(0, 12))

        # Disclaimer: what this tool does NOT clean
        ttk.Label(
            main,
            text="注：本工具清除文件属性元数据，不清除文档正文中的批注与修订记录。",
            font=("", 11, "bold"),
            foreground="#c0392b",
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

        # Two-column layout: left = file list, right = metadata panel
        self.paned = ttk.PanedWindow(main, orient=tk.HORIZONTAL)

        # === Left column: file list ===
        left = ttk.Frame(self.paned)
        self.paned.add(left, weight=1)

        list_frame = ttk.LabelFrame(left, text="文件列表", padding="6")
        list_frame.pack(fill=tk.BOTH, expand=True)

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

        # Right-click context menu
        self.listbox.bind("<Button-2>" if sys.platform == "darwin" else "<Button-3>",
                          self._on_right_click)

        # Drag-and-drop registration
        if HAS_DND:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)
            if hasattr(self.root, 'drop_target_register'):
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)

        # === Right column: metadata panel ===
        right = ttk.Frame(self.paned)
        self.paned.add(right, weight=2)

        meta_frame = ttk.LabelFrame(right, text="元数据信息", padding="6")
        meta_frame.pack(fill=tk.BOTH, expand=True)

        meta_inner = ttk.Frame(meta_frame)
        meta_inner.pack(fill=tk.BOTH, expand=True)

        meta_sb = ttk.Scrollbar(meta_inner)
        meta_sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.meta_text = tk.Text(
            meta_inner, wrap=tk.WORD, font=("", 11),
            padx=8, pady=8, state=tk.DISABLED,
            yscrollcommand=meta_sb.set,
        )
        self.meta_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        meta_sb.config(command=self.meta_text.yview)

        # Metadata display tags
        self.meta_text.tag_configure("cat_header", foreground="#B45309", font=("", 11, "bold"), spacing3=4)
        self.meta_text.tag_configure("cat_warn", foreground="#C2410C", font=("", 11, "bold"), spacing3=4)
        self.meta_text.tag_configure("label", foreground="#334155")
        self.meta_text.tag_configure("value", foreground="#0F172A")
        self.meta_text.tag_configure("value_warn", foreground="#B91C1C")
        self.meta_text.tag_configure("error", foreground="#DC2626", font=("", 11, "bold"))

        # Placeholder
        self.meta_text.configure(state=tk.NORMAL)
        self.meta_text.insert(tk.END, "请在左侧选择文件以查看元数据", "label")
        self.meta_text.configure(state=tk.DISABLED)

        # Auto-update on selection change
        self.listbox.bind("<<ListboxSelect>>", self._on_selection_change)

        # === Bottom bar: controls spanning full width ===
        # Pack bottom-to-top with side=BOTTOM so shrinking the window
        # squeezes the paned first, keeping buttons always visible.

        ready_text = "就绪 — 请添加文件"
        if not HAS_PDF_SUPPORT:
            ready_text += "  (PDF 功能需安装 pypdf)"
        self.status_var = tk.StringVar(value=ready_text)
        ttk.Label(main, textvariable=self.status_var, font=("", 9)).pack(side=tk.BOTTOM, anchor=tk.W)

        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 4))

        btn_frame = ttk.Frame(main)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 10))

        self.btn_add = ttk.Button(btn_frame, text="添加文件", command=self._add_files)
        self.btn_add.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_remove = ttk.Button(btn_frame, text="移除选中", command=self._remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_clear = ttk.Button(btn_frame, text="清空列表", command=self._clear_files)
        self.btn_clear.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_clean = tk.Label(
            btn_frame,
            text="  清除元数据  ",
            bg="#e74c3c", fg="white",
            font=("", 13, "bold"),
            padx=20, pady=8,
            relief=tk.RAISED,
            borderwidth=2,
            cursor="hand2",
        )
        self.btn_clean.pack(side=tk.RIGHT)
        self.btn_clean.bind("<Button-1>", lambda e: self._on_clean_click())
        self.btn_clean.bind("<Enter>", lambda e: self.btn_clean.config(bg="#c0392b"))
        self.btn_clean.bind("<Leave>", lambda e: self.btn_clean.config(bg="#e74c3c"))

        rand_frame = ttk.LabelFrame(main, text="处理模式（可选）", padding="8")
        rand_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 8))

        self.randomize_var = tk.BooleanVar(value=False)
        self.chk_random = ttk.Checkbutton(
            rand_frame,
            text="随机化：替换元数据为随机仿真值（人名、公司名、时间戳等）。不勾选 = 清空元数据",
            variable=self.randomize_var,
        )
        self.chk_random.pack(side=tk.LEFT)

        # Pack paned last so it fills remaining space above the bottom bar
        self.paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 6))

    # -- Button callbacks -------------------------------------------------

    def _add_files(self):
        types = [
            ("支持的文件", "*.docx *.xlsx *.pptx *.wps *.et *.dps *.pdf *.png *.jpg *.jpeg *.gif *.bmp *.tiff *.tif *.webp"),
            ("Word / WPS 文档", "*.docx *.wps"),
            ("Excel / WPS 表格", "*.xlsx *.et"),
            ("PowerPoint / WPS 演示", "*.pptx *.dps"),
            ("PDF 文件", "*.pdf"),
            ("图片文件", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff *.tif *.webp"),
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

    def _on_right_click(self, event):
        """Show context menu on right-click."""
        idx = self.listbox.nearest(event.y)
        if idx >= 0 and idx < len(self.files):
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)

        menu = tk.Menu(self.root, tearoff=0)
        if not self._cleaning:
            menu.add_command(label="移除选中", command=self._remove_selected)
            menu.add_command(label="清空列表", command=self._clear_files)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _on_selection_change(self, event=None):
        """Update the right-side metadata panel when a file is selected."""
        selected = self.listbox.curselection()

        self.meta_text.configure(state=tk.NORMAL)
        try:
            self.meta_text.delete("1.0", tk.END)

            if not selected:
                self.meta_text.insert(tk.END, "请在左侧选择文件以查看元数据", "label")
                return

            idx = selected[0]
            filepath = self.files[idx]
            fname = os.path.basename(filepath)

            if len(selected) > 1:
                self.meta_text.insert(tk.END, f"已选 {len(selected)} 个文件，仅显示首项\n\n", "label")

            try:
                metadata = read_metadata(filepath)
            except Exception as e:
                metadata = {"错误": {"读取失败": str(e)}}

            WARN_CATEGORIES = {
                "图片源路径", "链接 URL", "批注", "修订",
                "嵌入图片元数据", "批注和修订",
            }

            def _is_warn_cat(cat):
                return cat in WARN_CATEGORIES or cat.startswith("图片XMP-") or cat.startswith("嵌入图片")

            self.meta_text.insert(tk.END, f"{fname}\n", "cat_header")

            if "错误" in metadata:
                for key, val in metadata["错误"].items():
                    self.meta_text.insert(tk.END, f"\n{key}: ", "error")
                    self.meta_text.insert(tk.END, f"{val}\n", "error")
            else:
                for category, fields in metadata.items():
                    header_tag = "cat_warn" if _is_warn_cat(category) else "cat_header"
                    self.meta_text.insert(tk.END, f"\n── {category} ──\n\n", header_tag)
                    for label, value in fields.items():
                        display_val = str(value)
                        lbl_key = label + ":  " if label else ""
                        val_tag = "value_warn" if _is_warn_cat(category) else "value"
                        self.meta_text.insert(tk.END, f"  {lbl_key}", "label")
                        self.meta_text.insert(tk.END, f"{display_val}\n", val_tag)
        finally:
            self.meta_text.configure(state=tk.DISABLED)

    def _on_drop(self, event):
        """Handle file drop from OS file manager (requires tkinterdnd2)."""
        paths = self.root.tk.splitlist(event.data)

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

    def _show_about(self):
        """Show about dialog with program info and usage help."""
        dlg = tk.Toplevel(self.root)
        dlg.title("说明")
        dlg.geometry("450x680")
        dlg.transient(self.root)
        dlg.grab_set()

        content = ttk.Frame(dlg, padding="16")
        content.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(content, wrap=tk.WORD, font=("", 11), state=tk.DISABLED, height=32)
        sb = ttk.Scrollbar(content, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        txt.tag_configure("title", font=("", 14, "bold"), foreground="#1a1a1a")
        txt.tag_configure("heading", font=("", 11, "bold"), foreground="#333333", spacing3=4)
        txt.tag_configure("body", foreground="#444444", spacing3=1)
        txt.tag_configure("highlight", foreground="#c0392b", font=("", 11, "bold"), spacing3=4)

        txt.configure(state=tk.NORMAL)

        # Title
        txt.insert(tk.END, "元数据清除工具\n", "title")
        txt.insert(tk.END, f"版本 {__version__}\n\n", "body")

        # Copyright
        txt.insert(tk.END, "版权声明\n", "heading")
        txt.insert(tk.END, "本程序由 Asura 开发，保留所有权利。\n\n", "body")

        # Supported formats
        txt.insert(tk.END, "支持文件类型\n", "heading")
        txt.insert(tk.END, "Microsoft Office: .docx / .xlsx / .pptx\n", "body")
        txt.insert(tk.END, "WPS Office: .wps / .et / .dps\n", "body")
        txt.insert(tk.END, "PDF 文档\n", "body")
        txt.insert(tk.END, "图片文件: .jpg / .png / .gif / .bmp / .tiff / .webp / .heic\n\n", "body")

        # Cleanable metadata — clear mode
        txt.insert(tk.END, "清空模式（默认）\n", "heading")
        txt.insert(tk.END, "仅删除暴露身份的字段，其余保留原值：\n", "body")
        txt.insert(tk.END, "• 删除 — 作者、最后修改者、公司、管理者\n", "body")
        txt.insert(tk.END, "• 保留 — 时间戳、修订号、标题、模板、编辑时长等\n", "body")
        txt.insert(tk.END, "• 嵌入图片 EXIF 元数据 — 微信/QQ/截图工具粘贴的图片携带的信息\n", "body")
        txt.insert(tk.END, "• 图片源路径 — document.xml descr 属性中最隐蔽的泄露点\n\n", "body")

        # Randomization feature
        txt.insert(tk.END, "随机化模式（可选）\n", "heading")
        txt.insert(tk.END, "勾选后，所有元数据替换为随机仿真值：\n", "body")
        txt.insert(tk.END, "• 人名 — 拼音连写/英文名/键盘缩写等多风格混合\n", "body")
        txt.insert(tk.END, "• 公司 — 科技行业品牌简称，无城市前缀，极少后缀\n", "body")
        txt.insert(tk.END, "• 时间戳 — 基于文件实际时间锚定，偏差 ≤ 2 天，避开凌晨\n", "body")
        txt.insert(tk.END, "• 批次唯一保障 — 同批文件的人名和公司名绝不重复\n", "body")
        txt.insert(tk.END, "• 抗取证设计 — custom.xml 随机填充、时间戳散列\n", "body")
        txt.insert(tk.END, "• 字段随机留空 — 模仿真实文件不完全填满的状态\n\n", "body")

        # Privacy highlight
        txt.insert(tk.END, "★★★ 全程本地处理，文件不会上传到网络 ★★★\n\n", "highlight")

        # Notes
        txt.insert(tk.END, "注意事项\n", "heading")
        txt.insert(tk.END, "本工具清除文件属性元数据，不清除文档正文中的批注与修订记录。\n", "body")
        txt.insert(tk.END, "旧版二进制格式（.doc/.xls/.ppt）不支持，请另存为新版格式后再处理。\n\n", "body")

        # macOS ctime note
        txt.insert(tk.END, "macOS 用户注意\n", "heading")
        txt.insert(tk.END,
            "macOS 受内核限制无法修改文件 ctime（状态变更时间），\n"
            "批量处理后所有文件的 ctime 会集中在同一时刻。\n"
            "建议分批处理文件（每批少量，间隔手动停顿），\n"
            "以自然散列 ctime，避免被取证工具关联。\n"
            "Windows 平台无此限制，工具会自动完整散列所有时间戳。\n\n",
            "body",
        )

        # Usage
        txt.insert(tk.END, "使用方式\n", "heading")
        txt.insert(tk.END, "1. 点击「添加文件」选择文件（可多选批量处理）\n", "body")
        txt.insert(tk.END, "2. 可拖拽文件到窗口中\n", "body")
        txt.insert(tk.END, "3. 右键「查看元数据」预览文件中的隐藏信息\n", "body")
        txt.insert(tk.END, "4. （可选）勾选「随机化」以替换为随机仿真值\n", "body")
        txt.insert(tk.END, "5. 点击「清除元数据」执行清理，完成后弹窗展示报告\n", "body")

        txt.configure(state=tk.DISABLED)

        btn_frame = ttk.Frame(dlg, padding="8")
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="关闭", command=dlg.destroy).pack(side=tk.RIGHT)

    # -- Cleaning flow ----------------------------------------------------

    def _on_clean_click(self):
        """Guard for label-based button: ignore clicks while cleaning."""
        if not self._cleaning:
            self._start_cleaning()

    def _start_cleaning(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加要处理的文件。")
            return

        randomize = self.randomize_var.get()

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

        mode_text = "随机化" if randomize else "清除"
        ok = messagebox.askokcancel(
            "确认操作",
            f"即将{mode_text} {len(self.files)} 个文件的元数据。\n"
            "确认继续？",
        )
        if not ok:
            return

        # Pre-generate unique identities for the batch (iron rule)
        identities = generate_unique_identities(len(self.files)) if randomize else []

        self._cleaning = True
        self._toggle_buttons(tk.DISABLED)
        self.progress["value"] = 0
        self.status_var.set("正在处理...")

        t = Thread(target=self._clean_all, args=(randomize, identities), daemon=True)
        t.start()

    def _clean_all(self, randomize: bool, identities: list[dict]):
        total = len(self.files)
        results = []
        metadata_map: dict[str, dict] = {}
        t_start = time.time()

        for i, fp in enumerate(self.files):
            fname = os.path.basename(fp)
            self._ui_update(i, total, f"处理中: {fname}")

            identity = identities[i] if randomize else None
            ok, err = clean_file(fp, randomize, identity)
            results.append((fname, ok, err))

            # Re-read metadata after cleaning for the report
            if ok:
                try:
                    meta = read_metadata(fp)
                    metadata_map[fname] = meta
                except Exception:
                    metadata_map[fname] = {}

            self._ui_update(i + 1, total, f"完成: {fname}")

        elapsed = time.time() - t_start
        self.root.after(0, lambda: self._show_results(results, metadata_map,
                                                       randomize, elapsed))

    def _ui_update(self, cur, total, msg):
        self.root.after(0, lambda: self._do_update(cur, total, msg))

    def _do_update(self, cur, total, msg):
        self.progress["value"] = (cur / total) * 100 if total else 0
        self.status_var.set(msg)

    def _show_results(self, results, metadata_map, randomize, elapsed):
        self._cleaning = False
        self._toggle_buttons(tk.NORMAL)
        self.progress["value"] = 100

        ok_count = sum(1 for _, s, _ in results if s)
        bad_count = len(results) - ok_count
        mode_label = "随机化" if randomize else "清除"

        if bad_count == 0:
            self.status_var.set(f"全部成功 — 已{mode_label} {ok_count} 个文件的元数据")
        else:
            self.status_var.set(f"完成：{ok_count} 成功, {bad_count} 失败")

        # Build and show the metadata report popup
        report = format_metadata_report(results, metadata_map, randomize, elapsed)
        ReportPopup(self.root, report)

        # Clear file list on full success
        if bad_count == 0:
            self._clear_files()

    def _toggle_buttons(self, state):
        for b in (self.btn_add, self.btn_remove, self.btn_clear, self.btn_about, self.btn_clean, self.chk_random):
            b.config(state=state)

    def _on_close(self):
        if self._cleaning:
            messagebox.showwarning("提示", "正在处理中，请等待完成后再关闭窗口。")
        else:
            self._save_config()
            self.root.destroy()

    def _load_config(self):
        """Restore saved window size and column split ratio."""
        try:
            with open(CONFIG_PATH, "r") as f:
                cfg = json.load(f)
            if "geometry" in cfg:
                self.root.geometry(cfg["geometry"])
            if "sash_pos" in cfg:
                self.root.after(100, lambda p=cfg["sash_pos"]: self.paned.sashpos(0, p) if self.paned.winfo_exists() else None)
        except Exception:
            pass

    def _save_config(self):
        """Save window size and column split ratio."""
        try:
            cfg = {
                "geometry": self.root.geometry(),
                "sash_pos": self.paned.sashpos(0),
            }
            with open(CONFIG_PATH, "w") as f:
                json.dump(cfg, f)
        except Exception:
            pass


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
    global HAS_DND
    if len(sys.argv) > 1:
        cli_mode(sys.argv[1:])
    else:
        root = None
        if HAS_DND:
            if sys.platform == "darwin":
                # On macOS, TkinterDnD.Tk() creates an extra blank window.
                # Use standard tk.Tk() and manually load tkdnd's Tcl library.
                root = tk.Tk()
                try:
                    import tkinterdnd2
                    import platform
                    arch = 'osx-arm64' if platform.machine() == 'arm64' else 'osx-x64'
                    tkdnd_path = os.path.join(os.path.dirname(tkinterdnd2.__file__), 'tkdnd', arch)
                    root.tk.call('lappend', 'auto_path', tkdnd_path)
                    root.tk.call('package', 'require', 'tkdnd')
                except tk.TclError:
                    HAS_DND = False
            else:
                try:
                    root = TkinterDnD.Tk()
                except RuntimeError:
                    HAS_DND = False
                    root = tk.Tk()

        if root is None:
            root = tk.Tk()

        MetadataCleanerApp(root)
        root.mainloop()


if __name__ == "__main__":
    main()
