"""
Random value generation for metadata privacy enhancement.

Generates realistic-looking random identities (names, companies, dates)
for replacing sensitive metadata. Ensures batch uniqueness — no two files
in the same batch ever share the same person name or company name.
"""

import random
import secrets
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import ttk

# ============================================================
# Name pools
#   Combined:  80 first × 80 last = 6,400 name pairs
#   Standalone: 52 names × ~4 case variants ≈ 200
#   Total: ~6,600 distinct names
# Last names are pinyin of top Chinese surnames (百家姓).
# ============================================================

_FIRST_NAMES = [
    # Chinese pinyin given names (40)
    "Wei", "Jing", "Lei", "Tao", "Yan", "Min", "Jun", "Hong",
    "Fang", "Ping", "Hua", "Qiang", "Na", "Ling", "Jie",
    "Ming", "Li", "Hao", "Yu", "Bo", "Chao", "Yong", "Xin",
    "Wen", "Hui", "Rui", "Lan", "Bin", "Ying", "Mei", "Xue",
    "Feng", "Yun", "Kai", "Chen", "Yi", "Ning", "Ran", "Sheng",
    # English given names common in China (42)
    "David", "James", "Michael", "Jason", "Kevin", "Eric", "Jack", "Tom",
    "Andy", "Henry", "Tony", "Leo", "Mike", "Peter", "Alex", "Steven",
    "Daniel", "Justin", "Ryan", "Allen", "Frank", "George", "Sam", "Robert",
    "Linda", "Lisa", "Amy", "Emily", "Alice", "Anna", "Tina", "Wendy",
    "Jenny", "Helen", "Lucy", "Sarah", "Julia", "Grace", "Emma", "Cathy",
    "Jerry", "Bob",
]

# Standalone names — complete identities, never combined with a last name.
# Generic/system names and short letter combos that real users type as-is.
_STANDALONE_NAMES = [
    # Generic / system names (20) — common defaults on Chinese office PCs
    "User", "Admin", "Administrator", "Operator", "Guest", "Owner",
    "admin", "user", "operator", "test", "default", "office",
    "IT", "Web", "System", "info", "PC", "public",
    # Numeric keyboard smashing — extremely common for non-technical users
    "123", "123456", "111", "000", "888", "666",
    # Brand PC default account names (Lenovo/Dell/HP ship with these)
    "Lenovo", "lenovo", "ThinkPad", "ASUS", "Huawei", "Acer", "HP", "Dell",
    # Random short letter combos — lazy/informal typing
    "ab", "cc", "dd", "ff", "gg", "hh", "jj", "kk", "ll", "mm",
    "nn", "pp", "qq", "rr", "ss", "tt", "ww", "xx", "yy", "zz",
    "ly", "aj", "dj", "jc", "jd", "jp", "jr", "kc", "lj", "mj",
    "pj", "rj", "tj", "cj", "bj", "sj",
    "abc", "xyz", "qwe", "asd", "zxc", "rty", "fgh", "vbn", "uio", "jkl",
    "mno", "plm", "okn", "wsx", "edc", "rfv", "tgb", "yhn", "ujm", "ikm",
]

# Top 80 Chinese surnames (百家姓) in pinyin
_LAST_NAMES = [
    "Wang", "Li", "Zhang", "Liu", "Chen", "Yang", "Huang", "Zhao",
    "Wu", "Zhou", "Xu", "Sun", "Ma", "Zhu", "Hu", "Guo",
    "He", "Gao", "Lin", "Luo", "Zheng", "Liang", "Xie", "Song",
    "Tang", "Han", "Cao", "Deng", "Xiao", "Feng", "Zeng", "Cheng",
    "Cai", "Peng", "Pan", "Yuan", "Yu", "Dong", "Su", "Ye",
    "Lu", "Wei", "Jiang", "Tian", "Du", "Ding", "Shen", "Ren",
    "Yao", "Fan", "Fang", "Shi", "Fu", "Xia", "Tan", "Liao",
    "Zou", "Xiong", "Jin", "Kong", "Bai", "Cui", "Kang", "Mao",
    "Qiu", "Qin", "Gu", "Hou", "Shao", "Meng", "Long", "Wan",
    "Duan", "Lei", "Qian", "Yin", "Yi", "Chang", "Lai", "Gong",
]

# ============================================================
# Company name pools — tech-focused Chinese-style names
# Pattern: [Brand] [Industry] (suffix is rare, applied by style)
# 40 brand × 8 industry = 320 base combinations
# ============================================================

_COMPANY_BRANDS = [
    "HuaSheng", "LiXiang", "RuiHeng", "JinHui", "TianYu", "BoFeng",
    "XinRui", "JiaHe", "WanTai", "KaiYuan", "RongChang", "ZhiCheng",
    "YiAn", "FengHua", "LongTeng", "HaiTian", "ZhongHui", "JiaYuan",
    "XinCheng", "TongDa", "GuangLian", "HongYuan", "MingYang", "HuiZhong",
    "RuiFeng", "ChengXin", "AnDa", "BaiSheng", "ZhongCheng", "XinDa",
    "TianCheng", "RongXin", "KangHui", "FengShou", "HengXin", "JiaCheng",
    "ShengRui", "HuiTong", "ZhanXin", "YuanFeng", "RuiXin",
]

_COMPANY_INDUSTRIES = [
    "Technology", "Information Technology", "Network Technology",
    "Software", "Electronics", "Internet Technology",
    "Computer Technology", "Data Technology",
]

# ============================================================
# ============================================================
# Random generation functions
# ============================================================


def _avoid_night(dt: datetime) -> datetime:
    """Shift timestamps away from 00:00–07:00 (Beijing time) with 85% probability.

    Chinese office workers rarely create/edit documents in the middle of the night.
    A small proportion of night hours is kept for realism (overtime, night shifts).
    """
    if 0 <= dt.hour <= 6 and random.random() < 0.85:
        return dt.replace(hour=random.randint(8, 22))
    return dt


def randomize_times(actual_mtime: float | None = None) -> tuple[str, str, str]:
    """Return (created, last_printed, modified) ISO-8601 timestamps.

    When actual_mtime is provided (file's real modification timestamp),
    generated times are anchored to it:
      - modified:  within 2 days BEFORE actual_mtime (never later)
      - created:   between modified-7d and modified-5h
      - printed:   between created and modified

    Without actual_mtime, falls back to the last 2 years.

    All timestamps are biased away from 00:00–07:00 (Beijing time) since
    real office documents are rarely produced in those hours.
    """
    if actual_mtime is not None:
        actual_dt = datetime.fromtimestamp(actual_mtime)
        # modified: random in [actual - 2 days, actual] (never later than real)
        modified = actual_dt - timedelta(
            seconds=random.randint(0, 172800)  # 0 to 48 hours before
        )
        # created: random in [modified - 7 days, modified - 5 hours]
        created = modified - timedelta(
            seconds=random.randint(18000, 604800)  # 5h to 7d before modified
        )
        # printed: random between created and modified
        printed = created + timedelta(
            seconds=random.randint(0, int((modified - created).total_seconds()))
        )

        # Avoid midnight–7am (Beijing time), then fix ordering if needed
        created = _avoid_night(created)
        printed = _avoid_night(printed)
        modified = _avoid_night(modified)
        # Ensure chronological order is preserved after hour shifting
        if printed < created:
            printed = created + timedelta(minutes=random.randint(10, 120))
        if modified < printed:
            modified = printed + timedelta(minutes=random.randint(5, 60))

        return (
            created.strftime("%Y-%m-%dT%H:%M:%SZ"),
            printed.strftime("%Y-%m-%dT%H:%M:%SZ"),
            modified.strftime("%Y-%m-%dT%H:%M:%SZ"),
        )

    # Fallback: no real file time available
    now = datetime.now()
    earliest = now - timedelta(days=730)

    created_ts = random.randint(int(earliest.timestamp()), int(now.timestamp()) - 86400)
    created = _avoid_night(datetime.fromtimestamp(created_ts))

    last_printed_ts = random.randint(int(created.timestamp()) + 3600, int(now.timestamp()) - 7200)
    last_printed = _avoid_night(datetime.fromtimestamp(last_printed_ts))

    modified_ts = random.randint(int(last_printed.timestamp()) + 60, int(now.timestamp()) - 60)
    modified = _avoid_night(datetime.fromtimestamp(modified_ts))

    # Ensure chronological order after hour shifting
    if last_printed < created:
        last_printed = created + timedelta(minutes=random.randint(10, 120))
    if modified < last_printed:
        modified = last_printed + timedelta(minutes=random.randint(5, 60))

    return (
        created.strftime("%Y-%m-%dT%H:%M:%SZ"),
        last_printed.strftime("%Y-%m-%dT%H:%M:%SZ"),
        modified.strftime("%Y-%m-%dT%H:%M:%SZ"),
    )


def random_template(author_name: str | None = None, ext: str = ".docx") -> str:
    """Generate a realistic random template path.

    When author_name is provided, ~40% of user-dir paths use a username
    derived from the author (e.g. "Wei Zhang" → "wei"), so the template
    path doesn't contradict who supposedly created the file.

    Template filename is chosen based on file extension:
    .docx/.wps → Normal.dotm/dotx, .xlsx/.et → Book.xltx/Sheet.xltx,
    .pptx/.dps → Blank.potx.
    """
    _WIN_USERS = ["user", "Default", "Administrator", "Owner", "Admin",
                  "Public", "Lenovo", "ThinkPad", "HP", "ASUS", "Dell",
                  "zhangsan", "lisi", "wangwu", "xiaoming"]
    _MAC_USERS = ["user", "Shared", "admin", "mac", "apple"]

    # Standard Office templates by file type
    _TMPL_WORD = ["Normal.dotm", "Normal.dotx"]
    _TMPL_EXCEL = ["Book.xltx", "Sheet.xltx"]
    _TMPL_PPT = ["Blank.potx"]

    if ext in (".xlsx", ".et"):
        _TMPL = _TMPL_EXCEL
    elif ext in (".pptx", ".dps"):
        _TMPL = _TMPL_PPT
    else:
        _TMPL = _TMPL_WORD

    # Derive a candidate username from author (lowercase first token)
    _author_user = None
    if author_name:
        first_token = author_name.strip().split()[0].lower()
        if 2 <= len(first_token) <= 15 and first_token.isalpha():
            _author_user = first_token

    def _pick_user(users: list[str]) -> str:
        if _author_user and random.random() < 0.4:
            return _author_user
        return random.choice(users)

    if random.random() < 0.88:  # 88% Windows
        u = _pick_user(_WIN_USERS)
        t = random.choice(_TMPL)
        variant = random.random()
        if variant < 0.35:
            # Per-user template folder (most common)
            return f"C:\\Users\\{u}\\AppData\\Roaming\\Microsoft\\Templates\\{t}"
        elif variant < 0.60:
            # Custom Office Templates (Office 2013+ default)
            return f"C:\\Users\\{u}\\Documents\\Custom Office Templates\\{t}"
        elif variant < 0.80:
            # System-wide Click-to-Run templates (Office 365/2019/2021)
            ver = random.choice(["16.0", "15.0"])
            return (f"C:\\Program Files\\Microsoft Office\\root\\Office{ver}"
                    f"\\Templates\\1033\\{t}")
        elif variant < 0.93:
            # System-wide MSI templates (Office 2016/2013/2010)
            arch = random.choice(["Program Files", "Program Files (x86)"])
            ver = random.choice(["Office16", "Office15", "Office14"])
            return f"C:\\{arch}\\Microsoft Office\\{ver}\\Templates\\1033\\{t}"
        else:
            # Public / shared machine template folder
            return f"C:\\Users\\Public\\Documents\\Custom Office Templates\\{t}"

    else:  # 12% macOS
        u = _pick_user(_MAC_USERS)
        variant = random.random()
        if variant < 0.5:
            return (f"/Users/{u}/Library/Group Containers"
                    f"/UBF8T346G9.Office/User Content"
                    f"/Templates/Normal.dotm")
        elif variant < 0.8:
            return (f"/Users/{u}/Library/Application Support"
                    f"/Microsoft/Office/User Templates"
                    f"/Normal.dotm")
        else:
            return (f"/Users/{u}/Documents/Microsoft User Data"
                    f"/Office Templates/Normal.dotm")


def random_total_time() -> str:
    """Random editing duration in minutes, weighted towards realistic values.

    Minimum 120 minutes — shorter than that is unrealistic for real documents.
    60% medium (2-5h), 30% long (5-10h), 10% very long (10-40h)."""
    roll = random.random()
    if roll < 0.6:
        minutes = random.randint(120, 300)     # 2 – 5 hours
    elif roll < 0.9:
        minutes = random.randint(300, 600)     # 5 – 10 hours
    else:
        minutes = random.randint(600, 2400)    # 10 – 40 hours
    return str(minutes)


def random_rsid_mapping(unique_rsids: set[bytes]) -> dict[bytes, bytes]:
    """Build a replacement mapping for RSID values.
    Each unique RSID maps to a new random hex string of the same length.
    Same document → same mapping (consistent). Different documents → different.
    """
    return {
        rsid: secrets.token_hex(len(rsid) // 2).upper().encode()
        for rsid in unique_rsids
    }


# ============================================================
# Batch identity generation (iron rule: no duplicates within batch)
# ============================================================


def _random_name_style(first: str, last: str) -> str:
    """Combine first + last name with a randomly chosen casing/format style.

    Multiple styles prevent batch-processed files from sharing a uniform
    formatting pattern that would make them recognisable as a group.
    """
    styles = [
        # Most common: all-lowercase pinyin "zhangwei" — domain account / lazy typing — 35%
        lambda f, l: f"{l.lower()}{f.lower()}",
        lambda f, l: f"{l.lower()}{f.lower()}",
        lambda f, l: f"{l.lower()}{f.lower()}",
        lambda f, l: f"{l.lower()}{f.lower()}",
        # Western order: "Wei Zhang" — common in international firms — 25%
        lambda f, l: f"{f} {l}",
        lambda f, l: f"{f} {l}",
        lambda f, l: f"{f} {l}",
        # Only given name: "Wei" or "David" — some systems only capture first name — 18%
        lambda f, l: f"{f}",
        lambda f, l: f"{f}",
        # Chinese order with space: "Zhang Wei" — rare but exists — 10%
        lambda f, l: f"{l} {f}",
        # Chinese order no-space capitalized: "Zhangwei" — 7%
        lambda f, l: f"{l}{f}",
        # Lowercase with space: "zhang wei" — 5%
        lambda f, l: f"{l.lower()} {f.lower()}",
    ]
    return random.choice(styles)(first, last)


def _random_standalone_name(name: str) -> str:
    """Apply a random casing style to a standalone name (no last name).

    Standalone names like "admin", "IT", "ab" are complete identities
    and should never be combined with a surname.
    """
    styles = [
        lambda n: n,                        # original: "admin" — 30%
        lambda n: n,                        # (dup weight)
        lambda n: n,                        # (dup weight)
        lambda n: n.lower(),                # lower: "admin" — 20%
        lambda n: n.lower(),                # (dup weight)
        lambda n: n.upper(),                # UPPER: "ADMIN" — 20%
        lambda n: n.upper(),                # (dup weight)
        lambda n: n.title(),                # Title: "Admin" — 20%
        lambda n: n.title(),                # (dup weight)
        lambda n: n.lower().replace(" ", ""),  # compact — 10%
    ]
    return random.choice(styles)(name)


def _random_company_style(name: str) -> str:
    """Apply a random casing/format style to a company name.

    The base name is 'Brand Industry' (no city, no suffix).
    This function varies presentation — most common is a short abbreviation,
    mimicking how real users rarely type full formal company names.
    """
    def _short(_s: str) -> str:
        """Abbreviate industry: Technology→Tech, Information Technology→InfoTech, etc."""
        return (_s.replace("Technology", "Tech")
                 .replace("Information Technology", "InfoTech")
                 .replace("Network Technology", "NetTech")
                 .replace("Internet Technology", "NetTech")
                 .replace("Computer Technology", "CompTech")
                 .replace("Electronics", "Elec")
                 .replace("Software", "Soft")
                 .replace("Data Technology", "DataTech"))

    def _random_abbrev(_s: str) -> str:
        """Replace entire company name with 2-4 random uppercase letters."""
        return "".join(random.choices("ABCDEFGHIJKLMNOPQRSTUVWXYZ",
                                      k=random.randint(2, 4)))

    # Weighted styles: most people use short names or random letters
    style_weights = [
        # Short random letters (most common real-world pattern)
        (_random_abbrev,                                 35),
        (lambda _s: _random_abbrev(_s).lower(),          18),
        # Short form: "HuaSheng Tech"
        (_short,                                         15),
        (lambda s: _short(s).lower(),                     8),
        # Full form: "HuaSheng Technology"
        (lambda s: s,                                     8),
        (lambda s: s.lower(),                             4),
        # Rarely: add Co., Ltd. suffix
        (lambda s: f"{_short(s)} Co., Ltd.",              3),
        (lambda s: f"{_short(s)} Co. Ltd.",               2),
        (lambda s: f"{s} Co., Ltd.",                      2),
        # Oddballs
        (lambda s: s.upper(),                             2),
        (_short,                                          1),  # extra weight
        (lambda s: f"{_short(s).lower()}",                2),
    ]
    styles, weights = zip(*style_weights)
    return random.choices(styles, weights=weights, k=1)[0](name)


def generate_unique_identities(count: int) -> list[dict[str, str]]:
    """Pre-generate `count` unique (author_name, company_name) pairs.

    Name pool mixes two types (interleaved randomly, not appended):
      - Combined: 79 first × 80 last = 6,320 pairs → ~88% of output
      - Standalone: 81 generic/short-letter names → ~12% of output

    Standalone names like "admin" or "ab" are complete identities — they are
    never combined with a surname.  Both types are shuffled together so batch
    output has natural variety regardless of batch size.

    A used-names set guards against style-induced collisions.  Overflow adds
    a numeric suffix to guarantee uniqueness beyond pool size.
    """
    # Combined-name pool: (first, last) tuples
    name_pairs = [(f, l) for f in _FIRST_NAMES for l in _LAST_NAMES]
    random.shuffle(name_pairs)

    # Standalone-name pool: strings (no last name)
    standalone = list(_STANDALONE_NAMES)
    random.shuffle(standalone)

    total_name_pool = len(name_pairs) + len(standalone)
    paired_idx = 0
    standalone_idx = 0

    # Build all company combos: 40 brands × 8 industries = 320 base names
    # Short form: [Brand] [Industry]  (no city, no Co.,Ltd suffix)
    all_companies = []
    for brand in _COMPANY_BRANDS:
        for industry in _COMPANY_INDUSTRIES:
            all_companies.append(f"{brand} {industry}")
    random.shuffle(all_companies)
    company_pool = len(all_companies)

    identities = []
    used_names: set[str] = set()
    used_companies: set[str] = set()

    for i in range(count):
        # --- Name ---
        suffix = f"_{i // total_name_pool}" if i >= total_name_pool else ""

        # ~12% standalone, interleaved randomly so small batches also get variety
        paired_exhausted = paired_idx >= len(name_pairs)
        standalone_exhausted = standalone_idx >= len(standalone)
        use_standalone = (not standalone_exhausted
                          and (paired_exhausted or random.random() < 0.12))

        if use_standalone:
            base = standalone[standalone_idx % len(standalone)]
            standalone_idx += 1
            for _ in range(30):
                candidate = _random_standalone_name(base) + suffix
                if candidate not in used_names:
                    break
            else:
                candidate = f"{base}{suffix}"
                extra = 1
                while candidate in used_names:
                    candidate = f"{base}{suffix}_{extra}"
                    extra += 1
        else:
            first, last = name_pairs[paired_idx % len(name_pairs)]
            paired_idx += 1
            for _ in range(30):
                candidate = _random_name_style(first, last) + suffix
                if candidate not in used_names:
                    break
            else:
                candidate = f"{first} {last}{suffix}"
                extra = 1
                while candidate in used_names:
                    candidate = f"{first} {last}{suffix}_{extra}"
                    extra += 1
        used_names.add(candidate)

        # --- Company ---
        company_base = all_companies[i % company_pool]
        csuffix = f"_{i // company_pool}" if i >= company_pool else ""

        for _ in range(30):
            candidate_company = _random_company_style(company_base) + csuffix
            if candidate_company not in used_companies:
                break
        else:
            candidate_company = f"{company_base}{csuffix}"
            extra = 1
            while candidate_company in used_companies:
                candidate_company = f"{company_base}{csuffix}_{extra}"
                extra += 1
        used_companies.add(candidate_company)

        identities.append({
            "author_name": candidate,
            "company_name": candidate_company,
        })
    return identities


# ============================================================
# Report generation
# ============================================================


def format_metadata_report(
    results: list[tuple[str, bool, str]],
    metadata_map: dict[str, dict],
    randomize_mode: bool,
    elapsed_sec: float,
) -> str:
    """Build a human-readable report of post-cleaning metadata.

    Args:
        results: list of (filename, success_bool, error_message)
        metadata_map: {filename: {category: {label: value}}} from read_metadata()
        randomize_mode: whether randomization was enabled
        elapsed_sec: total processing time
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mode_label = "随机化" if randomize_mode else "清空"
    total = len(results)
    ok_count = sum(1 for _, ok, _ in results if ok)

    lines = []
    lines.append("═" * 56)
    lines.append("元数据清洗报告")
    lines.append(f"处理时间: {now}  |  模式: {mode_label}  |  文件数: {total}")
    lines.append("═" * 56)

    for i, (fname, ok, err) in enumerate(results, 1):
        lines.append("")
        lines.append(f"┌─ {i}. {fname} {'─' * max(1, 48 - len(fname) - len(str(i)))}")

        if not ok:
            lines.append(f"│ 状态: ❌ 失败")
            lines.append(f"│ 原因: {err}")
            lines.append("└" + "─" * 48)
            continue

        lines.append("│ 状态: ✅ 成功")

        meta = metadata_map.get(fname, {})
        if not meta:
            lines.append("│ (无可显示的元数据 — 已全部清除)")
            lines.append("└" + "─" * 48)
            continue

        # Print each category
        category_order = ["文档属性 (core.xml)", "扩展属性 (app.xml)",
                          "自定义属性 (custom.xml)", "嵌入图片元数据",
                          "图片来源路径 (document.xml)", "错误"]
        for cat in category_order:
            if cat in meta:
                lines.append(f"│")
                lines.append(f"│ ▸ {cat}:")
                for label, value in meta[cat].items():
                    # Truncate long values (e.g. image paths)
                    display_val = str(value)
                    lines.append(f"│   {label}: {display_val}")

        lines.append("└" + "─" * 48)

    lines.append("")
    lines.append("═" * 56)
    if ok_count == total:
        lines.append(f"✅ 全部 {total} 个文件处理成功")
    else:
        bad_count = total - ok_count
        lines.append(f"⚠️ 成功 {ok_count} 个, 失败 {bad_count} 个")
    lines.append(f"耗时: {elapsed_sec:.1f} 秒")
    lines.append("═" * 56)

    return "\n".join(lines)


# ============================================================
# Report popup window
# ============================================================


class ReportPopup:
    """Scrollable read-only popup displaying post-cleaning metadata report."""

    def __init__(self, parent, report_text: str):
        self._win = tk.Toplevel(parent)
        self._win.title("元数据清洗报告")
        self._win.geometry("720x540")
        self._win.minsize(500, 360)

        # Make modal
        self._win.transient(parent)
        self._win.grab_set()

        # Text area with scrollbar
        frame = ttk.Frame(self._win, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        sb = ttk.Scrollbar(frame)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self._text = tk.Text(
            frame,
            wrap=tk.NONE,
            font=("Courier New", 12),
            yscrollcommand=sb.set,
            state=tk.DISABLED,
        )
        self._text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=self._text.yview)

        # Populate text
        self._text.config(state=tk.NORMAL)
        self._text.insert("1.0", report_text)
        self._text.config(state=tk.DISABLED)

        # Close button
        btn_frame = ttk.Frame(self._win, padding="8")
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="关闭", command=self._win.destroy).pack(side=tk.RIGHT)

        # Ctrl+W / Cmd+W to close
        self._win.bind("<Control-w>", lambda _e: self._win.destroy())
        self._win.bind("<Command-w>", lambda _e: self._win.destroy())

        self._win.focus_set()
