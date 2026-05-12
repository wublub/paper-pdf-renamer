"""PDF 文献一键命名工具

特点:
  - 自动从 PDF 中提取 DOI, 通过 CrossRef API 获取准确元数据
  - 支持字段: Title / Journal / Journal-Abbr / Year / Accepted-Date /
              Published-Date / First-Author / DOI
  - 命名顺序、分隔符可在 GUI 中自由调整
  - 配置自动保存到 ~/.pdf_renamer_config.json

依赖: pypdf (若未安装会自动提示安装)
"""

import json
import re
import subprocess
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader  # type: ignore
    except ImportError:
        PdfReader = None  # type: ignore


CONFIG_PATH = Path.home() / ".pdf_renamer_config.json"

FIELDS = {
    "title":          "Title  标题",
    "journal":        "Journal  期刊全名",
    "journal_short":  "Journal-Abbr  期刊缩写",
    "year":           "Year  年份",
    "accepted_date":  "Accepted-Date  接受日期",
    "published_date": "Published-Date  出版日期",
    "first_author":   "First-Author  第一作者",
    "doi":            "DOI",
}


# ---------- 元数据 ----------

@dataclass
class PaperMeta:
    title: str = ""
    journal: str = ""
    journal_short: str = ""
    year: str = ""
    accepted_date: str = ""
    published_date: str = ""
    first_author: str = ""
    doi: str = ""
    source_file: str = ""

    def get(self, name: str) -> str:
        return getattr(self, name, "") or "Unknown"


DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.IGNORECASE)
ACCEPTED_RE = re.compile(
    r"Accepted[^\d\n]{0,30}?"
    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}"
    r"|[A-Za-z]+\s+\d{1,2},?\s+\d{4}"
    r"|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})",
    re.IGNORECASE,
)
MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
    "january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7,
    "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
}


def extract_text(pdf: Path, pages: int = 3) -> str:
    if PdfReader is None:
        return ""
    try:
        r = PdfReader(str(pdf))
        n = min(len(r.pages), pages)
        return "\n".join((r.pages[i].extract_text() or "") for i in range(n))
    except Exception:
        return ""


def find_doi(text: str) -> str:
    if not text:
        return ""
    m = DOI_RE.search(text)
    if not m:
        return ""
    return m.group(0).rstrip(".,;)]}>")


def query_crossref(doi: str, timeout: int = 10) -> Optional[dict]:
    url = f"https://api.crossref.org/works/{urllib.parse.quote(doi)}"
    req = urllib.request.Request(
        url, headers={"User-Agent": "PDF-Renamer/1.0 (academic use)"}
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.load(resp).get("message")
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError):
        return None


def date_parts_to_str(parts) -> str:
    if not parts:
        return ""
    p = parts[0] if isinstance(parts, list) and parts and isinstance(parts[0], list) else parts
    if not p:
        return ""
    try:
        if len(p) >= 3:
            return f"{int(p[0]):04d}-{int(p[1]):02d}-{int(p[2]):02d}"
        if len(p) == 2:
            return f"{int(p[0]):04d}-{int(p[1]):02d}"
        return f"{int(p[0]):04d}"
    except (TypeError, ValueError):
        return ""


def normalize_date(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    if re.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}", s):
        m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})[-\s]+([A-Za-z]+)[-\s,]+(\d{4})", s)
    if m and MONTHS.get(m.group(2).lower()):
        return (
            f"{int(m.group(3)):04d}-{MONTHS[m.group(2).lower()]:02d}-{int(m.group(1)):02d}"
        )
    m = re.match(r"^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})", s)
    if m and MONTHS.get(m.group(1).lower()):
        return (
            f"{int(m.group(3)):04d}-{MONTHS[m.group(1).lower()]:02d}-{int(m.group(2)):02d}"
        )
    return s


def extract_metadata(pdf: Path) -> PaperMeta:
    meta = PaperMeta(source_file=str(pdf))
    text = extract_text(pdf, pages=3)

    # 1) DOI -> CrossRef (最准确)
    doi = find_doi(text)
    if doi:
        meta.doi = doi
        cr = query_crossref(doi)
        if cr:
            if cr.get("title"):
                meta.title = cr["title"][0]
            ct = [t for t in (cr.get("container-title") or []) if t]
            sct = [t for t in (cr.get("short-container-title") or []) if t]
            if ct:
                # CrossRef 有时给多条, 最长的通常是完整全称
                meta.journal = max(ct, key=len)
            if sct:
                meta.journal_short = sct[0]
            elif len(ct) > 1:
                meta.journal_short = min(ct, key=len)

            for key in ("published-print", "published-online", "issued"):
                d = cr.get(key, {}).get("date-parts")
                if d:
                    meta.published_date = date_parts_to_str(d)
                    meta.year = meta.published_date[:4]
                    break

            acc = cr.get("accepted", {}).get("date-parts")
            if acc:
                meta.accepted_date = date_parts_to_str(acc)
            else:
                for a in cr.get("assertion") or []:
                    label = (str(a.get("label", "")) + str(a.get("name", ""))).lower()
                    if "accept" in label:
                        v = a.get("value", "")
                        if v:
                            meta.accepted_date = normalize_date(v)
                            break

            authors = cr.get("author") or []
            if authors:
                fam = authors[0].get("family") or ""
                giv = authors[0].get("given") or ""
                meta.first_author = fam or giv

    # 2) 文本兜底找 Accepted
    if not meta.accepted_date:
        m = ACCEPTED_RE.search(text)
        if m:
            meta.accepted_date = normalize_date(m.group(1))

    # 3) PDF 元数据兜底找标题
    if not meta.title and PdfReader is not None:
        try:
            info = PdfReader(str(pdf)).metadata or {}
            t = info.get("/Title", "") if hasattr(info, "get") else ""
            if t and len(str(t)) > 5:
                meta.title = str(t)
        except Exception:
            pass

    return meta


# ---------- 文件名构造 ----------

def sanitize(s: str, max_len: int = 100) -> str:
    if not s:
        return ""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", str(s))
    s = re.sub(r"\s+", " ", s).strip()
    return s[:max_len].rstrip()


def build_filename(meta: PaperMeta, tokens: list, ext: str = ".pdf") -> str:
    parts = []
    for t in tokens:
        if t["type"] == "field":
            parts.append(sanitize(meta.get(t["value"])))
        else:
            parts.append(t["value"])
    name = "".join(parts).strip(" _-.") or "Unnamed"
    return name + ext


# ---------- GUI ----------

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("PDF 文献一键命名工具")
        root.geometry("1180x780")

        self.folder = tk.StringVar()
        self.tokens = self._load_config()
        self.metas: dict[str, PaperMeta] = {}

        self._build_ui()
        self._refresh_pattern_display()

    # config
    def _load_config(self) -> list:
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return data
            except Exception:
                pass
        return [
            {"type": "field", "value": "year"},
            {"type": "sep",   "value": "_"},
            {"type": "field", "value": "journal"},
            {"type": "sep",   "value": "_"},
            {"type": "field", "value": "title"},
        ]

    def _save_config(self):
        try:
            CONFIG_PATH.write_text(
                json.dumps(self.tokens, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception:
            pass

    # build UI
    def _build_ui(self):
        # top
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="选择文件夹", command=self._choose_folder).pack(side="left")
        ttk.Entry(top, textvariable=self.folder).pack(
            side="left", fill="x", expand=True, padx=8
        )
        ttk.Button(top, text="扫描并提取", command=self._scan).pack(side="left")

        # pattern builder
        pat = ttk.LabelFrame(
            self.root,
            text="命名规则（左侧选中字段→点[添加字段]；分隔符自填→点[添加分隔符]；右侧选中可上移/下移/删除）",
            padding=8,
        )
        pat.pack(fill="x", padx=8, pady=4)

        left = ttk.Frame(pat)
        left.pack(side="left", fill="y")
        ttk.Label(left, text="可用字段").pack(anchor="w")
        self.field_box = tk.Listbox(left, height=9, width=28, exportselection=False)
        for label in FIELDS.values():
            self.field_box.insert("end", label)
        self.field_box.pack()
        self.field_box.bind("<Double-Button-1>", lambda _e: self._add_field())

        mid = ttk.Frame(pat)
        mid.pack(side="left", padx=10)
        ttk.Button(mid, text="→ 添加字段", command=self._add_field).pack(fill="x", pady=2)
        ttk.Label(mid, text="分隔符:").pack(anchor="w", pady=(8, 0))
        self.sep_var = tk.StringVar(value="_")
        ttk.Entry(mid, textvariable=self.sep_var, width=10).pack(fill="x")
        ttk.Button(mid, text="→ 添加分隔符", command=self._add_sep).pack(fill="x", pady=2)
        ttk.Separator(mid, orient="horizontal").pack(fill="x", pady=8)
        ttk.Button(mid, text="↑ 上移", command=lambda: self._move(-1)).pack(fill="x", pady=2)
        ttk.Button(mid, text="↓ 下移", command=lambda: self._move(1)).pack(fill="x", pady=2)
        ttk.Button(mid, text="× 删除", command=self._remove).pack(fill="x", pady=2)
        ttk.Button(mid, text="清空全部", command=self._clear).pack(fill="x", pady=2)

        right = ttk.Frame(pat)
        right.pack(side="left", fill="both", expand=True)
        ttk.Label(right, text="当前命名顺序").pack(anchor="w")
        self.order_box = tk.Listbox(right, height=9, exportselection=False)
        self.order_box.pack(fill="both", expand=True)

        # pattern preview line
        pp = ttk.Frame(self.root, padding=(8, 4))
        pp.pack(fill="x")
        ttk.Label(pp, text="规则预览：").pack(side="left")
        self.pattern_lbl = ttk.Label(pp, text="", foreground="#0066cc",
                                     font=("Consolas", 10, "bold"))
        self.pattern_lbl.pack(side="left")
        ttk.Button(pp, text="刷新文件名预览", command=self._refresh_preview).pack(side="right")

        # table
        tbl = ttk.LabelFrame(self.root, text="文件预览", padding=4)
        tbl.pack(fill="both", expand=True, padx=8, pady=4)
        cols = ("orig", "title", "journal", "year", "accepted", "newname")
        widths = (180, 240, 150, 60, 95, 320)
        headers = {
            "orig": "原文件名", "title": "Title", "journal": "Journal",
            "year": "Year", "accepted": "Accepted", "newname": "新文件名 (预览)",
        }
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings")
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=w, stretch=True)
        vsb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # bottom
        bot = ttk.Frame(self.root, padding=8)
        bot.pack(fill="x")
        self.status = ttk.Label(bot, text="就绪")
        self.status.pack(side="left")
        ttk.Button(bot, text="执行重命名", command=self._do_rename).pack(side="right")

    # pattern ops
    def _refresh_pattern_display(self):
        self.order_box.delete(0, "end")
        for t in self.tokens:
            if t["type"] == "field":
                self.order_box.insert("end", f"[字段] {FIELDS.get(t['value'], t['value'])}")
            else:
                self.order_box.insert("end", f"[分隔] \"{t['value']}\"")
        text = "".join(
            "{" + t["value"] + "}" if t["type"] == "field" else t["value"]
            for t in self.tokens
        ) or "(空)"
        self.pattern_lbl.config(text=text)

    def _add_field(self):
        sel = self.field_box.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先在左侧选择一个字段")
            return
        key = list(FIELDS.keys())[sel[0]]
        self.tokens.append({"type": "field", "value": key})
        self._refresh_pattern_display()
        self._save_config()

    def _add_sep(self):
        s = self.sep_var.get()
        if not s:
            return
        self.tokens.append({"type": "sep", "value": s})
        self._refresh_pattern_display()
        self._save_config()

    def _move(self, delta: int):
        sel = self.order_box.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if 0 <= j < len(self.tokens):
            self.tokens[i], self.tokens[j] = self.tokens[j], self.tokens[i]
            self._refresh_pattern_display()
            self.order_box.selection_set(j)
            self._save_config()

    def _remove(self):
        sel = self.order_box.curselection()
        if not sel:
            return
        del self.tokens[sel[0]]
        self._refresh_pattern_display()
        self._save_config()

    def _clear(self):
        if messagebox.askyesno("确认", "清空全部命名规则?"):
            self.tokens = []
            self._refresh_pattern_display()
            self._save_config()

    # actions
    def _choose_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.folder.set(d)

    def _scan(self):
        d = self.folder.get()
        if not d or not Path(d).is_dir():
            messagebox.showerror("错误", "请先选择有效文件夹")
            return
        pdfs = sorted(Path(d).glob("*.pdf"))
        if not pdfs:
            messagebox.showinfo("提示", "该文件夹下未找到 PDF")
            return
        self.tree.delete(*self.tree.get_children())
        self.metas.clear()

        def worker():
            for i, p in enumerate(pdfs, 1):
                self.status.config(text=f"提取中 {i}/{len(pdfs)}: {p.name}")
                meta = extract_metadata(p)
                self.metas[str(p)] = meta
                new_name = build_filename(meta, self.tokens)
                self.tree.insert(
                    "", "end", iid=str(p),
                    values=(p.name, meta.title[:80], meta.journal, meta.year,
                            meta.accepted_date, new_name),
                )
            self.status.config(text=f"完成,共 {len(pdfs)} 个文件")

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_preview(self):
        for path, meta in self.metas.items():
            try:
                self.tree.item(
                    path,
                    values=(Path(path).name, meta.title[:80], meta.journal,
                            meta.year, meta.accepted_date,
                            build_filename(meta, self.tokens)),
                )
            except tk.TclError:
                pass

    def _do_rename(self):
        if not self.metas:
            messagebox.showinfo("提示", "请先扫描文件夹")
            return
        if not self.tokens:
            messagebox.showerror("错误", "命名规则为空")
            return
        if not messagebox.askyesno("确认", f"对 {len(self.metas)} 个文件执行重命名?"):
            return

        ok, errs = 0, []
        new_metas = {}
        for path, meta in self.metas.items():
            p = Path(path)
            if not p.exists():
                continue
            new_path = p.with_name(build_filename(meta, self.tokens))
            if new_path == p:
                new_metas[str(p)] = meta
                continue
            # 防覆盖
            if new_path.exists():
                stem, suf = new_path.stem, new_path.suffix
                k = 1
                while new_path.exists():
                    new_path = p.with_name(f"{stem} ({k}){suf}")
                    k += 1
            try:
                p.rename(new_path)
                new_metas[str(new_path)] = meta
                ok += 1
            except OSError as e:
                errs.append(f"{p.name}: {e}")
                new_metas[str(p)] = meta

        self.metas = new_metas
        self._scan_after_rename()
        msg = f"完成,成功重命名 {ok} 个"
        if errs:
            msg += f"\n失败 {len(errs)}:\n" + "\n".join(errs[:8])
        messagebox.showinfo("结果", msg)

    def _scan_after_rename(self):
        self.tree.delete(*self.tree.get_children())
        for path, meta in self.metas.items():
            self.tree.insert(
                "", "end", iid=path,
                values=(Path(path).name, meta.title[:80], meta.journal,
                        meta.year, meta.accepted_date,
                        build_filename(meta, self.tokens)),
            )


def main():
    if PdfReader is None:
        root = tk.Tk()
        root.withdraw()
        if messagebox.askyesno(
            "缺少依赖",
            "未检测到 pypdf 库, 是否现在自动安装?\n(将运行: pip install pypdf)",
        ):
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
                messagebox.showinfo("提示", "安装完成, 请重新启动本程序")
            except subprocess.CalledProcessError as e:
                messagebox.showerror("失败", f"安装失败: {e}\n请手动执行 pip install pypdf")
        return

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
