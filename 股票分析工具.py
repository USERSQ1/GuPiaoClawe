# -*- coding: utf-8 -*-
"""
股票连续涨 / 连续跌分析工具

用法：
  1. 选择“涨”或“跌”文件夹（里面是每天一个 CSV，文件名形如 20260928.csv）。
  2. 选择开始日期和结束日期。
  3. 点击“开始分析”，找出区间内每一个交易日都出现在该文件夹里的股票
     （即连续上涨 / 连续下跌的股票）。

结果表格：
  - 表头可点击排序（点一下从小到大，再点一下从大到小）。
  - 除每天涨跌幅外，还统计：最新价、换手率、量比、市盈率(动态)、
    区间累计涨跌、区间最大单日涨幅/跌幅。

结果可一键导出为 CSV。
"""

import csv
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 界面配色（与采集工具一致）
C = {
    "bg": "#eef1f6",
    "card": "#ffffff",
    "primary": "#2f6bed",
    "primary_hover": "#2458c6",
    "primary_active": "#1d47a2",
    "primary_disabled": "#a9c2f0",
    "secondary_hover": "#eaf1fe",
    "text": "#1f2937",
    "muted": "#8a94a6",
    "border": "#e2e6ee",
    "log_bg": "#f8fafc",
}

STATUS_COLORS = {
    "idle": "#8a94a6",
    "info": "#2563eb",
    "success": "#16a34a",
    "running": "#f59e0b",
    "error": "#dc2626",
}

FONT = "Microsoft YaHei UI"
DATE_RE = re.compile(r"^(\d{8})\.csv$")  # 匹配 20260928.csv

# 固定列（分析字段），具体名称在分析时根据“涨/跌”动态替换最后一列
BASE_HEADERS = ["序号", "代码", "名称", "连续天数", "区间累计涨跌",
                "最新价", "换手率", "量比", "市盈率(动态)", "区间最大单日涨跌幅"]

FIXED_WIDTH = {
    "序号": 56, "代码": 78, "名称": 110, "连续天数": 74,
    "区间累计涨跌": 100, "最新价": 78, "换手率": 78, "量比": 66,
    "市盈率(动态)": 92, "区间最大单日涨幅": 122, "区间最大单日跌幅": 122,
}


# ---------------- 数据读取 ----------------

def scan_folder(folder):
    """扫描文件夹，返回 [(日期字符串, 文件路径)]，按日期升序。"""
    result = []
    try:
        names = os.listdir(folder)
    except OSError:
        return result
    for n in names:
        m = DATE_RE.match(n)
        if m:
            result.append((m.group(1), os.path.join(folder, n)))
    result.sort()
    return result


def read_stock_file(filepath):
    """读取一个 CSV，返回 {代码: 字段字典}。"""
    data = {}
    with open(filepath, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            code = (row.get("代码") or "").strip()
            if not code:
                continue
            data[code] = {
                "name": (row.get("名称") or "").strip(),
                "pct": (row.get("涨跌幅") or "").strip(),
                "price": (row.get("最新价") or "").strip(),
                "turnover": (row.get("换手率") or "").strip(),
                "vr": (row.get("量比") or "").strip(),
                "pe": (row.get("市盈率(动态)") or "").strip(),
            }
    return data


def parse_pct(s):
    """把 '20.00%' / '-3.5%' / '19.80' 等转成 float，失败返回 None。"""
    s = (s or "").strip().replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def longest_streak(present):
    """返回布尔序列中最长连续 True 的 (长度, 起始下标, 结束下标(不含))。
    并列时取最晚出现的一段。"""
    best_len = best_start = best_end = 0
    i, n = 0, len(present)
    while i < n:
        if not present[i]:
            i += 1
            continue
        j = i
        while j < n and present[j]:
            j += 1
        if j - i >= best_len:
            best_len, best_start, best_end = j - i, i, j
        i = j
    return best_len, best_start, best_end


# ---------------- 界面 ----------------

class AnalyzerApp:
    def __init__(self, root):
        self.root = root
        root.title("股票连续涨跌分析工具")
        root.geometry("1100x620")
        root.minsize(800, 520)
        root.configure(bg=C["bg"])

        self.files = []            # [(date, path)]
        self.fixed_headers = []    # 固定字段列名
        self.date_headers = []     # 日期列名
        self.result_headers = []   # 全部列名
        self.result_records = []   # [{"values": [...], "keys": [...]}]
        self.sort_col = None       # 当前排序列
        self.sort_desc = False     # 是否降序

        self._setup_style()
        self._build_ui()

    # ---------- 样式 ----------

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=(FONT, 10), background=C["bg"])
        style.configure("Card.TLabelframe", background=C["card"],
                        bordercolor=C["border"], borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background=C["card"],
                        foreground=C["text"], font=(FONT, 10, "bold"))
        style.configure("Card.TLabel", background=C["card"], foreground=C["text"])
        style.configure("CardMuted.TLabel", background=C["card"], foreground=C["muted"])

        style.configure("TEntry", fieldbackground="#ffffff", foreground=C["text"],
                        bordercolor=C["border"], lightcolor=C["border"],
                        darkcolor=C["border"], padding=6)
        style.configure("TCombobox", fieldbackground="#ffffff", foreground=C["text"],
                        bordercolor=C["border"], padding=5)
        style.map("TCombobox", fieldbackground=[("readonly", "#ffffff")],
                  selectbackground=[("readonly", "#ffffff")],
                  selectforeground=[("readonly", C["text"])])

        style.configure("Primary.TButton", background=C["primary"], foreground="#ffffff",
                        borderwidth=0, focusthickness=0, padding=(20, 9),
                        font=(FONT, 10, "bold"))
        style.map("Primary.TButton",
                  background=[("active", C["primary_hover"]),
                              ("pressed", C["primary_active"]),
                              ("disabled", C["primary_disabled"])],
                  foreground=[("disabled", "#eaf1fe")])

        style.configure("Secondary.TButton", background="#ffffff", foreground=C["primary"],
                        borderwidth=0, focusthickness=0, padding=(14, 7), font=(FONT, 10))
        style.map("Secondary.TButton",
                  background=[("active", C["secondary_hover"]),
                              ("pressed", "#d8e5fb"),
                              ("disabled", "#f1f3f8")],
                  foreground=[("disabled", "#b6bfcf")])

        style.configure("Treeview", rowheight=26, font=(FONT, 9),
                        fieldbackground="#ffffff", background="#ffffff",
                        foreground=C["text"])
        style.configure("Treeview.Heading", font=(FONT, 9, "bold"),
                        background="#eef2f9", foreground=C["text"],
                        relief="flat", padding=(6, 6))

    # ---------- 界面搭建 ----------

    def _build_ui(self):
        header = tk.Frame(self.root, bg=C["primary"])
        header.pack(fill="x")
        tk.Label(header, text="股票连续涨跌分析工具", bg=C["primary"], fg="#ffffff",
                 font=(FONT, 17, "bold")).pack(anchor="w", padx=20, pady=(18, 2))
        tk.Label(header, text="选择涨/跌文件夹 → 选择日期区间 → 分析区间内连续上涨或下跌的股票（点击表头可排序）",
                 bg=C["primary"], fg="#d8e4ff", font=(FONT, 9)).pack(
            anchor="w", padx=20, pady=(0, 16))

        # 设置卡片
        setting = ttk.LabelFrame(self.root, text="  分析设置  ", style="Card.TLabelframe",
                                 padding=14)
        setting.pack(fill="x", padx=16, pady=(14, 8))

        self.folder_var = tk.StringVar(value="")
        self.start_var = tk.StringVar(value="")
        self.end_var = tk.StringVar(value="")
        self.min_days_var = tk.StringVar(value="≥1 天")

        ttk.Label(setting, text="数据文件夹", style="Card.TLabel").grid(
            row=0, column=0, sticky="e", pady=6)
        ttk.Entry(setting, textvariable=self.folder_var).grid(
            row=0, column=1, columnspan=3, sticky="we", padx=(10, 6))
        ttk.Button(setting, text="选择文件夹…", style="Secondary.TButton",
                   cursor="hand2", command=self._choose_folder).grid(
            row=0, column=4, padx=(6, 0))

        ttk.Label(setting, text="开始日期", style="Card.TLabel").grid(
            row=1, column=0, sticky="e", pady=6)
        self.start_combo = ttk.Combobox(setting, textvariable=self.start_var,
                                        state="readonly", width=14)
        self.start_combo.grid(row=1, column=1, sticky="w", padx=(10, 0))

        ttk.Label(setting, text="结束日期", style="Card.TLabel").grid(
            row=1, column=2, sticky="e", padx=(24, 0))
        self.end_combo = ttk.Combobox(setting, textvariable=self.end_var,
                                      state="readonly", width=14)
        self.end_combo.grid(row=1, column=3, sticky="w", padx=(10, 0))

        ttk.Label(setting, text="最少连续天数", style="Card.TLabel").grid(
            row=1, column=4, sticky="e", padx=(28, 0))
        self.min_days_combo = ttk.Combobox(setting, textvariable=self.min_days_var,
                                           state="readonly", width=10)
        self.min_days_combo.grid(row=1, column=5, sticky="w", padx=(10, 0))
        self.min_days_combo.bind("<<ComboboxSelected>>", self._on_min_days_change)
        self.start_combo.bind("<<ComboboxSelected>>", self._update_min_days_options)
        self.end_combo.bind("<<ComboboxSelected>>", self._update_min_days_options)

        self.info_var = tk.StringVar(value="请先选择数据文件夹")
        ttk.Label(setting, textvariable=self.info_var, style="CardMuted.TLabel").grid(
            row=2, column=1, columnspan=5, sticky="w", padx=(10, 0), pady=(4, 0))

        setting.columnconfigure(1, weight=1)

        # 操作按钮
        btn = tk.Frame(self.root, bg=C["bg"])
        btn.pack(fill="x", padx=16, pady=6)

        self.btn_run = ttk.Button(btn, text="▶  开始分析", style="Primary.TButton",
                                  cursor="hand2", command=self._analyze)
        self.btn_run.pack(side="left", padx=(0, 10))

        self.btn_export = ttk.Button(btn, text="导出结果 CSV", style="Secondary.TButton",
                                     cursor="hand2", state="disabled",
                                     command=self._export)
        self.btn_export.pack(side="left")

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = tk.Label(self.root, textvariable=self.status_var, anchor="w",
                                     bg=STATUS_COLORS["idle"], fg="#ffffff",
                                     font=(FONT, 10), padx=14, pady=7)
        self.status_label.pack(fill="x", padx=16, pady=(6, 4))

        # 结果卡片
        result = ttk.LabelFrame(self.root, text="  分析结果（点击表头排序）  ",
                                style="Card.TLabelframe", padding=8)
        result.pack(fill="both", expand=True, padx=16, pady=(4, 14))

        self.tree = ttk.Treeview(result, columns=BASE_HEADERS, show="headings", height=12)
        for c in BASE_HEADERS:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=FIXED_WIDTH.get(c, 90), anchor="center")

        vsb = ttk.Scrollbar(result, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(result, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        result.rowconfigure(0, weight=1)
        result.columnconfigure(0, weight=1)

    # ---------- 交互 ----------

    def _choose_folder(self):
        path = filedialog.askdirectory(initialdir=self.folder_var.get() or os.getcwd())
        if not path:
            return
        self.folder_var.set(path)
        self.files = scan_folder(path)
        dates = [d for d, _ in self.files]

        if not dates:
            self.start_combo.configure(values=[])
            self.end_combo.configure(values=[])
            self.min_days_combo.configure(values=[])
            self.start_var.set("")
            self.end_var.set("")
            self.info_var.set("该文件夹内没有找到日期文件（文件名应为 8 位数字.csv）")
            return

        self.start_combo.configure(values=dates)
        self.end_combo.configure(values=dates)
        self.start_var.set(dates[0])
        self.end_var.set(dates[-1])
        self._update_min_days_options()

        target = self._target_name()
        self.info_var.set(f"共 {len(dates)} 个交易日（{dates[0]} ~ {dates[-1]}），"
                          f"分析目标：{target}的股票")

    def _target_name(self):
        """根据文件夹名判断是分析连续涨还是连续跌。"""
        base = os.path.basename(self.folder_var.get().strip())
        return "连续下跌" if "跌" in base else "连续上涨"

    def _min_days_options(self, n):
        """生成 1~n 天的“≥N 天”选项。"""
        return [f"≥{i} 天" for i in range(1, max(n, 1) + 1)]

    def _update_min_days_options(self, event=None):
        """根据当前选择的日期区间刷新“最少连续天数”下拉选项。"""
        dates = [d for d, _ in self.files]
        start, end = self.start_var.get(), self.end_var.get()
        if start in dates and end in dates:
            i0, i1 = dates.index(start), dates.index(end)
            if i0 > i1:
                i0, i1 = i1, i0
            opts = self._min_days_options(i1 - i0 + 1)
            self.min_days_combo.configure(values=opts)
            if self.min_days_var.get() not in opts:
                self.min_days_var.set(opts[0])
        else:
            self.min_days_combo.configure(values=[])

    def _min_days_value(self):
        """解析当前“最少连续天数”为整数，解析失败返回 1。"""
        m = re.match(r"≥\s*(\d+)", self.min_days_var.get())
        return int(m.group(1)) if m else 1

    def _on_min_days_change(self, event=None):
        """切换最少连续天数后直接刷新表格，不重新读取文件。"""
        if not self.result_records:
            return
        min_days = self._min_days_value()
        shown = len(self._filtered_records())
        self._render_tree()
        self._set_status(f"连续 ≥ {min_days} 天，共 {shown} 只", "info")

    def _analyze(self):
        if not self.files:
            messagebox.showwarning("提示", "请先选择包含日期文件的文件夹。")
            return
        dates = [d for d, _ in self.files]
        start = self.start_var.get()
        end = self.end_var.get()
        if start not in dates or end not in dates:
            messagebox.showwarning("提示", "请选择有效的开始和结束日期。")
            return

        i0, i1 = dates.index(start), dates.index(end)
        if i0 > i1:
            i0, i1 = i1, i0
        selected = self.files[i0:i1 + 1]

        self._update_min_days_options()
        min_days = self._min_days_value()

        self._set_status("分析中…", "running")
        self.root.update_idletasks()

        # 读取区间内每天的数据
        per_date = []
        for d, path in selected:
            per_date.append((d, read_stock_file(path)))

        # 统计区间内出现过的全部股票（不再要求每天都出现），
        # 每只股票的“连续天数”在 _build_result 中按最长连续交易日计算。
        codes = set()
        for _, data in per_date:
            codes |= set(data.keys())

        self._build_result(codes, per_date, self._target_name())
        shown = sum(1 for r in self.result_records if r["streak"] >= min_days)
        self._set_status(f"分析完成：{shown} 只{self._target_name()}股票"
                         f"（连续 ≥ {min_days} 天，区间共 {len(per_date)} 个交易日）",
                         "success")

    def _build_result(self, codes, per_date, target):
        """生成结果（含字段与排序键）并渲染到表格。"""
        date_labels = [d for d, _ in per_date]
        n_days = len(per_date)

        # 动态确定“区间最大单日涨/跌幅”字段名
        if "跌" in target:
            extreme_label = "区间最大单日跌幅"
        else:
            extreme_label = "区间最大单日涨幅"

        self.fixed_headers = ["序号", "代码", "名称", "连续天数", "区间累计涨跌",
                              "最新价", "换手率", "量比", "市盈率(动态)", extreme_label]
        self.date_headers = date_labels
        self.result_headers = self.fixed_headers + self.date_headers
        self.result_records = []

        for i, code in enumerate(sorted(codes), start=1):
            # 该股票在区间内每个交易日是否出现
            present = [code in per_date[j][1] for j in range(n_days)]
            # 每个交易日的涨跌幅（未出现的交易日记为 None，显示为 —）
            pcts = [per_date[j][1][code]["pct"] if present[j] else None
                    for j in range(n_days)]
            nums = [parse_pct(p) for p in pcts]

            # 最长连续上涨/下跌交易日（并列时取最晚一段）
            streak, win_start, win_end = longest_streak(present)

            # 区间累计涨跌 = 在最长连续区间内 ∏(1+p/100) - 1
            cum = None
            if all(nums[j] is not None for j in range(win_start, win_end)):
                acc = 1.0
                for j in range(win_start, win_end):
                    acc *= (1 + nums[j] / 100)
                cum = (acc - 1) * 100

            # 区间最大单日涨/跌幅（同样取自最长连续区间）
            seg = [nums[j] for j in range(win_start, win_end) if nums[j] is not None]
            if "跌" in target:
                mn = min(seg, default=None)
                extreme = -mn if mn is not None else None
            else:
                extreme = max(seg, default=None)

            # 最新数据取该股票最后一次出现的交易日
            last_idx = max(j for j in range(n_days) if present[j])
            last = per_date[last_idx][1][code]
            price = last["price"]
            turnover = last["turnover"]
            vr = last["vr"]
            pe = last["pe"]

            price_num = parse_pct(price)
            turnover_num = parse_pct(turnover)
            vr_num = parse_pct(vr)
            pe_num = parse_pct(pe)

            values = [
                str(i), code, last["name"], f"{streak} 天",
                f"{cum:+.2f}%" if cum is not None else "—",
                price or "—", turnover or "—", vr or "—", pe or "—",
                f"{extreme:.2f}%" if extreme is not None else "—",
            ] + [p if p is not None else "—" for p in pcts]

            keys = [
                i, code, last["name"], streak,
                cum if cum is not None else float("-inf"),
                price_num if price_num is not None else float("-inf"),
                turnover_num if turnover_num is not None else float("-inf"),
                vr_num if vr_num is not None else float("-inf"),
                pe_num if pe_num is not None else float("-inf"),
                extreme if extreme is not None else float("-inf"),
            ] + [n if n is not None else float("-inf") for n in nums]

            self.result_records.append(
                {"values": values, "keys": keys, "streak": streak})

        self.sort_col = None
        self.sort_desc = False
        self.btn_export.config(state="normal")
        self._render_tree()

    def _filtered_records(self):
        """按当前“最少连续天数”阈值过滤并重新编号，返回展示行列表。"""
        min_days = self._min_days_value()
        rows = []
        seq = 0
        for rec in self.result_records:
            if rec["streak"] < min_days:
                continue
            seq += 1
            values = list(rec["values"])
            values[0] = str(seq)
            rows.append(values)
        return rows

    def _render_tree(self):
        """按当前列与排序状态渲染表格（重建表头并绑定点击排序，并应用连续天数筛选）。"""
        tree = self.tree
        tree.configure(columns=self.result_headers)
        for c in self.result_headers:
            text = c
            if c == self.sort_col:
                text = c + ("  ▼" if self.sort_desc else "  ▲")
            tree.heading(c, text=text, command=lambda col=c: self._sort_by(col))

            if c in self.date_headers:
                width, anchor, stretch = 74, "center", False
            elif c == "名称":
                width, anchor, stretch = 110, "w", True
            else:
                width, anchor, stretch = FIXED_WIDTH.get(c, 92), "center", False
            tree.column(c, width=width, anchor=anchor, stretch=stretch)

        tree.delete(*tree.get_children())
        for values in self._filtered_records():
            tree.insert("", "end", values=values)

    def _sort_by(self, col):
        """点击表头：同列切换升/降序，换列则从升序开始。"""
        if col not in self.result_headers:
            return
        idx = self.result_headers.index(col)
        if self.sort_col == col:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_col = col
            self.sort_desc = False
        self.result_records.sort(key=lambda r: r["keys"][idx], reverse=self.sort_desc)
        self._render_tree()

    def _export(self):
        if not self.result_records:
            messagebox.showwarning("提示", "没有可导出的结果。")
            return
        rows = self._filtered_records()
        if not rows:
            messagebox.showwarning("提示", "当前筛选条件下没有可导出的结果。")
            return
        default_name = (f"{self._target_name()}分析_连续≥{self._min_days_value()}天_"
                        f"{self.start_var.get()}_至_{self.end_var.get()}.csv")
        path = filedialog.asksaveasfilename(
            initialdir=os.path.dirname(self.folder_var.get()) or os.getcwd(),
            initialfile=default_name, defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(self.result_headers)
                for row in rows:
                    writer.writerow(row)
            self._set_status(f"已导出：{path}", "success")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _set_status(self, text, level="idle"):
        self.status_var.set(text)
        self.status_label.configure(bg=STATUS_COLORS.get(level, STATUS_COLORS["idle"]))


def main():
    root = tk.Tk()
    AnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
