# -*- coding: utf-8 -*-
"""
东方财富 A 股行情数据抓取 —— 图形界面版

功能：
  1. 点击“打开浏览器”启动 Chrome 并加载行情页，可在浏览器中自行设置（筛选、排序等）。
  2. 点击“选择文件夹”指定数据保存目录，可自定义文件名，数据统一保存为 CSV。
  3. 点击“开始采集”自动翻页抓取全部数据，实时显示进度日志，可随时停止。

依赖：DrissionPage（pip install DrissionPage）。
"""

import csv
import os
import queue
import threading
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

from DrissionPage import ChromiumPage

# 表头（列名）
HEADERS = [
    "代码", "名称", "最新价", "涨跌幅", "涨跌额", "成交量(手)", "成交额",
    "振幅", "最高", "最低", "今开", "昨收", "量比", "换手率", "市盈率(动态)", "5分钟涨跌",
]

# 每一行 <td> 中需要的数据列索引（跳过：序号=0、相关链接=3、加自选=18）
COL_INDEX = [1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]

URL = "https://quote.eastmoney.com/center/gridlist.html#hs_a_board"

# 界面配色
C = {
    "bg": "#eef1f6",           # 窗口背景
    "card": "#ffffff",         # 卡片背景
    "primary": "#2f6bed",      # 主色
    "primary_hover": "#2458c6",
    "primary_active": "#1d47a2",
    "primary_disabled": "#a9c2f0",
    "secondary_hover": "#eaf1fe",
    "danger": "#ef4444",
    "danger_hover": "#dc2626",
    "text": "#1f2937",
    "muted": "#8a94a6",
    "border": "#e2e6ee",
    "log_bg": "#f8fafc",
}

STATUS_COLORS = {
    "idle": "#8a94a6",      # 灰
    "info": "#2563eb",      # 蓝
    "success": "#16a34a",   # 绿
    "running": "#f59e0b",   # 橙
    "error": "#dc2626",     # 红
}

FONT = "Microsoft YaHei UI"


# ---------------- 采集核心逻辑（与原脚本一致） ----------------

def extract_rows(page):
    """提取当前页表格中的所有股票数据，返回二维列表。"""
    rows = []
    trs = page.eles("css:div.quotetable table tbody tr")
    for tr in trs:
        tds = tr.eles("tag:td")
        if len(tds) < 18:  # 数据行共有 19 列，小于 18 说明不是有效数据行
            continue
        rows.append([tds[i].text.strip() for i in COL_INDEX])
    return rows


def current_page_num(page):
    """获取当前页码（分页条里 class=acitve 的那个链接，注意东财的拼写是 acitve）。"""
    active = page.ele("css:div.qtpager a.acitve")
    if active and active.text.strip().isdigit():
        return int(active.text.strip())
    return 0


def total_pages(page):
    """获取总页数（分页条里所有数字链接的最大值）。"""
    nums = []
    for a in page.eles("css:div.qtpager a"):
        t = a.text.strip()
        if t.isdigit():
            nums.append(int(t))
    return max(nums) if nums else 0


def save_csv(rows, filename):
    """保存为 CSV（utf-8-sig 编码，Excel 直接打开不乱码）。"""
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        writer.writerows(rows)


# ---------------- 图形界面 ----------------

class CollectorApp:
    def __init__(self, root):
        self.root = root
        root.title("东方财富 A 股行情采集")
        root.geometry("760x580")
        root.minsize(640, 520)
        root.configure(bg=C["bg"])

        # 状态变量
        self.dp = None                 # ChromiumPage 控制器
        self.browser_thread = None     # 浏览器工作线程
        self.start_event = threading.Event()  # 通知线程开始采集
        self.stop_collect = False      # 手动停止采集
        self.stop_all = False          # 关闭程序
        self.page_limit = None         # 采集页数（None = 全部）
        self.log_queue = queue.Queue() # 线程 -> 界面的消息队列

        self._setup_style()
        self._build_ui()
        self._poll_queue()             # 定期从队列取消息刷新界面

        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- 样式 ----------

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=(FONT, 10), background=C["bg"])

        # 卡片
        style.configure("Card.TLabelframe", background=C["card"],
                        bordercolor=C["border"], borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background=C["card"],
                        foreground=C["text"], font=(FONT, 10, "bold"))
        style.configure("Card.TLabel", background=C["card"], foreground=C["text"])
        style.configure("CardMuted.TLabel", background=C["card"], foreground=C["muted"])

        # 输入框 / 下拉框
        style.configure("TEntry", fieldbackground="#ffffff",
                        foreground=C["text"], bordercolor=C["border"],
                        lightcolor=C["border"], darkcolor=C["border"], padding=6)
        style.configure("TCombobox", fieldbackground="#ffffff",
                        foreground=C["text"], bordercolor=C["border"], padding=5)
        style.map("TCombobox",
                  fieldbackground=[("readonly", "#ffffff")],
                  selectbackground=[("readonly", "#ffffff")],
                  selectforeground=[("readonly", C["text"])])

        # 主按钮（蓝）
        style.configure("Primary.TButton", background=C["primary"],
                        foreground="#ffffff", borderwidth=0, focusthickness=0,
                        padding=(20, 9), font=(FONT, 10, "bold"))
        style.map("Primary.TButton",
                  background=[("active", C["primary_hover"]),
                              ("pressed", C["primary_active"]),
                              ("disabled", C["primary_disabled"])],
                  foreground=[("disabled", "#eaf1fe")])

        # 次按钮（白底蓝字）
        style.configure("Secondary.TButton", background="#ffffff",
                        foreground=C["primary"], borderwidth=0, focusthickness=0,
                        padding=(14, 7), font=(FONT, 10))
        style.map("Secondary.TButton",
                  background=[("active", C["secondary_hover"]),
                              ("pressed", "#d8e5fb"),
                              ("disabled", "#f1f3f8")],
                  foreground=[("disabled", "#b6bfcf")])

        # 危险按钮（红）
        style.configure("Danger.TButton", background=C["danger"],
                        foreground="#ffffff", borderwidth=0, focusthickness=0,
                        padding=(20, 9), font=(FONT, 10, "bold"))
        style.map("Danger.TButton",
                  background=[("active", C["danger_hover"]),
                              ("pressed", "#b91c1c"),
                              ("disabled", "#f3c1c1")],
                  foreground=[("disabled", "#fdecec")])

    # ---------- 界面搭建 ----------

    def _build_ui(self):
        # 顶部标题栏
        header = tk.Frame(self.root, bg=C["primary"])
        header.pack(fill="x")
        tk.Label(header, text="东方财富 A 股行情采集", bg=C["primary"], fg="#ffffff",
                 font=(FONT, 17, "bold")).pack(anchor="w", padx=20, pady=(18, 2))
        tk.Label(header, text="打开浏览器 → 在浏览器中设置 → 开始采集 → 选择文件夹保存",
                 bg=C["primary"], fg="#d8e4ff", font=(FONT, 9)).pack(
            anchor="w", padx=20, pady=(0, 16))

        # 保存设置卡片
        setting = ttk.LabelFrame(self.root, text="  保存设置  ", style="Card.TLabelframe",
                                 padding=14)
        setting.pack(fill="x", padx=16, pady=(14, 8))

        self.folder_var = tk.StringVar(value=os.getcwd())

        ttk.Label(setting, text="保存文件夹", style="Card.TLabel").grid(
            row=0, column=0, sticky="e", pady=6)
        ttk.Entry(setting, textvariable=self.folder_var).grid(
            row=0, column=1, columnspan=2, sticky="we", padx=(10, 6))
        ttk.Button(setting, text="选择文件夹…", style="Secondary.TButton",
                   cursor="hand2", command=self._choose_folder).grid(
            row=0, column=3, padx=(6, 0))

        self.page_var = tk.StringVar(value="全部")
        ttk.Label(setting, text="采集页数", style="Card.TLabel").grid(
            row=1, column=0, sticky="e", pady=6)
        ttk.Combobox(setting, textvariable=self.page_var, values=["全部"], width=12).grid(
            row=1, column=1, sticky="w", padx=(10, 0))
        ttk.Label(setting, text="（填数字采集前 N 页，填“全部”采集所有页）",
                  style="CardMuted.TLabel").grid(row=1, column=2, columnspan=2,
                                                 sticky="w", padx=(10, 0))

        setting.columnconfigure(1, weight=1)

        # 操作按钮区
        btn = tk.Frame(self.root, bg=C["bg"])
        btn.pack(fill="x", padx=16, pady=6)

        self.btn_open = ttk.Button(btn, text="🌐  打开浏览器", style="Primary.TButton",
                                   cursor="hand2", command=self._open_browser)
        self.btn_open.pack(side="left", padx=(0, 10))

        self.btn_start = ttk.Button(btn, text="▶  开始采集", style="Primary.TButton",
                                    cursor="hand2", state="disabled",
                                    command=self._start_collect)
        self.btn_start.pack(side="left", padx=(0, 10))

        self.btn_stop = ttk.Button(btn, text="■  停止采集", style="Danger.TButton",
                                   cursor="hand2", state="disabled",
                                   command=self._stop_collect)
        self.btn_stop.pack(side="left")

        # 状态栏
        self.status_var = tk.StringVar(value="未打开浏览器")
        self.status_label = tk.Label(self.root, textvariable=self.status_var,
                                     anchor="w", bg=STATUS_COLORS["idle"], fg="#ffffff",
                                     font=(FONT, 10), padx=14, pady=7)
        self.status_label.pack(fill="x", padx=16, pady=(6, 4))

        # 日志卡片
        logcard = ttk.LabelFrame(self.root, text="  运行日志  ", style="Card.TLabelframe",
                                 padding=8)
        logcard.pack(fill="both", expand=True, padx=16, pady=(4, 14))

        self.log_text = scrolledtext.ScrolledText(
            logcard, state="disabled", wrap="word", font=("Consolas", 9),
            bg=C["log_bg"], fg=C["text"], insertbackground=C["primary"],
            relief="flat", borderwidth=0, padx=10, pady=8, selectbackground="#cfe0fd")
        self.log_text.pack(fill="both", expand=True)

    # ---------- 按钮回调（主线程） ----------

    def _choose_folder(self):
        path = filedialog.askdirectory(initialdir=self.folder_var.get())
        if path:
            self.folder_var.set(path)

    def _open_browser(self):
        if self.browser_thread and self.browser_thread.is_alive():
            self._log("浏览器已经打开。")
            return
        self.btn_open.config(state="disabled")
        self._set_status("正在打开浏览器…", "info")
        # 在子线程中启动浏览器并等待采集命令
        self.browser_thread = threading.Thread(target=self._browser_worker, daemon=True)
        self.browser_thread.start()

    def _start_collect(self):
        try:
            self.page_limit = self._parse_page_limit()
        except ValueError:
            messagebox.showwarning("提示", "采集页数请输入正整数，或填“全部”。")
            return
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择有效的保存文件夹。")
            return
        if not self.dp:
            messagebox.showwarning("提示", "请先点击“打开浏览器”。")
            return
        self._set_ui("running")
        self.start_event.set()  # 唤醒工作线程开始采集

    def _stop_collect(self):
        self.stop_collect = True
        self._log("正在停止采集…")

    def _parse_page_limit(self):
        """解析采集页数：空/“全部”返回 None（采集全部），正整数返回 N，否则报错。"""
        raw = self.page_var.get().strip()
        if raw in ("", "全部", "所有"):
            return None
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        raise ValueError("无效的采集页数")

    # ---------- 工作线程：负责所有浏览器操作 ----------

    def _browser_worker(self):
        try:
            self._log("正在启动浏览器…")
            self.dp = ChromiumPage()
            self.dp.get(URL)
            time.sleep(3)  # 等待首页数据加载完成
            self._log("浏览器已打开并加载行情页，可在浏览器中设置后点击“开始采集”。")
            self._set_status("浏览器已打开，可开始采集", "success")
            self._set_ui("idle")
        except Exception as e:
            self._log(f"打开浏览器失败：{e}")
            self._set_status("浏览器打开失败", "error")
            self._set_ui("idle")
            return

        # 循环等待采集命令
        while not self.stop_all:
            self.start_event.wait()
            if self.stop_all:
                break
            self.start_event.clear()
            try:
                self._collect(self.dp)
            except Exception as e:
                self._log(f"采集出错：{e}")
                self._set_status("采集出错", "error")
            if not self.stop_all:
                self._set_ui("idle")

    def _collect(self, dp):
        """执行完整采集流程（在工作线程中运行）。"""
        self.stop_collect = False
        self._set_status("采集中…", "running")

        limit = self.page_limit
        total = total_pages(dp)
        if limit:
            target = min(limit, total) if total else limit
            self._log(f"共 {total} 页，本次采集前 {target} 页，开始抓取…")
        else:
            self._log(f"共 {total} 页，开始抓取…")

        all_rows = []
        page_num = current_page_num(dp) or 1

        while not self.stop_collect:
            rows = extract_rows(dp)
            if not rows:  # 可能还没加载好，重试一次
                time.sleep(2)
                rows = extract_rows(dp)

            if rows:
                all_rows.extend(rows)
                self._log(f"第 {page_num}/{total} 页：{len(rows)} 条，累计 {len(all_rows)} 条")
            else:
                self._log(f"第 {page_num} 页未抓到数据，结束。")
                break

            # 达到设定页数则停止
            if limit and page_num >= limit:
                self._log(f"已采集到设定的 {limit} 页，采集结束。")
                break

            # 点击“下一页”
            next_btn = dp.ele('tag:a@title=下一页')
            if next_btn is None:
                self._log("没有下一页按钮，抓取结束。")
                break

            before = page_num
            next_btn.click()
            time.sleep(2)  # 等待新一页数据加载

            after = current_page_num(dp)
            if after == 0 or after <= before:  # 页码没变化，说明已到最后一页
                self._log("已到最后一页，抓取结束。")
                break
            page_num = after

        if self.stop_all:
            return

        if self.stop_collect:
            self._log(f"已手动停止采集，当前累计 {len(all_rows)} 条。")
        else:
            self._log(f"抓取完成，共 {len(all_rows)} 条数据。")

        if all_rows:
            self._save(all_rows)

        self._set_status(f"采集结束，共 {len(all_rows)} 条", "success")

    def _save(self, rows):
        """把数据保存为 CSV 到指定文件夹，文件名用当天的年月日。"""
        folder = self.folder_var.get().strip() or os.getcwd()
        name = time.strftime("%Y%m%d", time.localtime())  # 例：20260928
        path = os.path.join(folder, name + ".csv")

        try:
            save_csv(rows, path)
            self._log(f"已保存 CSV：{path}")
        except Exception as e:
            self._log(f"保存失败：{e}")

    # ---------- 线程安全的界面更新 ----------

    def _log(self, msg):
        self.log_queue.put(("log", msg))

    def _set_status(self, text, level="idle"):
        color = STATUS_COLORS.get(level, STATUS_COLORS["idle"])
        self.log_queue.put(("status", (text, color)))

    def _set_ui(self, state):
        """state: running（采集中）/ idle（空闲）。"""
        self.log_queue.put(("ui", state))

    def _poll_queue(self):
        """每隔 100ms 从队列取出消息刷新界面（仅主线程操作界面）。"""
        try:
            while True:
                kind, val = self.log_queue.get_nowait()
                if kind == "log":
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", val + "\n")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                elif kind == "status":
                    text, color = val
                    self.status_var.set(text)
                    self.status_label.configure(bg=color)
                elif kind == "ui":
                    if val == "running":
                        self.btn_open.config(state="disabled")
                        self.btn_start.config(state="disabled")
                        self.btn_stop.config(state="normal")
                    else:  # idle
                        self.btn_open.config(state="normal")
                        self.btn_start.config(state="normal")
                        self.btn_stop.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ---------- 关闭程序 ----------

    def _on_close(self):
        self.stop_all = True
        self.stop_collect = True
        self.start_event.set()  # 唤醒工作线程以便退出
        if self.dp:
            try:
                self.dp.quit()
            except Exception:
                pass
        self.root.destroy()


def main():
    root = tk.Tk()
    CollectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
