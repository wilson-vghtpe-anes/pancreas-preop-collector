# -*- coding: utf-8 -*-
"""
胰頭手術術前資料收集 —— 操作介面

雙擊「執行.bat」啟動。
  1. 左邊貼上病歷號（可直接從 Excel 整欄複製貼上）
  2. 右邊勾選要查的術前檢驗：生化（CHEM）與血液（CBC/DC／凝血）兩個分頁，
     候選清單取自兩份累積報告的表頭，見 lab_items.py
  3. 按「開始收集」，跑完會產生 Excel

查詢邏輯都在 收集胰頭手術資料.py，這支只負責介面。
勾選的項目、速度設定會記在 介面設定.json，下次打開自動帶入。
"""

from __future__ import annotations

import json
import os
import queue
import sys
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import eip_session                       # noqa: E402
import lab_items                         # noqa: E402
import 收集胰頭手術資料 as core           # noqa: E402

SETTINGS_FILE = os.path.join(BASE_DIR, "介面設定.json")
FONT = "Microsoft JhengHei"

# 各分頁勾選清單的排法：每欄放哪幾個分類（名稱要跟 lab_items.REPORTS 一致）。
# 沒列在這裡的報告，分類會依序排成三欄。
GROUP_COLUMNS = {
    "DCHEM": [["營養／蛋白", "肝膽"],
              ["腎功能", "電解質"],
              ["血糖／血脂", "心肌／發炎", "鐵"]],
    "DCBC": [["血球計數", "紅血球指數"],
             ["白血球分類 DC（%）"],
             ["凝血"]],
}


def _columns_for(rtype):
    if rtype in GROUP_COLUMNS:
        return GROUP_COLUMNS[rtype]
    names = [g for g, _ in lab_items.GROUPS[rtype]]
    return [names[i::3] for i in range(min(3, len(names)))]


def load_settings() -> dict:
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data: dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


# ============================ EIP 帳號對話框 ============================

class CredentialDialog(tk.Toplevel):
    """輸入 EIP 帳號密碼。result = (帳號, 密碼, 是否記住) 或 None（取消）。"""

    def __init__(self, master, acc="", pwd=""):
        super().__init__(master)
        self.title("EIP 帳號")
        self.resizable(False, False)
        self.transient(master)
        self.result = None

        f = ttk.Frame(self, padding=14)
        f.grid()
        ttk.Label(f, text="登入 EIP 用的帳號密碼").grid(row=0, column=0, columnspan=2,
                                                     sticky="w", pady=(0, 8))
        ttk.Label(f, text="帳號").grid(row=1, column=0, sticky="w", pady=3)
        self.v_acc = tk.StringVar(value=acc)
        e_acc = ttk.Entry(f, textvariable=self.v_acc, width=24)
        e_acc.grid(row=1, column=1, sticky="w")
        ttk.Label(f, text="密碼").grid(row=2, column=0, sticky="w", pady=3)
        self.v_pwd = tk.StringVar(value=pwd)
        ttk.Entry(f, textvariable=self.v_pwd, width=24, show="●").grid(row=2, column=1, sticky="w")
        self.v_keep = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="記住（存進 eip帳號.txt，下次不用再輸入）",
                        variable=self.v_keep).grid(row=3, column=0, columnspan=2,
                                                   sticky="w", pady=(8, 0))
        box = ttk.Frame(f)
        box.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(box, text="確定", command=self._ok).pack(side="left")
        ttk.Button(box, text="取消", command=self.destroy).pack(side="left", padx=(6, 0))

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        e_acc.focus_set()
        self.grab_set()
        self.wait_window()

    def _ok(self):
        acc, pwd = self.v_acc.get().strip(), self.v_pwd.get().strip()
        if not (acc and pwd):
            messagebox.showwarning("EIP 帳號", "帳號和密碼都要填。", parent=self)
            return
        self.result = (acc, pwd, self.v_keep.get())
        self.destroy()


# ============================ 主視窗 ============================

class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)
        self.master = master
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        self.q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.result_path = None
        self.session_cred = None          # 不記住帳密時，只存在這次開著的期間

        st = load_settings()
        self.lab_vars: dict[str, tk.BooleanVar] = {}
        self._build(st)
        self._load_histnos()
        self._refresh()
        self.after(120, self._drain)
        master.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------ 版面 ------------------------------
    def _build(self, st):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(4, weight=1)

        ttk.Label(self, text="胰頭手術術前資料收集",
                  font=(FONT, 14, "bold")).grid(row=0, column=0, columnspan=2,
                                                sticky="w", pady=(0, 10))

        # ---- 左：病歷號 ----
        left = ttk.LabelFrame(self, text="病歷號", padding=8)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left.rowconfigure(1, weight=1)
        ttk.Label(left, text="一行一個，可從 Excel 貼上",
                  foreground="#666").grid(row=0, column=0, columnspan=2, sticky="w")
        self.txt_ids = tk.Text(left, width=16, height=18, font=("Consolas", 10), undo=True)
        self.txt_ids.grid(row=1, column=0, sticky="nsew", pady=(4, 4))
        sb = ttk.Scrollbar(left, orient="vertical", command=self.txt_ids.yview)
        sb.grid(row=1, column=1, sticky="ns", pady=(4, 4))
        self.txt_ids.configure(yscrollcommand=sb.set)
        self.txt_ids.bind("<<Modified>>", self._on_ids_modified)
        self.lbl_count = ttk.Label(left, text="")
        self.lbl_count.grid(row=2, column=0, columnspan=2, sticky="w")
        box = ttk.Frame(left)
        box.grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Button(box, text="從檔案載入…", command=self._import_ids).pack(side="left")
        ttk.Button(box, text="清空", width=5,
                   command=lambda: self.txt_ids.delete("1.0", "end")).pack(side="left", padx=(4, 0))

        # ---- 右：檢驗項目（每份累積報告一個分頁） ----
        right = ttk.LabelFrame(self, text="術前檢驗項目（EMR 累積報告）", padding=8)
        right.grid(row=1, column=1, sticky="nsew")

        picked = set(lab_items.normalize(st.get("labs", lab_items.DEFAULT_LABS)))
        old_extra = st.get("extra", {})
        if isinstance(old_extra, str):          # 舊版設定：只有一格（生化）
            old_extra = {lab_items.REPORT_TYPES[0]: old_extra}
        self.extra_vars: dict[str, tk.StringVar] = {}
        self.extra_lbls: dict[str, ttk.Label] = {}
        self.tabs: dict[str, ttk.Frame] = {}

        self.nb = ttk.Notebook(right)
        self.nb.grid(row=0, column=0, sticky="nw")
        for rtype in lab_items.REPORT_TYPES:
            tab = ttk.Frame(self.nb, padding=8)
            self.nb.add(tab, text=lab_items.REPORT_TITLE[rtype])
            self.tabs[rtype] = tab
            self._build_report_tab(tab, rtype, picked, old_extra.get(rtype, ""), st)

        self.lbl_picked = ttk.Label(right, text="", foreground="#0a5")
        self.lbl_picked.grid(row=1, column=0, sticky="w", pady=(6, 0))

        # 兩份報告都依月數查：resdtmonth = 手術至今的月數＋往前多查幾個月
        box = ttk.Frame(right)
        box.grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Label(box, text="檢驗查詢範圍：手術至今的月數，再往前多查").pack(side="left")
        self.var_extra_months = tk.IntVar(value=int(st.get("extra_months", core.EXTRA_MONTHS)))
        ttk.Spinbox(box, from_=0, to=60, increment=1, width=4,
                    textvariable=self.var_extra_months).pack(side="left", padx=(4, 2))
        ttk.Label(box, text="個月（比這更早的檢驗不採用）", foreground="#666").pack(side="left")

        self.var_vitals = tk.BooleanVar(value=st.get("vitals", True))
        self.var_vitals.trace_add("write", lambda *_: self._refresh())
        ttk.Checkbutton(right, text="另外查術前身高、體重，並計算 BMI（每位多開一頁）",
                        variable=self.var_vitals).grid(row=3, column=0, sticky="w", pady=(4, 0))
        ttk.Label(right, text="同一個分頁的項目在同一頁，多勾幾項不會變慢；兩個分頁都有勾，"
                              "每位多開一頁。每項取手術開始前最新的一筆。",
                  foreground="#666").grid(row=4, column=0, sticky="w", pady=(4, 0))

        # ---- 速度與選項 ----
        opt = ttk.LabelFrame(self, text="速度與選項", padding=8)
        opt.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Label(opt, text="每開一頁至少間隔").grid(row=0, column=0, sticky="w")
        self.var_interval = tk.DoubleVar(value=st.get("interval", eip_session.REQUEST_INTERVAL))
        ttk.Spinbox(opt, from_=1, to=60, increment=0.5, width=5,
                    textvariable=self.var_interval,
                    command=self._refresh).grid(row=0, column=1, padx=(4, 2))
        ttk.Label(opt, text="秒　　每位病人之間再停").grid(row=0, column=2, sticky="w")
        self.var_pdelay = tk.DoubleVar(value=st.get("patient_delay", core.PATIENT_DELAY))
        ttk.Spinbox(opt, from_=0, to=300, increment=1, width=5,
                    textvariable=self.var_pdelay,
                    command=self._refresh).grid(row=0, column=3, padx=(4, 2))
        ttk.Label(opt, text="秒").grid(row=0, column=4, sticky="w")
        self.lbl_est = ttk.Label(opt, text="", foreground="#0a5", font=(FONT, 10, "bold"))
        self.lbl_est.grid(row=0, column=5, sticky="w", padx=(16, 0))
        for v in (self.var_interval, self.var_pdelay):
            v.trace_add("write", lambda *_: self._refresh())

        box = ttk.Frame(opt)
        box.grid(row=1, column=0, columnspan=7, sticky="w", pady=(6, 0))
        self.var_restart = tk.BooleanVar(value=False)
        ttk.Checkbutton(box, text="全部重查（不接續上次進度）",
                        variable=self.var_restart).pack(side="left")
        self.var_visible = tk.BooleanVar(value=st.get("visible", False))
        ttk.Checkbutton(box, text="顯示 Chrome 畫面",
                        variable=self.var_visible).pack(side="left", padx=(14, 0))
        self.var_save_html = tk.BooleanVar(value=st.get("save_html", False))
        ttk.Checkbutton(box, text="存下網頁（除錯用，含病人資料）",
                        variable=self.var_save_html).pack(side="left", padx=(14, 0))
        ttk.Button(box, text="EIP 帳號…", command=self._edit_credentials).pack(side="left",
                                                                             padx=(14, 0))

        # ---- 動作列 ----
        act = ttk.Frame(self)
        act.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 6))
        act.columnconfigure(3, weight=1)
        self.btn_start = ttk.Button(act, text="開始收集", command=self._start)
        self.btn_start.grid(row=0, column=0)
        self.btn_stop = ttk.Button(act, text="停止", command=self._stop, state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=(6, 0))
        self.lbl_status = ttk.Label(act, text="", foreground="#666")
        self.lbl_status.grid(row=0, column=2, padx=(12, 0), sticky="w")
        self.pbar = ttk.Progressbar(act, mode="determinate", length=200)
        self.pbar.grid(row=0, column=3, sticky="ew", padx=(12, 12))
        self.btn_open = ttk.Button(act, text="開啟結果", command=self._open_result,
                                   state="disabled")
        self.btn_open.grid(row=0, column=4)
        ttk.Button(act, text="開啟資料夾",
                   command=lambda: self._startfile(BASE_DIR)).grid(row=0, column=5, padx=(6, 0))

        # ---- 訊息區 ----
        logf = ttk.Frame(self)
        logf.grid(row=4, column=0, columnspan=2, sticky="nsew")
        logf.columnconfigure(0, weight=1)
        logf.rowconfigure(0, weight=1)
        self.log = tk.Text(logf, height=10, wrap="word", font=("Consolas", 9),
                           state="disabled", background="#1e1e1e", foreground="#ddd")
        self.log.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(logf, orient="vertical", command=self.log.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=sb.set)

    def _build_report_tab(self, tab, rtype, picked, extra_text, st):
        """一份報告的分頁：分類勾選框、快速按鈕、其他代碼。"""
        groups = dict(lab_items.GROUPS[rtype])
        grid = ttk.Frame(tab)
        grid.grid(row=0, column=0, columnspan=3, sticky="nw")
        for col, names in enumerate(_columns_for(rtype)):
            colf = ttk.Frame(grid)
            colf.grid(row=0, column=col, sticky="n", padx=(0 if col == 0 else 8, 0))
            for gname in names:
                gf = ttk.LabelFrame(colf, text=gname, padding=(6, 2))
                gf.pack(fill="x", anchor="n", pady=(0, 6))
                for code, name in groups[gname]:
                    lid = f"{rtype}:{code}"
                    v = tk.BooleanVar(value=lid in picked)
                    v.trace_add("write", lambda *_: self._refresh())
                    self.lab_vars[lid] = v
                    ttk.Checkbutton(gf, text=f"{code:<6} {name}", variable=v).pack(anchor="w")

        box = ttk.Frame(tab)
        box.grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 6))
        for label, codes in (("預設", lab_items.DEFAULTS.get(rtype)),
                             ("常用", lab_items.COMMON.get(rtype))):
            if codes:
                ttk.Button(box, text=f"{label}（{'／'.join(codes)}）",
                           command=lambda r=rtype, c=codes: self._set_labs(r, c)
                           ).pack(side="left", padx=(0, 4))
        ttk.Button(box, text="全選",
                   command=lambda r=rtype: self._set_labs(r, lab_items.CODES[r])
                   ).pack(side="left", padx=(0, 4))
        ttk.Button(box, text="全不選",
                   command=lambda r=rtype: self._set_labs(r, ())).pack(side="left")

        ttk.Label(tab, text="其他代碼").grid(row=2, column=0, sticky="w")
        var = tk.StringVar(value=extra_text)
        var.trace_add("write", lambda *_: self._refresh())
        self.extra_vars[rtype] = var
        ttk.Entry(tab, textvariable=var, width=26).grid(row=2, column=1, sticky="w", padx=(6, 0))
        lbl = ttk.Label(tab, text="", foreground="#999")
        lbl.grid(row=2, column=2, sticky="w", padx=(6, 0))
        self.extra_lbls[rtype] = lbl


    # ------------------------------ 讀取狀態 ------------------------------
    def _labs(self) -> list[str]:
        checked = [lid for lid, v in self.lab_vars.items() if v.get()]
        for rtype, var in self.extra_vars.items():
            checked += lab_items.normalize(var.get(), default_report=rtype)
        return lab_items.normalize(checked)

    def _histnos(self) -> list[str]:
        return core.parse_histno_text(self.txt_ids.get("1.0", "end"))

    def _num(self, var, default):
        try:
            return max(0.0, float(var.get()))
        except (tk.TclError, ValueError):
            return default

    def _settings(self) -> dict:
        return {
            "labs": [lid for lid, v in self.lab_vars.items() if v.get()],
            "extra": {r: v.get().strip() for r, v in self.extra_vars.items()},
            "vitals": self.var_vitals.get(),
            "interval": self._num(self.var_interval, eip_session.REQUEST_INTERVAL),
            "patient_delay": self._num(self.var_pdelay, core.PATIENT_DELAY),
            "visible": self.var_visible.get(),
            "save_html": self.var_save_html.get(),
            "extra_months": self._extra_months(),
        }

    def _extra_months(self) -> int:
        try:
            return max(0, int(self.var_extra_months.get()))
        except (AttributeError, tk.TclError, ValueError):
            return core.EXTRA_MONTHS

    # ------------------------------ 互動 ------------------------------
    def _set_labs(self, rtype, codes):
        """只改這一份報告的勾選，其他分頁不動。"""
        want = {f"{rtype}:{c}" for c in codes}
        for lid, v in self.lab_vars.items():
            if lid.startswith(rtype + ":"):
                v.set(lid in want)

    def _on_ids_modified(self, _event=None):
        if self.txt_ids.edit_modified():
            self.txt_ids.edit_modified(False)
            self._refresh()

    def _refresh(self):
        n = len(self._histnos())
        self.lbl_count.configure(text=f"共 {n} 位")
        labs = self._labs()
        per = lab_items.by_report(labs)
        for rtype, tab in self.tabs.items():
            k = len(per.get(rtype, ()))
            self.nb.tab(tab, text=lab_items.REPORT_TITLE[rtype] + (f"（{k}）" if k else ""))
            bad = lab_items.invalid_tokens(self.extra_vars[rtype].get())
            if bad:
                self.extra_lbls[rtype].configure(text=f"看不懂、會略過：{'、'.join(bad)}",
                                                 foreground="#c00")
            else:
                self.extra_lbls[rtype].configure(text="清單沒有的項目，打代碼、逗號分隔",
                                                 foreground="#999")
        parts = [f"{lab_items.REPORT_TITLE[r]} {len(c)} 項" for r, c in per.items()]
        txt = f"已選 {len(labs)} 項" + (f"：{'、'.join(parts)}" if parts else "")
        if per:
            txt += f"（每位開 {len(per)} 頁檢驗報告）"
        self.lbl_picked.configure(text=txt)

        interval = self._num(self.var_interval, eip_session.REQUEST_INTERVAL)
        pdelay = self._num(self.var_pdelay, core.PATIENT_DELAY)
        mins = core.estimate_minutes(n, labs, self.var_vitals.get(), interval, pdelay)
        if n == 0:
            est = ""
        elif mins < 90:
            est = f"{n} 位預估約 {max(1, round(mins))} 分鐘"
        else:
            est = f"{n} 位預估約 {mins / 60:.1f} 小時"
        warn = interval < 2
        self.lbl_est.configure(text=est + ("　（間隔太短容易被鎖）" if warn else ""),
                               foreground="#c00" if warn else "#0a5")

    def _load_histnos(self):
        if os.path.exists(core.DEFAULT_LIST):
            try:
                self.txt_ids.insert("1.0", core.read_text_any(core.DEFAULT_LIST))
            except Exception:
                pass
        self.txt_ids.edit_modified(False)

    def _import_ids(self):
        path = filedialog.askopenfilename(
            title="選擇病歷號清單", initialdir=BASE_DIR,
            filetypes=[("文字檔", "*.txt *.csv"), ("所有檔案", "*.*")])
        if not path:
            return
        ids = core.read_histno_list(path)
        if not ids:
            messagebox.showwarning("載入病歷號", "這個檔案裡沒有讀到病歷號。")
            return
        self.txt_ids.delete("1.0", "end")
        self.txt_ids.insert("1.0", "\n".join(ids) + "\n")
        self._write(f"[清單] 從 {os.path.basename(path)} 載入 {len(ids)} 位")

    def _edit_credentials(self):
        acc, pwd = self.session_cred or eip_session.read_credentials()
        dlg = CredentialDialog(self.master, acc, pwd)
        if dlg.result:
            self._apply_credentials(*dlg.result)

    def _apply_credentials(self, acc, pwd, keep):
        self.session_cred = (acc, pwd)
        if keep:
            eip_session.save_credentials(acc, pwd)
            self._write("[帳號] 已存進 eip帳號.txt")

    def _get_credentials(self):
        if self.session_cred:
            return self.session_cred
        acc, pwd = eip_session.read_credentials()
        if acc and pwd:
            return acc, pwd
        dlg = CredentialDialog(self.master, acc, pwd)
        if not dlg.result:
            return None
        self._apply_credentials(*dlg.result)
        return dlg.result[0], dlg.result[1]

    def _write(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", str(msg) + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    @staticmethod
    def _startfile(path):
        try:
            os.startfile(path)            # Windows
        except Exception:
            pass

    def _open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            self._startfile(self.result_path)

    # ------------------------------ 執行 ------------------------------
    def _start(self):
        if self.worker and self.worker.is_alive():
            return
        ids = self._histnos()
        if not ids:
            messagebox.showwarning("開始收集", "請先在左邊輸入病歷號。")
            return
        labs, vitals = self._labs(), self.var_vitals.get()
        if not labs and not vitals:
            if not messagebox.askyesno(
                    "開始收集", "沒有勾任何檢驗，也不查身高體重。\n"
                                "這樣只會收集術前診斷、手術日期、手術名稱與年齡，確定嗎？"):
                return
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            messagebox.showerror("缺少套件", "缺少 openpyxl，請先執行 install_requirement.py。")
            return
        cred = self._get_credentials()
        if not cred:
            return

        st = self._settings()
        save_settings(st)
        try:                               # 清單存回檔案，下次打開還在、續跑也對得上
            with open(core.DEFAULT_LIST, "w", encoding="utf-8") as f:
                f.write(self.txt_ids.get("1.0", "end").rstrip() + "\n")
        except Exception:
            pass

        out = os.path.join(BASE_DIR, f"胰頭手術術前資料_{datetime.now():%Y%m%d_%H%M}.xlsx")
        args = dict(
            histnos=ids, out_path=out,
            headless=not st["visible"],
            save_dir=os.path.join(BASE_DIR, "除錯HTML") if st["save_html"] else None,
            restart=self.var_restart.get(),
            labs=labs, vitals=vitals,
            acc=cred[0], pwd=cred[1],
            request_interval=st["interval"], patient_delay=st["patient_delay"],
            extra_months=st["extra_months"],
        )

        self.stop_event.clear()
        self.result_path = None
        self.btn_start.configure(state="disabled", text="收集中…")
        self.btn_stop.configure(state="normal")
        self.btn_open.configure(state="disabled")
        self.pbar.configure(value=0, maximum=1)
        self.lbl_status.configure(text="登入中…")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.worker = threading.Thread(target=self._work, args=(args,), daemon=True)
        self.worker.start()

    def _work(self, args):
        try:
            rows, path = core.run(
                log=self.q.put, stop_event=self.stop_event,
                progress=lambda i, n, h: self.q.put(("progress", i, n, h)), **args)
            self.q.put(("done", rows, path))
        except Exception:
            self.q.put("\n========== 執行失敗 ==========")
            self.q.put(traceback.format_exc())
            self.q.put("================================")
            self.q.put(("done", None, None))

    def _stop(self):
        self.stop_event.set()
        self.btn_stop.configure(state="disabled")
        self.lbl_status.configure(text="停止中…（查完目前這位就停）")

    def _drain(self):
        try:
            while True:
                item = self.q.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "progress":
                    _, i, n, h = item
                    self.pbar.configure(maximum=max(n, 1), value=i)
                    if not self.stop_event.is_set():
                        self.lbl_status.configure(
                            text=f"{i + 1}/{n}　{h}" if h else f"{n}/{n} 完成")
                elif isinstance(item, tuple) and item and item[0] == "done":
                    self._done(item[1], item[2])
                else:
                    self._write(item)
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _done(self, rows, path):
        self.btn_start.configure(state="normal", text="開始收集")
        self.btn_stop.configure(state="disabled")
        self.result_path = path
        if path and os.path.exists(path):
            self.btn_open.configure(state="normal")
        if rows is None:
            self.lbl_status.configure(text="執行失敗，請看下方訊息")
            return
        ok = sum(1 for r in rows if r.get("手術日期"))
        pending = sum(1 for r in rows if core.needs_retry(r))
        msg = f"完成：{ok} / {len(rows)} 位找到胰頭手術"
        if pending:
            msg += f"，{pending} 位沒查完（再按一次開始會接著查）"
        self.lbl_status.configure(text=msg)

    def _on_close(self):
        save_settings(self._settings())
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(
                    "關閉", "正在收集中。\n確定要關閉嗎？已查到的會保存，下次打開可以接著查。"):
                return
            self._stop()
            self._wait_and_close()
            return
        self.master.destroy()

    def _wait_and_close(self):
        if self.worker and self.worker.is_alive():
            self.after(300, self._wait_and_close)
        else:
            self.master.destroy()


def main():
    root = tk.Tk()
    root.title("胰頭手術術前資料收集")
    try:
        root.call("tk", "scaling", 1.2)
    except tk.TclError:
        pass
    App(root)
    root.minsize(860, 720)
    root.mainloop()


if __name__ == "__main__":
    main()
