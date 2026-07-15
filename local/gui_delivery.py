# -*- coding: utf-8 -*-
"""
桌面 GUI 版送货单识别（Tkinter，零额外依赖，兼容 Win7 / Python 3.9）
======================================================================
纯本地窗口程序：下拉选模型 → 窗口内填 Key → 选图 → 开始识别 → 表格预览 → 导出 Excel。
识别 / Excel / 解析逻辑全部复用同目录的 ocr_delivery.py（不重复造轮子）。

运行（开发期）:
  python gui_delivery.py
打包 exe（Win7 需在 Python 3.9 下，见 打包说明.md）:
  pyinstaller --onefile --noconsole gui_delivery.py
"""
import os
import sys
import time
import threading
from pathlib import Path

# 确保能 import 同目录的 ocr_delivery
HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import ocr_delivery as core

IMG_EXTS = core.IMG_EXTS


class App:
    def __init__(self, root):
        self.root = root
        self.images = []
        self.workbook = None
        self.running = False
        self.t0 = 0

        root.title("送货单 → Excel 识别（本地版）")
        root.geometry("900x640")

        self._build_controls()
        self._build_log()
        self._build_table()
        self._build_export()

        # 初始化模型下拉与 Key
        self.prov_var.set("zhipu")
        self.on_provider_change()

    # ---------------- UI 构建 ----------------
    def _build_controls(self):
        f = ttk.Frame(self.root, padding=10)
        f.pack(fill="x")

        ttk.Label(f, text="模型预设：").grid(row=0, column=0, sticky="w")
        self.prov_var = tk.StringVar()
        names = {k: core.MODEL_PRESETS[k]["name"] for k in core.MODEL_PRESETS}
        self.prov_combo = ttk.Combobox(f, textvariable=self.prov_var,
                                       values=list(names.keys()),
                                       state="readonly", width=34)
        self.prov_combo.grid(row=0, column=1, sticky="w", padx=(4, 0))
        self.prov_combo.bind("<<ComboboxSelected>>", self.on_provider_change)

        self.key_label_var = tk.StringVar(value="API Key")
        ttk.Label(f, textvariable=self.key_label_var).grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.key_var = tk.StringVar()
        self.key_entry = ttk.Entry(f, textvariable=self.key_var, show="*", width=48)
        self.key_entry.grid(row=1, column=1, sticky="w", padx=(4, 0), pady=(6, 0))
        ttk.Button(f, text="存Key到.env", command=self.save_key).grid(row=1, column=2, padx=(6, 0), pady=(6, 0))

        # 选图按钮行
        bf = ttk.Frame(self.root, padding=(10, 0, 10, 6))
        bf.pack(fill="x")
        ttk.Button(bf, text="选择图片", command=self.choose_files).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="选择文件夹", command=self.choose_folder).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="清空选择", command=self.clear_images).pack(side="left", padx=(0, 6))
        self.img_label = ttk.Label(bf, text="尚未选择图片")
        self.img_label.pack(side="left", padx=(10, 0))

        ttk.Button(bf, text="▶ 开始识别", command=self.start).pack(side="right", padx=(6, 0))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bf, textvariable=self.status_var).pack(side="right", padx=(6, 0))

    def _build_log(self):
        lf = ttk.LabelFrame(self.root, text="运行日志", padding=6)
        lf.pack(fill="x", padx=10, pady=(0, 6))
        self.log = scrolledtext.ScrolledText(lf, height=7, font=("Consolas", 10))
        self.log.pack(fill="both", expand=True)

    def _build_table(self):
        tf = ttk.LabelFrame(self.root, text="识别结果预览", padding=6)
        tf.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.tree = ttk.Treeview(tf, show="headings")
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tf.grid_rowconfigure(0, weight=1)
        tf.grid_columnconfigure(0, weight=1)

    def _build_export(self):
        ef = ttk.Frame(self.root, padding=(10, 0, 10, 8))
        ef.pack(fill="x")
        self.export_btn = ttk.Button(ef, text="💾 导出 Excel", command=self.export, state="disabled")
        self.export_btn.pack(side="right")

    # ---------------- 交互 ----------------
    def on_provider_change(self, *a):
        p = self.prov_var.get()
        cfg = core.MODEL_PRESETS.get(p)
        if not cfg:
            return
        self.key_label_var.set("API Key（%s）" % cfg["key_env"])
        self.key_var.set(os.getenv(cfg["key_env"], ""))  # 从 .env 预填

    def save_key(self):
        p = self.prov_var.get()
        key_env = core.MODEL_PRESETS[p]["key_env"]
        v = self.key_var.get().strip()
        if not v:
            messagebox.showwarning("提示", "请先填写 Key 再保存")
            return
        self._write_env(key_env, v)
        messagebox.showinfo("已保存", "%s 已写入同目录 .env" % key_env)

    def _write_env(self, key_env, value):
        env_path = core.BASE / ".env"
        lines = env_path.read_text(encoding="utf-8").splitlines() if env_path.exists() else []
        found = False
        for i, ln in enumerate(lines):
            if ln.strip().startswith(key_env + "="):
                lines[i] = "%s=%s" % (key_env, value)
                found = True
                break
        if not found:
            lines.append("%s=%s" % (key_env, value))
        env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        os.environ[key_env] = value

    def choose_files(self):
        fs = filedialog.askopenfilenames(
            title="选择送货单图片",
            filetypes=[("图片", "*.jpg *.jpeg *.png *.bmp *.gif *.webp")])
        self._add_images(list(fs))

    def choose_folder(self):
        d = filedialog.askdirectory(title="选择包含送货单图片的文件夹")
        if not d:
            return
        fs = [os.path.join(d, fn) for fn in sorted(os.listdir(d))
              if Path(fn).suffix.lower() in IMG_EXTS]
        self._add_images(fs)

    def _add_images(self, fs):
        for f in fs:
            if f not in self.images:
                self.images.append(f)
        self._update_img_label()

    def clear_images(self):
        self.images = []
        self._update_img_label()

    def _update_img_label(self):
        n = len(self.images)
        if n == 0:
            self.img_label.config(text="尚未选择图片")
        else:
            self.img_label.config(text="已选 %d 张" % n)

    # ---------------- 识别流程（线程执行，避免界面卡死）----------------
    def start(self):
        if self.running:
            return
        provider = self.prov_var.get()
        cfg = core.MODEL_PRESETS.get(provider)
        if not cfg:
            return
        api_key = self.key_var.get().strip() or os.getenv(cfg["key_env"], "")
        if not api_key:
            messagebox.showerror("缺少 Key", "请填写 %s 后再开始" % cfg["key_env"])
            return
        if not self.images:
            messagebox.showerror("未选图片", "请先选择送货单图片")
            return

        self.running = True
        self.t0 = time.time()
        self.workbook = None
        self.export_btn.config(state="disabled")
        self._clear_tree()
        self.log.delete("1.0", "end")
        self._log("开始识别（模型：%s）…" % cfg["name"])

        threading.Thread(target=self._worker, args=(provider, api_key), daemon=True).start()
        self._tick()

    def _worker(self, provider, api_key):
        try:
            cfg = core.MODEL_PRESETS[provider]
            n = len(self.images)
            max_dim = 1200 if n > 2 else 2000
            quality = 70 if n > 2 else 85
            self._log("压缩 %d 张图片（max_dim=%d, quality=%d）…" % (n, max_dim, quality))
            b64 = [core.compress_and_encode(p, max_dim, quality) for p in self.images]
            self._log("调用模型：%s …" % cfg["name"])
            raw = core.call_model(cfg, api_key, b64)
            self._log("模型返回（前 200 字）：\n" + raw[:200])
            rows = core.extract_json(raw)
            if not isinstance(rows, list):
                raise ValueError("返回不是数组")
            wb = core.build_workbook(rows)
            elapsed = time.time() - self.t0
            self.root.after(0, self._done, rows, wb, len(rows), elapsed)
        except Exception as e:
            self.root.after(0, self._error, str(e))

    def _tick(self):
        if not self.running:
            return
        self.status_var.set("识别中… 耗时 %.1fs" % (time.time() - self.t0))
        self.root.after(200, self._tick)

    def _done(self, rows, wb, count, elapsed):
        self.running = False
        self.workbook = wb
        self._show_table(rows)
        self.export_btn.config(state="normal")
        self.status_var.set("完成：%d 行，耗时 %.1fs" % (count, elapsed))
        self._log("完成：识别 %d 行，耗时 %.1fs" % (count, elapsed))

    def _error(self, msg):
        self.running = False
        self.status_var.set("失败")
        self._log("出错：" + msg)

    # ---------------- 表格 / 日志 ----------------
    def _show_table(self, rows):
        self._clear_tree()
        self.tree["columns"] = core.EXCEL_HEADERS
        for h in core.EXCEL_HEADERS:
            self.tree.heading(h, text=h)
            self.tree.column(h, width=100, anchor="w", minwidth=60)
        for r in rows:
            self.tree.insert("", "end", values=[r.get(h, "") for h in core.EXCEL_HEADERS])

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _log(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def export(self):
        if self.workbook is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="送货单_" + time.strftime("%Y-%m-%d") + ".xlsx")
        if path:
            self.workbook.save(path)
            messagebox.showinfo("已保存", path)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
