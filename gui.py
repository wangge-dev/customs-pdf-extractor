"""本地报关单批量提取程序的 Windows 图形界面。"""

from __future__ import annotations

import os
import queue
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import backend


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("报关单 PDF 批量提取")
        self.geometry("760x535")
        self.minsize(650, 460)
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status = tk.StringVar(value="请选择 PDF 文件夹和结果保存位置。")
        self.events = queue.Queue()
        self.running = False
        self.last_result_dir = None
        self.last_result = {}
        self._make_widgets()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.after(100, self._drain_events)

    def _make_widgets(self):
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(5, weight=1)

        ttk.Label(outer, text="报关单 PDF 批量提取", font=("Microsoft YaHei UI", 16, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 14)
        )
        ttk.Label(outer, text="PDF 文件夹").grid(row=1, column=0, sticky="w", pady=5)
        self.input_entry = ttk.Entry(outer, textvariable=self.input_path)
        self.input_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=5)
        self.input_button = ttk.Button(outer, text="选择文件夹", command=self._choose_input)
        self.input_button.grid(row=1, column=2, pady=5)
        ttk.Label(outer, text="结果保存到").grid(row=2, column=0, sticky="w", pady=5)
        self.output_entry = ttk.Entry(outer, textvariable=self.output_path)
        self.output_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=5)
        self.output_button = ttk.Button(outer, text="选择文件夹", command=self._choose_output)
        self.output_button.grid(row=2, column=2, pady=5)

        buttons = ttk.Frame(outer)
        buttons.grid(row=3, column=0, columnspan=3, sticky="w", pady=(12, 8))
        self.start_button = ttk.Button(buttons, text="开始提取", command=self._start)
        self.start_button.pack(side="left", padx=(0, 10))
        self.open_button = ttk.Button(buttons, text="打开结果文件夹", command=self._open_output)
        self.open_button.pack(side="left", padx=(0, 10))
        self.report_button = ttk.Button(buttons, text="打开核对报告", command=self._open_report, state="disabled")
        self.report_button.pack(side="left", padx=(0, 10))
        self.images_button = ttk.Button(buttons, text="打开原图核对", command=self._open_images, state="disabled")
        self.images_button.pack(side="left")
        ttk.Label(outer, textvariable=self.status, wraplength=700).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(2, 6)
        )

        log_frame = ttk.Frame(outer)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=12, state="disabled", wrap="word", font=("Microsoft YaHei UI", 9))
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)
        ttk.Label(outer, text="适用于同类有文字层报关单；扫描件或版式异常须核对原图。", foreground="#555555").grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )

    def _choose_input(self):
        selected = filedialog.askdirectory(title="选择存放 PDF 的文件夹", parent=self)
        if selected:
            self.input_path.set(selected)

    def _choose_output(self):
        selected = filedialog.askdirectory(title="选择结果保存文件夹", parent=self, mustexist=True)
        if selected:
            self.output_path.set(selected)

    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", str(text).rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy):
        self.running = busy
        for widget in (self.input_entry, self.output_entry, self.input_button, self.output_button, self.start_button):
            widget.configure(state="disabled" if busy else "normal")

    def _start(self):
        input_dir = self.input_path.get().strip().strip('"')
        output_dir = self.output_path.get().strip().strip('"')
        if not input_dir or not Path(input_dir).is_dir():
            messagebox.showwarning("请选择文件夹", "请选择存在的 PDF 文件夹。", parent=self)
            return
        if not output_dir or not Path(output_dir).is_dir():
            messagebox.showwarning("请选择文件夹", "请选择存在的结果保存文件夹。", parent=self)
            return
        self._set_busy(True)
        self.last_result_dir = None
        self.last_result = {}
        self._update_result_buttons()
        self.status.set("正在处理，请保持窗口打开。")
        self._append_log("开始处理：" + input_dir)
        self._append_log("结果保存到：" + output_dir)

        def run():
            try:
                result = backend.run_batch(input_dir, output_dir, lambda message: self.events.put(("progress", str(message))))
                self.events.put(("done", result))
            except Exception as exc:
                self.events.put(("error", (str(exc) or type(exc).__name__, traceback.format_exc())))

        threading.Thread(target=run, daemon=True).start()

    def _drain_events(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "progress":
                    self.status.set(value)
                    self._append_log(value)
                elif kind == "done":
                    self._set_busy(False)
                    self._append_log("处理结束，请查看结果和核对报告。")
                    if isinstance(value, dict):
                        self.last_result = value
                        self.last_result_dir = value.get('结果文件夹')
                        for key, item in value.items():
                            if isinstance(item, (str, int, float)):
                                self._append_log(f"{key}：{item}")
                    self._update_result_buttons()
                    summary = self._result_summary(self.last_result)
                    self.status.set(summary)
                    messagebox.showinfo("处理结束", summary + "\n请打开核对报告，查看待核对项目。", parent=self)
                elif kind == "error":
                    self._set_busy(False)
                    self.status.set("处理失败，请查看下方信息。")
                    self._append_log("处理失败：" + value[0])
                    self._append_log(value[1])
                    messagebox.showerror("处理失败", value[0], parent=self)
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    @staticmethod
    def _result_summary(result):
        keys = ("PDF总数", "商品明细行数", "待核对文件数", "OCR未完成文件数")
        if not all(key in result for key in keys):
            return "处理结束，请查看结果和核对报告。"
        return (f"处理结束：{result['PDF总数']} 份 PDF，{result['商品明细行数']} 条商品明细；"
                f"{result['待核对文件数']} 份待核对，{result['OCR未完成文件数']} 份 OCR 未完成。")

    def _update_result_buttons(self):
        for widget, key in ((self.report_button, '核对报告'), (self.images_button, '原图核对页')):
            widget.configure(state="normal" if self.last_result.get(key) else "disabled")

    def _open_result_file(self, key):
        path = self.last_result.get(key)
        if not path or not Path(path).is_file():
            messagebox.showwarning("文件不存在", "找不到本次结果文件，请检查是否已移动或删除。", parent=self)
            return
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror("无法打开文件", str(exc), parent=self)

    def _open_report(self):
        self._open_result_file('核对报告')

    def _open_images(self):
        self._open_result_file('原图核对页')

    def _open_output(self):
        path = self.last_result_dir or self.output_path.get().strip().strip('"')
        if not path or not Path(path).is_dir():
            messagebox.showwarning("文件夹不存在", "请先选择存在的结果保存文件夹。", parent=self)
            return
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror("无法打开文件夹", str(exc), parent=self)

    def _close(self):
        if self.running:
            messagebox.showinfo("正在处理", "程序正在处理，请等待处理结束后再关闭。", parent=self)
            return
        self.destroy()


if __name__ == "__main__":
    Application().mainloop()
