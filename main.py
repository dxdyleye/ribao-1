# -*- coding: utf-8 -*-
"""蚊媒监测数据处理工具 —— 图形界面（Tkinter）。

控件：输入文件选择 / 年月日下拉 / 排除字段 / 输出目录 / 开始处理 / 日志区。
后台线程执行，界面不卡顿；错误弹窗提示后停止本次处理，不退出程序。
"""
import calendar
import os
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk

from mosquito.pipeline import ProcessingError
from mosquito.runner import process_file


class App(object):
    def __init__(self, root):
        self.root = root
        root.title('蚊媒监测数据处理工具')
        root.geometry('760x600')
        self.q = queue.Queue()
        self.running = False

        pad = {'padx': 8, 'pady': 4}
        frm = ttk.Frame(root, padding=10)
        frm.pack(fill='both', expand=True)

        # 输入文件
        ttk.Label(frm, text='输入Excel文件：').grid(row=0, column=0, sticky='w', **pad)
        self.entry_file = ttk.Entry(frm, width=62)
        self.entry_file.grid(row=0, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self.choose_file).grid(row=0, column=2, **pad)

        # 目标日期
        ttk.Label(frm, text='目标日期：').grid(row=1, column=0, sticky='w', **pad)
        d = tk.Frame(frm)
        d.grid(row=1, column=1, sticky='w', **pad)
        now = date.today()
        self.var_year = tk.StringVar(value=str(now.year))
        self.var_month = tk.StringVar(value=str(now.month))
        self.var_day = tk.StringVar(value=str(now.day))
        ttk.Spinbox(d, from_=2000, to=2100, textvariable=self.var_year, width=6).pack(side='left')
        ttk.Label(d, text='年').pack(side='left')
        ttk.Spinbox(d, from_=1, to=12, textvariable=self.var_month, width=4).pack(side='left')
        ttk.Label(d, text='月').pack(side='left')
        ttk.Spinbox(d, from_=1, to=31, textvariable=self.var_day, width=4).pack(side='left')
        ttk.Label(d, text='日').pack(side='left')

        # 排除字段
        ttk.Label(frm, text='排除字段（可选）：').grid(row=2, column=0, sticky='w', **pad)
        self.entry_excl = ttk.Entry(frm, width=62)
        self.entry_excl.grid(row=2, column=1, sticky='we', **pad)
        ttk.Label(frm, text='如：荔湾区（删除地市-区/县/市-街道/乡/镇中包含该字段的记录）',
                  foreground='gray').grid(row=3, column=1, sticky='w', padx=8)

        # 输出目录
        ttk.Label(frm, text='输出目录：').grid(row=4, column=0, sticky='w', **pad)
        self.entry_out = ttk.Entry(frm, width=62)
        self.entry_out.grid(row=4, column=1, sticky='we', **pad)
        ttk.Button(frm, text='选择…', command=self.choose_dir).grid(row=4, column=2, **pad)

        # 开始按钮
        self.btn = ttk.Button(frm, text='开始处理', command=self.on_start)
        self.btn.grid(row=5, column=1, sticky='w', **pad)

        # 日志区
        ttk.Label(frm, text='处理日志：').grid(row=6, column=0, sticky='nw', **pad)
        self.txt = tk.Text(frm, height=18, state='disabled', font=('Consolas', 9))
        self.txt.grid(row=7, column=0, columnspan=3, sticky='nsew', **pad)
        frm.rowconfigure(7, weight=1)
        frm.columnconfigure(1, weight=1)

        self.root.after(100, self._poll)

    # ---------------- 界面动作 ----------------

    def log(self, msg):
        self.txt.configure(state='normal')
        self.txt.insert('end', msg + '\n')
        self.txt.see('end')
        self.txt.configure(state='disabled')

    def choose_file(self):
        p = filedialog.askopenfilename(
            title='选择输入Excel文件',
            filetypes=[('Excel 文件', '*.xlsx *.xls'), ('所有文件', '*.*')])
        if p:
            self.entry_file.delete(0, 'end')
            self.entry_file.insert(0, p)
            if not self.entry_out.get().strip():
                self.entry_out.delete(0, 'end')
                self.entry_out.insert(0, os.path.dirname(p))

    def choose_dir(self):
        p = filedialog.askdirectory(title='选择输出目录')
        if p:
            self.entry_out.delete(0, 'end')
            self.entry_out.insert(0, p)

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'log':
                    self.log(payload)
                elif kind == 'done':
                    self.log('处理完成。')
                    messagebox.showinfo('完成', '已生成 4 份文件：\n\n' + '\n'.join(payload))
                    if payload:
                        self._open_folder(os.path.dirname(payload[0]))
                    self.running = False
                    self.btn.configure(state='normal')
                elif kind == 'error':
                    self.log('错误：' + payload)
                    messagebox.showerror('处理失败', payload)
                    self.running = False
                    self.btn.configure(state='normal')
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    @staticmethod
    def _open_folder(path):
        try:
            if os.path.isdir(path):
                if sys.platform.startswith('win'):
                    os.startfile(path)
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', path])
                else:
                    subprocess.Popen(['xdg-open', path])
        except Exception:
            pass

    def on_start(self):
        if self.running:
            return
        f = self.entry_file.get().strip()
        out = self.entry_out.get().strip()
        if not f or not os.path.isfile(f):
            messagebox.showerror('输入错误', '请选择有效的输入Excel文件。')
            return
        if not f.lower().endswith(('.xlsx', '.xls')):
            messagebox.showerror('输入错误', '输入文件必须为 Excel（.xlsx/.xls）。')
            return
        if not out:
            out = os.path.dirname(f)
            self.entry_out.delete(0, 'end')
            self.entry_out.insert(0, out)
        try:
            y = int(self.var_year.get())
            m = int(self.var_month.get())
            d = int(self.var_day.get())
            if not 1 <= m <= 12:
                raise ValueError('月份须在 1-12 之间')
            last = calendar.monthrange(y, m)[1]
            if not 1 <= d <= last:
                raise ValueError('%d月只有 %d 天' % (m, last))
            target = date(y, m, d)
        except (ValueError, TypeError) as e:
            messagebox.showerror('日期错误', '日期不合法：%s' % e)
            return
        ex = self.entry_excl.get().strip() or None

        self.running = True
        self.btn.configure(state='disabled')
        self.log('开始处理…')
        threading.Thread(target=self.worker, args=(f, out, target, ex), daemon=True).start()

    def worker(self, f, out, target, ex):
        def log(msg):
            self.q.put(('log', msg))
        try:
            paths = process_file(f, out, target.year, target.month, target.day, ex, log=log)
            self.q.put(('done', paths))
        except ProcessingError as e:
            self.q.put(('error', '处理中止：%s' % e))
        except Exception:
            self.q.put(('error', '发生未预期错误：\n%s' % traceback.format_exc()))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
