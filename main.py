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
        root.geometry('780x660')
        self.q = queue.Queue()
        self.running = False

        pad = {'padx': 8, 'pady': 4}
        frm = ttk.Frame(root, padding=10)
        frm.pack(fill='both', expand=True)

        # 总库表文件（唯一输入）
        ttk.Label(frm, text='输入总库表文件：').grid(row=0, column=0, sticky='w', **pad)
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

        # 排除字段（支持多个；“添加排除字段”按钮逐个添加输入框，框间可选“和/或”连接）
        ttk.Label(frm, text='排除字段（可选）：').grid(row=2, column=0, sticky='nw', **pad)
        self.excl_frame = ttk.Frame(frm)
        self.excl_frame.grid(row=2, column=1, columnspan=2, sticky='we', **pad)
        self.excl_rows = []
        bar = ttk.Frame(self.excl_frame)
        bar.pack(side='bottom', fill='x', pady=(3, 0))
        ttk.Button(bar, text='添加排除字段', command=self._add_exclude_row).pack(side='left')
        ttk.Label(bar, text='多个字段间可选“和/或”：和=须同时包含前后两个字段才删除，或=包含其一即删除'
                            '（匹配“地市-区/县/市-街道/乡/镇”与“社区/村居”两列，任一列命中即删除）',
                  foreground='gray').pack(side='left', padx=(8, 0))
        self._add_exclude_row(first=True)

        # 飞行监测表（可选，需求一）
        ttk.Label(frm, text='飞行监测表（可选）：').grid(row=4, column=0, sticky='w', **pad)
        self.entry_flight = ttk.Entry(frm, width=62)
        self.entry_flight.grid(row=4, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self.choose_flight).grid(row=4, column=2, **pad)

        # 输出目录
        ttk.Label(frm, text='输出目录：').grid(row=5, column=0, sticky='w', **pad)
        self.entry_out = ttk.Entry(frm, width=62)
        self.entry_out.grid(row=5, column=1, sticky='we', **pad)
        ttk.Button(frm, text='选择…', command=self.choose_dir).grid(row=5, column=2, **pad)

        # 开始按钮
        self.btn = ttk.Button(frm, text='开始处理', command=self.on_start)
        self.btn.grid(row=6, column=1, sticky='w', **pad)

        # 日志区
        ttk.Label(frm, text='处理日志：').grid(row=7, column=0, sticky='nw', **pad)
        self.txt = tk.Text(frm, height=18, state='disabled', font=('Consolas', 9))
        self.txt.grid(row=8, column=0, columnspan=3, sticky='nsew', **pad)
        frm.rowconfigure(8, weight=1)
        frm.columnconfigure(1, weight=1)

        self.root.after(100, self._poll)

    # ---------------- 排除字段行管理（需求三） ----------------

    def _add_exclude_row(self, first=False):
        """按一下“添加排除字段”即添加一个排除字段输入框；除第一个外，
        输入框前带“和/或”下拉（连接前一个字段）。"""
        row = ttk.Frame(self.excl_frame)
        row.pack(fill='x', pady=1)
        connector = None
        if not first:
            connector = ttk.Combobox(row, values=('和', '或'), width=3, state='readonly')
            connector.set('或')
            connector.pack(side='left', padx=(0, 4))
        entry = ttk.Entry(row, width=50)
        entry.pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='删除', width=4,
                   command=lambda r=row: self._remove_exclude_row(r)).pack(side='left', padx=(4, 0))
        self.excl_rows.append({'frame': row, 'entry': entry, 'connector': connector})

    def _remove_exclude_row(self, frame):
        idx = next((i for i, r in enumerate(self.excl_rows) if r['frame'] is frame), None)
        if idx is None:
            return
        frame.destroy()
        del self.excl_rows[idx]
        # 若删掉的是第一行，新的第一行不再需要“和/或”连接词
        if self.excl_rows and self.excl_rows[0]['connector'] is not None:
            self.excl_rows[0]['connector'].destroy()
            self.excl_rows[0]['connector'] = None

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

    def choose_flight(self):
        p = filedialog.askopenfilename(
            title='选择飞行监测表Excel文件',
            filetypes=[('Excel 文件', '*.xlsx *.xls'), ('所有文件', '*.*')])
        if p:
            self.entry_flight.delete(0, 'end')
            self.entry_flight.insert(0, p)

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
                    messagebox.showinfo('完成', '已生成 %d 份文件：\n\n%s'
                                        % (len(payload), '\n'.join(payload)))
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
        ex_terms = []
        for r in self.excl_rows:
            f = r['entry'].get().strip()
            if f:
                conn = r['connector'].get() if r['connector'] is not None else None
                ex_terms.append({'field': f, 'connector': conn})
        exclude = ex_terms if ex_terms else None
        fp = self.entry_flight.get().strip() or None
        if fp and not os.path.isfile(fp):
            messagebox.showerror('输入错误', '飞行监测表文件不存在。')
            return
        if fp and not fp.lower().endswith(('.xlsx', '.xls')):
            messagebox.showerror('输入错误', '飞行监测表必须为 Excel（.xlsx/.xls）。')
            return

        self.running = True
        self.btn.configure(state='disabled')
        self.log('开始处理…')
        threading.Thread(target=self.worker, args=(f, out, target, exclude, fp), daemon=True).start()

    def worker(self, f, out, target, exclude, flight_path):
        def log(msg):
            self.q.put(('log', msg))
        try:
            paths = process_file(f, out, target.year, target.month, target.day,
                                 exclude, flight_path, log=log)
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
