# -*- coding: utf-8 -*-
"""处理编排：读 Excel -> 流水线 -> 生成 4 份输出文件"""
import os
from datetime import date

import pandas as pd

from . import config as C
from . import excel_output, word_output
from .pipeline import run_pipeline


def process_file(input_path, output_dir, year, month, day, exclude=None, log=None):
    """返回生成的 3 个文件完整路径列表（总库表为唯一输入）。"""
    def logmsg(s):
        if log:
            log(s)

    logmsg('正在读取总库表文件…')
    source = pd.read_excel(input_path)
    target = date(year, month, day)

    logmsg('正在预处理数据（日期筛选/空值/排除字段/距末例天数/防控区类型）…')
    res = run_pipeline(source, target, exclude)
    logmsg('基础数据集 %d 条；最终BI表 %d 条；最终ADI表 %d 条'
           % (len(res.base), len(res.bi_final), len(res.adi_final)))

    os.makedirs(output_dir, exist_ok=True)
    paths = []

    # 1 计算过程 Excel（9 个 Sheet）
    logmsg('正在生成计算过程Excel…')
    p = os.path.join(output_dir, C.calc_xlsx_name(year, month, day))
    excel_output.write_calc_workbook(p, res.calc_sheets)
    paths.append(p)

    # 2 日报 Word（叙述版）
    logmsg('正在生成日报Word…')
    bi_sec = word_output.build_section(res.bi_final, exclude, res.excluded_cities, 'BI')
    adi_sec = word_output.build_section(res.adi_final, exclude, res.excluded_cities, 'ADI')
    p = os.path.join(output_dir, C.daily_docx_name(year, month, day))
    word_output.write_daily_report(p, target, bi_sec, adi_sec)
    paths.append(p)

    # 3 监测点汇总 Excel（村居一览表）
    logmsg('正在生成村居一览表Excel…')
    p = os.path.join(output_dir, C.summary_xlsx_name(year, month, day))
    excel_output.write_monitoring_workbook(p, res.bi_final, res.adi_final, res.deletions)
    paths.append(p)

    logmsg('全部完成。')
    return paths
