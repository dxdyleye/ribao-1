# -*- coding: utf-8 -*-
"""处理编排：读 Excel -> 流水线 -> 生成 4 份输出文件"""
import os
from datetime import date

import pandas as pd

from . import config as C
from . import excel_output, word_output
from .pipeline import ProcessingError, run_pipeline


def _file_attr_note(path):
    """Windows 下读取文件属性，报告加密/重解析（云占位）等标志，用于排查铠大师/EFS/网盘加密。"""
    try:
        st = os.stat(path)
        attrs = getattr(st, 'st_file_attributes', 0) or 0
        flags = []
        if attrs & 0x4000:
            flags.append('加密属性 ENCRYPTED（可能被文档加密软件/EFS 加密）')
        if attrs & 0x400:
            flags.append('重解析点 REPARSE_POINT（云占位/加密虚拟文件）')
        if attrs & 0x2:
            flags.append('隐藏')
        if flags:
            return '文件属性检测：%s' % '；'.join(flags)
    except OSError:
        pass
    return None


def _check_input_file(input_path, label):
    """读取前校验输入文件：区分“路径不存在”与“云占位/加密文件不可读”，并列出同目录文件辅助排查。"""
    if not input_path:
        raise ProcessingError('未选择%s文件。' % label)
    if os.path.isfile(input_path):
        return
    msg = ['%s文件无法读取：\n%s' % (label, input_path)]
    if not os.path.exists(input_path):
        msg.append('原因：该路径下找不到文件（可能已被移动/改名，或路径被改动，'
                   '或被文档加密软件/安全软件拦截了程序对该文件的访问）。')
    else:
        msg.append('原因：文件存在但无法作为普通文件读取（可能是网盘/OneDrive 在线占位文件，'
                   '或被文档加密软件如铠大师加密、仅授权软件可读）。\n'
                   '请先在资源管理器中右键该文件 →“始终保留在此设备上 / 下载”，或把文件解密/另存为明文后再选择。')
    attr = _file_attr_note(input_path)
    if attr:
        msg.append(attr)
    parent = os.path.dirname(input_path)
    try:
        names = sorted(os.listdir(parent))
        if names:
            shown = names[:15]
            msg.append('目录“%s”下的文件（前 %d 个，共 %d 个）：\n%s%s'
                       % (parent, min(15, len(names)), len(names),
                          '\n'.join('  · ' + n for n in shown),
                          '\n  …' if len(names) > 15 else ''))
        else:
            msg.append('目录“%s”为空。' % parent)
    except OSError:
        msg.append('无法列出目录“%s”（目录可能不存在或无权访问）。' % parent)
    raise ProcessingError('\n'.join(msg))


def process_file(input_path, output_dir, year, month, day, exclude=None, flight_path=None, log=None):
    """返回生成的 3 个文件完整路径列表（总库表为唯一输入；飞行监测表可选）。"""
    def logmsg(s):
        if log:
            log(s)

    _check_input_file(input_path, '输入总库表')
    logmsg('正在读取总库表文件…')
    try:
        source = pd.read_excel(input_path)
    except OSError as e:
        note = _file_attr_note(input_path)
        raise ProcessingError('读取总库表文件失败（%s）：\n%s%s'
                              % (e, input_path, ('\n' + note) if note else ''))
    target = date(year, month, day)

    flight_df = None
    if flight_path:
        _check_input_file(flight_path, '飞行监测表')
        logmsg('正在读取飞行监测表文件…')
        try:
            flight_df = pd.read_excel(flight_path)
        except OSError as e:
            note = _file_attr_note(flight_path)
            raise ProcessingError('读取飞行监测表文件失败（%s）：\n%s%s'
                                  % (e, flight_path, ('\n' + note) if note else ''))

    logmsg('正在预处理数据（日期筛选/空值/排除字段/距末例天数/防控区类型）…')
    res = run_pipeline(source, target, exclude, flight_df=flight_df)
    logmsg('基础数据集 %d 条；最终BI表 %d 条；最终ADI表 %d 条'
           % (len(res.base), len(res.bi_final), len(res.adi_final)))
    if res.flight_bi_count or res.flight_adi_count:
        logmsg('飞行监测表：整合入BI %d 条、整合入ADI %d 条'
               % (res.flight_bi_count, res.flight_adi_count))

    # 在目标目录下建立输出文件夹：省蚊媒监测日报（输入的日期），输出文件均放入其中
    os.makedirs(output_dir, exist_ok=True)
    out_dir = os.path.join(output_dir, '省蚊媒监测日报（%d月%d日）' % (month, day))
    os.makedirs(out_dir, exist_ok=True)
    logmsg('输出文件夹：%s' % out_dir)
    paths = []

    # 1 计算过程 Excel（9 个 Sheet）
    logmsg('正在生成计算过程Excel…')
    p = os.path.join(out_dir, C.calc_xlsx_name(year, month, day))
    excel_output.write_calc_workbook(p, res.calc_sheets)
    paths.append(p)

    # 2 日报 Word（叙述版）
    logmsg('正在生成日报Word…')
    bi_sec = word_output.build_section(res.bi_final, res.exclude_display, res.excluded_cities, 'BI')
    adi_sec = word_output.build_section(res.adi_final, res.exclude_display, res.excluded_cities, 'ADI')
    p = os.path.join(out_dir, C.daily_docx_name(year, month, day))
    word_output.write_daily_report(p, target, bi_sec, adi_sec)
    paths.append(p)

    # 3 监测点汇总 Excel（村居一览表）
    logmsg('正在生成村居一览表Excel…')
    p = os.path.join(out_dir, C.summary_xlsx_name(year, month, day))
    excel_output.write_monitoring_workbook(p, res.bi_final, res.adi_final, res.deletions)
    paths.append(p)

    # 4 基础数据集 Excel（BI/ADI基础数据集 + 乱码处理）
    logmsg('正在生成基础数据集Excel…')
    p = os.path.join(out_dir, C.base_xlsx_name(year, month, day))
    excel_output.write_base_workbook(p, res.base_bi, res.base_adi, res.messy_bi, res.messy_adi)
    paths.append(p)

    logmsg('全部完成。')
    return paths
