# -*- coding: utf-8 -*-
"""Excel 输出：计算过程 Excel（9 个 Sheet）与 监测点汇总 Excel（3 个 Sheet）。

使用 openpyxl 写入单元格填充色（黄/红标记、风险等级背景色）。

监测点汇总 BI表/ADI表（3.7 节）：
- 仅“风险水平*”一列按数值着色，其余列不着色；
- “地市”列纵向合并连续相同的单元格；
- 所有单元格水平 + 垂直居中。
"""
import os
import re
import shutil
import zipfile

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import config as C
from .parser import district_display_sheet

_INTERNAL_COLS = ('_yellow', '_deleted', '_modified', '_K', '_conv', '_orig', '_in_bi', '_src')
_NUM_COLS = ('监测指标值', 'BI*', 'ADI', '原BI值', '原SSI值', '转换后的SSI值')
_CENTER = Alignment(horizontal='center', vertical='center')


def _cell_font(size=None):
    """中文字体 仿宋_GB2312，英文字体 Times New Roman（字号可选）。
    eastAsia 在 _apply_east_asia_font 中写入 styles.xml。"""
    f = Font(name=C.FONT_EN)
    if size:
        f.size = size
    return f


def _apply_east_asia_font(path):
    """openpyxl 无法直接写 eastAsia 字体，保存后改写 styles.xml：
    将所有 Times New Roman 字体的 <name> 替换为带 eastAsia="仿宋_GB2312" 的 <rFonts>。"""
    tmp = path + '.tmp'
    with zipfile.ZipFile(path, 'r') as zin, \
            zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/styles.xml':
                text = data.decode('utf-8')
                text = text.replace(
                    '<name val="Times New Roman"/>',
                    '<rFonts ascii="Times New Roman" hAnsi="Times New Roman" eastAsia="%s"/>' % C.FONT_CN)
                data = text.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, path)


def _strip_internal(df):
    return df.drop(columns=[c for c in _INTERNAL_COLS if c in df.columns])


def _merge_same_values(ws, col_idx):
    """纵向合并第 col_idx 列中连续相同的单元格（数据区从第 2 行起）"""
    max_row = ws.max_row
    if max_row <= 2:
        return
    start = 2
    prev = ws.cell(row=2, column=col_idx).value
    for r in range(3, max_row + 1):
        cur = ws.cell(row=r, column=col_idx).value
        if cur != prev:
            if r - 1 > start:
                ws.merge_cells(start_row=start, start_column=col_idx,
                               end_row=r - 1, end_column=col_idx)
            start = r
            prev = cur
    if max_row > start:
        ws.merge_cells(start_row=start, start_column=col_idx,
                       end_row=max_row, end_column=col_idx)


def _write_sheet(writer, name, df, flag_col=None, risk_cols=None,
                 merge_cols=None, center=False, header_rename=None, font_size=None):
    """写入一个 Sheet。

    - flag_col：真值列，整行按 FILL_YELLOW/FILL_RED 填充（计算过程表标记用）；
    - risk_cols：按风险等级只给这些列着色（监测点 BI/ADI/整合表，仅风险水平列）；
    - merge_cols：纵向合并这些列中连续相同的单元格；
    - center：所有单元格水平 + 垂直居中；
    - header_rename：{df列名: 新表头}，写表头后改名（用于整合表两个“风险水平*”）；
    - font_size：单元格字号（None 用默认）；所有单元格中文字体 仿宋_GB2312、英文 Times New Roman。
    """
    out = _strip_internal(df)
    out.to_excel(writer, sheet_name=name, index=False)
    wb = writer.book
    ws = wb[name]
    ncols = len(out.columns)

    # 表头改名（先于 header 映射构建）
    if header_rename:
        for df_col, new_name in header_rename.items():
            if df_col in list(out.columns):
                ws.cell(row=1, column=list(out.columns).index(df_col) + 1).value = new_name

    header = {c.value: c.column for c in ws[1]}

    # 数值列格式
    for col_name in _NUM_COLS:
        col_idx = header.get(col_name)
        if col_idx:
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.0'

    # 字体（中文 仿宋_GB2312 / 英文 Times New Roman）+ 居中
    font = _cell_font(font_size)
    for r in range(1, ws.max_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = font
            if center:
                cell.alignment = _CENTER

    # 填充色
    if flag_col:
        for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
            if flag_col in df.columns and bool(row.get(flag_col)):
                color = C.FILL_RED if flag_col.startswith('_deleted') else C.FILL_YELLOW
                fill = PatternFill('solid', fgColor=color)
                for c_idx in range(1, ncols + 1):
                    ws.cell(row=r_idx, column=c_idx).fill = fill
    if risk_cols:
        for df_col in risk_cols:
            if df_col not in list(out.columns):
                continue
            wcol = list(out.columns).index(df_col) + 1
            for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
                risk = row.get(df_col)
                color = C.RISK_FILLS.get(risk)
                if color:
                    ws.cell(row=r_idx, column=wcol).fill = PatternFill('solid', fgColor=color)

    # 纵向合并连续相同单元格
    if merge_cols:
        for col_name in merge_cols:
            col_idx = header.get(col_name)
            if col_idx:
                _merge_same_values(ws, col_idx)

    # 列宽（按内容自适应，上限 60）
    for c_idx in range(1, ncols + 1):
        letter = get_column_letter(c_idx)
        maxlen = len(str(out.columns[c_idx - 1]))
        for v in out.iloc[:, c_idx - 1].astype(str).head(200):
            maxlen = max(maxlen, len(v))
        ws.column_dimensions[letter].width = min(maxlen + 2, 60)


def write_calc_workbook(path, calc_sheets):
    """计算过程 Excel：BI_ADI_计算过程_MM月DD日.xlsx（9 个 Sheet）"""
    with pd_writer(path) as writer:
        for name, (df, flag) in calc_sheets.items():
            _write_sheet(writer, name, df, flag_col=flag)
    _apply_east_asia_font(path)


def write_zongku_processing(path, sheet1, sheet2, sheet3, sheet4):
    """总库表处理 Excel：总库-广州市-提取 / 广州市表-提取 / 广州市-整合 / 整合后的总库"""
    with pd_writer(path) as writer:
        sheet1.to_excel(writer, sheet_name='总库-广州市-提取', index=False)
        sheet2.to_excel(writer, sheet_name='广州市表-提取', index=False)
        sheet3.to_excel(writer, sheet_name='广州市-整合', index=False)
        sheet4.to_excel(writer, sheet_name='整合后的总库', index=False)


def write_monitoring_workbook(path, bi_final, adi_final, deletions):
    """监测点汇总 Excel（村居一览表）：Sheet1 BI表 / Sheet2 ADI表 / Sheet3 BI+ADI整合表 / Sheet4 删除数据情况说明

    字号：整合表五号(10.5)，其余 sheet 14；字体：中文 仿宋_GB2312、英文 Times New Roman。
    """
    bi = _display_frame(bi_final, 'BI*')
    adi = _display_frame(adi_final, 'ADI')
    integrated = _build_integrated_frame(bi, adi)
    with pd_writer(path) as writer:
        _write_sheet(writer, 'BI表', bi, risk_cols=['风险水平*'],
                     merge_cols=['地市'], center=True, font_size=C.SIZE_14)
        _write_sheet(writer, 'ADI表', adi, risk_cols=['风险水平*'],
                     merge_cols=['地市'], center=True, font_size=C.SIZE_14)
        _write_sheet(writer, 'BI+ADI整合表', integrated,
                     risk_cols=['ADI风险', 'BI风险'],
                     merge_cols=['地市'], center=True, font_size=C.SIZE_WUHAO,
                     header_rename={'ADI风险': '风险水平*', 'BI风险': '风险水平*'})
        _write_sheet(writer, '删除数据情况说明', deletions, font_size=C.SIZE_14)
    _apply_east_asia_font(path)


def _build_integrated_frame(bi, adi):
    """BI 表与 ADI 表整合（规则同 Word 一览表）：键 = (地市, 区县, 街道, 监测地点)，
    缺失项写 '/'；列序 = 地市/区县/街道/监测地点/ADI/ADI风险/BI*/BI风险。"""
    bi_m = bi.rename(columns={'风险水平*': 'BI风险'})
    adi_m = adi.rename(columns={'风险水平*': 'ADI风险'})
    m = bi_m.merge(adi_m, on=['地市', '区县', '街道', '监测地点'], how='outer')
    for col in ('ADI', 'ADI风险', 'BI*', 'BI风险'):
        if col not in m.columns:
            m[col] = '/'
        m[col] = m[col].where(m[col].notna(), '/')
    m['_city_idx'] = m['地市'].map(C.CITY_INDEX).fillna(99).astype(int)
    m = m.sort_values(['_city_idx', '区县', '街道', '监测地点'], kind='stable').reset_index(drop=True)
    return m[['地市', '区县', '街道', '监测地点', 'ADI', 'ADI风险', 'BI*', 'BI风险']]


def _display_frame(final, value_col):
    """最终表 -> 村居一览表显示口径（区县去后缀、市辖区->/），按 3.7 排序：
    地市固定顺序 → 区县升序 → 街道升序 → 监测地点升序"""
    df = final.copy()
    df['区县'] = df['区县'].map(district_display_sheet)
    df = df.sort_values(
        ['_city_idx', '区县', '街道', '监测地点'],
        kind='stable',
    ).reset_index(drop=True)
    return df[['地市', '区县', '街道', '监测地点', value_col, '风险水平*']]


def pd_writer(path):
    """兼容 pandas 2.x 的 ExcelWriter 上下文（openpyxl 引擎）"""
    import pandas as pd
    return pd.ExcelWriter(path, engine='openpyxl')
