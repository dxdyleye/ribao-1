# -*- coding: utf-8 -*-
"""工作簿整合工具：参照需求文档 §3.7「村居一览表 Excel → Sheet3 BI+ADI整合表」的处理规则，
将 Excel 工作簿的 Sheet1（BI 表）与 Sheet2（ADI 表）整合，输出为 Sheet3（写回原文件，保留 Sheet1/Sheet2）。

处理规则（§3.7 + v4.0 D51）：
- 合并键 = (地市, 区县, 街道, 监测地点)，BI 与 ADI 外连接；
- 仅出现于 BI 表 → ADI 数值写 '/'；仅出现于 ADI 表 → BI* 数值写 '/'；
- 防御处理：同键在一侧仍有多行时，取监测指标值最大的行（§3.7.1）；
- 输出列 = 地市 | 区县 | 街道 | 监测地点 | ADI* | BI*（去除“风险水平*”列）；
- 排序：地市固定顺序 → 区县升序 → 街道升序 → 监测地点升序（GBK/拼音序）；
- 格式：所有单元格水平+垂直居中；地市列纵向合并连续相同单元格；
  ADI/BI* 数值列按风险等级着色（安全 92D050 / 低 FFFF00 / 中 FFC000 / 高 FF0000）；字号五号(10.5)。

用法：python3 整合工作簿.py [工作簿.xlsx]
"""
import sys

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from mosquito import config as C
from mosquito.excel_output import _build_integrated_frame, _has_cjk, _pinyin_key

_KEYS = ['地市', '区县', '街道', '监测地点']
_OUT_COLS = ['地市', '区县', '街道', '监测地点', 'ADI*', 'BI*']
_CENTER = Alignment(horizontal='center', vertical='center')


def _keep_max_per_key(frame, value_col):
    """防御处理：同键 (地市,区县,街道,监测地点) 在一侧仍有多行时，取监测指标值最大的行"""
    frame = frame.copy()
    dup = frame.duplicated(subset=_KEYS, keep=False)
    if not dup.any():
        return frame
    keep = frame.groupby(_KEYS)[value_col].idxmax()
    return frame.loc[keep].sort_index()


def integrate_sheets(bi_df, adi_df):
    """Sheet1(BI) + Sheet2(ADI) -> 整合表 DataFrame（含内部风险列，供着色）"""
    bi = _keep_max_per_key(bi_df, 'BI*')
    adi = adi_df.copy()
    if 'ADI' in adi.columns and 'ADI*' not in adi.columns:
        adi = adi.rename(columns={'ADI': 'ADI*'})      # 指标列名统一为 ADI*
    adi = _keep_max_per_key(adi, 'ADI*')
    # 复用程序 §3.7 的整合逻辑（合并键、缺失写'/'、排序、内部风险列）
    return _build_integrated_frame(bi, adi)


def write_sheet3(wb, frame, sheet_name='Sheet3'):
    """把整合表写入工作簿的 Sheet3（保留原有 Sheet1/Sheet2 及其格式）"""
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    font_size = C.SIZE_XIAOSI      # 小四

    # 表头（加粗）
    for j, name in enumerate(_OUT_COLS, 1):
        cell = ws.cell(row=1, column=j, value=name)
        cell.font = Font(name=C.FONT_CN, size=font_size, bold=True)
        cell.alignment = _CENTER

    # 数据行 + 风险着色
    for i, (_, r) in enumerate(frame.iterrows(), start=2):
        for j, col in enumerate(_OUT_COLS, 1):
            v = r[col]
            cell = ws.cell(row=i, column=j, value=v)
            cell.font = Font(name=C.FONT_CN if _has_cjk(v) else C.FONT_EN, size=font_size)
            cell.alignment = _CENTER
            if col in ('ADI*', 'BI*') and isinstance(v, (int, float)):
                cell.number_format = '0.0'
        for col, risk_col in (('ADI*', 'ADI风险'), ('BI*', 'BI风险')):
            color = C.RISK_FILLS.get(r.get(risk_col))
            if color:
                ws.cell(row=i, column=_OUT_COLS.index(col) + 1).fill = PatternFill('solid', fgColor=color)

    # 地市列纵向合并连续相同单元格（数据区自第 2 行起）
    max_row = ws.max_row
    if max_row > 2:
        start = 2
        prev = ws.cell(row=2, column=1).value
        for row in range(3, max_row + 1):
            cur = ws.cell(row=row, column=1).value
            if cur != prev:
                if row - 1 > start:
                    ws.merge_cells(start_row=start, start_column=1, end_row=row - 1, end_column=1)
                start = row
                prev = cur
        if max_row > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=max_row, end_column=1)

    # 有内容的单元格（含表头）显示全部框线（细线）
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(_OUT_COLS)):
        for cell in row:
            cell.border = border

    # 列宽
    for j, col in enumerate(_OUT_COLS, 1):
        width = max(len(col), *(len(str(v)) for v in frame[col].head(200)))
        ws.column_dimensions[chr(64 + j)].width = min(width + 4, 60)
    return ws


def main(path):
    bi = pd.read_excel(path, sheet_name='Sheet1')
    adi = pd.read_excel(path, sheet_name='Sheet2')
    # 输入校验：Sheet1 需含 BI*、Sheet2 需含 ADI 或 ADI*
    if 'BI*' not in bi.columns or not ('ADI' in adi.columns or 'ADI*' in adi.columns):
        raise SystemExit('Sheet1/Sheet2 列不匹配：Sheet1 需含 BI*，Sheet2 需含 ADI（或 ADI*）。')

    frame = integrate_sheets(bi, adi)
    wb = load_workbook(path)
    write_sheet3(wb, frame)
    wb.save(path)

    print('整合完成：Sheet1(%d 行 BI) + Sheet2(%d 行 ADI) -> Sheet3(%d 行)'
          % (len(bi), len(adi), len(frame)))
    print('输出文件：%s' % path)
    print('Sheet3 列：%s' % ' | '.join(_OUT_COLS))
    n_bi_only = int((frame['ADI*'] == '/').sum())
    n_adi_only = int((frame['BI*'] == '/').sum())
    print('仅BI侧 %d 行（ADI* 写 /）；仅ADI侧 %d 行（BI* 写 /）' % (n_bi_only, n_adi_only))


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else '工作簿1.xlsx')
