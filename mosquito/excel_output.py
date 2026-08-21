# -*- coding: utf-8 -*-
"""Excel 输出：计算过程 Excel（9 个 Sheet）与 监测点汇总 Excel（3 个 Sheet）。

使用 openpyxl 写入单元格填充色（黄/红标记、风险等级背景色）。
"""
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from . import config as C
from .parser import district_display_sheet

_INTERNAL_COLS = ('_yellow', '_deleted', '_modified', '_K', '_conv', '_orig', '_in_bi', '_src')
_NUM_COLS = ('监测指标值', 'BI*', 'ADI', '原BI值', '原SSI值', '转换后的SSI值')


def _strip_internal(df):
    return df.drop(columns=[c for c in _INTERNAL_COLS if c in df.columns])


def _write_sheet(writer, name, df, flag_col=None, risk_col=None, fill_all=False):
    """写入一个 Sheet；flag_col 为真值列时按 FILL_YELLOW/FILL_RED 填充整行；
    risk_col 非空时按风险等级填充整行背景色。"""
    out = _strip_internal(df)
    out.to_excel(writer, sheet_name=name, index=False)
    wb = writer.book
    ws = wb[name]

    # 数值列格式
    header = {c.value: c.column for c in ws[1]}
    for col_name in _NUM_COLS:
        col_idx = header.get(col_name)
        if col_idx:
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.0'

    # 填充色
    if flag_col or risk_col:
        for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
            if flag_col and flag_col in df.columns and bool(row.get(flag_col)):
                color = C.FILL_RED if flag_col.startswith('_deleted') else C.FILL_YELLOW
                fill = PatternFill('solid', fgColor=color)
                for c_idx in range(1, len(out.columns) + 1):
                    ws.cell(row=r_idx, column=c_idx).fill = fill
            elif risk_col and risk_col in df.columns:
                risk = row.get(risk_col)
                color = C.RISK_FILLS.get(risk)
                if color:
                    fill = PatternFill('solid', fgColor=color)
                    for c_idx in range(1, len(out.columns) + 1):
                        ws.cell(row=r_idx, column=c_idx).fill = fill

    # 列宽（按内容自适应，上限 60）
    for c_idx in range(1, len(out.columns) + 1):
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


def write_monitoring_workbook(path, bi_final, adi_final, deletions):
    """监测点汇总 Excel：Sheet1 BI表 / Sheet2 ADI表 / Sheet3 删除数据情况说明"""
    bi = _display_frame(bi_final, 'BI*')
    adi = _display_frame(adi_final, 'ADI')
    with pd_writer(path) as writer:
        _write_sheet(writer, 'BI表', bi, risk_col='风险水平*')
        _write_sheet(writer, 'ADI表', adi, risk_col='风险水平*')
        _write_sheet(writer, '删除数据情况说明', deletions)


def _display_frame(final, value_col):
    """最终表 -> 监测点汇总表显示口径（区县去后缀、市辖区->/），按 3.7 排序"""
    df = final.copy()
    df['区县'] = df['区县'].map(district_display_sheet)
    df['_type'] = df[C.COL_TYPE]
    df = df.sort_values(
        ['_city_idx', '区县', '街道', '监测地点', '_type'],
        kind='stable',
    ).reset_index(drop=True)
    return df[['地市', '区县', '街道', '监测地点', value_col, '风险水平*']]


def pd_writer(path):
    """兼容 pandas 2.x 的 ExcelWriter 上下文（openpyxl 引擎）"""
    import pandas as pd
    return pd.ExcelWriter(path, engine='openpyxl')
