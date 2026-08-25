# -*- coding: utf-8 -*-
"""Excel 输出：计算过程 Excel（9 个 Sheet）与 村居一览表 Excel（4 个 Sheet）。

使用 openpyxl 写入单元格填充色（黄/红标记、风险等级背景色）。

监测点汇总 BI表/ADI表（3.7 节）：
- 仅“风险水平*”列按数值着色，其余列不着色；
- “地市”列纵向合并连续相同的单元格；
- 所有单元格水平 + 垂直居中；
- 字体：含中文的单元格用 仿宋_GB2312，纯英文/数字单元格用 Times New Roman
  （xlsx 单格仅支持一个字体名，无法像 Word 那样按脚本分别设置）。
"""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import config as C
from .parser import district_display_sheet
from .pipeline import round1

import re

_INTERNAL_COLS = ('_yellow', '_deleted', '_modified', '_K', '_conv', '_orig', '_in_bi', '_src')
_NUM_COLS = ('监测指标值', 'BI*', 'ADI*', '原BI值', '原SSI值', '转换后的SSI值')
_CENTER = Alignment(horizontal='center', vertical='center')


def _has_cjk(text):
    if text is None:
        return False
    return any('\u4e00' <= ch <= '\u9fff' for ch in str(text))


def _pinyin_key(s):
    """中文按拼音（字母）升序的排序键：GB2312 一级汉字按拼音排序，
    因此 GBK 字节序近似拼音序（无第三方依赖）。无法编码的字符退回 Unicode 码点。"""
    if s is None:
        return ''
    s = str(s)
    try:
        return s.encode('gbk').hex()
    except UnicodeEncodeError:
        return s


_TYPE_ORDER = {'核心区': 0, '警戒区': 1}   # 同社区内防控区类型排列：核心区 → 警戒区 → 其他
_TYPE_SUFFIX_RE = re.compile(r'（\d+）\s*$')
_TYPE_RE = re.compile(r'（([^（）]*?)）$')


def _type_sort_key(mp):
    """从监测地点提取防控区类型作为排序键（去掉冲突序号（1）与“，飞行监测”标记）：
    核心区=0、警戒区=1、其他类型=2。"""
    s = str(mp)
    s = _TYPE_SUFFIX_RE.sub('', s)                       # 去掉（1）（2）… 冲突序号
    m = _TYPE_RE.search(s)
    t = m.group(1) if m else ''
    t = t.replace('，飞行监测', '')
    return _TYPE_ORDER.get(t, 2)


def _community_sort_key(mp):
    """从监测地点提取社区/村居作为排序键（取第一个“（”之前的部分，即社区名）"""
    s = str(mp)
    i = s.find('（')
    return _pinyin_key(s[:i] if i >= 0 else s)


def _cell_font(size, text, bold=False):
    """含中文 -> 仿宋_GB2312；纯英文/数字 -> Times New Roman（xlsx 每格单一字体名）"""
    f = Font(name=C.FONT_CN if _has_cjk(text) else C.FONT_EN)
    if size:
        f.size = size
    if bold:
        f.bold = True
    return f


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


def _write_sheet(writer, name, df, flag_col=None, risk_cols=None, risk_src=None,
                 drop_cols=None, merge_cols=None, center=False, header_rename=None, font_size=None,
                 borders=False, header_bold=False):
    """写入一个 Sheet。

    - flag_col：真值列，整行按 FILL_YELLOW/FILL_RED 填充（计算过程表标记用）；
    - risk_cols：按风险等级只给这些列着色（风险水平文字列，如“安全/低风险/…”）；
    - risk_src：{输出列: 风险水平文字列}，按风险等级给输出列着色（需求四：颜色套用到 BI*/ADI* 列，
      风险文字列本身不输出）；
    - drop_cols：写入前删除的 df 列（如被 risk_src 借用、不输出的“风险水平*”）；
    - merge_cols：纵向合并这些列中连续相同的单元格；
    - center：所有单元格水平 + 垂直居中；
    - header_rename：{df列名: 新表头}，写表头后改名（用于整合表两个“风险水平*”）；
    - font_size：单元格字号（None 用默认）；所有单元格中文字体 仿宋_GB2312、英文 Times New Roman；
    - borders：为有内容的单元格（含表头）显示全部框线（细线）；
    - header_bold：标题行（表头）加粗。
    """
    out = _strip_internal(df)
    if drop_cols:
        out = out.drop(columns=[c for c in drop_cols if c in out.columns])
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

    # 字体（含中文 仿宋_GB2312 / 纯英文数字 Times New Roman）+ 居中
    for r in range(1, ws.max_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = _cell_font(font_size, cell.value)
            if center:
                cell.alignment = _CENTER

    # 标题行（表头）加粗
    if header_bold:
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = _cell_font(font_size, cell.value, bold=True)

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
    if risk_src:
        for out_col, src_col in risk_src.items():
            if out_col not in list(out.columns):
                continue
            wcol = list(out.columns).index(out_col) + 1
            for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
                risk = row.get(src_col)
                color = C.RISK_FILLS.get(risk)
                if color:
                    ws.cell(row=r_idx, column=wcol).fill = PatternFill('solid', fgColor=color)

    # 纵向合并连续相同单元格
    if merge_cols:
        for col_name in merge_cols:
            col_idx = header.get(col_name)
            if col_idx:
                _merge_same_values(ws, col_idx)

    # 有内容的单元格（含表头）显示全部框线（细线）
    if borders:
        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for r in range(1, ws.max_row + 1):
            for c in range(1, ncols + 1):
                ws.cell(row=r, column=c).border = border

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



def write_monitoring_workbook(path, bi_final, adi_final, deletions):
    """监测点汇总 Excel（村居一览表）：Sheet1 BI表 / Sheet2 ADI表 / Sheet3 BI+ADI整合表 / Sheet4 删除数据情况说明

    需求四：去除“风险水平*”列，原来的风险背景色（绿/黄/橘/红）套用到 BI*/ADI* 数值列；
    整合表缺失项（'/'）背景同安全绿色（92D050）。
    字号：整合表小四(12)，其余 sheet 14；整合表有内容的单元格显示全部框线；
    字体：中文 仿宋_GB2312、英文 Times New Roman。
    """
    bi = _display_frame(bi_final, 'BI*')
    adi = _display_frame(adi_final, 'ADI').rename(columns={'ADI': 'ADI*'})   # 指标列名用 ADI*
    integrated = _build_integrated_frame(bi, adi)
    with pd_writer(path) as writer:
        _write_sheet(writer, 'BI表', bi, risk_src={'BI*': '风险水平*'},
                     drop_cols=['风险水平*'], merge_cols=['地市'], center=True, font_size=C.SIZE_14)
        _write_sheet(writer, 'ADI表', adi, risk_src={'ADI*': '风险水平*'},
                     drop_cols=['风险水平*'], merge_cols=['地市'], center=True, font_size=C.SIZE_14)
        _write_sheet(writer, 'BI+ADI整合表', integrated, risk_src={'ADI*': 'ADI风险', 'BI*': 'BI风险'},
                     drop_cols=['ADI风险', 'BI风险'], merge_cols=['地市'], center=True,
                     font_size=C.SIZE_XIAOSI, borders=True, header_bold=True)
        _write_sheet(writer, '删除数据情况说明', deletions, font_size=C.SIZE_14)


def _build_integrated_frame(bi, adi):
    """BI 表与 ADI 表整合（规则同 Word 一览表）：键 = (地市, 区县, 街道, 监测地点)，
    缺失项数值写 '/'；输出列 = 地市/区县/街道/监测地点/ADI*/BI*（ADI 指标列名用 ADI*；
    风险水平列仅内部用于着色）。缺失项（'/'）内部风险按“安全”处理 → 背景同为安全绿色（92D050）。"""
    bi_m = bi.rename(columns={'风险水平*': 'BI风险'})
    adi_m = adi.rename(columns={'风险水平*': 'ADI风险'})
    m = bi_m.merge(adi_m, on=['地市', '区县', '街道', '监测地点'], how='outer')
    for col in ('ADI*', 'BI*'):
        if col not in m.columns:
            m[col] = '/'
        m[col] = m[col].where(m[col].notna(), '/')
    for col in ('ADI风险', 'BI风险'):
        if col not in m.columns:
            m[col] = '安全'
        m[col] = m[col].where(m[col].notna(), '安全')
    m['_city_idx'] = m['地市'].map(C.CITY_INDEX).fillna(99).astype(int)
    m['_k1'] = m['区县'].map(_pinyin_key)
    m['_k2'] = m['街道'].map(_pinyin_key)
    m['_kcomm'] = m['监测地点'].map(_community_sort_key)   # 同一社区聚在一起
    m['_ktype'] = m['监测地点'].map(_type_sort_key)        # 社区内 核心区 → 警戒区
    m['_k3'] = m['监测地点'].map(_pinyin_key)
    m = m.sort_values(['_city_idx', '_k1', '_k2', '_kcomm', '_ktype', '_k3'], kind='stable') \
        .drop(columns=['_k1', '_k2', '_kcomm', '_ktype', '_k3']).reset_index(drop=True)
    return m[['地市', '区县', '街道', '监测地点', 'ADI*', 'ADI风险', 'BI*', 'BI风险']]


def _display_frame(final, value_col):
    """最终表 -> 村居一览表显示口径（区县去后缀、市辖区->-、数值四舍五入1位），
    排序：地市固定顺序 → 区县升序（拼音）→ 街道升序（拼音）→ 社区/村居 → 防控区类型（核心区→警戒区）→ 监测地点升序（拼音）"""
    df = final.copy()
    df['区县'] = df['区县'].map(district_display_sheet)
    df[value_col] = df[value_col].map(round1)       # 与参考一览表一致：四舍五入保留 1 位
    df['_k1'] = df['区县'].map(_pinyin_key)
    df['_k2'] = df['街道'].map(_pinyin_key)
    df['_kcomm'] = df['监测地点'].map(_community_sort_key)   # 同一社区聚在一起
    df['_ktype'] = df['监测地点'].map(_type_sort_key)        # 社区内 核心区 → 警戒区
    df['_k3'] = df['监测地点'].map(_pinyin_key)
    df = df.sort_values(['_city_idx', '_k1', '_k2', '_kcomm', '_ktype', '_k3'], kind='stable') \
        .drop(columns=['_k1', '_k2', '_kcomm', '_ktype', '_k3']).reset_index(drop=True)
    return df[['地市', '区县', '街道', '监测地点', value_col, '风险水平*']]


def pd_writer(path):
    """兼容 pandas 2.x 的 ExcelWriter 上下文（openpyxl 引擎）"""
    import pandas as pd
    return pd.ExcelWriter(path, engine='openpyxl')
