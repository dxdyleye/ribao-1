# -*- coding: utf-8 -*-
"""Word 输出：日报（叙述版，3.6/3.6.1）与 一览表（BI+ADI 整合表，3.7.1）。

排版参数按示例文档实测：
- 日报：A4，页边距 上3.70/下3.50/左2.80/右2.60；标题方正小标宋简体二号居中；副题仿宋三号居中；
  章节黑体三号首行缩进2字符；正文仿宋三号两端对齐、首行缩进2字符、行距28磅固定。
- 一览表：A4，页边距 2.54/2.54/3.17/3.17；标题方正小标宋二号居中两行；副题楷体三号居中；
  表头/数据仿宋五号居中；表后注释宋体/仿宋小四。
"""
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import config as C
from .parser import district_display_word
from .pipeline import round1

_JUSTIFY = WD_ALIGN_PARAGRAPH.JUSTIFY
_CENTER = WD_ALIGN_PARAGRAPH.CENTER

_NOTES = [
    '1.省疾控中心飞行监测及市、县（区）疾控中心在两热疫点镇（街道）开展入户调查、蚊媒应急监测评估等工作。',
    '2.广州市幼虫监测如为标准间指数SSI按1:2换算成布雷图指数BI。',
    '3.风险水平分级标准为：',
    '安全范围：BI<5，或成蚊密度ADI≤2，符合蚊媒防控基本要求；',
    '低度风险：5≤BI<10，或成蚊密度2＜ADI≤5，存在传播潜在风险；',
    '中度风险：10≤BI<20，或成蚊密度5＜ADI≤10，提示可能引发本地疫情暴发；',
    '高度风险：BI≥20，或成蚊密度ADI＞10，发现病例易导致疫情暴发流行。',
]


# ---------------- 基础工具 ----------------

def _set_run(r, east, size, bold=None):
    r.font.name = 'Times New Roman'
    rPr = r._element.get_or_add_rPr()
    rPr.get_or_add_rFonts().set(qn('w:eastAsia'), east)
    r.font.size = Pt(size)
    if bold is not None:
        r.font.bold = bold


def _add_par(doc, text, east, size, align=None, indent_chars=0, line_pt=None, bold=None, keep_next=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    if line_pt:
        pf.line_spacing = Pt(line_pt)          # 固定行距（exact）
    if keep_next:
        pf.keep_with_next = True
    if indent_chars:
        pPr = p._p.get_or_add_pPr()
        ind = pPr.get_or_add_ind()
        ind.set(qn('w:firstLineChars'), str(indent_chars * 100))
        ind.set(qn('w:firstLine'), str(int(indent_chars * size * 20)))
    if text:
        r = p.add_run(text)
        _set_run(r, east, size, bold)
    return p


def _setup_page(doc, top, bottom, left, right):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(top)
    sec.bottom_margin = Cm(bottom)
    sec.left_margin = Cm(left)
    sec.right_margin = Cm(right)


# ---------------- 日报（3.6） ----------------

def build_section(final_df, exclude_field, excluded_cities, metric):
    """从最终表构建 Word 章节统计（metric = 'BI' 或 'ADI'）"""
    df = final_df
    value_col = 'BI*' if metric == 'BI' else 'ADI'

    # 地市列表（固定顺序、名称不带市、最后一个保留市、排除注记）
    cities = [c for c in C.CITY_ORDER if c in set(df['地市'])]
    if not cities:
        city_list = ''
    elif len(cities) == 1:
        c0 = cities[0]
        note = '（%s除外）' % exclude_field if exclude_field and c0 in excluded_cities else ''
        city_list = c0 + note + '市'
    else:
        names = []
        for c in cities:
            nm = c
            if exclude_field and c in excluded_cities:
                nm = '%s（%s除外）' % (c, exclude_field)
            names.append(nm)
        last = cities[-1]
        if exclude_field and last in excluded_cities:
            last_name = '%s市（%s除外）' % (last, exclude_field)
        else:
            last_name = names[-1] + '市'
        city_list = '、'.join(names[:-1]) + '和' + last_name

    total = len(df)
    district_num = df.groupby(['地市', '区县']).ngroups
    street_num = df.groupby(['地市', '街道']).ngroups
    point_num = df['监测地点'].nunique()
    if metric == 'BI':
        ok_num = int((df[value_col] < 5).sum())
        risk_num = int((df[value_col] >= 5).sum())
    else:
        ok_num = int((df[value_col] <= 2).sum())
        risk_num = int((df[value_col] > 2).sum())
    ok_rate = round1(ok_num * 100.0 / total) if total else 0.0
    risk_rate = round1(risk_num * 100.0 / total) if total else 0.0

    # 风险列表（高→中→低，组内按值降序；数量为 0 不输出）
    def element(row):
        dist = row['区县'] if row['区县'] != '市辖区' else ''
        return '%s市%s%s%s（%.1f）' % (row['地市'], dist, row['街道'], row['监测地点'], row[value_col])

    if metric == 'BI':
        high = df[df[value_col] >= 20]
        mid = df[(df[value_col] >= 10) & (df[value_col] < 20)]
        low = df[(df[value_col] >= 5) & (df[value_col] < 10)]
    else:
        high = df[df[value_col] > 10]
        mid = df[(df[value_col] > 5) & (df[value_col] <= 10)]
        low = df[(df[value_col] > 2) & (df[value_col] <= 5)]

    risk_lists = {
        '高风险': [element(r) for _, r in high.sort_values(value_col, ascending=False).iterrows()],
        '中风险': [element(r) for _, r in mid.sort_values(value_col, ascending=False).iterrows()],
        '低风险': [element(r) for _, r in low.sort_values(value_col, ascending=False).iterrows()],
    }

    return {
        'metric': metric, 'city_list': city_list, 'city_num': len(cities),
        'district_num': district_num, 'street_num': street_num, 'point_num': point_num,
        'total': total, 'ok_num': ok_num, 'risk_num': risk_num,
        'ok_rate': ok_rate, 'risk_rate': risk_rate, 'risk_lists': risk_lists,
    }


def _write_section(doc, heading, sec, m, d):
    _add_par(doc, heading, '黑体', 16, indent_chars=2, line_pt=28, keep_next=True)
    metric = sec['metric']
    if sec['total'] == 0:
        _add_par(doc, '%d月%d日，当日无符合条件（距末例天数≤5或>40000、防控区类型为核心区/警戒区）的%s监测数据。'
                 % (m, d, metric), '仿宋_GB2312', 16, _JUSTIFY, indent_chars=2, line_pt=28)
        return
    overview = ('%d月%d日，对%s共%d个地市%d个县区%d个镇街%d个村居/监测点开展了蚊媒密度应急监测。'
                '监测结果显示：%d个监测点%s达标准要求，达标率为%.1f%%（%d/%d）；'
                '%d个监测点%s为低中高风险，占%.1f%%（%d/%d）。') % (
        m, d, sec['city_list'], sec['city_num'], sec['district_num'], sec['street_num'],
        sec['point_num'], sec['ok_num'], metric, sec['ok_rate'], sec['ok_num'], sec['total'],
        sec['risk_num'], metric, sec['risk_rate'], sec['risk_num'], sec['total'])
    _add_par(doc, overview, '仿宋_GB2312', 16, _JUSTIFY, indent_chars=2, line_pt=28)
    for level in ('高风险', '中风险', '低风险'):
        elems = sec['risk_lists'][level]
        if elems:
            _add_par(doc, '%s%d个：%s。' % (level, len(elems), '、'.join(elems)),
                     '仿宋_GB2312', 16, _JUSTIFY, indent_chars=2, line_pt=28)


def write_daily_report(path, target_date, bi_section, adi_section):
    doc = Document()
    _setup_page(doc, 3.70, 3.50, 2.80, 2.60)
    y, m, d = target_date.year, target_date.month, target_date.day

    _add_par(doc, '全省媒介伊蚊传染病相关蚊媒监测情况', '方正小标宋简体', 22, _CENTER, line_pt=36)
    _add_par(doc, '', '仿宋_GB2312', 16, line_pt=15)
    _add_par(doc, '（省疾控中心  %d年%d月%d日20:00）' % (y, m, d), '仿宋_GB2312', 16, _CENTER, line_pt=28)
    _add_par(doc, '', '仿宋_GB2312', 16, indent_chars=2, line_pt=28)

    _write_section(doc, '一、布雷图指数（BI）调查评估。', bi_section, m, d)
    _write_section(doc, '二、成蚊密度（ADI）快速评估。', adi_section, m, d)

    doc.save(path)


# ---------------- 一览表 Word（3.7.1） ----------------

def build_merged_table(bi_final, adi_final):
    """合并最终 BI 表与 ADI 表：键 = (地市, 区县, 街道, 监测地点)；缺失项写 '/'"""
    bi = bi_final[['地市', '区县', '街道', '监测地点', 'BI*', '风险水平*', '_city_idx']].copy()
    bi = bi.rename(columns={'风险水平*': 'BI风险'})
    adi = adi_final[['地市', '区县', '街道', '监测地点', 'ADI', '风险水平*', '_city_idx']].copy()
    adi = adi.rename(columns={'风险水平*': 'ADI风险'})

    m = bi.merge(adi, on=['地市', '区县', '街道', '监测地点'], how='outer', suffixes=('', '_adi'))
    for col in ('ADI', 'ADI风险', 'BI*', 'BI风险'):
        if col not in m.columns:
            m[col] = '/'
        m[col] = m[col].where(m[col].notna(), '/')

    m['区县'] = m['区县'].map(district_display_word)
    if '_city_idx_adi' in m.columns:
        m['_city_idx'] = m['_city_idx'].fillna(m['_city_idx_adi'])
    m['_city_idx'] = m['_city_idx'].fillna(99).astype(int)
    m = m.sort_values(['_city_idx', '区县', '街道', '监测地点'], kind='stable').reset_index(drop=True)
    return m


def _scale_widths(base):
    """列宽按版心（21.0 - 3.17*2 = 14.66 cm）等比缩放"""
    total = sum(base)
    scale = 14.66 / total
    widths = [round(w * scale, 2) for w in base]
    widths[-1] = round(14.66 - sum(widths[:-1]), 2)
    return widths


def write_summary_word(path, target_date, merged):
    doc = Document()
    _setup_page(doc, 2.54, 2.54, 3.17, 3.17)
    y, m, d = target_date.year, target_date.month, target_date.day

    _add_par(doc, '全省媒介伊蚊传染病疫点', '方正小标宋简体', 22, _CENTER)
    _add_par(doc, '蚊媒密度监测情况', '方正小标宋简体', 22, _CENTER)
    _add_par(doc, '（省疾控中心  %d年%d月%d日20:00）' % (y, m, d), '楷体_GB2312', 16, _CENTER)
    _add_par(doc, '', '仿宋_GB2312', 10.5)

    header = ['地市', '区县', '街道', '监测地点', 'ADI', '风险水平*', 'BI*', '风险水平*']
    table = doc.add_table(rows=1 + len(merged), cols=len(header))
    table.style = 'Table Grid'
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)

    widths = _scale_widths([1.2, 1.4, 2.5, 5.0, 1.4, 2.1, 1.2, 2.2])
    grid = table._tbl.find(qn('w:tblGrid'))
    for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
        gc.set(qn('w:w'), str(int(w * 567)))  # cm -> twips

    for j, h in enumerate(header):
        cell = table.rows[0].cells[j]
        cell.width = Cm(widths[j])
        p = cell.paragraphs[0]
        p.alignment = _CENTER
        _set_run(p.add_run(h), '仿宋_GB2312', 10.5, bold=True)

    data_cols = ['地市', '区县', '街道', '监测地点', 'ADI', 'ADI风险', 'BI*', 'BI风险']
    for i, (_, row) in enumerate(merged.iterrows(), start=1):
        for j, col in enumerate(data_cols):
            cell = table.rows[i].cells[j]
            cell.width = Cm(widths[j])
            p = cell.paragraphs[0]
            p.alignment = _CENTER
            v = row[col]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                txt = '%.1f' % v
            else:
                txt = str(v)
            _set_run(p.add_run(txt), '仿宋_GB2312', 10.5)

    _add_par(doc, '* 注：', '宋体', 12, _JUSTIFY, bold=True)
    for note in _NOTES:
        _add_par(doc, note, '仿宋_GB2312', 12, _JUSTIFY)

    doc.save(path)
