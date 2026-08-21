# -*- coding: utf-8 -*-
"""数据处理流水线（需求文档 v3.2 + 金标准 BI-报出整理/ADI-报出整理校准）。

关键规则（已按金标准实测校准）：
- 数值保留源精度（不四舍五入到 1 位小数）。
- BI 表 = BI 方法行 + “无同键 BI”的 SSI 行（×2 换算并入）；有同键 BI 的 SSI 行丢弃（不取大合并）。
- 按完整键 K=(地市-区/县/市-街道/乡/镇, 社区/村居, 地址1, 地址2, 防控区类型) 去重，保留最大值。
- 地址区分：仅对 (地市-区/县/市-街道/乡/镇, 社区/村居, 防控区类型) 组内存在 >=2 个不同地址的行改名。
"""
import re
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd

from . import config as C
from .parser import parse_location


class ProcessingError(Exception):
    """业务处理异常（GUI 弹出提示后停止本次处理，不退出程序）"""


def round1(x):
    """四舍五入保留 1 位小数（Decimal ROUND_HALF_UP），用于百分比/显示"""
    if x is None or (isinstance(x, float) and x != x):
        return x
    return float(Decimal(str(x)).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP))


def validate_columns(df):
    missing = [c for c in C.REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ProcessingError('源文件缺少必需列：' + '、'.join(missing))


def _make_keys(frame):
    """键值 K = (地市-区/县/市-街道/乡/镇, 社区/村居, 地址1, 地址2, 防控区类型)；空值视为空字符串"""
    cols = [C.COL_LOC, C.COL_COMMUNITY, C.COL_ADDR1, C.COL_ADDR2, C.COL_TYPE]
    t = frame[cols].copy()
    t = t.where(t.notna(), '')
    return t.astype(str).apply(tuple, axis=1)


# ---------------- 最小区分地址（3.8 节，金标准校准版） ----------------

_ADDR_SEP = re.compile(r'[，,、。\s]+')


def _clean_addr(addr, loc, community, type_):
    """清理地址：依次去除与 完整地市串/街道/区县/社区/地市名/防控区类型 重复的连续子串（先长后短）"""
    s = str(addr)
    parsed = parse_location(loc)
    city = parsed[0] if parsed else None
    district = parsed[1] if parsed else None
    street = parsed[2] if parsed else None
    pats = [p for p in (loc, street, district, community,
                        city + '市' if city else None, type_) if p]
    for pat in sorted(set(pats), key=len, reverse=True):
        while pat in s:
            s = s.replace(pat, '')
    return s


def apply_address_distinction(frame):
    """3.8 地址区分：组 = (地市-区/县/市-街道/乡/镇, 社区/村居, 防控区类型)。

    仅当组内存在 >=2 个不同地址时，逐行以“社区/村居+最小区分地址（类型）”改名；
    单地址组不改名（与金标准一致）。返回新增 '_modified' 标记的副本。
    """
    frame = frame.copy()
    frame['_modified'] = False
    group_cols = [C.COL_LOC, C.COL_COMMUNITY, C.COL_TYPE]

    def addr_of(r):
        a = r[C.COL_ADDR1]
        if a is None or (isinstance(a, float) and a != a) or str(a).strip() == '':
            a = r[C.COL_ADDR2]
        if a is None or (isinstance(a, float) and a != a) or str(a).strip() == '':
            return None
        return str(a).strip()

    for key, idx in frame.groupby(group_cols).groups.items():
        rows = frame.loc[idx]
        addrs = rows.apply(addr_of, axis=1)
        if len(set(a for a in addrs if a)) < 2:
            continue
        # 每行候选片段（地址优先取地图定位版，空则手填）
        cand = {}
        maxlen = 0
        for i, r in rows.iterrows():
            a = addr_of(r)
            if a is None:
                cand[i] = None
                continue
            cleaned = _clean_addr(a, r[C.COL_LOC], r[C.COL_COMMUNITY], r[C.COL_TYPE])
            segs = [s for s in _ADDR_SEP.split(cleaned) if s]
            cand[i] = segs if segs else None
            if segs:
                maxlen = max(maxlen, len(segs))
        # 逐片段扩展直到组内候选互异
        cur = {}
        for n in range(1, maxlen + 1):
            cur = {}
            for i, segs in cand.items():
                cur[i] = ''.join(segs[:n]) if segs else None
            vals = [v for v in cur.values() if v]
            if len(vals) == len(set(vals)):
                break
        # 冲突/无地址行：保持原监测地点，追加序号
        groups = defaultdict(list)
        for i, v in cur.items():
            groups[v].append(i)
        for v, idxs in groups.items():
            if len(idxs) > 1:
                for j, i in enumerate(idxs, 1):
                    new = '%s（%s）（%d）' % (frame.at[i, C.COL_COMMUNITY], frame.at[i, C.COL_TYPE], j)
                    if new != frame.at[i, '监测地点']:
                        frame.loc[i, '监测地点'] = new
                        frame.loc[i, '_modified'] = True
        # 正常改名
        for i, v in cur.items():
            if v is None:
                continue
            comm = frame.at[i, C.COL_COMMUNITY]
            typ = frame.at[i, C.COL_TYPE]
            while v.startswith(comm):
                v = v[len(comm):]
            new = '%s%s（%s）' % (comm, v, typ)
            if new != frame.at[i, '监测地点']:
                frame.loc[i, '监测地点'] = new
                frame.loc[i, '_modified'] = True
    return frame


# ---------------- 流水线结果 ----------------

class PipelineResult(object):
    def __init__(self):
        self.base = None
        self.bi_final = None      # 最终BI表（含隐藏列 防控区类型、_city_idx）
        self.adi_final = None     # 最终ADI表
        self.calc_sheets = {}     # 计算过程Excel: sheet名 -> (DataFrame, 标记列或None)
        self.deletions = None     # 删除数据情况说明
        self.excluded_cities = set()


# ---------------- 主流程 ----------------

def run_pipeline(source_df, target_date, exclude=None):
    validate_columns(source_df)

    df = source_df.copy()
    del_rows = []
    excluded_cities = set()

    def record(idx_list, reason):
        for i in idx_list:
            del_rows.append((i, reason))

    # ---- P1 日期解析 + P2 监测时间为空 ----
    dpart = pd.to_datetime(df[C.COL_TIME], errors='coerce').dt.date
    mask_time_nan = dpart.isna()
    record(list(df.index[mask_time_nan]), '监测时间为空')
    df = df[~mask_time_nan]
    dpart = dpart[~mask_time_nan]
    df = df[dpart == target_date]      # 日期不符：不记录（D2）
    if df.empty:
        raise ProcessingError('该日期无监测数据')

    # ---- P3 监测指标值为空 ----
    vals = pd.to_numeric(df[C.COL_VALUE], errors='coerce')
    mask_val_nan = vals.isna()
    record(list(df.index[mask_val_nan]), '监测指标值为空')
    df = df[~mask_val_nan]
    df[C.COL_VALUE] = vals[~mask_val_nan]      # 保留原精度，不取整

    # ---- P5 排除字段 ----
    if exclude and str(exclude).strip():
        ex = str(exclude).strip()
        mask_ex = df[C.COL_LOC].astype(str).str.contains(ex, regex=False)
        for i in df.index[mask_ex]:
            del_rows.append((i, '被排除字段过滤'))
            parsed = parse_location(df.at[i, C.COL_LOC])
            if parsed is not None:
                excluded_cities.add(parsed[0])
        df = df[~mask_ex]

    # ---- P6 距末例天数（保留 <=5 或 >40000） ----
    days = pd.to_numeric(df[C.COL_DAYS], errors='coerce')
    keep_days = days.notna() & ((days <= C.DAYS_LOW) | (days > C.DAYS_HIGH))
    record(list(df.index[~keep_days]), '距末例天数不在范围内')
    df = df[keep_days]

    # ---- P7 防控区类型 ----
    mask_type = ~df[C.COL_TYPE].isin(C.VALID_TYPES)
    record(list(df.index[mask_type]), '防控区类型不符合')
    df = df[~mask_type]

    # ---- P8 监测地点 ----
    df['监测地点'] = df[C.COL_COMMUNITY].astype(str) + '（' + df[C.COL_TYPE].astype(str) + '）'

    # ---- P9 解析地市/区县/街道 ----
    parsed = df[C.COL_LOC].map(parse_location)
    mask_parse = parsed.isna()
    record(list(df.index[mask_parse]), '行政区划无法解析')
    df = df[~mask_parse]
    df['地市'] = [p[0] for p in parsed[~mask_parse]]
    df['区县'] = [p[1] for p in parsed[~mask_parse]]
    df['街道'] = [p[2] for p in parsed[~mask_parse]]

    base = df.copy()      # P10 基础数据集

    # ---- 方法归属：非 BI/SSI/ADI 记入删除说明 ----
    mask_method = ~df[C.COL_METHOD].isin(C.ALL_PROCESSED_METHODS)
    record(list(df.index[mask_method]), '监测方法不属于BI/SSI/ADI（不参与计算）')
    df = df[~mask_method]

    bi_ssi = df[df[C.COL_METHOD].isin(C.BI_SSI_METHODS)].copy()
    adi_raw = df[df[C.COL_METHOD] == C.METHOD_ADI].copy()

    bi_final, bi_sheets = _bi_ssi_pipeline(bi_ssi, del_rows)
    adi_final, adi_sheets = _adi_pipeline(adi_raw, del_rows)

    # ---- 删除数据情况说明（3.9） ----
    src_cols = list(source_df.columns)
    rows = []
    for i, reason in del_rows:
        row = {c: source_df.at[i, c] for c in src_cols}
        row['删除原因'] = reason
        rows.append(row)
    deletions = pd.DataFrame(rows, columns=src_cols + ['删除原因']) if rows else \
        pd.DataFrame(columns=src_cols + ['删除原因'])

    res = PipelineResult()
    res.base = base
    res.bi_final = bi_final
    res.adi_final = adi_final
    res.deletions = deletions
    res.excluded_cities = excluded_cities
    res.calc_sheets = {}
    res.calc_sheets.update(bi_sheets)
    res.calc_sheets.update(adi_sheets)
    return res


# ---------------- BI + SSI 专项（金标准校准版） ----------------

_BASE_SHEET_COLS = ['地市', '区县', '街道', C.COL_LOC, C.COL_COMMUNITY,
                    C.COL_ADDR1, C.COL_ADDR2, C.COL_TYPE, '监测地点',
                    C.COL_METHOD, C.COL_VALUE]


def _dedup_keep_max(frame, del_rows, reason='重复数据保留最大值'):
    """按完整键 K 去重，每组保留监测指标值最大的行；被删行记入删除说明并打 '_deleted' 标记"""
    frame = frame.copy()
    frame['_deleted'] = False
    keep = frame.groupby('_K')[C.COL_VALUE].idxmax()
    dropped = frame.index.difference(keep)
    frame.loc[dropped, '_deleted'] = True
    for i in dropped:
        del_rows.append((i, reason))
    return frame


def _bi_ssi_pipeline(bi_ssi, del_rows):
    ssi = bi_ssi[bi_ssi[C.COL_METHOD] == C.METHOD_SSI].copy()
    bi = bi_ssi[bi_ssi[C.COL_METHOD] == C.METHOD_BI].copy()

    # SSI ×2 换算（保留原精度）
    ssi['_conv'] = ssi[C.COL_VALUE] * 2
    ssi['_orig'] = ssi[C.COL_VALUE]
    ssi['_K'] = _make_keys(ssi)
    bi['_K'] = _make_keys(bi)
    bi_keys = set(bi['_K'])
    ssi['_in_bi'] = ssi['_K'].isin(bi_keys)

    # ---- Sheet1「SSI表」：全部 SSI 行（×2 后） ----
    sheet1 = ssi[_BASE_SHEET_COLS].copy()
    sheet1[C.COL_VALUE] = ssi['_conv']

    # ---- 并入：无同键 BI 的 SSI（×2，方法改 BI，标黄）；有同键 BI 的 SSI 丢弃 ----
    ssi_only = ssi[~ssi['_in_bi']].copy()
    ssi_only[C.COL_METHOD] = C.METHOD_BI
    ssi_only[C.COL_VALUE] = ssi_only['_conv']
    ssi_only['_yellow'] = True

    ssi_matched = ssi[ssi['_in_bi']].copy()
    record_list = []
    for i in ssi_matched.index:
        record_list.append((i, 'SSI记录与同键BI记录重复（保留BI）'))
    for i, reason in record_list:
        del_rows.append((i, reason))

    # ---- Sheet2「BI+SSI(不重复)」：BI 行 + 并入的 SSI 行（黄） ----
    bi2 = bi.copy()
    bi2['_yellow'] = False
    sheet2 = pd.concat([bi2, ssi_only])[_BASE_SHEET_COLS + ['_yellow']]

    # ---- Sheet3「BI+SSI(重复)+取较大值处理」：被丢弃的同键 SSI 行（红）+ 追溯列 ----
    rows3 = []
    for k, grp in ssi_matched.groupby('_K'):
        bgrp = bi[bi['_K'] == k]
        max_bi = bgrp[C.COL_VALUE].max()
        best = grp.loc[grp['_conv'].idxmax()]
        row = best[_BASE_SHEET_COLS].copy()
        row['原BI值'] = max_bi
        row['原SSI值'] = best['_orig']
        row['转换后的SSI值'] = best['_conv']
        rows3.append(row)
    sheet3 = pd.DataFrame(rows3) if rows3 else \
        pd.DataFrame(columns=_BASE_SHEET_COLS + ['原BI值', '原SSI值', '转换后的SSI值'])
    sheet3['_deleted'] = True

    # ---- Sheet4「重复数据删除(BI)」：去重（被删行红） ----
    bi_d = _dedup_keep_max(bi, del_rows)
    ssi_d = _dedup_keep_max(ssi_only, del_rows)
    sheet4 = pd.concat([bi_d, ssi_d])[_BASE_SHEET_COLS + ['_deleted']]

    # ---- Sheet5「地址区分处理(BI)」：去重后行，被修改行标黄 ----
    kept = pd.concat([bi_d[~bi_d['_deleted']], ssi_d[~ssi_d['_deleted']]])
    kept = apply_address_distinction(kept)
    sheet5 = kept[_BASE_SHEET_COLS + ['_modified']]

    # ---- Sheet6「最终表(BI)」 ----
    bi_final = kept[['地市', '区县', '街道', '监测地点', C.COL_VALUE, C.COL_TYPE]].copy()
    bi_final['BI*'] = bi_final[C.COL_VALUE]
    bi_final['风险水平*'] = bi_final[C.COL_VALUE].map(C.grade_bi)
    bi_final['_city_idx'] = bi_final['地市'].map(C.CITY_INDEX)

    res = {}
    res['SSI表'] = (sheet1, None)
    res['BI+SSI(不重复)'] = (sheet2, '_yellow')
    res['BI+SSI(重复)+取较大值处理'] = (sheet3, '_deleted')
    res['重复数据删除(BI)'] = (sheet4, '_deleted')
    res['地址区分处理(BI)'] = (sheet5, '_modified')
    res['最终表(BI)'] = (bi_final[['地市', '区县', '街道', '监测地点', 'BI*', '风险水平*']], None)
    return bi_final, res


# ---------------- ADI 专项（金标准校准版） ----------------

def _adi_pipeline(adi_raw, del_rows):
    adi = adi_raw.copy()
    adi['_K'] = _make_keys(adi)

    # Sheet1「重复数据删除(ADI)」：全部行，被删行红
    adi = _dedup_keep_max(adi, del_rows)
    sheet1 = adi[_BASE_SHEET_COLS + ['_deleted']]

    # Sheet2「地址区分处理(ADI)」
    kept = adi[~adi['_deleted']].copy()
    kept = apply_address_distinction(kept)
    sheet2 = kept[_BASE_SHEET_COLS + ['_modified']]

    # Sheet3「最终表(ADI)」——指标列名用 ADI（与金标准/一览表一致）
    adi_final = kept[['地市', '区县', '街道', '监测地点', C.COL_VALUE, C.COL_TYPE]].copy()
    adi_final['ADI'] = adi_final[C.COL_VALUE]
    adi_final['风险水平*'] = adi_final[C.COL_VALUE].map(C.grade_adi)
    adi_final['_city_idx'] = adi_final['地市'].map(C.CITY_INDEX)

    res = {}
    res['重复数据删除(ADI)'] = (sheet1, '_deleted')
    res['地址区分处理(ADI)'] = (sheet2, '_modified')
    res['最终表(ADI)'] = (adi_final[['地市', '区县', '街道', '监测地点', 'ADI', '风险水平*']], None)
    return adi_final, res
