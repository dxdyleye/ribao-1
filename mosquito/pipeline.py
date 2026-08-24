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


# ---------------- 排除字段（需求三：多个字段，和/或 连接） ----------------

def _exclude_matches(loc_str, exclude):
    """多字段排除匹配：单字符串 = 子串匹配；
    列表 = [{field, connector}]，connector('和'/'或')连接前一个字段：
    和 = 须同时包含前后两个字段才删除；或 = 包含其一即删除。从左到右求值。"""
    s = str(loc_str)
    if isinstance(exclude, str):
        f = exclude.strip()
        return bool(f) and f in s
    res = None
    for i, t in enumerate(exclude):
        f = str(t.get('field', '')).strip()
        cur = bool(f) and f in s
        if res is None:
            res = cur
        else:
            res = (res and cur) if t.get('connector') == '和' else (res or cur)
    return bool(res)


def _exclude_matches_any(loc_str, comm_str, exclude):
    """排除匹配扩展（D53）：对“地市-区/县/市-街道/乡/镇”或“社区/村居”两列分别匹配，
    任一列命中排除表达式即返回 True。"""
    return _exclude_matches(loc_str, exclude) or _exclude_matches(comm_str, exclude)


def exclude_display(exclude):
    """排除字段的展示串（供 Word 注记），如 “荔湾区” / “荔湾区或越秀区”；无则 None"""
    if not exclude:
        return None
    if isinstance(exclude, str):
        return exclude.strip() or None
    parts = []
    for i, t in enumerate(exclude):
        f = str(t.get('field', '')).strip()
        if not f:
            continue
        if i > 0:
            parts.append('和' if t.get('connector') == '和' else '或')
        parts.append(f)
    return ''.join(parts) or None


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

    仅当组内存在 >=2 个不同地址时，逐行以“社区/村居（最小区分地址）（类型）”改名；
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
        # 正常改名：监测地点 = {社区/村居}（{最小区分地址}）（{防控区类型}）（3.8 修订）
        for i, v in cur.items():
            if v is None:
                continue
            comm = frame.at[i, C.COL_COMMUNITY]
            typ = frame.at[i, C.COL_TYPE]
            while v.startswith(comm):
                v = v[len(comm):]
            new = '%s（%s）（%s）' % (comm, v, typ)
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
        self.exclude_display = None   # 排除字段展示串（Word 注记用）
        self.flight_bi_count = 0      # 飞行监测表整合入 BI 最终表的条数
        self.flight_adi_count = 0     # 飞行监测表整合入 ADI 最终表的条数


# ---------------- 主流程 ----------------

def run_pipeline(source_df, target_date, exclude=None, flight_df=None):
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

    # ---- P5 排除字段（支持多个字段，“和/或”连接；单字符串保持子串匹配。
    #      D53：对“地市-区/县/市-街道/乡/镇”与“社区/村居”两列分别匹配，任一列命中即删除） ----
    if exclude:
        loc_s = df[C.COL_LOC].astype(str)
        comm_s = df[C.COL_COMMUNITY].astype(str)
        mask_ex = pd.Series([_exclude_matches_any(a, b, exclude) for a, b in zip(loc_s, comm_s)],
                            index=df.index)
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

    # ---- 飞行监测表（需求一，可选）：参考 P1/P2/P3/P8 处理 + A2 去重，
    #      提取 ADI值/BI值（自动计算）分别整合入 ADI/BI 最终表 ----
    flight_bi_count = flight_adi_count = 0
    if flight_df is not None:
        bi_extra, adi_extra = process_flight_monitoring(flight_df, target_date)
        if bi_extra is not None and not bi_extra.empty:
            flight_bi_count = len(bi_extra)
            bi_final = pd.concat([bi_final, bi_extra], ignore_index=True)
            bi_sheets['最终表(BI)'] = (bi_final[['地市', '区县', '街道', '监测地点', 'BI*', '风险水平*']], None)
        if adi_extra is not None and not adi_extra.empty:
            flight_adi_count = len(adi_extra)
            adi_final = pd.concat([adi_final, adi_extra], ignore_index=True)
            adi_sheets['最终表(ADI)'] = (adi_final[['地市', '区县', '街道', '监测地点', 'ADI', '风险水平*']]
                                        .rename(columns={'ADI': 'ADI*'}), None)

    # ---- 删除数据情况说明（3.9 修订：列同计算过程表、仅输入日期数据、按删除原因排序、含监测日期） ----
    del_cols = ['地市', '区县', '街道', C.COL_LOC, C.COL_COMMUNITY,
                C.COL_ADDR1, C.COL_ADDR2, C.COL_TYPE, '监测地点',
                C.COL_METHOD, C.COL_VALUE]

    def _is_na(v):
        return v is None or (isinstance(v, float) and v != v)

    rows = []
    for i, reason in del_rows:
        if reason == '监测时间为空':
            continue                      # 仅说明输入日期的数据的处理情况（无日期行不计入）
        r = source_df.loc[i]
        loc = r.get(C.COL_LOC)
        parsed = parse_location(loc) if not _is_na(loc) and str(loc).strip() else None
        comm = r.get(C.COL_COMMUNITY)
        typ = r.get(C.COL_TYPE)
        mp = None
        if not _is_na(comm) and not _is_na(typ) and str(comm).strip() and str(typ).strip():
            mp = '%s（%s）' % (str(comm).strip(), str(typ).strip())
        rows.append({
            '地市': parsed[0] if parsed else None,
            '区县': parsed[1] if parsed else None,
            '街道': parsed[2] if parsed else None,
            C.COL_LOC: r.get(C.COL_LOC),
            C.COL_COMMUNITY: comm,
            C.COL_ADDR1: r.get(C.COL_ADDR1),
            C.COL_ADDR2: r.get(C.COL_ADDR2),
            C.COL_TYPE: typ,
            '监测地点': mp,
            C.COL_METHOD: r.get(C.COL_METHOD),
            C.COL_VALUE: r.get(C.COL_VALUE),
            '监测日期': target_date,
            '删除原因': reason,
        })
    deletions = pd.DataFrame(rows, columns=del_cols + ['监测日期', '删除原因']) if rows else \
        pd.DataFrame(columns=del_cols + ['监测日期', '删除原因'])
    deletions = deletions.sort_values('删除原因', kind='stable').reset_index(drop=True)

    res = PipelineResult()
    res.base = base
    res.bi_final = bi_final
    res.adi_final = adi_final
    res.deletions = deletions
    res.excluded_cities = excluded_cities
    res.exclude_display = exclude_display(exclude)
    res.flight_bi_count = flight_bi_count
    res.flight_adi_count = flight_adi_count
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
    res['BI+SSI(重复)+取较大值处理'] = (sheet3, None)   # 背景无色（D40）
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

    # Sheet3「最终表(ADI)」——指标列名用 ADI*（与金标准/一览表一致）
    adi_final = kept[['地市', '区县', '街道', '监测地点', C.COL_VALUE, C.COL_TYPE]].copy()
    adi_final['ADI'] = adi_final[C.COL_VALUE]
    adi_final['风险水平*'] = adi_final[C.COL_VALUE].map(C.grade_adi)
    adi_final['_city_idx'] = adi_final['地市'].map(C.CITY_INDEX)

    res = {}
    res['重复数据删除(ADI)'] = (sheet1, '_deleted')
    res['地址区分处理(ADI)'] = (sheet2, '_modified')
    res['最终表(ADI)'] = (adi_final[['地市', '区县', '街道', '监测地点', 'ADI', '风险水平*']]
                          .rename(columns={'ADI': 'ADI*'}), None)
    return adi_final, res


# ---------------- 飞行监测表专项处理（需求一） ----------------

def process_flight_monitoring(flight_df, target_date):
    """飞行监测表（可选输入，范表“(2026疫点飞行监测)飞行监测-0821”）：
    - 参考 3.1 P1/P2/P3/P8：筛选当天日期、删除监测时间为空、删除指标值为空，
      新增列 监测地点 = `{社区/村居}（{防控区类型}，飞行监测）`；
    - 参考 3.4 A2：按完整键 K 去重，每组保留指标值最大的行；
    - 提取 ADI值（自动计算）、BI值（自动计算）两列，分别生成与最终表列结构一致的
      (bi_extra, adi_extra)，供整合入 BI/ADI 最终表。

    返回 (bi_extra, adi_extra)；某流无数据或当天无数据返回 None。
    """
    required = [C.COL_TIME, C.COL_LOC, C.COL_COMMUNITY, C.COL_TYPE,
                C.COL_ADDR1, C.COL_ADDR2, C.FLIGHT_COL_BI, C.FLIGHT_COL_ADI]
    missing = [c for c in required if c not in flight_df.columns]
    if missing:
        raise ProcessingError('飞行监测表缺少必需列：' + '、'.join(missing))

    df = flight_df.copy()

    # P1 + P2：监测时间解析；删除监测时间为空；筛选当天日期
    dpart = pd.to_datetime(df[C.COL_TIME], errors='coerce').dt.date
    mask_time_ok = dpart.notna()
    df = df[mask_time_ok]
    dpart = dpart[mask_time_ok]
    df = df[dpart == target_date]
    if df.empty:
        return None, None

    # 行政区划解析（最终表需要 地市/区县/街道；解析失败行删除）
    parsed = df[C.COL_LOC].map(parse_location)
    mask_parsed = parsed.notna()
    df = df[mask_parsed]
    parsed = parsed[mask_parsed]
    df['地市'] = [p[0] for p in parsed]
    df['区县'] = [p[1] for p in parsed]
    df['街道'] = [p[2] for p in parsed]

    # P8 监测地点
    df['监测地点'] = (df[C.COL_COMMUNITY].astype(str)
                     + '（' + df[C.COL_TYPE].fillna('').astype(str) + '，飞行监测）')

    def stream(value_col):
        # P3：指标值为空删除；A2：完整键去重保留最大值
        s = df.copy()
        s['_val'] = pd.to_numeric(s[value_col], errors='coerce')
        s = s[s['_val'].notna()]
        if s.empty:
            return None
        s['_K'] = _make_keys(s)
        keep = s.groupby('_K')['_val'].idxmax()
        return s.loc[keep]

    def to_extra(s, value_col, grade):
        if s is None or s.empty:
            return None
        e = s[['地市', '区县', '街道', '监测地点', '_val', C.COL_TYPE]].copy()
        e[C.COL_VALUE] = e['_val']
        e[value_col] = e['_val']
        e['风险水平*'] = e['_val'].map(grade)
        e['_city_idx'] = e['地市'].map(C.CITY_INDEX)
        return e.drop(columns=['_val'])

    bi_s = stream(C.FLIGHT_COL_BI)
    adi_s = stream(C.FLIGHT_COL_ADI)
    return to_extra(bi_s, 'BI*', C.grade_bi), to_extra(adi_s, 'ADI', C.grade_adi)
