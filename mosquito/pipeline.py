# -*- coding: utf-8 -*-
"""数据处理流水线（需求文档 v3.2 + 金标准 BI-报出整理/ADI-报出整理校准）。

关键规则（已按金标准实测校准）：
- 数值保留源精度（不四舍五入到 1 位小数）。
- BI 表 = BI 方法行 + “无同键 BI”的 SSI 行（×2 换算并入）；有同键 BI 的 SSI 行不并入，
  但转换后 SSI 值 ≥ 5 时与原 BI 值取大者作为最终值（D55）。
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


def _is_new_exclude(exclude):
    """判断是否为新的排除字段格式（列表项含 loc/comm 两列）"""
    return isinstance(exclude, list) and bool(exclude) and any(
        ('loc' in t or 'comm' in t) for t in exclude)


def _exclude_matches_any(loc_str, comm_str, exclude):
    """排除匹配：
    - 新格式（列表项含 loc/comm）：每个排除字段 = 区/县/市-街道/乡/镇 与 社区/村居 两列；
      字段内两列都填时为“与”（两列均命中才删除），只填一列时按该列匹配，都不填则该字段无效；
      多个排除字段之间为“或”（任一字段命中即删除）。
    - 旧格式（field/connector）保持兼容。"""
    if _is_new_exclude(exclude):
        for f in exclude:
            loc_f = str(f.get('loc', '')).strip()
            comm_f = str(f.get('comm', '')).strip()
            if not loc_f and not comm_f:
                continue
            loc_ok = (not loc_f) or (loc_f in str(loc_str))
            comm_ok = (not comm_f) or (comm_f in str(comm_str))
            if loc_ok and comm_ok:
                return True
        return False
    return _exclude_matches(loc_str, exclude) or _exclude_matches(comm_str, exclude)


def _common_prefix(strs):
    """多个字符串的最长公共前缀"""
    if not strs:
        return ''
    p = str(strs[0])
    for s in strs[1:]:
        s = str(s)
        while not s.startswith(p):
            p = p[:-1]
            if not p:
                return ''
    return p


def exclude_display(exclude):
    """排除字段展示串（供 Word 注记）：
    - 新格式：每个字段 = 区/县/市-街道/乡/镇 + 社区/村居 拼接；多个字段提取所有字段均有的
      公共前缀放在最前（第一个字段保留），其后各字段去掉公共前缀，以、相隔。
      例：罗定市素龙街道平南村委、罗定市素龙街道龙岗花园、罗定市罗城街道区屋居委
          → 罗定市素龙街道平南村委、素龙街道龙岗花园、罗城街道区屋居委
    - 旧格式（field/connector）保持兼容。"""
    if not exclude:
        return None
    if isinstance(exclude, str):
        return exclude.strip() or None
    if _is_new_exclude(exclude):
        fields = []
        for f in exclude:
            s = (str(f.get('loc', '')).strip() + str(f.get('comm', '')).strip()).strip()
            if s:
                fields.append(s)
        if not fields:
            return None
        if len(fields) > 1:
            common = _common_prefix(fields)
            if common:
                fields = [fields[0]] + [f[len(common):] for f in fields[1:]]
        return '、'.join(fields)
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


def _addr_of_pref(col_primary, col_secondary):
    """地址取值函数：优先取 col_primary，为空则取 col_secondary；皆空返回 None"""
    def pref(r):
        a = r[col_primary]
        if a is None or (isinstance(a, float) and a != a) or str(a).strip() == '':
            a = r[col_secondary]
        if a is None or (isinstance(a, float) and a != a) or str(a).strip() == '':
            return None
        return str(a).strip()
    return pref


def _group_candidates(frame, idx, addr_pref):
    """按给定地址来源计算组内每行的候选片段（清理地址后切分）；返回 {index: segs|None}, maxlen"""
    rows = frame.loc[idx]
    cand = {}
    maxlen = 0
    for i, r in rows.iterrows():
        a = addr_pref(r)
        if a is None:
            cand[i] = None
            continue
        cleaned = _clean_addr(a, r[C.COL_LOC], r[C.COL_COMMUNITY], r[C.COL_TYPE])
        segs = [s for s in _ADDR_SEP.split(cleaned) if s]
        cand[i] = segs if segs else None
        if segs:
            maxlen = max(maxlen, len(segs))
    return cand, maxlen


def _extend_distinct(cand, maxlen):
    """原规则第 4-5 步：逐步追加下一片段直至组内候选互异；返回 {index: addr|None}"""
    cur = {}
    for n in range(1, maxlen + 1):
        cur = {}
        for i, segs in cand.items():
            cur[i] = ''.join(segs[:n]) if segs else None
        vals = [v for v in cur.values() if v]
        if len(vals) == len(set(vals)):
            break
    return cur


def apply_address_distinction(frame):
    """3.8 地址区分：组 = (地市-区/县/市-街道/乡/镇, 社区/村居, 防控区类型)。

    分别用“监测地址（地图定位版）优先”与“监测地址（手填）优先”两种来源按原规则
    计算最小区分地址，取字段较短者作为最终值（等长时取地图定位版）；仍冲突的追加
    （1）（2）… 序号。仅当组内存在 >=2 个不同地址时改名。返回新增 '_modified' 标记的副本。
    """
    frame = frame.copy()
    frame['_modified'] = False
    group_cols = [C.COL_LOC, C.COL_COMMUNITY, C.COL_TYPE]
    pref_map = _addr_of_pref(C.COL_ADDR1, C.COL_ADDR2)   # 地图定位版优先（原规则）
    pref_hand = _addr_of_pref(C.COL_ADDR2, C.COL_ADDR1)  # 手填优先

    for key, idx in frame.groupby(group_cols).groups.items():
        rows = frame.loc[idx]
        if len(set(a for a in rows.apply(pref_map, axis=1) if a)) < 2 \
                and len(set(a for a in rows.apply(pref_hand, axis=1) if a)) < 2:
            continue

        cur_map, maxlen1 = _group_candidates(frame, idx, pref_map)
        cur1 = _extend_distinct(cur_map, maxlen1)
        cur_hand, maxlen2 = _group_candidates(frame, idx, pref_hand)
        cur2 = _extend_distinct(cur_hand, maxlen2)

        # 取两来源中字段较短的最小区分地址（等长取地图定位版）
        chosen = {}
        for i in idx:
            a1 = cur1.get(i)
            a2 = cur2.get(i)
            if a1 is None and a2 is None:
                chosen[i] = None
            elif a1 is None:
                chosen[i] = a2
            elif a2 is None:
                chosen[i] = a1
            else:
                chosen[i] = a1 if len(a1) <= len(a2) else a2

        # 冲突处理：同值组内按“候选少者优先”为每条分配互异候选（取较短者）；
        # 任一记录无法分配到互异候选时，该组整组按冲突处理（追加序号，与 3.8 第 6 步一致）
        groups = defaultdict(list)
        for i, v in chosen.items():
            groups[v].append(i)
        final = {}
        for v, idxs in groups.items():
            if v is None or len(idxs) == 1:
                for i in idxs:
                    final[i] = v
                continue
            opts = {}
            for i in idxs:
                cands = []
                for x in (cur1.get(i), cur2.get(i)):
                    if x and x not in cands:
                        cands.append(x)
                opts[i] = sorted(cands, key=len)      # 短者优先
            used_in = set()
            tmp = {}
            ok = True
            for i in sorted(opts, key=lambda i: (len(opts[i]), i)):
                pick = next((x for x in opts[i] if x not in used_in), None)
                if pick is None:
                    ok = False
                    break
                tmp[i] = pick
                used_in.add(pick)
            if ok:
                final.update(tmp)
            else:
                for i in idxs:
                    final[i] = None                   # 整组冲突

        # 正常改名：监测地点 = {社区/村居}（{最小区分地址}）（{防控区类型}）（3.8 修订）
        for i in idx:
            comm = frame.at[i, C.COL_COMMUNITY]
            typ = frame.at[i, C.COL_TYPE]
            v = final.get(i)
            if v is None:
                continue
            while v.startswith(comm):
                v = v[len(comm):]
            new = '%s（%s）（%s）' % (comm, v, typ)
            if new != frame.at[i, '监测地点']:
                frame.loc[i, '监测地点'] = new
                frame.loc[i, '_modified'] = True
        # 冲突/无地址行：保持原监测地点，追加序号（（1）（2）…）
        none_idx = [i for i in idx if final.get(i) is None]
        if len(none_idx) > 1:
            for j, i in enumerate(none_idx, 1):
                new = '%s（%s）（%d）' % (frame.at[i, C.COL_COMMUNITY], frame.at[i, C.COL_TYPE], j)
                if new != frame.at[i, '监测地点']:
                    frame.loc[i, '监测地点'] = new
                    frame.loc[i, '_modified'] = True
    return frame


# ---------------- 流水线结果 ----------------

# 基础数据集输出列（Sheet1/Sheet2；不含 街道，街道仅内部用于排序）
_BASE_SHEET_COLS_FULL = ['地市', '区县', '街道', C.COL_LOC, C.COL_COMMUNITY, '监测地点',
                         C.COL_ADDR1, C.COL_ADDR2, C.COL_TIME, C.COL_METHOD, C.COL_VALUE]
_BASE_OUT_COLS = ['地市', C.COL_LOC, '区县', C.COL_COMMUNITY, '监测地点',
                  C.COL_ADDR1, C.COL_ADDR2, C.COL_TIME, C.COL_METHOD, C.COL_VALUE]


def _compute_first_monitor(source_df):
    """按 (地市-区/县/市-街道/乡/镇, 社区/村居, 监测方法) 计算每个监测点的“最初监测日期”：
    从 P1 处理前的数据中，取该点**对应监测方法**且**监测指标值不为空**的记录，
    在其监测日期序列中，取最近一次中断（相邻日期差 > 1 天）之后重新开始的日期；
    无中断则取最早监测日期。返回 {(loc, comm, method): date}。"""
    out = {}
    df = source_df[[C.COL_LOC, C.COL_COMMUNITY, C.COL_METHOD, C.COL_VALUE, C.COL_TIME]].copy()
    df[C.COL_TIME] = pd.to_datetime(df[C.COL_TIME], errors='coerce')
    df[C.COL_VALUE] = pd.to_numeric(df[C.COL_VALUE], errors='coerce')
    df = df.dropna(subset=[C.COL_TIME, C.COL_VALUE])     # 监测指标值不为空
    df[C.COL_LOC] = df[C.COL_LOC].fillna('').astype(str).str.strip()
    df[C.COL_COMMUNITY] = df[C.COL_COMMUNITY].fillna('').astype(str).str.strip()
    df[C.COL_METHOD] = df[C.COL_METHOD].fillna('').astype(str).str.strip()
    for (loc, comm, method), grp in df.groupby([C.COL_LOC, C.COL_COMMUNITY, C.COL_METHOD]):
        dates = sorted(grp[C.COL_TIME].dt.date.unique())
        if not dates:
            continue
        start = dates[0]
        for a, b in zip(dates, dates[1:]):
            if (b - a).days > 1:
                start = b
        out[(loc, comm, method)] = start
    return out


def _build_messy_frames(messy_rows):
    """把捕获的“乱码”记录（距末例天数>40000 且 距首例天数>40000）整理为 BI/ADI 两个 DataFrame：
    列 = 基础数据集列 + 最初监测日期、距输入日期天数、_dropped（供浅红标记，内部列）。"""
    bi_rows, adi_rows = [], []
    for m in messy_rows:
        parsed = parse_location(m['loc'])
        if parsed is None:
            continue
        city, district, street = parsed
        mp = '%s（%s）' % (str(m['comm']).strip(), str(m['type']).strip())
        rec = {
            '地市': city, '区县': district, '街道': street,
            C.COL_LOC: m['loc'], C.COL_COMMUNITY: m['comm'],
            '监测地点': mp, C.COL_ADDR1: m['addr1'], C.COL_ADDR2: m['addr2'],
            C.COL_TIME: m['time'], C.COL_METHOD: m['method'], C.COL_VALUE: m['value'],
            '最初监测日期': m['first_monitor'], '距输入日期天数': m['interval'],
            '_dropped': m['dropped'],
        }
        if m['method'] in C.BI_SSI_METHODS:
            bi_rows.append(rec)
        elif m['method'] == C.METHOD_ADI:
            adi_rows.append(rec)
    cols = _BASE_SHEET_COLS_FULL + ['最初监测日期', '距输入日期天数', '_dropped']
    bi = pd.DataFrame(bi_rows, columns=cols) if bi_rows else pd.DataFrame(columns=cols)
    adi = pd.DataFrame(adi_rows, columns=cols) if adi_rows else pd.DataFrame(columns=cols)
    return bi, adi


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
        self.base_bi = None           # 基础数据集（BI 部分）
        self.base_adi = None          # 基础数据集（ADI 部分）
        self.messy_bi = None          # 乱码处理（BI）：末例>40000且首例>40000 的记录
        self.messy_adi = None         # 乱码处理（ADI）


# ---------------- 主流程 ----------------

def run_pipeline(source_df, target_date, exclude=None, flight_df=None):
    validate_columns(source_df)

    df = source_df.copy()
    del_rows = []
    excluded_cities = set()
    # 该点最初监测日期映射（用于 P6 乱码规则：末例>40000 且 首例>40000 的记录）
    first_monitor = _compute_first_monitor(source_df)

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

    # ---- P6 距末例天数（保留 <=5 或 >40000；其中 >40000 的记录仅保留 距首例天数 <=5 或 >40000 的，D56） ----
    # 新增（乱码规则）：距末例天数>40000 且 距首例天数>40000 的记录，按 (地市-区/县/市-街道/乡/镇,
    # 社区/村居) 确定该点“最初监测日期”（最近一次中断后重新开始的日期），
    # 计算到输入日期的间隔天数：<=5 纳入、>5 放弃；这些记录另输出至“乱码处理”sheet。
    days = pd.to_numeric(df[C.COL_DAYS], errors='coerce')
    keep_days = days.notna() & (days <= C.DAYS_LOW)
    messy_rows = []          # 乱码记录（末例>40000 且 首例>40000），供 基础数据集 乱码 sheet
    if C.COL_FIRST_DAYS in df.columns:
        first_days = pd.to_numeric(df[C.COL_FIRST_DAYS], errors='coerce')
        mask_high = days > C.DAYS_HIGH
        first_le5 = first_days.notna() & (first_days <= C.DAYS_LOW)
        first_high = first_days.notna() & (first_days > C.DAYS_HIGH)
        messy_mask = mask_high & first_high              # 末例>40000 且 首例>40000
        keep_days = keep_days | (mask_high & first_le5)  # 首例<=5 保留
        messy_interval = pd.Series([None] * len(df), index=df.index, dtype=object)
        if messy_mask.any():
            locs = df.loc[messy_mask, C.COL_LOC].fillna('').astype(str).str.strip()
            comms = df.loc[messy_mask, C.COL_COMMUNITY].fillna('').astype(str).str.strip()
            methods = df.loc[messy_mask, C.COL_METHOD].fillna('').astype(str).str.strip()
            starts = [first_monitor.get((l, c, m)) for l, c, m in zip(locs, comms, methods)]
            messy_interval.loc[messy_mask] = [
                (target_date - st).days if st is not None else None for st in starts]
        interval_ok = pd.Series([False] * len(df), index=df.index)
        has_iv = messy_interval.notna()
        interval_ok[has_iv] = pd.to_numeric(messy_interval[has_iv]) <= C.DAYS_LOW
        keep_messy = messy_mask & interval_ok
        keep_days = keep_days | keep_messy
        # 捕获乱码记录（含被放弃的，供 乱码 sheet；放弃的记浅红）
        if messy_mask.any():
            for i in df.index[messy_mask]:
                r = df.loc[i]
                messy_rows.append({
                    'loc': r[C.COL_LOC], 'comm': r[C.COL_COMMUNITY], 'type': r[C.COL_TYPE],
                    'addr1': r[C.COL_ADDR1], 'addr2': r[C.COL_ADDR2],
                    'time': r[C.COL_TIME], 'method': r[C.COL_METHOD], 'value': r[C.COL_VALUE],
                    'first_monitor': first_monitor.get(
                        (str(r[C.COL_LOC]).strip(), str(r[C.COL_COMMUNITY]).strip(),
                         str(r[C.COL_METHOD]).strip())),
                    'interval': messy_interval.at[i],
                    'dropped': not bool(interval_ok.at[i]),
                })
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

    # 基础数据集拆分（Sheet1 BI基础数据集 / Sheet2 ADI基础数据集）与乱码处理表
    base_bi = base[base[C.COL_METHOD].isin(C.BI_SSI_METHODS)][_BASE_SHEET_COLS_FULL].copy()
    base_adi = base[base[C.COL_METHOD] == C.METHOD_ADI][_BASE_SHEET_COLS_FULL].copy()
    messy_bi, messy_adi = _build_messy_frames(messy_rows)

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
    res.base_bi = base_bi
    res.base_adi = base_adi
    res.messy_bi = messy_bi
    res.messy_adi = messy_adi
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
        record_list.append((i, 'SSI记录与同键BI记录重复'))
    for i, reason in record_list:
        del_rows.append((i, reason))

    # ---- 同键 SSI 处理（新规则）：SSI×2 < 5 → 不并入（保留原BI值）；
    #      SSI×2 >= 5 → 与原BI值取大者作为最终值，并提升该K的BI值 ----
    rows3 = []
    bi_raise = {}          # K -> 最终BI值（仅当 SSI×2 >= 5 时提升）
    for k, grp in ssi_matched.groupby('_K'):
        bgrp = bi[bi['_K'] == k]
        max_bi = bgrp[C.COL_VALUE].max()        # 原BI值（同键BI最大值，D8）
        best = grp.loc[grp['_conv'].idxmax()]   # 多条SSI取转换后最大值（D8）
        conv = best['_conv']
        if conv >= 5:
            final_val = max(max_bi, conv)
            bi_raise[k] = final_val
        else:
            final_val = max_bi
        row = best[_BASE_SHEET_COLS].copy()
        row['原BI值'] = max_bi
        row['原SSI值'] = best['_orig']
        row['转换后的SSI值'] = conv
        row['最终BI值'] = final_val
        rows3.append(row)
    # 提升同键 BI 的最终值（SSI×2 >= 5 且大于原BI值时）
    if bi_raise:
        for k, final_val in bi_raise.items():
            bgrp = bi[bi['_K'] == k]
            max_idx = bgrp[C.COL_VALUE].idxmax()
            bi.loc[max_idx, C.COL_VALUE] = final_val

    # ---- Sheet2「BI+SSI(不重复)」：BI 行（含提升后的最终值）+ 并入的 SSI 行（黄） ----
    bi2 = bi.copy()
    bi2['_yellow'] = False
    sheet2 = pd.concat([bi2, ssi_only])[_BASE_SHEET_COLS + ['_yellow']]

    # ---- Sheet3「BI与SSI重复处理」：被丢弃的同键 SSI 行 + 追溯列（含最终BI值） ----
    sheet3 = pd.DataFrame(rows3) if rows3 else \
        pd.DataFrame(columns=_BASE_SHEET_COLS + ['原BI值', '原SSI值', '转换后的SSI值', '最终BI值'])
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
    res['BI与SSI重复处理'] = (sheet3, None)   # 背景无色（D40）
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
