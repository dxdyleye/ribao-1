# -*- coding: utf-8 -*-
"""行政区划解析（3.2 节）。

真实数据格式（无分隔符，直接拼接）：
    广东省{地市}市{区县}{区/县/市}{街道}{街道/镇/乡}
例：
    广东省广州市白云区人和镇
    广东省东莞市市辖区黄江镇
    广东省肇庆市端州区端州区黄岗街道   （端州区在街道部分重复出现）
"""
import re

from .config import CITY_ORDER

_PROV = '广东省'
_DISTRICT_RE = re.compile(r'^(市辖区|.+?区|.+?县|.+?市)')


def _is_nan(v):
    return v is None or (isinstance(v, float) and v != v)


def parse_location(s):
    """解析一行 '地市-区/县/市-街道/乡/镇' -> (地市, 区县, 街道)；无法解析返回 None

    D64：主规则 `_DISTRICT_RE` 按“区→县→市”优先匹配以该字结尾的最短前缀，但当街道名
    含“区/县/市”字（如“……高新技术产业开发区”）时可能把整串误当区县、街道为空而失败；
    此时回退：取去掉地市后剩余串中**最短的“区/县/市”结尾前缀**作为区县，
    街道 = 剩余字段（D18 仍去除街道前重复的区县前缀）。
    """
    if _is_nan(s):
        return None
    s = str(s).strip()
    if not s.startswith(_PROV):
        return None
    rest = s[len(_PROV):]

    city = None
    for c in CITY_ORDER:
        if rest.startswith(c + '市'):
            city = c
            rest = rest[len(c) + 1:]
            break
    if city is None:
        return None

    def _split(district, street):
        """街道去重复的区县前缀（D18）后返回；街道为空则失败"""
        if street.startswith(district):
            street = street[len(district):]
        if not street:
            return None
        return city, district, street

    m = _DISTRICT_RE.match(rest)
    if m:
        r = _split(m.group(1), rest[len(m.group(1)):])
        if r is not None:
            return r

    # D64 回退：取最短“区/县/市”结尾前缀为区县，其余为街道
    if rest.startswith('市辖区'):
        return _split('市辖区', rest[len('市辖区'):])
    for i in range(1, len(rest) + 1):
        if rest[i - 1] in '区县市' and len(rest[:i]) >= 3:
            r = _split(rest[:i], rest[i:])
            if r is not None:
                return r
    return None


def district_display_sheet(district):
    """村居一览表显示：市辖区->-（东莞/中山，与参考一致），去掉末尾 市/县/区 字；
    例外（D68）：区县为“城区”（汕尾市城区等）时保留“城区”，不去掉“区”字。"""
    if district == '市辖区':
        return '-'
    if district == '城区':
        return '城区'
    return district.rstrip('市县区')
