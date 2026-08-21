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
    """解析一行 '地市-区/县/市-街道/乡/镇' -> (地市, 区县, 街道)；无法解析返回 None"""
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

    m = _DISTRICT_RE.match(rest)
    if not m:
        return None
    district = m.group(1)
    street = rest[len(district):]

    # D18：街道以区县名开头则去掉该前缀（端州区特例一般化）
    if street.startswith(district):
        street = street[len(district):]

    if not street:
        return None
    return city, district, street


def district_display_sheet(district):
    """监测点汇总 Excel 显示：市辖区->/，去掉末尾 市/县/区 字"""
    if district == '市辖区':
        return '/'
    return district.rstrip('市县区')


def district_display_word(district):
    """一览表 Word 显示：市辖区->-（示例如此），去掉末尾 市/县/区 字"""
    if district == '市辖区':
        return '-'
    return district.rstrip('市县区')
