# -*- coding: utf-8 -*-
"""全局配置：列名、固定地市顺序、风险分级、颜色、输出文件命名。"""

# ---------- 必需列（10 列，列名完全匹配） ----------
COL_TIME = '监测时间（年/月/日）'
COL_VALUE = '监测指标值'
COL_LOC = '地市-区/县/市-街道/乡/镇'
COL_DAYS = '距末例天数（自动计算）'
COL_TYPE = '防控区类型'
COL_COMMUNITY = '社区/村居'
COL_METHOD = '监测方法（如BI/RI/MOI/ADI等）'
COL_ADDR1 = '监测地址（地图定位版）'
COL_ADDR2 = '监测地址（如“监测地址”定位字段不可用，可手填；如定位可用，不需要重复填写）'

REQUIRED_COLUMNS = [COL_TIME, COL_VALUE, COL_LOC, COL_DAYS, COL_TYPE,
                    COL_COMMUNITY, COL_METHOD, COL_ADDR1, COL_ADDR2]

# ---------- 监测方法 ----------
METHOD_BI = '布雷图指数BI'
METHOD_SSI = '标准间指数SSI'
METHOD_ADI = '成蚊密度指数法ADI'
BI_SSI_METHODS = (METHOD_BI, METHOD_SSI)
ALL_PROCESSED_METHODS = (METHOD_BI, METHOD_SSI, METHOD_ADI)

# ---------- 防控区类型 ----------
VALID_TYPES = ('核心区', '警戒区')

# ---------- 距末例天数 ----------
DAYS_LOW = 5          # 保留 距末例天数 <= 5
DAYS_HIGH = 40000     # 或 > 40000

# ---------- 21 地市固定顺序 ----------
CITY_ORDER = ['广州', '深圳', '珠海', '汕头', '佛山', '韶关', '河源', '梅州', '惠州',
              '汕尾', '东莞', '中山', '江门', '阳江', '湛江', '茂名', '肇庆', '清远',
              '潮州', '揭阳', '云浮']
CITY_INDEX = {c: i for i, c in enumerate(CITY_ORDER)}


# ---------- 风险分级（3.5 节） ----------
def grade_bi(x):
    """布雷图指数 BI：<5 安全；5-10 低；10-20 中；>=20 高"""
    if x < 5:
        return '安全'
    if x < 10:
        return '低风险'
    if x < 20:
        return '中风险'
    return '高风险'


def grade_adi(x):
    """成蚊密度 ADI：<=2 安全；2-5 低；5-10 中；>10 高"""
    if x <= 2:
        return '安全'
    if x <= 5:
        return '低风险'
    if x <= 10:
        return '中风险'
    return '高风险'


# ---------- 颜色（风险水平背景色，参照金标准样式表中的 Excel 标准色） ----------
FILL_SAFE = 'C6EFCE'   # 安全-绿
FILL_LOW = 'FFEB9C'    # 低风险-黄
FILL_MID = 'FFC000'    # 中风险-橘黄
FILL_HIGH = 'FFC7CE'   # 高风险-红（Excel 标准“差”样式色）
RISK_FILLS = {'安全': FILL_SAFE, '低风险': FILL_LOW, '中风险': FILL_MID, '高风险': FILL_HIGH}
FILL_YELLOW = 'FFFF00'  # 标记黄（SSI来源 / 被修改行）
FILL_RED = 'FF0000'     # 标记红（被删除行）

# ---------- 字体（所有输出文件：中文 仿宋_GB2312，英文 Times New Roman） ----------
FONT_CN = '仿宋_GB2312'
FONT_EN = 'Times New Roman'
SIZE_WUHAO = 10.5       # 五号（整合表）
SIZE_14 = 14            # 其他 sheet 字号（村居一览表）

# ---------- 输出文件命名（月日不补零，全角冒号；计算过程表两位补零） ----------
def daily_docx_name(year, month, day):
    return '全省媒介伊蚊传染病疫情蚊媒监测情况（%d月%d日20：00）.docx' % (month, day)


def summary_xlsx_name(year, month, day):
    return '村居一览表（%d月%d日20：00）.xlsx' % (month, day)


def calc_xlsx_name(year, month, day):
    return 'BI_ADI_计算过程_%02d月%02d日.xlsx' % (month, day)
