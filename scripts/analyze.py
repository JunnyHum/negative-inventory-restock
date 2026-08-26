#!/usr/bin/env python3
"""
负库存翻单到货分析
- 支持新旧两种仓库库存格式（自动识别）
- 新格式：系统库存导出_黄之俊_*.xlsx（2026-05-19起统一）
- 旧格式（已停用）：实体库存_黄之俊*.xlsx
"""
import openpyxl
import argparse
import json
import logging
import sys
from collections import defaultdict
from pathlib import Path
import re
import os
import glob
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta, date
import datetime as dt

CONFIG_PATH = Path(__file__).parent.parent / 'config.json'

DEFAULT_CONFIG = {
    'paths': {
        'warehouse_dir':       '~/Desktop/淘宝店铺/仓库商品库存表/',
        'biz_root_dir':        '~/Desktop/淘宝店铺/生意参谋数据/',
        'biz_dir':             '~/Desktop/淘宝店铺/生意参谋数据/淘宝网红店（手动）/',
        'restock_dir':         '~/Desktop/淘宝店铺/生产部补单表/',
        'output_dir':          '~/Desktop/淘宝店铺/归档/窗口期到货_负库存/',
        'product_table':       '~/Desktop/淘宝店铺/店铺商品表/SG网红店商品表.xlsx',
        'wechat_file_dir':     '/Users/junny/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_iqz0imhyx4m812_8da3/msg/file',
    },
    'file_patterns': {
        'warehouse':       '系统库存导出_',
        'warehouse_old':   '实体库存_黄之俊',
        'sg_restock':      'SG品牌进度表*.xlsx',
        'nba_restock':     'SG.NBA电商专供订单进度*.xlsx',
        'biz_advisor':     '*生意参谋*.xls*',
    },
    'shop': {
        'name': 'SG网红店',
        'biz_keywords': ['网红店', 'sg网红']
    }
}

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

SIZE_KEYS = ['XS', 'S', 'M', 'L', 'XL', '2XL', '均码']
SIZE_LAST_MAP = {'4': 'S', '5': 'M', '6': 'L', '7': 'XL', '8': '2XL', '0': '均码'}

# 本店铺未上架商品黑名单/排除款号（用户指定不纳入分析记录的非本店铺销售款式）
EXCLUDED_UNOFFICIAL_CODES = {'WD920217'}

# 人工指定的款式到货交期校准字典（针对散表中遗留历史日期或业务最新改期调整）
MANUAL_DELIVERY_OVERRIDES = {
    'WE133321': datetime(2026, 9, 2),
}

def toInt(v) -> int:
    """将单元格值安全转换为整数"""
    if v is None or v == '' or v == '-':
        return 0
    return int(str(v).replace(',', ''))

def inferYear(month: int) -> int:
    """根据当前月份推断日期所属年份，处理跨年场景"""
    now = datetime.now()
    currentMonth = now.month
    year = now.year
    # 当前11-12月看到1-2月 → 视为明年
    if currentMonth >= 11 and month <= 2:
        year += 1
    # 当前1-2月看到11-12月 → 视为去年
    elif currentMonth <= 2 and month >= 11:
        year -= 1
    return year

def load_config():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding='utf-8')
        logger.info("已生成默认配置: %s", CONFIG_PATH)
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = json.load(f)
    paths = config.get('paths', {})
    for key, path in paths.items():
        expanded = os.path.expanduser(path)
        paths[key] = expanded  # 始终展开 ~，不依赖目录是否已存在
    return config

def get_latest_file(directory, pattern='实体库存_黄之俊'):
    if not os.path.exists(directory):
        return None
    files = sorted(glob.glob(os.path.join(directory, '*.xlsx')), key=os.path.getmtime, reverse=True)
    for f in files:
        if pattern in os.path.basename(f):
            return f
    return None  # 不匹配则返回 None，不随意 fallback

def is_yellow_cell(cell):
    """检查单元格背景填充是否为黄色系（生产部在大货进度表中涂黄表示出清）"""
    if not cell or not hasattr(cell, 'fill') or not cell.fill or cell.fill.fill_type in [None, 'none']:
        return False
    fg = cell.fill.fgColor
    bg = cell.fill.bgColor
    rgb = None
    if fg and fg.rgb: rgb = str(fg.rgb)
    elif bg and bg.rgb: rgb = str(bg.rgb)
    if rgb:
        rgb_u = rgb.upper()
        if any(y in rgb_u for y in ['FFFF00', 'FFF2CC', 'FFE699', 'FFD966', 'FFFFCC']):
            return True
    return False

def parse_spec_new(spec):
    """
    新格式规格编码：货号(8位) + 颜色(2位) + 尺码(2位) = 12位
    例：NE2219014814 → 货号=NE221901, 颜色=48, 尺码=S
        NE2219017918 → 货号=NE221901, 颜色=79, 尺码=2XL
    返回: (货号, 颜色码, 尺码名)
    """
    if not spec or len(spec) < 12:
        return None, None, None
    code = spec[:8]
    color = spec[8:10]
    size_last = spec[11]
    size_name = SIZE_LAST_MAP.get(size_last, None)
    return code, color, size_name

def parse_size_name_from_text(size_text):
    """从 S(165/78A) 或 S 或 均码 等文本提取尺码名"""
    if not size_text:
        return None
    s = str(size_text).strip().upper()
    # 均码
    if '均码' in s or '均' in s:
        return '均码'
    # 匹配 S(165/78A) 等格式
    m = re.match(r'^([A-Z]+)', s)
    if m:
        key = m.group(1)
        if key in ['XS', 'S', 'M', 'L', 'XL', '2XL', 'XXL']:
            return key if key != 'XXL' else '2XL'
    return None

def excel_serial_to_date(serial):
    if serial is None: return None
    if isinstance(serial, datetime): return serial
    try:
        serial = int(float(serial))
        if serial >= 60: serial -= 1
        return datetime(1899, 12, 31) + timedelta(days=serial)
    except (ValueError, TypeError, OverflowError):
        return None

def get_date_value(val):
    # 支持 M/D-M/D (如 "6/6-6/8") 和 M/D (如 "6/8")
    if isinstance(val, str):
        s = val.strip()
        # M/D-M/D 格式 (如 "6/6-6/8")
        m1 = re.match(r'^(\d+)/(\d+)-(\d+)/(\d+)$', s)
        if m1:
            month2 = int(m1.group(3))
            day2 = int(m1.group(4))
            year = inferYear(month2)
            return datetime(year, month2, day2)  # 取最后一天
        # M/D-D 格式 (如 "4/15-16"，同月)
        m = re.match(r'^(\d+)/(\d+)-(\d+)$', s)
        if m:
            month = int(m.group(1))
            year = inferYear(month)
            return datetime(year, month, int(m.group(3)))  # 取最后一天
        # M.D-M.D 格式 (如 "3.31-4.10"，跨月带点)
        m = re.match(r'^(\d+)\.(\d+)-(\d+)\.(\d+)$', s)
        if m:
            month2 = int(m.group(3))
            day2 = int(m.group(4))
            year = inferYear(month2)
            return datetime(year, month2, day2)  # 取最后一天
        # 单日 M/D 格式 (如 "6/8")
        m = re.match(r'^(\d+)/(\d+)$', s)
        if m:
            month = int(m.group(1))
            year = inferYear(month)
            return datetime(year, month, int(m.group(2)))
        # 尝试标准年-月-日或月-日等格式解析
        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%m-%d', '%m/%d']:
            try:
                dt_val = datetime.strptime(s, fmt)
                if fmt in ['%m-%d', '%m/%d']:
                    dt_val = dt_val.replace(year=inferYear(dt_val.month))
                return dt_val
            except ValueError:
                pass
        return None

    if val is None: return None
    if isinstance(val, datetime): return val
    if isinstance(val, (int, float)):
        return excel_serial_to_date(val)
    s = str(val).strip()
    if not s or s in ['-', 'None', '']: return None
    # 对于已经是标准的年-月-日字符串，直接尝试解析
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%m-%d', '%m/%d']:
        try:
            dt_val = datetime.strptime(s, fmt)
            if fmt in ['%m-%d', '%m/%d']:
                dt_val = dt_val.replace(year=inferYear(dt_val.month))
            return dt_val
        except ValueError:
            pass
    return None

def split_row_by_newline(rd, indexes):
    """
    根据 plan_arrival/actual_arrival 以及 total_qty 列的换行符拆分行数据
    返回列表 [rd1, rd2, ...]
    """
    del_idx = indexes['plan_arrival'] if indexes['plan_arrival'] is not None else indexes['actual_arrival']
    qty_idx = indexes['total_qty']
    
    del_text = str(rd[del_idx]).strip() if del_idx is not None and len(rd) > del_idx and rd[del_idx] is not None else ''
    qty_text = str(rd[qty_idx]).strip() if qty_idx is not None and len(rd) > qty_idx and rd[qty_idx] is not None else ''
    
    # 只要日期或下单量含有换行符，且它们并不是简单相同的重复，就进行分裂
    if '\n' in del_text or '\n' in qty_text:
        del_parts = [p.strip() for p in del_text.split('\n') if p.strip()]
        qty_parts = [p.strip() for p in qty_text.split('\n') if p.strip()]
        
        # 补全列表长度使两者一致
        max_parts = max(len(del_parts), len(qty_parts))
        sub_rds = []
        for i in range(max_parts):
            d_part = del_parts[i] if i < len(del_parts) else (del_parts[-1] if del_parts else '')
            q_part = qty_parts[i] if i < len(qty_parts) else (qty_parts[-1] if qty_parts else '')
            
            new_rd = list(rd)
            if del_idx is not None and len(new_rd) > del_idx:
                new_rd[del_idx] = d_part
            if qty_idx is not None and len(new_rd) > qty_idx:
                new_rd[qty_idx] = q_part
            sub_rds.append(new_rd)
        return sub_rds
    return [rd]

# ── 库存加载（新格式）───────────────────────────────────────────────────────
def findWarehouseColumnIndexes(headerCells: list) -> dict:
    """仓库库存导出表的动态列映射"""
    indexes = {
        'productCode': None,  # 商品编码（货号）
        'specCode': None,     # 规格编码
        'available': None,    # 可销数
        'stock': None,        # 库存数（实体）
        'sizeName': None,     # 尺码
        'colorName': None,    # 颜色
        'virtualWhCode': None, # 虚拟仓编码
        'virtualWhName': None, # 虚拟仓名称
    }
    for idx, cell in enumerate(headerCells):
        if cell is None:
            continue
        val = str(cell).strip()
        if val == '商品编码':
            indexes['productCode'] = idx
        elif val == '规格编码':
            indexes['specCode'] = idx
        elif val in ['可销数', '可销']:
            indexes['available'] = idx
        elif val in ['库存数', '实体库存', '库存']:
            indexes['stock'] = idx
        elif val in ['尺码', '规格']:
            indexes['sizeName'] = idx
        elif val == '颜色':
            indexes['colorName'] = idx
        elif val in ['虚拟仓编码', '虚拟仓']:
            indexes['virtualWhCode'] = idx
        elif val in ['虚拟仓名称', '虚拟仓名']:
            indexes['virtualWhName'] = idx
    return indexes

def load_warehouse(filepath):
    """
    新格式：系统库存导出_黄之俊_*.xlsx
    支持动态列识别，自动适配商品编码、规格编码、可销数、库存数、尺码、颜色列。
    规格编码：货号8位+颜色2位+尺码2位（共12位）
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb.active
    
    # 读取第一行作为表头
    first_row = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col_indexes = findWarehouseColumnIndexes(first_row)
    
    # 必须的核心列
    if col_indexes['productCode'] is None or col_indexes['specCode'] is None:
        wb.close()
        raise ValueError(f"仓库库存表 {filepath} 未识别到关键表头 (商品编码 / 规格编码)")
        
    logger.info("新格式库存文件: %s (%d行), 动态列配置: %s", 
                os.path.basename(filepath), ws.max_row, col_indexes)

    wh_neg_total  = defaultdict(float)   # skc -> 总可销
    wh_neg_size   = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS})  # skc -> {尺码: 可销}
    whEntityTotal = defaultdict(float)   # skc -> 总实体
    wh_colors_txt = {}   # skc -> 颜色文字（[48]藏青）
    wh_tw011_avail = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS}) # skc -> {尺码: 可销} (独享仓)
    wh_tw011_stock = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS}) # skc -> {尺码: 实体} (独享仓)
    wh_actual_sizes = defaultdict(set)   # code -> set of sizes seen

    # 尺码、颜色、可销、库存的列索引
    code_idx = col_indexes['productCode']
    spec_idx = col_indexes['specCode']
    avail_idx = col_indexes['available'] if col_indexes['available'] is not None else 10 # 默认10 (11列)
    stock_idx = col_indexes['stock'] if col_indexes['stock'] is not None else 11 # 默认11 (12列)
    size_idx = col_indexes['sizeName'] if col_indexes['sizeName'] is not None else 16 # 默认16 (17列)
    color_idx = col_indexes['colorName'] if col_indexes['colorName'] is not None else 17 # 默认17 (18列)
    wh_code_idx = col_indexes['virtualWhCode']
    wh_name_idx = col_indexes['virtualWhName']

    for row in ws.iter_rows(min_row=2):
        rd = [c.value for c in row]
        if len(rd) <= max(code_idx, spec_idx):
            continue
            
        code     = str(rd[code_idx]).strip() if rd[code_idx] else ''
        spec     = str(rd[spec_idx]).strip() if rd[spec_idx] else ''
        
        kexiao   = float(rd[avail_idx]) if len(rd) > avail_idx and rd[avail_idx] is not None else 0.0
        entity   = float(rd[stock_idx]) if len(rd) > stock_idx and rd[stock_idx] is not None else 0.0
        size_txt = rd[size_idx] if len(rd) > size_idx else None
        color_txt = rd[color_idx] if len(rd) > color_idx else None

        wh_code_val = str(rd[wh_code_idx]).strip() if wh_code_idx is not None and len(rd) > wh_code_idx and rd[wh_code_idx] is not None else ''
        wh_name_val = str(rd[wh_name_idx]).strip() if wh_name_idx is not None and len(rd) > wh_name_idx and rd[wh_name_idx] is not None else ''

        # 排除退货仓数据
        is_return_wh = wh_code_val.startswith('TW201') or ('退货' in wh_name_val)
        if is_return_wh:
            continue

        if not code or len(spec) < 12:
            continue

        skc_code, color_code, size_name = parse_spec_new(spec)
        if not skc_code:
            continue

        skc = skc_code + color_code  # 10位 SKC

        # 优先用尺码文字，其次用规格编码解析
        if size_txt:
            parsed_size = parse_size_name_from_text(str(size_txt))
            if parsed_size:
                size_name = parsed_size

        if code and size_name:
            wh_actual_sizes[code].add(size_name)

        if color_txt:
            wh_colors_txt[skc] = str(color_txt)

        wh_neg_total[skc]  += kexiao
        whEntityTotal[skc]  += entity
        if size_name:
            wh_neg_size[skc][size_name] += kexiao

        is_tw011 = (wh_code_val == 'TW011-SHARE-B') and ('S传统电商独享' in wh_name_val)
        if is_tw011 and size_name:
            wh_tw011_avail[skc][size_name] += kexiao
            wh_tw011_stock[skc][size_name] += entity

    wb.close()
    logger.info("新格式解析: %d个SKC", len(wh_neg_total))
    return wh_neg_total, wh_neg_size, whEntityTotal, wh_colors_txt, wh_tw011_avail, wh_tw011_stock, wh_actual_sizes

# ── 店铺商品表（全 Sheet 动态识别）───────────────────────────────────────────
def load_product_table(filepath):
    """扫描商品表文件中的所有 Sheet，动态定位款号、上架日期与商品状态，返回 {款号: {'上架日期': listDate, '商品状态': statusVal}}"""
    if not filepath or not os.path.exists(filepath):
        return {}
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        result = {}
        for sh in wb.sheetnames:
            ws = wb[sh]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            
            # 动态查找表头
            header_idx = -1
            code_col = None
            date_col = None
            status_col = None
            
            for r_i, r in enumerate(rows[:10]):
                if not r: continue
                r_str = [str(x).strip() if x is not None else '' for x in r]
                for c_i, val in enumerate(r_str):
                    if val in ['商品编码', '款号', '商家编码']:
                        code_col = c_i
                    elif val in ['上架日期', '上架时间']:
                        date_col = c_i
                    elif val in ['商品状态', '状态', '在售状态']:
                        status_col = c_i
                if code_col is not None:
                    header_idx = r_i
                    break
            
            if code_col is None:
                continue
                
            for row in rows[header_idx + 1:]:
                if not row or len(row) <= code_col:
                    continue
                productCode = str(row[code_col]).strip() if row[code_col] else ''
                if not productCode or productCode == 'None':
                    continue
                
                listDate = row[date_col] if (date_col is not None and len(row) > date_col) else None
                statusVal = str(row[status_col]).strip() if (status_col is not None and len(row) > status_col and row[status_col]) else ''
                
                if productCode not in result or (listDate and not result[productCode].get('上架日期')):
                    result[productCode] = {
                        '上架日期': listDate,
                        '商品状态': statusVal,
                    }
        wb.close()
        logger.info("店铺商品表全表匹配款号: %d个", len(result))
        return result
    except Exception as e:
        logger.warning("加载店铺商品表失败: %s", e)
        return {}

# ── 生意参谋 ─────────────────────────────────────────────────────────────────
def resolveBizDir(config: dict) -> str:
    """
    根据 config 中的店铺关键词，在生意参谋根目录下自动定位对应店铺的数据子目录。
    优先使用 biz_root_dir + 关键词匹配，若匹配失败则回退到 biz_dir。
    """
    paths = config.get('paths', {})
    rootDir = paths.get('biz_root_dir', '')
    fallbackDir = paths.get('biz_dir', '')
    
    if not rootDir or not os.path.isdir(rootDir):
        return fallbackDir
    
    shopConfig = config.get('shop', {})
    keywords = shopConfig.get('biz_keywords', [])
    
    if not keywords:
        return fallbackDir
    
    # 扫描子目录（包括符号链接）
    candidates = []
    try:
        for entry in os.listdir(rootDir):
            entryPath = os.path.join(rootDir, entry)
            if not os.path.isdir(entryPath):
                continue
            entryLower = entry.lower()
            for kw in keywords:
                if kw.lower() in entryLower:
                    candidates.append(entryPath)
                    break
    except Exception as e:
        logger.warning("扫描生意参谋根目录失败: %s", e)
        return fallbackDir
    
    if not candidates:
        logger.warning("生意参谋目录中未匹配到店铺关键词 %s，回退使用 %s", keywords, fallbackDir)
        return fallbackDir
    
    # 多个候选时，优先选择含有生意参谋数据文件的目录
    for c in candidates:
        # 检查目录本身或其子目录(如 7日商品数据/) 中是否有匹配文件
        if glob.glob(os.path.join(c, '**', '*生意参谋*'), recursive=True):
            logger.info("自动识别生意参谋目录: %s", c)
            return c
            
    logger.info("自动识别生意参谋目录(无数据文件): %s", candidates[0])
    return candidates[0]

def load_business_data(biz_dir, pattern):
    """
    在指定的生意参谋目录下加载最新文件，支持在子目录（如 7日商品数据/ 等）下递归查找
    """
    # 递归查找匹配的文件
    search_path = os.path.join(biz_dir, '**', pattern)
    files = sorted(glob.glob(search_path, recursive=True), key=os.path.getmtime, reverse=True)
    if not files:
        logger.info("在 %s 下未找到生意参谋数据", biz_dir)
        return {}
        
    logger.info("加载最新生意参谋文件: %s", files[0])
    try:
        prod_data = defaultdict(lambda: {'visitors': 0, 'pv': 0, 'cart': 0, 'orders': 0, 'pay': 0})
        
        is_xls = False
        try:
            # 1. 尝试用 openpyxl 打开 (xlsx)
            wb = openpyxl.load_workbook(files[0], data_only=True, read_only=True)
            ws = wb.active
            
            # 自动定位表头行
            header_row_idx = None
            headers = []
            row_idx = 0
            for row in ws.iter_rows(values_only=True):
                row_idx += 1
                if row_idx > 15: break
                row_str = [str(x).strip() for x in row if x is not None]
                if '货号' in row_str:
                    header_row_idx = row_idx
                    headers = [str(x).strip() if x is not None else '' for x in row]
                    break
                    
            if header_row_idx is not None:
                row_idx = 0
                for row in ws.iter_rows(values_only=True):
                    row_idx += 1
                    if row_idx <= header_row_idx:
                        continue
                    rd = dict(zip(headers, row))
                    code = str(rd.get('货号', '')).strip()
                    if not code or code == 'None' or code == '-': continue
                    prod_data[code]['visitors'] += toInt(rd.get('商品访客数'))
                    prod_data[code]['pv']       += toInt(rd.get('商品浏览量'))
                    prod_data[code]['cart']     += toInt(rd.get('商品加购件数'))
                    prod_data[code]['orders']   += toInt(rd.get('下单件数'))
                    prod_data[code]['pay']      += toInt(rd.get('支付件数'))
            wb.close()
        except Exception as ex:
            logger.info("openpyxl 打开生意参谋失败（可能是 xls 格式），将尝试 xlrd 打开。错误: %s", ex)
            is_xls = True
            
        # 2. 降级使用 xlrd 打开 (xls)
        if is_xls:
            import xlrd
            wb = xlrd.open_workbook(files[0])
            ws = wb.sheet_by_index(0)
            
            header_row_idx = None
            headers = []
            for r in range(min(15, ws.nrows)):
                row_vals = ws.row_values(r)
                row_str = [str(x).strip() for x in row_vals if x is not None]
                if '货号' in row_str:
                    header_row_idx = r
                    headers = [str(x).strip() if x is not None else '' for x in row_vals]
                    break
                    
            if header_row_idx is not None:
                for r in range(header_row_idx + 1, ws.nrows):
                    row_vals = ws.row_values(r)
                    rd = dict(zip(headers, row_vals))
                    code = str(rd.get('货号', '')).strip()
                    if not code or code == 'None' or code == '-': continue
                    prod_data[code]['visitors'] += toInt(rd.get('商品访客数'))
                    prod_data[code]['pv']       += toInt(rd.get('商品浏览量'))
                    prod_data[code]['cart']     += toInt(rd.get('商品加购件数'))
                    prod_data[code]['orders']   += toInt(rd.get('下单件数'))
                    prod_data[code]['pay']      += toInt(rd.get('支付件数'))
                    
        logger.info("生意参谋款号: %d", len(prod_data))
        return prod_data
    except Exception as e:
        logger.error("生意参谋读取失败: %s", e)
        return {}

def score_and_advise(unmatched_skcs, biz_data):
    """对无翻单负库存进行评分和建议"""
    codes = list(set(skc[:8] for skc in unmatched_skcs.keys()))
    if not biz_data:
        return {code: {'score': None, 'pay': 0, 'visitors': 0, 'cart': 0, 'advice': '⚪ 无数据'}
                for code in codes}
    max_v = max((biz_data[c]['visitors'] for c in codes if c in biz_data and biz_data[c]['visitors']), default=1)
    max_p = max((biz_data[c]['pay'] for c in codes if c in biz_data and biz_data[c]['pay']), default=1)
    max_c = max((biz_data[c]['cart'] for c in codes if c in biz_data and biz_data[c]['cart']), default=1)
    results = {}
    for code in codes:
        d = biz_data.get(code, {})
        visitors = d.get('visitors', 0)
        pay      = d.get('pay', 0)
        cart     = d.get('cart', 0)
        if visitors == 0 and pay == 0:
            advice = '❌ 不建议（无验证）'
            score = 0
        else:
            sv = visitors / max_v * 100 if max_v else 0
            sp = pay      / max_p * 100 if max_p else 0
            sc = cart     / max_c * 100 if max_c else 0
            score = sv * 0.3 + sp * 0.5 + sc * 0.2
            if   score > 50 and pay >= 3:  advice = '🔥 建议翻单'
            elif score >= 30 and pay >= 1: advice = '⚠️ 考虑翻单'
            elif pay == 0:                 advice = '❌ 不建议（0销量）'
            else:                          advice = '⚠️ 考虑翻单'
        results[code] = {'score': round(score, 1), 'pay': pay, 'visitors': visitors,
                         'cart': cart, 'advice': advice}
    return results

def has_skc_stock(skc, sizes, is_junma):
    """判断某个 SKC 在某天是否有库存（仅以总仓可销 > 0 为准）"""
    if is_junma:
        return (sizes.get('均码', 0) > 0)
    else:
        for s in ['S', 'M', 'L', 'XL', '2XL']:
            if (sizes.get(s, 0) > 0):
                return True
        return False

def findColumnIndexes(headerCells: list) -> dict:
    """
    根据表头单元格值列表，自动匹配各个关键字段的列索引（0-based）
    """
    indexes = {
        'code': None,            # 款号
        'color': None,           # 颜色
        'factory': None,         # 工厂
        'shipped': None,         # 总出货 / 已出货列
        'actual_arrival': None,  # 优先：实际到仓时间 / 实际出货时间 / 到仓时间列
        'plan_arrival': None,    # 备选：出货时间 / 货期列
        'status': None,          # 备注 / 生产状态列
        'cleared': None,         # 是否清完 / 结清列
        'total_qty': None,       # 总下单列
        'sizes': {}              # 尺码名 -> 列索引
    }
    
    # 尺码名与列名匹配正则
    size_patterns = {
        'XS': [r'^xs$'],
        'S': [r'^s$', r'^s/'],
        'M': [r'^m$', r'^m/'],
        'L': [r'^l$', r'^l/'],
        'XL': [r'^xl$', r'^xl/'],
        '2XL': [r'^xxl$', r'^xxl/', r'^2xl$', r'^2xl/'],
        '均码': [r'^均码$', r'^f/均码$', r'^f$']
    }

    for idx, cell in enumerate(headerCells):
        if cell is None:
            continue
        val = str(cell).strip().lower().replace('\n', '')
        
        if val == '款号':
            indexes['code'] = idx
        elif val == '颜色':
            indexes['color'] = idx
        elif val in ['条码', '规格编码', '条形码', 'sku条码']:
            # NOTE: 条码列（12位 SKC）可作为颜色编号提取的备选来源，第 8~10 位即颜色编号
            indexes['spec'] = idx
        elif val == '工厂':
            indexes['factory'] = idx
        elif val in ['出货数量', '总出货', '已出货']:
            indexes['shipped'] = idx
        elif val in ['实际到仓时间', '实际出货时间', '实际\n出货时间', '到仓时间']:
            indexes['actual_arrival'] = idx
        elif val in ['出货时间', '货期', '到货时间', '到货日期', '交货期']:
            indexes['plan_arrival'] = idx
        elif val in ['备注', '生产状态']:
            indexes['status'] = idx
        elif val in ['是否清完', '结清', '清完', '是否清']:
            indexes['cleared'] = idx
        elif val in ['总下单', '总下单数', '下单总量', '翻单数量', '翻单数', '翻下单数']:
            indexes['total_qty'] = idx
            
        for size_name, patterns in size_patterns.items():
            for pat in patterns:
                if re.match(pat, val):
                    indexes['sizes'][size_name] = idx
                    break
                    
    return indexes

def get_row_values_safely(ws, r_num):
    try:
        return [cell.value for cell in next(ws.iter_rows(min_row=r_num, max_row=r_num))]
    except StopIteration:
        return []

def getHeaderRowAndIndexes(ws) -> tuple[int, dict]:
    """
    在工作表中定位表头行并返回 (1-based行号, 列索引字典)
    """
    for r in range(1, 15):
        row_cells = get_row_values_safely(ws, r)
        if not any(row_cells):
            continue
        row_str = [str(x).strip() for x in row_cells if x is not None]
        if any('款号' in s for s in row_str) and any('颜色' in s for s in row_str):
            indexes = findColumnIndexes(row_cells)
            return r, indexes
            
    first_row = get_row_values_safely(ws, 1)
    return 1, findColumnIndexes(first_row)

# ── 色系分组（2026-06-04 完整版:从商品资料.基础信息 提取 11 系 100 色）──────────
# 不同色系是不同 SKU,不能互相替代。源数据:
#   ~/Desktop/店铺对账单/SPRAYGROUND商品资料.xlsx → Sheet: 基础信息
#   col 20: 2位色号, col 21: 颜色名, col 24: 色系分类
# 严格使用 SG 官方色系定义,未收录色号返回 None (保守拒绝近似匹配)。
COLOR_FAMILY = {
    # 白色系 (00-04): 透明/银/米白/本白/奶白
    '00': '白色系', '01': '白色系', '02': '白色系', '03': '白色系', '04': '白色系',
    # 杏卡色系 (05-09): 灰白/杏/卡其/沙/暖烟灰
    '05': '杏卡色系', '06': '杏卡色系', '07': '杏卡色系', '08': '杏卡色系', '09': '杏卡色系',
    # 黄色系 (10-19): 杏白/金/黄/橘粉/芥黄/柠檬黄/中黄/米色/暗金/土黄
    '10': '黄色系', '11': '黄色系', '12': '黄色系', '13': '黄色系', '14': '黄色系',
    '15': '黄色系', '16': '黄色系', '17': '黄色系', '18': '黄色系', '19': '黄色系',
    # 灰色系 (20-29): 银灰/浅灰/灰/中灰/深灰/浅紫灰/深紫灰/浅花灰/花灰/深花灰
    '20': '灰色系', '21': '灰色系', '22': '灰色系', '23': '灰色系', '24': '灰色系',
    '25': '灰色系', '26': '灰色系', '27': '灰色系', '28': '灰色系', '29': '灰色系',
    # 绿色系 (30-39): 浅水绿/湖绿/浅绿/绿/草绿/灰绿/深绿/军绿/茶绿/墨绿
    '30': '绿色系', '31': '绿色系', '32': '绿色系', '33': '绿色系', '34': '绿色系',
    '35': '绿色系', '36': '绿色系', '37': '绿色系', '38': '绿色系', '39': '绿色系',
    # 蓝色系 (40-49): 浅水蓝/湖蓝/深湖蓝/浅蓝/中蓝/灰蓝/宝蓝/深蓝/藏青/牛仔蓝
    '40': '蓝色系', '41': '蓝色系', '42': '蓝色系', '43': '蓝色系', '44': '蓝色系',
    '45': '蓝色系', '46': '蓝色系', '47': '蓝色系', '48': '蓝色系', '49': '蓝色系',
    # 紫色系 (50-59): 浅紫/粉紫/紫/蓝紫/深紫/丁香紫/黑紫/亮紫/茄紫/紫红
    '50': '紫色系', '51': '紫色系', '52': '紫色系', '53': '紫色系', '54': '紫色系',
    '55': '紫色系', '56': '紫色系', '57': '紫色系', '58': '紫色系', '59': '紫色系',
    # 红色系 (60-69): 粉红/桃红/大红/藕粉红/暗红/粉金/橙/橙红/酒红/玫红
    '60': '红色系', '61': '红色系', '62': '红色系', '63': '红色系', '64': '红色系',
    '65': '红色系', '66': '红色系', '67': '红色系', '68': '红色系', '69': '红色系',
    # 咖黑色系 (70-79): 浅咖/灰褐/古铜/深咖/棕/深褐/红褐/黑金/瓦黑/黑
    '70': '咖黑色系', '71': '咖黑色系', '72': '咖黑色系', '73': '咖黑色系', '74': '咖黑色系',
    '75': '咖黑色系', '76': '咖黑色系', '77': '咖黑色系', '78': '咖黑色系', '79': '咖黑色系',
    # 反光色系 (80-89): 反光银/金/绿/蓝/紫/红/黑 + 反光一二三
    '80': '反光色系', '81': '反光色系', '82': '反光色系', '83': '反光色系', '84': '反光色系',
    '85': '反光色系', '86': '反光色系', '87': '反光色系', '88': '反光色系', '89': '反光色系',
    # 格子系 (90-99): 豹纹/蛇纹/花色/格子绿/蓝/紫/红/黑/格子多色1/2
    '90': '格子系', '91': '格子系', '92': '格子系', '93': '格子系', '94': '格子系',
    '95': '格子系', '96': '格子系', '97': '格子系', '98': '格子系', '99': '格子系',
}


def _colorFamily(colorCode):
    """返回色系名。未知色号返回 None (拒绝近似匹配)。"""
    # 去掉可能的 [xx] 前缀
    if colorCode and colorCode.startswith('['):
        return None  # 异常色号格式,保守跳过
    return COLOR_FAMILY.get(colorCode)


def get_color_category_name(color_txt, skc=''):
    """
    根据颜色文本（如 '[03]本白' 或 '本白'）及 SKC 动态解析归纳颜色分类（色系）。
    专供客服与运营按色系（白色系/黑色系/蓝色系等）快速筛选。
    """
    if not color_txt:
        color_txt = ''
        
    # 1. 尝试从 [XX] 或 SKC 后2位提取数字编号映射 COLOR_FAMILY
    m = re.search(r'\[(\d+)\]', str(color_txt))
    num = m.group(1) if m else (skc[8:10] if len(skc) >= 10 and skc[8:10].isdigit() else None)
    if num and num in COLOR_FAMILY:
        return COLOR_FAMILY[num]
        
    # 2. 文本关键词智能归类
    txt = str(color_txt).lower()
    if any(k in txt for k in ['白', '米白', '本白', '象牙']):
        return '白色系'
    elif any(k in txt for k in ['黑', '瓦黑', '墨黑']):
        return '黑色系'
    elif any(k in txt for k in ['蓝', '牛仔', '藏青', '宝蓝', '蔚蓝']):
        return '蓝色系'
    elif any(k in txt for k in ['灰', '花灰', '碳灰', '银']):
        return '灰色系'
    elif any(k in txt for k in ['红', '粉', '桃', '玫', '橙']):
        return '红色系'
    elif any(k in txt for k in ['黄', '杏', '卡其', '驼', '燕麦']):
        return '黄色系'
    elif any(k in txt for k in ['绿', '薄荷', '军绿', '橄榄']):
        return '绿色系'
    elif any(k in txt for k in ['咖', '棕', '栗', '茶']):
        return '棕/咖色系'
    elif any(k in txt for k in ['紫', '薰衣草']):
        return '紫色系'
        
    return '其他色系'


def findApproxColorMatches(negSkcs, allRestock, skcHasAny):
    """
    对无精确翻单匹配的负库存 SKC,检查同款号下是否有同色系其他颜色的翻单记录。
    只有同色系才认为可近似(2026-06-04 Junny 修复)。
    """
    approxMatches = []
    
    for skc in negSkcs:
        if skc in skcHasAny:
            continue
            
        code = skc[:8]
        origColor = skc[8:]
        origFamily = _colorFamily(origColor)
        # 原色号不在色系映射内 → 不参与近似匹配
        if origFamily is None:
            continue
        
        # 收集同款号下所有有翻单记录的颜色
        sameCodeRestockColors = set()
        for otherSkc in skcHasAny:
            if otherSkc[:8] == code:
                sameCodeRestockColors.add(otherSkc[8:])
        
        # 只保留同色系的色号
        sameFamilyColors = {c for c in sameCodeRestockColors if _colorFamily(c) == origFamily}
        if not sameFamilyColors:
            continue
            
        for altColor in sameFamilyColors:
            altSkc = code + altColor
            for (rSkc, actual_date), d in allRestock.items():
                if rSkc == altSkc:
                    matchType = '★色号差异(同色系):原[{}]→参考[{}]'.format(origColor, altColor)
                    approxMatches.append({
                        'origSkc': skc,
                        'altSkc': altSkc,
                        'delivery': actual_date,
                        'matchType': matchType,
                        'restock_detail': d
                    })
    
    return approxMatches


def parse_individual_restock_file(filepath):
    """
    解析独立的行格式翻单表（例如 SG26年电商款翻单XXXX.xlsx）。
    表格式为一行一个 SKU（12位规格），包含尺码和颜色列，并使用 Forward Fill 自动向下填充日期。
    """
    filename = os.path.basename(filepath)
    if 'hoodrich' in filename.lower():
        logger.info("  [硬性过滤] 忽略 HOODRICH 品牌翻单文件: %s", filename)
        return []
    logger.info("解析独立翻单表: %s", filename)
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    records = []
    
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        
        # 1. 寻找表头行
        header_row = None
        for r in range(1, 15):
            row_iter = ws.iter_rows(min_row=r, max_row=r)
            first_item = next(row_iter, None)
            if not first_item:
                continue
            row_cells = [cell.value for cell in first_item]
            if not any(row_cells):
                continue
            row_str = [str(x).strip() for x in row_cells if x is not None]
            if any('款号' in s for s in row_str) and any('颜色' in s for s in row_str):
                header_row = r
                break
                
        if not header_row:
            continue
            
        header_item = next(ws.iter_rows(min_row=header_row, max_row=header_row), None)
        if not header_item:
            continue
        header = [cell.value for cell in header_item]
        
        # 2. 定位关键列
        indexes = {
            'code': None, 'spec': None, 'color': None, 'size': None,
            'qty': None, 'delivery': None, 'factory': None, 'status': None,
        }
        
        for idx, cell in enumerate(header):
            if cell is None: continue
            val = str(cell).strip().lower().replace('\n', '')
            if val == '款号':
                indexes['code'] = idx
            elif val in ['条码', '规格编码']:
                indexes['spec'] = idx
            elif val == '颜色':
                indexes['color'] = idx
            elif val == '尺码':
                indexes['size'] = idx
            elif val in ['下单', '下单数', '翻单数', '翻单数量']:
                indexes['qty'] = idx
            elif val in ['到货时间', '货期', '交货期']:
                indexes['delivery'] = idx
            elif val == '工厂':
                indexes['factory'] = idx
            elif val in ['生产部', '状态', '生产状态']:
                indexes['status'] = idx
            elif val == '备注':
                if indexes['status'] is None:
                    indexes['status'] = idx

        # 兜底匹配 qty 列
        if indexes['qty'] is None:
            for idx, cell in enumerate(header):
                if cell is None: continue
                val = str(cell).strip().lower().replace('\n', '')
                if '下单' in val:
                    indexes['qty'] = idx
                    break
                    
        if indexes['code'] is None or indexes['color'] is None or indexes['size'] is None:
            continue
            
        last_pc = None
        last_delivery = None
        last_factory = ''
        last_status = ''
        
        # 3. 遍历行
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            rd = list(row)
            if len(rd) <= max(indexes['code'], indexes['color'], indexes['size']):
                continue
                
            pc = str(rd[indexes['code']]).strip() if rd[indexes['code']] else ''
            if not pc or pc == 'None' or pc in EXCLUDED_UNOFFICIAL_CODES:
                continue
                
            color = str(rd[indexes['color']]) if rd[indexes['color']] else ''
            
            # NOTE: 双重验证颜色编号 —— 两路来源互为兜底，防止遗漏
            # 来源1: 颜色字段中的 [XX] 格式编号（如 [03]本白 -> '03'）
            # 来源2: 条码列（12位规格码 = 款号8位+颜色2位+尺码2位）第8~10位
            cc_candidates = set()
            mc = re.search(r'\[(\d+)\]', color)
            if mc:
                cc_candidates.add(mc.group(1))
            if indexes.get('spec') is not None and len(rd) > indexes['spec']:
                spec_raw = str(rd[indexes['spec']]).strip() if rd[indexes['spec']] else ''
                if len(spec_raw) == 12 and spec_raw[:8] == pc:
                    cc_candidates.add(spec_raw[8:10])
            # 两路都失败时保留 '?' 占位，便于后续排查
            if not cc_candidates:
                cc_candidates.add('?')
            
            size_raw = rd[indexes['size']]
            size_name = parse_size_name_from_text(size_raw)
            if not size_name:
                continue
                
            delivery_date = None
            if indexes['delivery'] is not None and len(rd) > indexes['delivery']:
                delivery_date = get_date_value(rd[indexes['delivery']])
                
            factory = ''
            if indexes['factory'] is not None and len(rd) > indexes['factory']:
                factory = str(rd[indexes['factory']]) if rd[indexes['factory']] else ''
                
            status = ''
            if indexes['status'] is not None and len(rd) > indexes['status']:
                status = str(rd[indexes['status']]) if rd[indexes['status']] else ''
                
            qty_val = 0
            if indexes['qty'] is not None and len(rd) > indexes['qty']:
                v = rd[indexes['qty']]
                if isinstance(v, (int, float)):
                    qty_val = float(v)
                    
            # Forward Fill 向下填充 (全局向下填充以支持跨货号合并单元格的交期、工厂、状态正常传递)
            if delivery_date is None:
                delivery_date = last_delivery
            else:
                last_delivery = delivery_date

            raw_deliv_orig = delivery_date

            # 业务指定/人工交期校准覆盖 (处理散表中遗留的历史旧交期或业务调整)
            if pc in MANUAL_DELIVERY_OVERRIDES:
                delivery_date = MANUAL_DELIVERY_OVERRIDES[pc]
                last_delivery = delivery_date
                
            if not factory:
                factory = last_factory
            else:
                last_factory = factory
                
            if not status:
                status = last_status
            else:
                last_status = status
                
            last_pc = pc
                
            if qty_val <= 0:
                continue
            
            # 对每个颜色编号来源各注册一条翻单记录（去重由外层 gk_sku_restock 字典保证）
            for cc in cc_candidates:
                skc = pc + cc
                records.append({
                    'skc': skc,
                    'size': size_name,
                    'qty': qty_val,
                    'delivery': delivery_date,
                    'raw_delivery': raw_deliv_orig,
                    'factory': factory,
                    'status': status
                })
            
    wb.close()
    return records


# ── 主分析 ────────────────────────────────────────────────────────────────────
def analyze():
    config = load_config()
    paths  = config['paths']
    pats   = config['file_patterns']

    # 仓库文件：系统库存导出_黄之俊*.xlsx
    # 优先报表目录，其次飞书附件目录
    wh_file = (
        get_latest_file(paths.get('warehouse_dir') or '', pats['warehouse']) or
        get_latest_file(paths.get('feishu_inbound') or '', pats['warehouse'])
    )
    # 检测最新库存文件是否是今天的
    need_download_wh = False
    today_date_str = datetime.now().strftime('%Y%m%d')
    if not wh_file:
        logger.info("未找到任何仓库库存文件，判定需要下载今天最新数据")
        need_download_wh = True
    else:
        # 检测文件名中是否包含今天的日期
        if today_date_str not in os.path.basename(wh_file):
            logger.info("最新仓库文件 %s 不是今天的，准备自动拉取今天最新数据", os.path.basename(wh_file))
            need_download_wh = True

    if need_download_wh:
        logger.info("⚡ 触发自动拉取今天最新 Jeoms 仓库库存数据流程...")
        try:
            # 找到 jeoms-inventory 技能脚本的路径
            jeoms_script_path = Path(__file__).parent.parent.parent / 'jeoms-inventory' / 'main.py'
            if jeoms_script_path.exists():
                import subprocess, shutil
                python_bin = shutil.which('python3') or sys.executable
                cmd = [python_bin, str(jeoms_script_path)]
                logger.info("运行命令: %s", " ".join(cmd))
                res = subprocess.run(cmd, capture_output=True, text=True)
                if res.returncode == 0:
                    logger.info("✓ 自动拉取今天最新库存数据成功")
                    # 重新获取最新的库存文件路径
                    wh_file = (
                        get_latest_file(paths.get('warehouse_dir') or '', pats['warehouse']) or
                        get_latest_file(paths.get('feishu_inbound') or '', pats['warehouse'])
                    )
                else:
                    logger.error("❌ 自动拉取今日库存数据失败: %s\n%s", res.stdout, res.stderr)
            else:
                logger.error("未找到 jeoms-inventory 下载脚本: %s", jeoms_script_path)
        except Exception as dl_err:
            logger.error("自动触发下载仓库库存数据异常: %s", dl_err)

    if not wh_file:
        raise FileNotFoundError("未找到仓库库存文件，请确认数据源路径")
    logger.info("仓库文件: %s", wh_file)

    # 优先从库存文件名中提取分析日期以避免物理服务器与业务模拟时空不一致的 Bug
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    m_wh = re.search(r'_(\d{8})\d*', os.path.basename(wh_file))
    if m_wh:
        try:
            today = datetime.strptime(m_wh.group(1), "%Y%m%d").replace(hour=0, minute=0, second=0, microsecond=0)
            logger.info("🎯 自动从库存文件名解析出目标运行日期: %s", today.strftime('%Y-%m-%d'))
        except ValueError:
            pass

    # 动态窗口期：今天前10天 ~ 今天后15天
    window_start = today - timedelta(days=10)
    window_end   = today + timedelta(days=15)
    logger.info("窗口期: %s ~ %s", window_start.strftime('%m/%d'), window_end.strftime('%m/%d'))

    def is_older_than_last_week(source_name, base_date=today):
        """
        判断该翻单表是否来源于上周或更早（>7天前发起）
        """
        if not source_name:
            return False
        bname = os.path.basename(str(source_name))
        m = re.search(r'(\d{2})(\d{2})', bname)
        if m:
            try:
                mm, dd = int(m.group(1)), int(m.group(2))
                if 1 <= mm <= 12 and 1 <= dd <= 31:
                    f_dt = datetime(base_date.year, mm, dd)
                    return f_dt < (base_date - timedelta(days=7))
            except Exception:
                pass
        return False

    # ── 1. 加载当前仓库 ──────────────────────────────────────────────
    wh_neg_total, wh_neg_size, whEntityTotal, wh_colors_txt, wh_tw011_avail, wh_tw011_stock, wh_actual_sizes = load_warehouse(wh_file)

    # 判断均码款
    product_is_junma = {}
    for skc in wh_neg_total:
        code = skc[:8]
        if code not in product_is_junma:
            has_junma = '均码' in wh_actual_sizes.get(code, set())
            product_is_junma[code] = has_junma

    # 筛选可销<10 of SKC
    neg_skcs = {}
    for skc, qty in wh_neg_total.items():
        code    = skc[:8]
        if code in EXCLUDED_UNOFFICIAL_CODES or skc in EXCLUDED_UNOFFICIAL_CODES:
            continue
        is_junma = product_is_junma.get(code, False)
        sizes   = wh_neg_size.get(skc, {})
        if is_junma:
            if sizes.get('均码', 0) < 10:
                neg_skcs[skc] = qty
        else:
            has_neg_size = False
            for s in ['S', 'M', 'L', 'XL', '2XL']:
                if sizes.get(s, 0) < 10:
                    has_neg_size = True
                    break
            if has_neg_size:
                neg_skcs[skc] = qty
    logger.info("可销<10的SKC: %d", len(neg_skcs))

    # ── 2. 加载前一天的仓库（用于Sheet2库存回补）────────────────────
    all_dirs = [
        (paths.get('warehouse_dir') or ''),
        (paths.get('feishu_inbound') or ''),
    ]
    all_wh_files = []
    for d in all_dirs:
        if d:
            all_wh_files += sorted(glob.glob(os.path.join(d, '*.xlsx')), key=os.path.getmtime, reverse=True)

    # 过滤出所有合法的系统库存文件
    valid_wh_files = []
    for f in all_wh_files:
        f_basename = os.path.basename(f)
        if '系统库存导出' in f_basename and not f_basename.startswith('~$') and not f_basename.startswith('.~'):
            if f not in valid_wh_files:
                valid_wh_files.append(f)

    # 寻找当前 wh_file 降序排列中的下一个文件
    prev_wh_file = None
    try:
        cur_idx = valid_wh_files.index(wh_file)
        if cur_idx + 1 < len(valid_wh_files):
            prev_wh_file = valid_wh_files[cur_idx + 1]
    except ValueError:
        # 降级：使用修改时间仅次于当前 wh_file 的前一个有效文件
        cur_mtime = os.path.getmtime(wh_file)
        for f in valid_wh_files:
            if os.path.getmtime(f) < cur_mtime and f != wh_file:
                prev_wh_file = f
                break

    prev_neg_total  = defaultdict(float)
    prev_neg_size   = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS})
    prev_tw011_avail = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS})
    prev_tw011_stock = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS})
    if prev_wh_file:
        logger.info("前日仓库文件: %s", prev_wh_file)
        prev_neg_total, prev_neg_size, _, _, prev_tw011_avail, prev_tw011_stock, _ = load_warehouse(prev_wh_file)

    product_table = load_product_table(paths.get('product_table'))
    today_date = datetime.now().date()



    # ── 3. 加载翻单记录 ──────────────────────────────────────────────
    restock_dir = paths['restock_dir']
    sg_files = sorted(glob.glob(os.path.join(restock_dir, pats['sg_restock'])), key=os.path.getmtime, reverse=True)
    nba_files = sorted(glob.glob(os.path.join(restock_dir, pats['nba_restock'])), key=os.path.getmtime, reverse=True)
    

    if not sg_files:
        raise FileNotFoundError(f"未找到SG品牌进度表")
    nba_file = nba_files[0] if nba_files else None

    all_restock = {}  # {(skc, date): {...}}
    skc_has_any = set()
    master_progress_skcs = set() # 存放所有在大货进度表中出现的 SKC 集合 (含历史已结清)
    master_progress_codes = set() # 存放所有在大货进度表中出现的款号集合
    master_active_skcs = set() # 仅存放在大货进度表中【在途/未结清】的 SKC 集合
    master_active_code_colors = set() # 仅存放在大货进度表中【在途/未结清】的 (款号, 颜色/颜色编号)
    master_records_db = [] # 存放两张大货总表解析出的全量四维翻单条目数据库 [{'code', 'cc', 'color', 'date', 'qty'}]

    # SG品牌进度表 (多文件与多 Sheet 自适应与日期回退)
    for sg_file in sg_files:
        wb1 = openpyxl.load_workbook(sg_file, data_only=True, read_only=True)
        for sheet in wb1.sheetnames:
            ws1 = wb1[sheet]
            header_num, sg_indexes = getHeaderRowAndIndexes(ws1)
            
            # 必须找到关键列，否则视为无效 Sheet 并自动跳过
            if sg_indexes['code'] is None or sg_indexes['color'] is None:
                continue
                
            logger.info("解析 SG 品牌进度表 Sheet: %s", sheet)
            for row_cells in ws1.iter_rows(min_row=header_num + 1):
                rd_raw = [c.value for c in row_cells]
                row_yellow_clean = any(is_yellow_cell(c) for c in row_cells[:15])
                if len(rd_raw) <= max(sg_indexes['code'], sg_indexes['color']):
                    continue
                pc = str(rd_raw[sg_indexes['code']]).strip() if rd_raw[sg_indexes['code']] else ''
                if not pc or pc in EXCLUDED_UNOFFICIAL_CODES:
                    continue
                # 颜色编号提取双源兼容：
                # 来源1: 颜色字段 [XX] 正则
                # 来源2: 条码列第 8~10 位（一个条码只代表一个尺码，可代表整款颜色）
                color_raw = str(rd_raw[sg_indexes['color']]) if rd_raw[sg_indexes['color']] else ''
                mc_raw = re.search(r'\[(\d+)\]', color_raw)
                cc_raw = mc_raw.group(1) if mc_raw else None
                if not cc_raw:
                    spec_idx_sg = sg_indexes.get('spec')
                    if spec_idx_sg is not None and len(rd_raw) > spec_idx_sg:
                        spec_raw_sg = str(rd_raw[spec_idx_sg]).strip() if rd_raw[spec_idx_sg] else ''
                        if len(spec_raw_sg) >= 10 and spec_raw_sg[8:10].isdigit():
                            cc_raw = spec_raw_sg[8:10]
                if cc_raw:
                    master_progress_skcs.add(pc + cc_raw)
                    
                # 全量款号登记 (供新品/铺货审计使用)
                master_progress_codes.add(pc)
                    
                # 自适应换行拆分
                sub_rds = split_row_by_newline(rd_raw, sg_indexes)
                for rd in sub_rds:
                    # 过滤已出货 (精细化判定：已清完，或黄色背景填充，或已出货量 >= 总下单量)
                    shipped_idx = sg_indexes['shipped']
                    cleared_idx = sg_indexes.get('cleared')
                    total_qty_idx = sg_indexes.get('total_qty')
                    
                    is_shipped_clean = False
                    
                    # 1. 优先通过 "是否清完" 列进行结清打标判定
                    if cleared_idx is not None and len(rd) > cleared_idx:
                        cleared_val = str(rd[cleared_idx]).strip() if rd[cleared_idx] is not None else ''
                        if any(x in cleared_val for x in ['清', '是', '已清', '完']):
                            is_shipped_clean = True
                    
                    # 2. 生产部视觉打标：单元格黄色背景填充表示出清
                    if not is_shipped_clean and row_yellow_clean:
                        is_shipped_clean = True
                    
                    # 2. 备选：必须同时存在总下单量且出货量 >= 总下单量，才判定为结清
                    if not is_shipped_clean and shipped_idx is not None and len(rd) > shipped_idx:
                        shipped = rd[shipped_idx]
                        if shipped is not None and shipped != '':
                            try:
                                shipped_val = float(shipped)
                                if total_qty_idx is not None and len(rd) > total_qty_idx and rd[total_qty_idx] is not None:
                                    try:
                                        total_qty_val = float(rd[total_qty_idx])
                                        if total_qty_val > 0 and shipped_val >= total_qty_val:
                                            is_shipped_clean = True
                                    except ValueError:
                                        pass
                            except ValueError:
                                pass
                                
                    color_clean = re.sub(r'\[\d+\]', '', color_raw).strip()
                    # 关键逻辑：在途/未结清的记录纳入 master_active_skcs 用于四维对账
                    if not is_shipped_clean:
                        if cc_raw:
                            master_active_skcs.add(pc + cc_raw)
                        if color_clean:
                            master_active_code_colors.add((pc, color_clean))
                            if cc_raw:
                                master_active_code_colors.add((pc, cc_raw))
                                
                    color = str(rd[sg_indexes['color']]) if rd[sg_indexes['color']] else ''
                    mc = re.search(r'\[(\d+)\]', color)
                    cc = mc.group(1) if mc else '?'
                    skc = pc + cc
                    
                    # 实际到仓时间 / 实际出货时间 提取（带回退）
                    actual_date = None
                    act_idx = sg_indexes['actual_arrival']
                    if act_idx is not None and len(rd) > act_idx:
                        actual_date = get_date_value(rd[act_idx])
                    if actual_date is None:
                        plan_idx = sg_indexes['plan_arrival']
                        if plan_idx is not None and len(rd) > plan_idx:
                            actual_date = get_date_value(rd[plan_idx])
                            
                    key = (skc, actual_date)
                    if key not in all_restock:
                        all_restock[key] = {'sizes': {k: 0 for k in SIZE_KEYS}, 'color': color, 'status': '', 'factory': ''}
                        
                    # 提取各个尺码数量并准备重新分摊
                    row_sizes = {}
                    row_sizes_sum = 0
                    for sz, idx in sg_indexes['sizes'].items():
                        val = 0
                        if idx is not None and len(rd) > idx:
                            raw_val = rd[idx]
                            if isinstance(raw_val, (int, float)):
                                val = max(0.0, float(raw_val))
                        row_sizes[sz] = val
                        row_sizes_sum += val
                        
                    # 提取下单总量
                    total_qty_val = 0
                    if total_qty_idx is not None and len(rd) > total_qty_idx and rd[total_qty_idx] is not None:
                        try:
                            total_qty_val = max(0.0, float(rd[total_qty_idx]))
                        except ValueError:
                            pass
                            
                    if total_qty_val > 0:
                        if abs(row_sizes_sum - total_qty_val) > 1.0:
                            if row_sizes_sum > 0:
                                # 按各码现有的比例分摊
                                temp_sum = 0
                                for sz in row_sizes:
                                    ratio = row_sizes[sz] / row_sizes_sum
                                    row_sizes[sz] = round(total_qty_val * ratio)
                                    temp_sum += row_sizes[sz]
                                diff = total_qty_val - temp_sum
                                if diff != 0 and row_sizes:
                                    max_sz = max(row_sizes, key=row_sizes.get)
                                    row_sizes[max_sz] += diff
                            else:
                                # 平摊
                                valid_sizes = [s for s in ['S', 'M', 'L', 'XL', '2XL'] if s in row_sizes]
                                if valid_sizes:
                                    base_qty = total_qty_val // len(valid_sizes)
                                    rem = total_qty_val % len(valid_sizes)
                                    for s in valid_sizes:
                                        row_sizes[s] = base_qty
                                    row_sizes[valid_sizes[0]] += rem
                                    
                    for sz, val in row_sizes.items():
                        all_restock[key]['sizes'][sz] += val
                        
                    status_idx = sg_indexes['status']
                    orig_status = str(rd[status_idx])[:30] if (status_idx is not None and len(rd) > status_idx and rd[status_idx]) else ''
                    if is_shipped_clean:
                        all_restock[key]['status'] = '已出清' if not orig_status else f"已出清({orig_status})"
                        all_restock[key]['is_cleared'] = True
                    elif orig_status:
                        all_restock[key]['status'] = orig_status
                        
                    factory_idx = sg_indexes['factory']
                    if factory_idx is not None and len(rd) > factory_idx and rd[factory_idx]:
                        all_restock[key]['factory'] = str(rd[factory_idx])
                        
                    if skc: 
                        skc_has_any.add(skc)
                        master_progress_skcs.add(skc)
                        
                    # 注册四维对账条目至数据库
                    master_records_db.append({
                        'code': pc,
                        'cc': cc_raw,
                        'color': color_clean,
                        'date': actual_date,
                        'qty': total_qty_val,
                        'is_cleared': is_shipped_clean
                    })
        wb1.close()

    # NBA翻单表 (多 Sheet 自适应与日期回退)
    if nba_file and os.path.exists(nba_file):
        wb2 = openpyxl.load_workbook(nba_file, data_only=True, read_only=True)
        for sheet in wb2.sheetnames:
            ws2 = wb2[sheet]
            header_num, nba_indexes = getHeaderRowAndIndexes(ws2)
            
            if nba_indexes['code'] is None or nba_indexes['color'] is None:
                continue
                
            logger.info("解析 NBA 翻单表 Sheet: %s", sheet)
            for row_cells in ws2.iter_rows(min_row=header_num + 1):
                rd = [c.value for c in row_cells]
                row_yellow_clean_nba = any(is_yellow_cell(c) for c in row_cells[:15])
                if len(rd) <= max(nba_indexes['code'], nba_indexes['color']):
                    continue
                pc = str(rd[nba_indexes['code']]).strip() if rd[nba_indexes['code']] else ''
                if not pc or pc in EXCLUDED_UNOFFICIAL_CODES:
                    continue
                # 颜色编号提取双源兼容：
                # 来源1: 颜色字段 [XX] 正则
                # 来源2: 条码列第 8~10 位（尺码无关，只代表整款颜色）
                color_raw2 = str(rd[nba_indexes['color']]) if rd[nba_indexes['color']] else ''
                mc_raw2 = re.search(r'\[(\d+)\]', color_raw2)
                cc_raw2 = mc_raw2.group(1) if mc_raw2 else None
                if not cc_raw2:
                    spec_idx_nba = nba_indexes.get('spec')
                    if spec_idx_nba is not None and len(rd) > spec_idx_nba:
                        spec_raw_nba = str(rd[spec_idx_nba]).strip() if rd[spec_idx_nba] else ''
                        if len(spec_raw_nba) >= 10 and spec_raw_nba[8:10].isdigit():
                            cc_raw2 = spec_raw_nba[8:10]
                if cc_raw2:
                    master_progress_skcs.add(pc + cc_raw2)
                    
                master_progress_codes.add(pc)
                    
                # 过滤已出货 (精细化判定：已清完，或黄色背景，或已出货量 >= 总下单量)
                shipped_idx = nba_indexes['shipped']
                cleared_idx = nba_indexes.get('cleared')
                total_qty_idx = nba_indexes.get('total_qty')
                
                is_shipped_clean = False
                
                # 1. 优先通过 "是否清完" 列进行结清打标判定
                if cleared_idx is not None and len(rd) > cleared_idx:
                    cleared_val = str(rd[cleared_idx]).strip() if rd[cleared_idx] is not None else ''
                    if any(x in cleared_val for x in ['清', '是', '已清', '完']):
                        is_shipped_clean = True
                
                # 2. 生产部视觉打标：单元格黄色背景填充表示出清
                if not is_shipped_clean and row_yellow_clean_nba:
                    is_shipped_clean = True
                
                # 2. 备选：必须同时存在总下单量且出货量 >= 总下单量，才判定为结清
                if not is_shipped_clean and shipped_idx is not None and len(rd) > shipped_idx:
                    shipped = rd[shipped_idx]
                    if shipped is not None and shipped != '':
                        try:
                            shipped_val = float(shipped)
                            if total_qty_idx is not None and len(rd) > total_qty_idx and rd[total_qty_idx] is not None:
                                try:
                                    total_qty_val = float(rd[total_qty_idx])
                                    if total_qty_val > 0 and shipped_val >= total_qty_val:
                                        is_shipped_clean = True
                                except ValueError:
                                    pass
                        except ValueError:
                            pass
                            
                color_clean2 = re.sub(r'\[\d+\]', '', color_raw2).strip()
                # 关键逻辑：在途/未结清的记录纳入 master_active_skcs 用于四维对账
                if not is_shipped_clean:
                    if cc_raw2:
                        master_active_skcs.add(pc + cc_raw2)
                    if color_clean2:
                        master_active_code_colors.add((pc, color_clean2))
                        if cc_raw2:
                            master_active_code_colors.add((pc, cc_raw2))
                            
                color = str(rd[nba_indexes['color']]) if rd[nba_indexes['color']] else ''
                mc = re.search(r'\[(\d+)\]', color)
                cc = mc.group(1) if mc else '?'
                skc = pc + cc
                
                # 实际到仓时间 / 实际出货时间 提取（带回退）
                actual_date = None
                act_idx = nba_indexes['actual_arrival']
                if act_idx is not None and len(rd) > act_idx:
                    actual_date = get_date_value(rd[act_idx])
                if actual_date is None:
                    plan_idx = nba_indexes['plan_arrival']
                    if plan_idx is not None and len(rd) > plan_idx:
                        actual_date = get_date_value(rd[plan_idx])
                        
                key = (skc, actual_date)
                if key not in all_restock:
                    all_restock[key] = {'sizes': {k: 0 for k in SIZE_KEYS}, 'color': color, 'status': '', 'factory': ''}
                    
                for sz, idx in nba_indexes['sizes'].items():
                    if idx is not None and len(rd) > idx:
                        val = rd[idx]
                        all_restock[key]['sizes'][sz] += val if isinstance(val, (int, float)) else 0
                        
                status_idx = nba_indexes['status']
                orig_status_nba = str(rd[status_idx])[:30] if (status_idx is not None and len(rd) > status_idx and rd[status_idx]) else ''
                if is_shipped_clean:
                    all_restock[key]['status'] = '已出清' if not orig_status_nba else f"已出清({orig_status_nba})"
                    all_restock[key]['is_cleared'] = True
                elif orig_status_nba:
                    all_restock[key]['status'] = orig_status_nba
                    
                factory_idx = nba_indexes['factory']
                if factory_idx is not None and len(rd) > factory_idx and rd[factory_idx]:
                    all_restock[key]['factory'] = str(rd[factory_idx])
                    
                if skc: 
                    skc_has_any.add(skc)
                    master_progress_skcs.add(skc)
                    
                # 注册四维对账条目至数据库
                total_qty_nba = 0
                if total_qty_idx is not None and len(rd) > total_qty_idx and rd[total_qty_idx] is not None:
                    try: total_qty_nba = float(rd[total_qty_idx])
                    except: pass
                master_records_db.append({
                    'code': pc,
                    'cc': cc_raw2,
                    'color': color_clean2,
                    'date': actual_date,
                    'qty': total_qty_nba,
                    'is_cleared': is_shipped_clean
                })
        wb2.close()

    # ── 3.5 扫描并加载独立翻单表 ──────────────────────────────────────
    try:
        desktop_dir = os.path.expanduser(paths.get('restock_dir') or '/Users/junny/Desktop/淘宝店铺/生产部补单表/')
        wechat_dir = os.path.expanduser(paths.get('wechat_file_dir') or '/Users/junny/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_iqz0imhyx4m812_8da3/msg/file')
        
        candidates = []
        candidates += glob.glob(os.path.join(desktop_dir, "*翻单*.xlsx"))
        if os.path.exists(wechat_dir):
            candidates += glob.glob(os.path.join(wechat_dir, "**", "*翻单*.xlsx"), recursive=True)
            
        now = today
        now_year, now_week, _ = now.isocalendar()
        
        def is_file_in_current_week(filepath):
            filename = os.path.basename(filepath)
            file_date = None
            
            # 尝试从文件名提取 4 位数日期
            m = re.search(r'\d{4}', filename)
            if m:
                date_str = m.group(0)
                try:
                    month = int(date_str[:2])
                    day = int(date_str[2:])
                    year = now.year
                    file_date = datetime(year, month, day)
                except ValueError:
                    pass
            
            if file_date is None:
                # 提取失败时使用文件修改时间兜底
                mtime = os.path.getmtime(filepath)
                file_date = datetime.fromtimestamp(mtime)
                
            file_year, file_week, _ = file_date.isocalendar()
            return (file_year == now_year) and (file_week == now_week)

        def get_group_key(filename):
            base = os.path.splitext(filename)[0]
            owner = 'default'
            m_owner = re.search(r'[\(（]([^\(\)（）]+)[\)）]', filename)
            if m_owner:
                owner = m_owner.group(1).strip()
                owner = re.sub(r'\d+', '', owner).strip()
                if not owner: owner = 'default'
                
            brand = 'unknown'
            filename_upper = filename.upper()
            if 'SG' in filename_upper:
                brand = 'SG'
            elif 'NBA' in filename_upper:
                brand = 'NBA'
            else:
                base_clean = re.sub(r'[\(（].*?[\)）]', '', base)
                base_clean = re.sub(r'\d+', '', base_clean)
                base_clean = base_clean.replace('翻单', '').replace('电商款', '').replace('年', '').replace(' ', '')
                brand = base_clean or 'default'
                
            return f"{brand}_{owner}"

        def get_version_family_key(filename, filepath=None):
            """提取版本家族 key，优先从 Excel 内容标题行读取真实日期（解决文件名与内容日期不一致问题）。
            例如：文件名 SG26年电商款翻单0728（桥）.xlsx，但内容第1行为 SG电商专供款翻单0729（桥），
            则应识别为 0729 而非 0728，避免与真实的 0728 文件竞争。"""
            base = os.path.splitext(filename)[0]
            gk = get_group_key(filename)
            date_str = None
            
            # NOTE: 优先尝试读取 Excel 内容标题行（第1行A1单元格）中的4位日期
            if filepath:
                try:
                    import openpyxl
                    _wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
                    _ws = _wb.active
                    first_row = next(_ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
                    _wb.close()
                    if first_row:
                        title_str = ' '.join(str(c) for c in first_row if c is not None)
                        m_title = re.search(r'\d{4}', title_str)
                        if m_title:
                            date_str = m_title.group(0)
                except Exception:
                    pass  # 读取失败时退回文件名解析
            
            # FALLBACK: 从文件名中提取4位数日期
            if not date_str:
                m_date = re.search(r'\d{4}', base)
                date_str = m_date.group(0) if m_date else 'nodate'
            
            return f"{gk}_{date_str}"

        # 跟踪每个负责人 gk 分组中，各 SKU 在各到货日期的最后一次写入值，从而实现覆盖去重
        # 格式：{(gk, skc, size, delivery_date): {'qty': qty, 'status': status, 'factory': factory, 'is_missing_delivery': bool}}
        gk_sku_restock = {}
        
        # 1. 过滤：滚动加载最近 14 天内修改或接收的翻单文件，以确保上周活跃翻单不会丢失
        limit_time = today - timedelta(days=14)
        raw_candidates = []
        for filepath in candidates:
            basename = os.path.basename(filepath)
            if basename.startswith('~$') or basename.startswith('.~'):
                continue
            if 'hoodrich' in basename.lower():
                logger.info("  [硬性过滤] 排除 HOODRICH 品牌翻单文件: %s", basename)
                continue
            mtime = os.path.getmtime(filepath)
            if datetime.fromtimestamp(mtime) >= limit_time:
                raw_candidates.append((filepath, mtime))
                
        # 2. 版本系列文件去重：同一品牌+负责人+日期的多个副本文件（如带有 (1) 或修改时间不同的同名表），只保留最新的那一份
        family_dict = {}
        for filepath, mtime in raw_candidates:
            basename = os.path.basename(filepath)
            # NOTE: 传入 filepath 使得函数可以读取 Excel 内容标题行来提取真实日期
            family_key = get_version_family_key(basename, filepath=filepath)
            if family_key not in family_dict or mtime > family_dict[family_key][1]:
                family_dict[family_key] = (filepath, mtime)
                
        valid_candidates = list(family_dict.values())
        valid_candidates.sort(key=lambda x: x[1])
        
        logger.info("自动过滤后待解析独立翻单文件共: %d个", len(valid_candidates))
        for filepath, mtime in valid_candidates:
            basename = os.path.basename(filepath)
            gk = get_group_key(basename)
            logger.info("  解析独立翻单表: %s (分组: %s, 修改时间: %s)", 
                        basename, gk, datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S'))
            
            try:
                records = parse_individual_restock_file(filepath)
                logger.info("    从 %s 中解析出 %d 条翻单记录", basename, len(records))
                for r in records:
                    # 滚动活跃池交期有效性校验：如果是本周以前的历史老文件，且其交期早于(today - 7天)，判定为已完全交付的陈旧记录，跳过不计入
                    if not is_file_in_current_week(filepath) and r['delivery'] is not None:
                        comp_del = r['delivery']
                        if isinstance(comp_del, dt.date) and not isinstance(comp_del, datetime):
                            comp_del = datetime.combine(comp_del, datetime.min.time())
                        if comp_del < today - timedelta(days=7):
                            continue

                    skc = r['skc']
                    delivery_date = r['delivery']
                    is_missing_delivery = False
                    if delivery_date is None:
                        is_missing_delivery = True
                        delivery_date = today
                        
                    if isinstance(delivery_date, dt.date) and not isinstance(delivery_date, datetime):
                        delivery_date = datetime.combine(delivery_date, datetime.min.time())
                        
                    sz = r['size']
                    # 相同 (gk, skc, sz, delivery_date) 在新文件里有就直接更新覆盖，而非直接叠加
                    gk_key = (gk, skc, sz, delivery_date)
                    gk_sku_restock[gk_key] = {
                        'qty': r['qty'],
                        'status': str(r['status'])[:30] if r['status'] else '',
                        'factory': str(r['factory']) if r['factory'] else '',
                        'is_missing_delivery': is_missing_delivery,
                        'raw_delivery': r.get('raw_delivery'),
                        'source': basename
                    }
                    
                    if skc:
                        skc_has_any.add(skc)
            except Exception as ex:
                logger.error("  解析独立翻单表 %s 失败: %s", basename, ex)
                
        # 3. 同一 (gk, skc, sz) 下各尺码数量完全一致但货期不同 → 合并取最晚货期
        # NOTE: 先按 (gk, skc) 聚合所有 (delivery_date, sz) 的 qty 签名，
        #       签名相同意味着"跨货期、各尺码下单量完全相等"，直接视为同一笔翻单被反复更新，
        #       只保留货期最晚的那一条，避免重复计入翻单数量。
        
        # Step 3a：以 (gk, skc) 为 key，收集该 SKC 所有货期下的 {sz: qty} 快照
        skc_date_sizes = defaultdict(dict)  # {(gk, skc, delivery_date): {sz: qty}}
        skc_date_meta  = {}                 # {(gk, skc, delivery_date): info}
        for (gk_k, skc_k, sz_k, del_k), info in gk_sku_restock.items():
            skc_date_sizes[(gk_k, skc_k, del_k)][sz_k] = info['qty']
            skc_date_meta [(gk_k, skc_k, del_k)] = info  # 同 (gk, skc, del_k) 内最后一次 sz 写入的 meta 共享
        
        # Step 3b：对每个 (gk, skc)，将所有货期版本的 qty_signature 分组
        gk_skc_dates = defaultdict(list)    # {(gk, skc): [(delivery_date, qty_sig)]}
        for (gk_k, skc_k, del_k), sizes in skc_date_sizes.items():
            # qty 签名：将每个 sz→qty 排序后转为不可变 tuple，数量相同则签名相同
            qty_sig = tuple(sorted(sizes.items()))
            gk_skc_dates[(gk_k, skc_k)].append((del_k, qty_sig, sizes))
        
        # Step 3c：签名相同的多个货期条目 → 只保留货期最晚的
        deduped_skc_date_sizes = {}  # {(gk, skc, delivery_date): {sz: qty}}  只剩代表条目
        for (gk_k, skc_k), date_list in gk_skc_dates.items():
            # 按签名分组
            sig_groups = defaultdict(list)
            for del_k, qty_sig, sizes in date_list:
                sig_groups[qty_sig].append((del_k, sizes))
            for qty_sig, same_sig_entries in sig_groups.items():
                # 取该签名分组中货期最晚的条目作为代表条目
                latest_del, latest_sizes = max(same_sig_entries, key=lambda x: x[0])
                deduped_skc_date_sizes[(gk_k, skc_k, latest_del)] = latest_sizes

        # Step 3d：写入去重后的翻单数据（各尺码数量直接赋值，不叠加——同一笔翻单取最晚货期后只保留一条）
        for (gk_k, skc_k, delivery_date), sizes in deduped_skc_date_sizes.items():
            meta = skc_date_meta.get((gk_k, skc_k, delivery_date), {})
            is_missing = meta.get('is_missing_delivery', False)
            key = (skc_k, delivery_date)
            # NOTE: 直接覆盖写入，不使用 +=，因为相同签名的多个货期版本已折叠为一条最晚货期记录
            sizes_filled = {k: 0 for k in SIZE_KEYS}
            for sz_k, qty_v in sizes.items():
                if sz_k in sizes_filled:
                    sizes_filled[sz_k] = qty_v
            all_restock[key] = {
                'sizes': sizes_filled,
                'color': '',
                'status': meta.get('status', ''),
                'factory': meta.get('factory', ''),
                'is_missing_delivery': is_missing,
                'source': meta.get('source', '')
            }
                
    except Exception as e:
        logger.error("扫描独立翻单表失败: %s", e)

    logger.info("翻单记录: %d条", len(all_restock))

    # ── 3.8 生成库存回补数据（Sheet2）─────────────────────────────────
    sheet2_data = []
    if prev_wh_file:
        for skc, cur_total in wh_neg_total.items():
            code = skc[:8]
            if code in EXCLUDED_UNOFFICIAL_CODES or skc in EXCLUDED_UNOFFICIAL_CODES:
                continue
            is_junma = product_is_junma.get(code, False)
            
            # 判断昨日与今日是否实际有库存（独享仓最高优先级）
            prev_sizes = prev_neg_size.get(skc, {k: 0.0 for k in SIZE_KEYS})
            cur_sizes = wh_neg_size.get(skc, {k: 0.0 for k in SIZE_KEYS})
            
            yesterday_has_stock = has_skc_stock(skc, prev_sizes, is_junma)
            today_has_stock = has_skc_stock(skc, cur_sizes, is_junma)
            
            if not yesterday_has_stock and today_has_stock:
                productCode = skc[:8]
                if productCode in product_table:
                    listDate = product_table[productCode].get('上架日期')
                    if listDate and hasattr(listDate, 'date') and listDate.date() > today_date:
                        remark = '🚀 自动铺货(新品预售)'
                    else:
                        remark = '📦 开启商品同步'
                elif productCode in master_progress_codes or skc in master_progress_skcs or skc in skc_has_any:
                    remark = '🌟 大货新品到仓(待上架)'
                else:
                    remark = '❓ 异常新品到仓(商品表未录)'
                    
                sheet2_data.append({
                    'skc': skc,
                    'color': wh_colors_txt.get(skc, ''),
                    'cur_total': cur_total,
                    'prev_total': prev_neg_total.get(skc, 0),
                    'prev_sizes': prev_sizes,
                    'cur_sizes':  cur_sizes,
                    '备注': remark,
                })
    logger.info("库存回补（Sheet2）: %d条", len(sheet2_data))

    # 跨源交叉审计：提取被两张大货进度表完全遗漏的微信活跃翻单
    # 黄金准则（严格四维匹配算法）：
    #   必须同时匹配 款号 + 颜色/编号 + 预计到货日期 + 下单数量！
    #   只有在大货总表中能找到在途/相近交期且数量吻合的翻单记录时，才判定为「已登记」；
    #   否则，凡是不一致或找不到对应翻单批次的，一律判定为「大货表遗漏/未录入」，提报预警！
    omitted_details = []
    seen_omitted_skcs = set()
    
    for (skc, actual_date), d in all_restock.items():
        if not skc or not d.get('source'):
            continue
            
        pc = skc[:8]
        cc = skc[8:10] if len(skc) >= 10 and skc[8:10].isdigit() else None
        color_txt = str(d.get('color', '')).strip()
        color_clean = re.sub(r'\[\d+\]', '', color_txt).strip()
        w_qty = sum(d['sizes'].values())
        if w_qty <= 0:
            continue
            
        # 在 master_records_db 中进行【下单时间/文件日期 + 货期/到货日期 + 款 + 色】严格比对
        is_fully_matched = False
        w_deliv_date = actual_date.date() if isinstance(actual_date, datetime) else (actual_date if isinstance(actual_date, date) else None)
        
        for m in master_records_db:
            # 1. 款匹配
            if m['code'] != pc:
                continue
            # 2. 色匹配
            c_match = False
            if cc and m['cc'] == cc:
                c_match = True
            elif color_clean and m['color'] == color_clean:
                c_match = True
            if not c_match:
                continue
            # 3. 排除已显式标记清完结清的历史旧批次
            if m['is_cleared']:
                continue
            # 4. 到货日期/货期匹配 (交期相差 25 天以内或其中一个为空)
            d_match = True
            m_deliv_date = m['date'].date() if isinstance(m['date'], datetime) else (m['date'] if isinstance(m['date'], date) else None)
            if w_deliv_date and m_deliv_date:
                if abs((w_deliv_date - m_deliv_date).days) > 25:
                    d_match = False
            if not d_match:
                continue
                
            # 【下单时间/货期 + 款 + 色】全符合，即使下单数量有轻微出入，也认定已在大货总表登记
            is_fully_matched = True
            break
            
        if not is_fully_matched:
            omitted_key = (skc, w_deliv_date)
            if omitted_key in seen_omitted_skcs:
                continue
            seen_omitted_skcs.add(omitted_key)
            
            owner = d.get('status') or ''
            omitted_details.append({
                'skc': skc,
                'code': pc,
                'color': wh_colors_txt.get(skc, color_txt),
                'delivery': actual_date,
                'qty': w_qty,
                'owner': owner,
                'factory': d.get('factory') or '',
                'source': d.get('source') or ''
            })

    # ── 4. 窗口期过滤 ─────────────────────────────────────────────────
    results = []
    customer_results = []
    customer_window_end = today + timedelta(days=20)
    logger.info("客服专用到货参考窗口期: %s ~ %s", window_start.strftime('%m/%d'), customer_window_end.strftime('%m/%d'))

    for (skc, actual_date), d in all_restock.items():
        total = sum(d['sizes'].values())
        if total <= 0: continue
        if actual_date is None: continue
        if not isinstance(actual_date, datetime):
            try:
                actual_date = datetime.strptime(str(actual_date), '%Y-%m-%d')
            except (ValueError, TypeError):
                continue

        # 客服专用参考：只要在窗口期内（window_start <= actual_date <= customer_window_end），即便是出清条目也记录到客服专用 Sheet
        if window_start <= actual_date <= customer_window_end:
            d_cust = dict(d)
            if d.get('is_missing_delivery'):
                d_cust['delivery'] = '交期未填/待定'
            else:
                d_cust['delivery'] = actual_date.strftime('%Y-%m-%d')
            d_cust['color'] = wh_colors_txt.get(skc, d.get('color', ''))
            customer_results.append((skc, actual_date, d_cust))

        # 标准窗口期到货（Sheet1 逻辑：排除已出清项，需在网红店商品表中登记，且在标准 window_end 内）
        if d.get('is_cleared', False): continue
        if skc not in neg_skcs: continue
        pc_code = skc[:8]
        if pc_code not in product_table: continue
        if window_start <= actual_date <= window_end:
            d2 = dict(d)
            if d.get('is_missing_delivery'):
                d2['delivery'] = '交期未填/待定'
            else:
                d2['delivery'] = actual_date.strftime('%Y-%m-%d')
            d2['color'] = wh_colors_txt.get(skc, d.get('color', ''))
            
            # 对大货表缺失项在 Sheet 1 生产状态打上显式高亮标记：
            # 区分「上周及更早散表在当周总表无记录(疑似未实际下单)」与「本周新翻单(大货表待登)」
            if skc not in master_progress_skcs:
                owner = d.get('status') or ''
                src = d.get('source') or ''
                if is_older_than_last_week(src, base_date=today):
                    d2['status'] = f"⚠️疑似未下单({owner}/上周散表)" if owner else "⚠️疑似未下单(上周散表未登总表)"
                else:
                    d2['status'] = f"⚠️大货表待登({owner})" if owner else "⚠️大货表待登(本周新单)"
                
            d2['isApprox'] = False
            results.append((skc, actual_date, d2))

    # ── 通用跨源尺码签名去重与交期智能融合引擎 ───────────────────────
    def dedup_and_merge_entries(entry_list, is_customer_sheet=False):
        """
        全局跨源尺码签名去重与交期融合引擎
        同一 SKC 下各尺码数量完全一致（sz_signature 相同）视为同一笔翻单：
        1. 交期裁决：剔除'交期未填/待定'，优先选取有具体交期的版本；若有多个具体交期，以最新校准（最晚有效交期）为准。
        2. 生产状态融合：优先保留包含具体生产进度描述（如'在绣花'、'没面料'、'正在印花中'等）的状态。
        3. 工厂融合：优先选取有效非空工厂名（排除 #N/A 和空）。
        4. 数据源融合：若为客服表，合并多源文件名（如 微信散表 + 大货表）。
        """
        groups = defaultdict(list)
        for skc, act_dt, d in entry_list:
            sz_tuple = tuple(round(float(d.get('sizes', {}).get(k, 0) or 0), 0) for k in SIZE_KEYS)
            groups[(skc, sz_tuple)].append((skc, act_dt, d))

        deduped = []
        for (skc, sz_tuple), item_list in groups.items():
            if len(item_list) == 1:
                deduped.append(item_list[0])
                continue

            # 存在重复项，执行智能裁决与融合
            # 1. 优先提取明确具体交期（过滤未填/待定/None）
            valid_items = []
            for it in item_list:
                it_dt = it[1]
                it_deliv_str = str(it[2].get('delivery', ''))
                if it_dt is not None and '待定' not in it_deliv_str and '未填' not in it_deliv_str and it_deliv_str != 'None':
                    valid_items.append((it_dt, it))

            if valid_items:
                # 按具体交期取最晚（即最新校准的货期）的记录作为主模板
                valid_items.sort(key=lambda x: x[0])
                best_dt, best_entry = valid_items[-1]
                final_dt = best_dt
                final_delivery_str = best_dt.strftime('%Y-%m-%d')
            else:
                best_entry = item_list[0]
                final_dt = best_entry[1]
                final_delivery_str = '交期未填/待定'

            # 2. 生产状态融合：优先保留包含具体生产进度的描述
            statuses = [it[2].get('status', '') for it in item_list if it[2].get('status') and str(it[2].get('status')) != 'None']
            detailed_status = ''
            for st in statuses:
                if any(kw in str(st) for kw in ['中', '在', '料', '定做', '订做', '改', '裁', '开板', '已出清', '缺失', '到仓']):
                    detailed_status = str(st)
                    break
            if not detailed_status and statuses:
                detailed_status = str(statuses[-1])

            # 3. 工厂融合：优先选取有效工厂名称（排除 #N/A、None、- 等无效占位）
            factories = [it[2].get('factory', '') for it in item_list if it[2].get('factory') and str(it[2].get('factory')).strip() not in ['#N/A', 'None', '', '-', 'None']]
            final_factory = str(factories[-1]) if factories else str(best_entry[2].get('factory') or '')

            merged_d = dict(best_entry[2])
            merged_d['delivery'] = final_delivery_str
            merged_d['status'] = detailed_status
            merged_d['factory'] = final_factory

            # 4. 数据源融合（主要用于客服专用表）
            if is_customer_sheet:
                sources = []
                for it in item_list:
                    src = it[2].get('source', '')
                    if src and src not in sources:
                        sources.append(src)
                if sources:
                    merged_d['source'] = ' + '.join(sources)

            deduped.append((skc, final_dt, merged_d))

        return deduped

    # ── 对 Sheet 1 (窗口期到货) 与 Sheet 4 (客服专用参考) 执行去重融合 ──
    results = dedup_and_merge_entries(results, is_customer_sheet=False)
    customer_results = dedup_and_merge_entries(customer_results, is_customer_sheet=True)

    # 客服专用参考按 (款号, 到货日期, SKC) 排序
    customer_results.sort(key=lambda x: (x[0][:8], x[2].get('delivery', ''), x[0]))
    logger.info("客服专用到货参考全量摘录(已跨源去重融合): %d条", len(customer_results))

    # ── 5. 近似颜色匹配（P0 功能） ──────────────────────────────────────
    approxMatches = findApproxColorMatches(neg_skcs, all_restock, skc_has_any)
    approx_count = 0
    for match in approxMatches:
        deliv = match['delivery']
        if deliv is None:
            continue
        if not isinstance(deliv, datetime):
            try:
                deliv = datetime.strptime(str(deliv), '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
        if window_start <= deliv <= window_end:
            d_alt = match['restock_detail']
            d2 = dict(d_alt)
            d2['delivery'] = deliv.strftime('%Y-%m-%d')
            d2['color'] = wh_colors_txt.get(match['origSkc'], d_alt.get('color', ''))
            # 备注匹配状态
            d2['status'] = match['matchType']
            d2['isApprox'] = True
            results.append((match['origSkc'], deliv, d2))
            approx_count += 1
            
    if approx_count > 0:
        logger.info("引入近似颜色匹配行: %d条", approx_count)

    results.sort(key=lambda x: (x[0][:8], x[2].get('delivery', ''), neg_skcs.get(x[0], 0)))
    logger.info("窗口期到货(含近似): %d条", len(results))

    cutoff = (today - timedelta(days=7)).date()
    restock_recent = set()
    for (skc, actual_date), d in all_restock.items():
        if sum(d['sizes'].values()) <= 0: continue
        if actual_date and actual_date.date() >= cutoff:
            restock_recent.add(skc)
            
    approxMatchedSkcs = set(m['origSkc'] for m in approxMatches)

    unmatched = {skc: qty for skc, qty in neg_skcs.items() if skc not in restock_recent and skc not in approxMatchedSkcs}
    logger.info("无翻单负库存: %d条", len(unmatched))

    # ── 7. 生意参谋评分 ────────────────────────────────────────────────
    # 生意参谋目录自适应定位
    biz_dir_resolved = resolveBizDir(config)
    biz_data = load_business_data(biz_dir_resolved, pats['biz_advisor'])
    scores   = score_and_advise(unmatched, biz_data)

    # ── 8. 归集全链路疑似错误与待核对项（供专属纠错 Sheet 使用）──────
    audit_errors = []
    seen_audit_keys = set()
    
    # 1. 历史旧交期 / 疑似遗漏交期修改（如 8月发起的翻单填成了 3月）
    for (gk_k, skc_k, sz_k, del_k), info in gk_sku_restock.items():
        pc = skc_k[:8]
        if pc in EXCLUDED_UNOFFICIAL_CODES or skc_k in EXCLUDED_UNOFFICIAL_CODES:
            continue
        raw_del = info.get('raw_delivery')
        # 如果填写的原始交期早于今天 20 天以上
        if raw_del and isinstance(raw_del, datetime) and raw_del < today - timedelta(days=20):
            audit_key = (skc_k, 'HISTORICAL_DELIVERY', raw_del.strftime('%Y-%m-%d'))
            if audit_key not in seen_audit_keys:
                seen_audit_keys.add(audit_key)
                src = info.get('source', '')
                audit_errors.append({
                    'code': pc,
                    'skc': skc_k,
                    'color': wh_colors_txt.get(skc_k, ''),
                    'type': '⚠️ 历史旧交期(疑似遗漏修改)',
                    'source': src,
                    'raw_delivery': raw_del.strftime('%Y-%m-%d'),
                    'qty': info.get('qty', 0),
                    'owner_factory': f"{info.get('status', '')} / {info.get('factory', '')}",
                    'suggestion': f"翻单表中填写的交期为 {raw_del.strftime('%Y-%m-%d')}（早于当前日期），疑似复制旧模板未更新交期，请核对最新预计到仓日"
                })
        # 原始交期未填/待定
        elif raw_del is None or info.get('is_missing_delivery'):
            audit_key = (skc_k, 'MISSING_DELIVERY')
            if audit_key not in seen_audit_keys:
                seen_audit_keys.add(audit_key)
                audit_errors.append({
                    'code': pc,
                    'skc': skc_k,
                    'color': wh_colors_txt.get(skc_k, ''),
                    'type': '❓ 交期未填/待定',
                    'source': info.get('source', ''),
                    'raw_delivery': '空白/未填',
                    'qty': info.get('qty', 0),
                    'owner_factory': f"{info.get('status', '')} / {info.get('factory', '')}",
                    'suggestion': "翻单表中有下单数量但未填写到货日期，请核实并补充明确交期"
                })
                
    # 2. 微信有翻单但大货总进度表遗漏未登记 (来自 omitted_details)
    for item in omitted_details:
        audit_key = (item['skc'], 'OMITTED_FROM_MASTER')
        if audit_key not in seen_audit_keys:
            seen_audit_keys.add(audit_key)
            deliv_str = item['delivery'].strftime('%Y-%m-%d') if isinstance(item['delivery'], datetime) else str(item['delivery'])
            src = item.get('source', '')
            is_old = is_older_than_last_week(src, base_date=today)
            if is_old:
                err_type = '⚠️ 疑似临时改动/未下单'
                sugg = f"该翻单来源于上周或更早散表（{src}），但在当周生产大货总表中查无登记，疑似临时改动且未实际向工厂下单，请核对是否已取消"
            else:
                err_type = '⚠️ 本周新单待总表登记'
                sugg = f"本周微信散表（{src}）最新下单，当周生产大货总表尚未同步录入，请生产部及时补登"

            audit_errors.append({
                'code': item['code'],
                'skc': item['skc'],
                'color': item['color'],
                'type': err_type,
                'source': src,
                'raw_delivery': deliv_str,
                'qty': item['qty'],
                'owner_factory': f"{item.get('owner', '')} / {item.get('factory', '')}",
                'suggestion': sugg
            })
            
    # 3. 翻单款但在店铺主商品表《SG网红店商品表.xlsx》中查无此款 (未建档)
    for (skc, act_dt), d in all_restock.items():
        pc = skc[:8]
        if pc in EXCLUDED_UNOFFICIAL_CODES or skc in EXCLUDED_UNOFFICIAL_CODES:
            continue
        if pc not in product_table:
            audit_key = (pc, 'NOT_IN_PRODUCT_TABLE')
            if audit_key not in seen_audit_keys:
                seen_audit_keys.add(audit_key)
                audit_errors.append({
                    'code': pc,
                    'skc': skc,
                    'color': wh_colors_txt.get(skc, d.get('color', '')),
                    'type': '⚠️ 商品表未录入(未登记款)',
                    'source': d.get('source', ''),
                    'raw_delivery': act_dt.strftime('%Y-%m-%d') if isinstance(act_dt, datetime) else str(act_dt),
                    'qty': sum(d['sizes'].values()),
                    'owner_factory': f"{d.get('status', '')} / {d.get('factory', '')}",
                    'suggestion': "已有翻单安排，但店铺主商品表中尚未建档录入该款，需核对是否为专供款或补录"
                })
                
    # 4. 仓库已扫码入库但商品表未录入 (来自 Sheet2 中的异常款)
    for d2 in sheet2_data:
        if '异常新品到仓' in d2.get('备注', ''):
            skc = d2['skc']
            pc = skc[:8]
            audit_key = (skc, 'WH_STOCK_NOT_IN_PRODUCT_TABLE')
            if audit_key not in seen_audit_keys:
                seen_audit_keys.add(audit_key)
                audit_errors.append({
                    'code': pc,
                    'skc': skc,
                    'color': d2.get('color', ''),
                    'type': '❓ 仓库有货但商品表未录',
                    'source': '系统库存导出(实际入库)',
                    'raw_delivery': '已在仓',
                    'qty': d2.get('cur_total', 0),
                    'owner_factory': '仓库实物库存',
                    'suggestion': "仓库已扫码入库产生实际可销库存，但商品表未登记该款，需尽快建档上架并开启同步"
                })

    logger.info("疑似错误待核对项归集: %d条", len(audit_errors))

    return results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched, scores, \
           paths['output_dir'], sheet2_data, whEntityTotal, product_table, wh_file, omitted_details, \
           customer_results, wh_neg_total, audit_errors

# ── Excel 输出 ───────────────────────────────────────────────────────────────
def to_excel(results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched, scores,
             output_dir, sheet2_data, whEntityTotal, product_table=None, wh_file=None, omitted_details=None,
             customer_results=None, wh_neg_total=None, audit_errors=None):
    from datetime import datetime
    import re
    
    time_suffix = ""
    if wh_file:
        basename = os.path.basename(wh_file)
        # 尝试提取文件名中的14位时间戳 (如 20260707085544)
        match_14 = re.search(r'(\d{14})', basename)
        if match_14:
            time_suffix = match_14.group(1)
        else:
            # 降级：提取8位日期并拼接当前时分秒
            match_8 = re.search(r'(\d{8})', basename)
            if match_8:
                time_suffix = match_8.group(1) + datetime.now().strftime('%H%M%S')
                
    if not time_suffix:
        time_suffix = datetime.now().strftime('%Y%m%d%H%M%S')
        
    out_path  = os.path.join(output_dir, f"窗口期到货_负库存_{time_suffix}.xlsx")
    os.makedirs(output_dir, exist_ok=True)

    wb = openpyxl.Workbook()

    hf   = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    hfont = Font(color='FFFFFF', bold=True, size=10)
    tb   = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),   bottom=Side(style='thin'))

    def draw_legend_box(ws, start_col_idx, legend_items, title="🎨 填充颜色含义告示板"):
        """在 worksheet 的右侧空列放置美观的图例告示区，方便用户和客服直接对照理解。"""
        header_fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True, size=10)
        thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                             top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        
        c1 = start_col_idx
        c2 = start_col_idx + 1
        
        # 标题行 (第1行合并)
        ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        t_cell = ws.cell(1, c1)
        t_cell.value = title
        t_cell.fill = header_fill
        t_cell.font = header_font
        t_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(1, c2).border = thin_border
        t_cell.border = thin_border
        
        ws.row_dimensions[1].height = 24
        
        for i, (fill_hex, label, desc) in enumerate(legend_items, start=2):
            cell_color = ws.cell(i, c1)
            cell_desc  = ws.cell(i, c2)
            
            cell_color.value = label
            cell_color.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type='solid')
            cell_color.font = Font(bold=True, size=9)
            cell_color.alignment = Alignment(horizontal='center', vertical='center')
            cell_color.border = thin_border
            
            cell_desc.value = desc
            cell_desc.font = Font(size=9)
            cell_desc.alignment = Alignment(horizontal='left', vertical='center')
            cell_desc.border = thin_border
            
            ws.row_dimensions[i].height = 20

        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(c1)].width = 24
        ws.column_dimensions[get_column_letter(c2)].width = 46

    # ── Sheet1: 窗口期到货 ───────────────────────────────────────────
    ws1 = wb.active
    ws1.title = '窗口期到货'
    headers1 = ['款号', 'SKC编码', '颜色', '仓库可销',
                 'S', 'M', 'L', 'XL', '2XL', '均码',
                 '批次数量', 'S', 'M', 'L', 'XL', '2XL', '均码',
                 '到货日期', '生产状态', '工厂']
    ws1.append(headers1)
    for ci, h in enumerate(headers1, 1):
        cell = ws1.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb

    for skc, actual_date, d in results:
        restock_total = sum(d['sizes'].values())
        wh_total  = neg_skcs.get(skc, 0)
        wh_sizes = wh_neg_size.get(skc, {k: 0 for k in SIZE_KEYS})
        wh_color = wh_colors_txt.get(skc, d.get('color', ''))
        entity   = whEntityTotal.get(skc, 0)

        row = [
            skc[:8], skc, wh_color,
            round(wh_total, 0),
            round(wh_sizes.get('S', 0), 0), round(wh_sizes.get('M', 0), 0),
            round(wh_sizes.get('L', 0), 0), round(wh_sizes.get('XL', 0), 0),
            round(wh_sizes.get('2XL', 0), 0), round(wh_sizes.get('均码', 0), 0),
            round(restock_total, 0),
            round(d['sizes'].get('S', 0), 0), round(d['sizes'].get('M', 0), 0),
            round(d['sizes'].get('L', 0), 0), round(d['sizes'].get('XL', 0), 0),
            round(d['sizes'].get('2XL', 0), 0), round(d['sizes'].get('均码', 0), 0),
            d.get('delivery', ''), d.get('status', ''), d.get('factory', '')
        ]
        ws1.append(row)
        rn = ws1.max_row
        
        isApprox = d.get('isApprox', False)
        if isApprox:
            fill_color = 'FFF2CC'
        else:
            fill_color = 'FCE4D6' if wh_total < 0 else 'DAEFCE'
            
        for ci in range(1, len(headers1) + 1):
            cell = ws1.cell(rn, ci)
            cell.border = tb
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw1 = [10, 14, 10, 10, 6, 6, 6, 6, 6, 6, 10, 6, 6, 6, 6, 6, 6, 12, 20, 8]
    for ci, w in enumerate(cw1, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.freeze_panes = 'A2'
    ws1.auto_filter.ref = f"A1:{get_column_letter(ws1.max_column)}{ws1.max_row}"

    # 绘制 Sheet1 右侧告示区
    draw_legend_box(ws1, 22, [
        ('DAEFCE', '🟢 正常到货', '仓库可销正常未断货，到货计划按期推进'),
        ('FCE4D6', '🔴 缺货加急到货', '仓库负库存/缺货到货，需生产与仓库重点加急入库'),
        ('FFF2CC', '🟡 同色系近似匹配', '原色号无翻单，基于同款号同色系参考的交期到货'),
        ('F8CBAD', '⚠️ 疑似未下单', '上周或更早散表但在当周总表中无记录，疑似临时改动未实际下单')
    ], title="🎨 到货跟进告示板")

    # ── Sheet2: 库存回补 ─────────────────────────────────────────────
    ws2 = wb.create_sheet('库存回补')
    headers2 = ['款号', 'SKC', '颜色',
                 '当前可销', '当前实体',
                 '前表可销', 'XS前', 'S前', 'M前', 'L前', 'XL前', '2XL前', '均码前',
                 '后表可销', 'XS后', 'S后', 'M后', 'L后', 'XL后', '2XL后', '均码后', '备注']
    ws2.append(headers2)
    for ci, h in enumerate(headers2, 1):
        cell = ws2.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb

    # 按运营优先级排序：1.开启商品同步 ➔ 2.大货新品到仓 ➔ 3.自动铺货 ➔ 4.异常新品到仓
    priority_map = {
        '📦 开启商品同步': 1,
        '开启商品同步': 1,
        '🌟 大货新品到仓(待上架)': 2,
        '🚀 自动铺货(新品预售)': 3,
        '自动铺货': 3,
        '❓ 异常新品到仓(商品表未录)': 4,
        '新款待上架': 4
    }
    sorted_sheet2 = sorted(sheet2_data, key=lambda x: priority_map.get(x.get('备注', ''), 99))

    for d2 in sorted_sheet2:
        skc        = d2['skc']
        cur_sizes  = d2['cur_sizes']
        prev_sizes = d2['prev_sizes']
        entity     = whEntityTotal.get(skc, 0)
        remark     = d2.get('备注', '')
        row = [
            skc[:8], skc, d2['color'],
            round(d2['cur_total'], 0), round(entity, 0),
            round(d2['prev_total'], 0),
            round(prev_sizes.get('XS', 0), 0), round(prev_sizes.get('S', 0), 0),
            round(prev_sizes.get('M', 0), 0),   round(prev_sizes.get('L', 0), 0),
            round(prev_sizes.get('XL', 0), 0),  round(prev_sizes.get('2XL', 0), 0),
            round(prev_sizes.get('均码', 0), 0),
            round(d2['cur_total'], 0),
            round(cur_sizes.get('XS', 0), 0), round(cur_sizes.get('S', 0), 0),
            round(cur_sizes.get('M', 0), 0),  round(cur_sizes.get('L', 0), 0),
            round(cur_sizes.get('XL', 0), 0), round(cur_sizes.get('2XL', 0), 0),
            round(cur_sizes.get('均码', 0), 0),
            remark,
        ]
        ws2.append(row)
        rn = ws2.max_row
        color_map = {
            '📦 开启商品同步': 'FFF2CC',         # 暖黄色高亮
            '🚀 自动铺货(新品预售)': 'DAEFCE',    # 清新浅绿
            '🌟 大货新品到仓(待上架)': 'E2EFDA',   # 柔和浅青绿
            '❓ 异常新品到仓(商品表未录)': 'F4CCCC', # 醒目浅红警告
            '开启商品同步': 'FFF2CC',
            '自动铺货': 'DAEFCE',
            '新款待上架': 'F4CCCC'
        }
        fill_color = color_map.get(remark, 'FFFFFF')
        for ci in range(1, len(headers2) + 1):
            cell = ws2.cell(rn, ci)
            cell.border = tb
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw2 = [10, 14, 10, 10, 10, 10, 6, 6, 6, 6, 6, 6, 6, 10, 6, 6, 6, 6, 6, 6, 6, 26]
    for ci, w in enumerate(cw2, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = 'A2'
    ws2.auto_filter.ref = f"A1:{get_column_letter(ws2.max_column)}{ws2.max_row}"

    # 绘制 Sheet2 右侧告示区
    draw_legend_box(ws2, 24, [
        ('FFF2CC', '🟡 📦 开启商品同步', '最高优先级！需立即去 JEOMS / 千牛后台开启自动上传同步'),
        ('E2EFDA', '🟢 🌟 大货新品到仓(待上架)', '仓库已实际扫码入库的大货新品，准备安排商品上架'),
        ('DAEFCE', '🌿 🚀 自动铺货(新品预售)', '新品预售商品到仓回补'),
        ('F4CCCC', '🔴 ❓ 异常新品到仓', '扫码入库但商品表中尚未查到信息的异常款，需运营核对')
    ], title="🎨 铺货与商品同步告示板")

    # ── Sheet3: 无翻单需评分 ─────────────────────────────────────────
    ws3 = wb.create_sheet('无翻单需评分')
    headers3 = ['款号', 'SKC', '仓库可销', '综合评分',
                '7天访客', '7天支付', '7天加购', '翻单建议']
    ws3.append(headers3)
    for ci, h in enumerate(headers3, 1):
        cell = ws3.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb

    for skc, qty in sorted(unmatched.items(), key=lambda x: neg_skcs[x[0]]):
        code   = skc[:8]
        sc     = scores.get(code, {})
        score  = sc.get('score')
        advice = sc.get('advice', '⚪ 无数据')
        fill_color = 'F4CCCC' if '❌' in advice else ('FFF2CC' if '⚠️' in advice else 'DAEFCE')
        row = [code, skc, round(qty, 0),
               str(score) if score is not None else '无数据',
               sc.get('visitors', 0), sc.get('pay', 0), sc.get('cart', 0), advice]
        ws3.append(row)
        rn = ws3.max_row
        for ci in range(1, len(headers3) + 1):
            cell = ws3.cell(rn, ci)
            cell.border = tb
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw3 = [10, 14, 10, 10, 10, 10, 10, 16]
    for ci, w in enumerate(cw3, 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = 'A2'
    ws3.auto_filter.ref = f"A1:{get_column_letter(ws3.max_column)}{ws3.max_row}"

    # ── Sheet4: 客服专用_到货参考 ───────────────────────────────────────
    if customer_results is None:
        customer_results = []
    if wh_neg_total is None:
        wh_neg_total = {}

    ws4 = wb.create_sheet('客服专用_到货参考')
    headers4 = ['款号', 'SKC编码', '颜色', '仓库可销',
                 'S', 'M', 'L', 'XL', '2XL', '均码',
                 '批次到货数量', 'S', 'M', 'L', 'XL', '2XL', '均码',
                 '预计到货日期', '生产状态', '工厂', '翻单数据源']
    ws4.append(headers4)
    for ci, h in enumerate(headers4, 1):
        cell = ws4.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb

    for skc, actual_date, d in customer_results:
        restock_total = sum(d['sizes'].values())
        wh_total  = wh_neg_total.get(skc, 0)
        wh_sizes = wh_neg_size.get(skc, {k: 0 for k in SIZE_KEYS})
        wh_color = wh_colors_txt.get(skc, d.get('color', ''))

        row = [
            skc[:8], skc, wh_color,
            round(wh_total, 0),
            round(wh_sizes.get('S', 0), 0), round(wh_sizes.get('M', 0), 0),
            round(wh_sizes.get('L', 0), 0), round(wh_sizes.get('XL', 0), 0),
            round(wh_sizes.get('2XL', 0), 0), round(wh_sizes.get('均码', 0), 0),
            round(restock_total, 0),
            round(d['sizes'].get('S', 0), 0), round(d['sizes'].get('M', 0), 0),
            round(d['sizes'].get('L', 0), 0), round(d['sizes'].get('XL', 0), 0),
            round(d['sizes'].get('2XL', 0), 0), round(d['sizes'].get('均码', 0), 0),
            d.get('delivery', ''), d.get('status', ''), d.get('factory', ''), d.get('source', '')
        ]
        ws4.append(row)
        rn = ws4.max_row
        
        fill_color = 'E2EFDA' if wh_total >= 10 else ('FCE4D6' if wh_total < 0 else 'FFF2CC')
        for ci in range(1, len(headers4) + 1):
            cell = ws4.cell(rn, ci)
            cell.border = tb
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw4 = [10, 14, 10, 10, 6, 6, 6, 6, 6, 6, 10, 6, 6, 6, 6, 6, 6, 12, 20, 10, 24]
    for ci, w in enumerate(cw4, 1):
        ws4.column_dimensions[get_column_letter(ci)].width = w
    ws4.freeze_panes = 'A2'
    ws4.auto_filter.ref = f"A1:{get_column_letter(ws4.max_column)}{ws4.max_row}"

    # 绘制 Sheet4 右侧客服专属告示区
    draw_legend_box(ws4, 23, [
        ('E2EFDA', '🟢 浅绿标示', '仓库现货充足 (可销 ≥ 10)，下单即可正常现货发货'),
        ('FFF2CC', '🟡 暖黄标示', '仓库现货偏紧 (0 ≤ 可销 ≤ 9)，接单需关注余量，参考预计到货期'),
        ('FCE4D6', '🔴 浅红标示', '仓库缺货断货 (可销 < 0)，需引导买家预售，参考预计到货期承诺发货')
    ], title="🎨 客服查货与预售指引告示区")

    # ── Sheet5: 疑似错误_待核对 ─────────────────────────────────────
    if audit_errors is None:
        audit_errors = []
        
    ws5 = wb.create_sheet('疑似错误_待核对')
    headers5 = ['款号', 'SKC', '颜色', '疑似异常类型', '涉及数据源文件', '表格原填交期', '涉及数量', '负责人/工厂', '疑问说明与核对建议']
    ws5.append(headers5)
    for ci, h in enumerate(headers5, 1):
        cell = ws5.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb

    color_audit_map = {
        '⚠️ 历史旧交期(疑似遗漏修改)': 'FFF2CC', # 暖黄预警
        '⚠️ 大货总表未登记': 'F4CCCC',        # 浅红警告
        '⚠️ 商品表未录入(未登记款)': 'E2EFDA', # 浅青色
        '❓ 仓库有货但商品表未录': 'FCE4D6',   # 浅粉红
        '❓ 交期未填/待定': 'FFF2CC'           # 暖黄预警
    }

    for item in audit_errors:
        err_type = item.get('type', '')
        fill_color = color_audit_map.get(err_type, 'FFF2CC')
        row = [
            item.get('code', ''),
            item.get('skc', ''),
            item.get('color', ''),
            err_type,
            item.get('source', ''),
            item.get('raw_delivery', ''),
            round(item.get('qty', 0), 0),
            item.get('owner_factory', ''),
            item.get('suggestion', '')
        ]
        ws5.append(row)
        rn = ws5.max_row
        for ci in range(1, len(headers5) + 1):
            cell = ws5.cell(rn, ci)
            cell.border = tb
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw5 = [10, 14, 10, 24, 26, 14, 10, 16, 46]
    for ci, w in enumerate(cw5, 1):
        ws5.column_dimensions[get_column_letter(ci)].width = w
    ws5.freeze_panes = 'A2'
    ws5.auto_filter.ref = f"A1:{get_column_letter(ws5.max_column)}{ws5.max_row}"

    # 绘制 Sheet5 右侧专属告示区
    draw_legend_box(ws5, 11, [
        ('FFF2CC', '🟡 ⚠️ 历史旧交期', '翻单表填写的交期早于当前日期，疑似复制旧模板未更新交期'),
        ('F4CCCC', '🔴 ⚠️ 大货表遗漏', '微信群中已有下单记录，但生产大货总表中未见此批翻单'),
        ('E2EFDA', '🟢 ⚠️ 商品表未录', '已有翻单安排，但《SG网红店商品表.xlsx》中尚未建档录入该款'),
        ('FCE4D6', '🌸 ❓ 仓库有货未录', '仓库已扫码入库产生实际可销库存，但商品表未登记该款')
    ], title="🎨 数据核对与纠错告示区")

    wb.save(out_path)
    logger.info("已保存: %s", out_path)
    return out_path

# ── 入口 ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='负库存翻单到货分析')
    parser.add_argument('--skc', '--款号', help='只查询指定SKC（支持模糊匹配）')
    args = parser.parse_args()

    results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched, scores, \
        output_dir, sheet2_data, whEntityTotal, product_table, wh_file, omitted_details, \
        customer_results, wh_neg_total, audit_errors = analyze()

    if args.skc:
        results  = [r for r in results  if args.skc in r[0]]
        unmatched = {k: v for k, v in unmatched.items() if args.skc in k}
        if omitted_details:
            omitted_details = [o for o in omitted_details if args.skc in o['skc']]
        if customer_results:
            customer_results = [c for c in customer_results if args.skc in c[0]]
        if audit_errors:
            audit_errors = [a for a in audit_errors if args.skc in a['skc'] or args.skc in a['code']]

    if results or unmatched or customer_results:
        path = to_excel(results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched,
                        scores, output_dir, sheet2_data, whEntityTotal, product_table, wh_file, omitted_details,
                        customer_results, wh_neg_total, audit_errors)
        logger.info("完成: 窗口期到货%d条 | 客服参考%d条 | 库存回补%d条 | 疑似错误核对%d条 | 无翻单%d条",
                    len(results), len(customer_results), len(sheet2_data), len(audit_errors), len(unmatched))
    else:
        logger.info("无数据")
