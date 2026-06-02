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
from collections import defaultdict
from pathlib import Path
import re
import os
import glob
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

CONFIG_PATH = Path(__file__).parent.parent / 'config.json'

DEFAULT_CONFIG = {
    'paths': {
        'warehouse_dir':       '~/Desktop/淘宝店铺/仓库商品库存表/',
        'biz_dir':             '~/Desktop/报表/c1报表/报表数据源/',
        'restock_dir':         '~/Desktop/淘宝店铺/生产部补单表/',
        'output_dir':          '~/Desktop/淘宝店铺/归档/窗口期到货_负库存/',
        'product_table':       '~/Desktop/淘宝店铺/店铺商品表/SG网红店商品表.xlsx',
    },
    'file_patterns': {
        'warehouse':       '系统库存导出_',
        'warehouse_old':   '实体库存_黄之俊',
        'sg_restock':      'SG品牌进度表*.xlsx',
        'nba_restock':     'SG.NBA电商专供订单进度*.xlsx',
        'biz_advisor':     '*生意参谋*.xls*',
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

def load_config():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding='utf-8')
        logger.info("已生成默认配置: %s", CONFIG_PATH)
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = json.load(f)
    paths = config.get('paths', {})
    for key, path in paths.items():
        expanded = os.path.expanduser(path)
        if os.path.exists(expanded):
            paths[key] = expanded
    return config

def get_latest_file(directory, pattern='实体库存_黄之俊'):
    if not os.path.exists(directory):
        return None
    files = sorted(glob.glob(directory + '*.xlsx'), key=os.path.getmtime, reverse=True)
    for f in files:
        if pattern in f:
            return f
    return None  # 不匹配则返回 None，不随意 fallback

def get_latest_inventory_file(dirs_and_patterns):
    """尝试多个目录+模式组合，返回第一个找到的文件"""
    for directory, pattern in dirs_and_patterns:
        f = get_latest_file(directory, pattern)
        if f:
            return f
    return None

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

def parse_color_name(color_str):
    """从 [48]藏青 格式提取颜色名"""
    if not color_str:
        return ''
    m = re.search(r'\[?\d+\]?(.+)', str(color_str))
    return m.group(1).strip() if m else str(color_str).strip()

def parse_size_name_from_text(size_text):
    """从 S(165/78A) 或 S 或 均码 等文本提取尺码名"""
    if not size_text:
        return None
    s = str(size_text).strip()
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
    except:
        return None

def get_date_value(val):
    # 支持 M/D-M/D (如 "6/6-6/8") 和 M/D (如 "6/8")
    if isinstance(val, str):
        s = val.strip()
        m1 = re.match(r'^(\d+)/(\d+)(?:-(\d+)/(\d+))?$', s)
        if m1:
            groups = m1.groups()
            if groups[2] is not None:  # M/D-M/D 格式
                return datetime(datetime.now().year, int(groups[0]), int(groups[1]))  # 取范围开始日
            else:  # M/D 格式
                return datetime(datetime.now().year, int(groups[0]), int(groups[1]))
        return None

    if val is None: return None
    if isinstance(val, datetime): return val
    if isinstance(val, (int, float)):
        return excel_serial_to_date(val)
    s = str(val).strip()
    if not s or s in ['-', 'None', '']: return None
    current_year = datetime.now().year
    m = re.match(r'^(\d+)/(\d+)-(\d+)$', s)
    if m: return datetime(current_year, int(m.group(1)), int(m.group(2)))
    m = re.match(r'^(\d+)/(\d+)$', s)
    if m: return datetime(current_year, int(m.group(1)), int(m.group(2)))
    return None

# ── 库存加载（新格式）───────────────────────────────────────────────────────
def load_warehouse(filepath):
    """
    新格式：系统库存导出_黄之俊_*.xlsx
    列（1-based）：1=商品名称 2=商品编码 3=实体仓编码 4=实体仓名称 5=虚拟仓编码
                  6=虚拟仓名称 7=仓库类型 8=商品简称 9=规格编码 10=季节
                  11=可销数 12=库存数 13=采购在途数 14=可用数
    规格编码：货号8位+颜色2位+尺码2位（共12位）
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb.active
    logger.info("新格式库存文件: %s (%d行)", filepath.split('/')[-1], ws.max_row)

    wh_neg_total  = defaultdict(float)   # skc -> 总可销
    wh_neg_size   = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS})  # skc -> {尺码: 可销}
    wh实体_total  = defaultdict(float)   # skc -> 总实体
    wh_colors_txt = {}   # skc -> 颜色文字（[48]藏青）
    wh_spec_names = {}   # skc -> 规格编码（用于参考）

    for row in ws.iter_rows(min_row=2):
        rd = [c.value for c in row]
        # 取关键列（1-based转0-based已在 rd 索引中）
        # rd[1]=商品编码 rd[8]=规格编码 rd[10]=可销数 rd[11]=库存数 rd[16]=尺码文字 rd[17]=颜色文字
        code     = str(rd[1]).strip() if rd[1] else ''
        spec     = str(rd[8]).strip() if rd[8] else ''
        kexiao   = float(rd[10]) if rd[10] is not None else 0.0
        entity   = float(rd[11]) if rd[11] is not None else 0.0
        size_txt = rd[16]  # 尺码文字 S(165/78A)、均码 等
        color_txt = rd[17]  # 颜色文字 [48]藏青、[79]黑色 等

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

        if color_txt:
            wh_colors_txt[skc] = str(color_txt)

        wh_neg_total[skc]  += kexiao
        wh实体_total[skc]  += entity
        if size_name:
            wh_neg_size[skc][size_name] += kexiao

    wb.close()
    logger.info("新格式解析: %d个SKC", len(wh_neg_total))
    return wh_neg_total, wh_neg_size, wh实体_total, wh_colors_txt


# ── 店铺商品表（正价Sheet）───────────────────────────────────────────────────
def load_product_table(filepath):
    """
    加载正价Sheet，返回 {款号: {上架日期, 商品状态}}
    用于判断库存回补是'未来上新'还是'老款回补'
    """
    if not filepath or not os.path.exists(filepath):
        return {}
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb['正价']
        result = {}
        for r in range(2, ws.max_row + 1):
            款号 = ws.cell(r, 2).value  # col2=商品编码
            if not 款号:
                continue
            款号 = str(款号).strip()
            上架日期 = ws.cell(r, 14).value  # col14=上架日期
            商品状态 = ws.cell(r, 4).value  # col4=商品状态
            result[款号] = {
                '上架日期': 上架日期,
                '商品状态': 商品状态,
            }
        wb.close()
        logger.info("正价表款号: %d个", len(result))
        return result
    except Exception as e:
        logger.warning("加载正价表失败: %s", e)
        return {}


# ── 生意参谋 ─────────────────────────────────────────────────────────────────
def load_business_data(biz_dir, pattern):
    files = sorted(glob.glob(biz_dir + pattern), key=os.path.getmtime, reverse=True)
    if not files:
        logger.info("未找到生意参谋数据")
        return {}
    try:
        ext = os.path.splitext(files[0])[1].lower()
        prod_data = defaultdict(lambda: {'visitors': 0, 'pv': 0, 'cart': 0, 'orders': 0, 'pay': 0})
        if ext == '.xls':
            import xlrd
            wb = xlrd.open_workbook(files[0])
            ws = wb.sheet_by_index(0)
            headers = ws.row_values(4)
            for row_idx in range(5, ws.nrows):
                row = dict(zip(headers, ws.row_values(row_idx)))
                code = row.get('货号', '')
                if not code: continue
                def to_int(v):
                    if v in ['', '-', None]: return 0
                    return int(str(v).replace(',', ''))
                prod_data[code]['visitors'] += to_int(row.get('商品访客数'))
                prod_data[code]['pv']       += to_int(row.get('商品浏览量'))
                prod_data[code]['cart']     += to_int(row.get('商品加购件数'))
                prod_data[code]['orders']   += to_int(row.get('下单件数'))
                prod_data[code]['pay']      += to_int(row.get('支付件数'))
        else:
            wb = openpyxl.load_workbook(files[0], data_only=True)
            ws = wb.active
            headers = [str(c.value) if c.value else '' for c in ws[4]]
            for row in ws.iter_rows(min_row=5, values_only=True):
                rd = dict(zip(headers, row))
                code = str(rd.get('货号', '')).strip()
                if not code: continue
                def to_int(v):
                    if v is None or v == '' or v == '-': return 0
                    return int(str(v).replace(',', ''))
                prod_data[code]['visitors'] += to_int(rd.get('商品访客数'))
                prod_data[code]['pv']       += to_int(rd.get('商品浏览量'))
                prod_data[code]['cart']     += to_int(rd.get('商品加购件数'))
                prod_data[code]['orders']   += to_int(rd.get('下单件数'))
                prod_data[code]['pay']      += to_int(rd.get('支付件数'))
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


def findColumnIndexes(headerCells: list) -> dict:
    """
    根据表头单元格值列表，自动匹配各个关键字段的列索引（0-based）
    """
    indexes = {
        'code': None,            # 款号
        'color': None,           # 颜色
        'factory': None,         # 工厂
        'shipped': None,         # 总出货 / 已出货列
        'actual_arrival': None,  # 优先：实际到仓时间 / 实际出货时间列
        'plan_arrival': None,    # 备选：到仓时间 / 出货时间 / 货期列
        'status': None,          # 备注 / 生产状态列
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
        elif val == '工厂':
            indexes['factory'] = idx
        elif val in ['出货数量', '总出货', '已出货']:
            indexes['shipped'] = idx
        elif val in ['实际到仓时间', '实际出货时间', '实际\n出货时间']:
            indexes['actual_arrival'] = idx
        elif val in ['到仓时间', '出货时间', '货期']:
            indexes['plan_arrival'] = idx
        elif val in ['备注', '生产状态']:
            indexes['status'] = idx
            
        for size_name, patterns in size_patterns.items():
            for pat in patterns:
                if re.match(pat, val):
                    indexes['sizes'][size_name] = idx
                    break
                    
    return indexes

def getHeaderRowAndIndexes(ws) -> tuple[int, dict]:
    """
    在工作表中定位表头行并返回 (1-based行号, 列索引字典)
    """
    for r in range(1, 15):
        row_cells = [cell.value for cell in next(ws.iter_rows(min_row=r, max_row=r))]
        if not any(row_cells):
            continue
        row_str = [str(x).strip() for x in row_cells if x is not None]
        if any('款号' in s for s in row_str) and any('颜色' in s for s in row_str):
            indexes = findColumnIndexes(row_cells)
            return r, indexes
            
    first_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    return 1, findColumnIndexes(first_row)


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
    if not wh_file:
        raise FileNotFoundError("未找到仓库库存文件，请确认数据源路径")
    logger.info("仓库文件: %s", wh_file)

    # 动态窗口期：今天前10天 ~ 今天后15天
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    window_start = today - timedelta(days=10)
    window_end   = today + timedelta(days=15)
    logger.info("窗口期: %s ~ %s", window_start.strftime('%m/%d'), window_end.strftime('%m/%d'))

    # ── 1. 加载当前仓库 ──────────────────────────────────────────────
    wh_neg_total, wh_neg_size, wh_entity_total, wh_colors_txt = load_warehouse(wh_file)

    # 判断均码款
    product_is_junma = {}
    for skc in wh_neg_total:
        code = skc[:8]
        if code not in product_is_junma:
            has_junma = any(
                s.startswith(code) and wh_neg_size.get(s, {}).get('均码', 0) > 0
                for s in wh_neg_size
            )
            product_is_junma[code] = has_junma

    # 筛选可销<10 of SKC
    neg_skcs = {}
    for skc, qty in wh_neg_total.items():
        code    = skc[:8]
        is_junma = product_is_junma.get(code, False)
        sizes   = wh_neg_size.get(skc, {})
        if is_junma:
            if sizes.get('均码', 0) < 10:
                neg_skcs[skc] = qty
        else:
            if any(sizes.get(s, 0) < 10 for s in ['S', 'M', 'L', 'XL', '2XL']):
                neg_skcs[skc] = qty
    logger.info("可销<10的SKC: %d", len(neg_skcs))

    # ── 2. 加载前一天的仓库（用于Sheet2库存回补）────────────────────
    import datetime as dt
    cur_basename = os.path.basename(wh_file)
    cur_date_str = re.search(r'(\d{8})', cur_basename)
    cur_date = None
    if cur_date_str:
        cur_date = dt.datetime.strptime(cur_date_str.group(1), '%Y%m%d').date()

    all_dirs = [
        (paths.get('warehouse_dir') or ''),
        (paths.get('feishu_inbound') or ''),
    ]
    all_wh_files = []
    for d in all_dirs:
        if d:
            all_wh_files += sorted(glob.glob(d + '*.xlsx'), key=os.path.getmtime, reverse=True)

    prev_wh_file = None
    if cur_date:
        for f in all_wh_files:
            f_basename = os.path.basename(f)
            f_date_str = re.search(r'(\d{8})', f_basename)
            f_date = None
            if f_date_str:
                f_date = dt.datetime.strptime(f_date_str.group(1), '%Y%m%d').date()
            if f_date and f_date < cur_date and f != wh_file and '系统库存导出' in f:
                prev_wh_file = f
                break

    prev_neg_total  = defaultdict(float)
    prev_neg_size   = defaultdict(lambda: {k: 0.0 for k in SIZE_KEYS})
    if prev_wh_file:
        logger.info("前日仓库文件: %s", prev_wh_file)
        prev_neg_total, prev_neg_size, _, _ = load_warehouse(prev_wh_file)

    product_table = load_product_table(paths.get('product_table'))
    today = datetime.now().date()

    sheet2_data = []
    if prev_wh_file:
        for skc, cur_total in wh_neg_total.items():
            prev_total = prev_neg_total.get(skc, 0)
            if prev_total <= 0 and cur_total > 0:
                款号 = skc[:8]
                if 款号 in product_table:
                    上架 = product_table[款号].get('上架日期')
                    if 上架 and hasattr(上架, 'date') and 上架.date() > today:
                        备注 = '自动铺货'
                    else:
                        备注 = '开启商品同步'
                else:
                    备注 = '新款待上架'
                sheet2_data.append({
                    'skc': skc,
                    'color': wh_colors_txt.get(skc, ''),
                    'cur_total': cur_total,
                    'prev_total': prev_total,
                    'prev_sizes': prev_neg_size.get(skc, {k: 0.0 for k in SIZE_KEYS}),
                    'cur_sizes':  wh_neg_size.get(skc,  {k: 0.0 for k in SIZE_KEYS}),
                    '备注': 备注,
                })
    logger.info("库存回补（Sheet2）: %d条", len(sheet2_data))

    # ── 3. 加载翻单记录 ──────────────────────────────────────────────
    restock_dir = paths['restock_dir']
    sg_files = sorted(glob.glob(restock_dir + pats['sg_restock']), key=os.path.getmtime, reverse=True)
    nba_files = sorted(glob.glob(restock_dir + pats['nba_restock']), key=os.path.getmtime, reverse=True)
    if not sg_files:
        raise FileNotFoundError(f"未找到SG品牌进度表")
    nba_file = nba_files[0] if nba_files else None

    all_restock = {}  # {(skc, date): {...}}
    skc_has_any = set()

    # SG品牌进度表 (多 Sheet 自适应与日期回退)
    wb1 = openpyxl.load_workbook(sg_files[0], data_only=True, read_only=True)
    for sheet in wb1.sheetnames:
        ws1 = wb1[sheet]
        header_num, sg_indexes = getHeaderRowAndIndexes(ws1)
        
        # 必须找到关键列，否则视为无效 Sheet 并自动跳过
        if sg_indexes['code'] is None or sg_indexes['color'] is None:
            continue
            
        logger.info("解析 SG 品牌进度表 Sheet: %s", sheet)
        for row in ws1.iter_rows(min_row=header_num + 1, values_only=True):
            rd = list(row)
            if len(rd) <= max(sg_indexes['code'], sg_indexes['color']):
                continue
            pc = str(rd[sg_indexes['code']]).strip() if rd[sg_indexes['code']] else ''
            if not pc:
                continue
                
            # 过滤已出货
            shipped_idx = sg_indexes['shipped']
            if shipped_idx is not None and len(rd) > shipped_idx:
                shipped = rd[shipped_idx]
                if shipped is not None and shipped != '':
                    try:
                        if float(shipped) > 0:
                            continue
                    except ValueError:
                        pass
                        
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
                
            for sz, idx in sg_indexes['sizes'].items():
                if idx is not None and len(rd) > idx:
                    val = rd[idx]
                    all_restock[key]['sizes'][sz] += val if isinstance(val, (int, float)) else 0
                    
            status_idx = sg_indexes['status']
            if status_idx is not None and len(rd) > status_idx and rd[status_idx]:
                all_restock[key]['status'] = str(rd[status_idx])[:30]
                
            factory_idx = sg_indexes['factory']
            if factory_idx is not None and len(rd) > factory_idx and rd[factory_idx]:
                all_restock[key]['factory'] = str(rd[factory_idx])
                
            if skc: 
                skc_has_any.add(skc)
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
            for row in ws2.iter_rows(min_row=header_num + 1, values_only=True):
                rd = list(row)
                if len(rd) <= max(nba_indexes['code'], nba_indexes['color']):
                    continue
                pc = str(rd[nba_indexes['code']]).strip() if rd[nba_indexes['code']] else ''
                if not pc:
                    continue
                    
                # 过滤已出货
                shipped_idx = nba_indexes['shipped']
                if shipped_idx is not None and len(rd) > shipped_idx:
                    shipped = rd[shipped_idx]
                    if shipped is not None and (isinstance(shipped, (int, float)) and shipped > 0):
                        continue
                        
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
                if status_idx is not None and len(rd) > status_idx and rd[status_idx]:
                    all_restock[key]['status'] = str(rd[status_idx])[:30]
                    
                factory_idx = nba_indexes['factory']
                if factory_idx is not None and len(rd) > factory_idx and rd[factory_idx]:
                    all_restock[key]['factory'] = str(rd[factory_idx])
                    
                if skc: 
                    skc_has_any.add(skc)
        wb2.close()

    logger.info("翻单记录: %d条", len(all_restock))

    # ── 4. 窗口期过滤 ─────────────────────────────────────────────────
    results = []
    for (skc, actual_date), d in all_restock.items():
        if skc not in neg_skcs: continue
        total = sum(d['sizes'].values())
        if total <= 0: continue
        if actual_date is None: continue
        if not isinstance(actual_date, datetime):
            try:
                actual_date = datetime.strptime(str(actual_date), '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
        if window_start <= actual_date <= window_end:
            d2 = dict(d)
            d2['delivery'] = actual_date.strftime('%Y-%m-%d')
            d2['color'] = wh_colors_txt.get(skc, d.get('color', ''))
            results.append((skc, actual_date, d2))

    results.sort(key=lambda x: (x[2].get('delivery', ''), neg_skcs.get(x[0], 0)))
    logger.info("窗口期到货: %d条", len(results))

    # ── 5. 无翻单负库存 ─────────────────────────────────────────────
    cutoff = today - timedelta(days=7)
    restock_recent = set()
    for (skc, actual_date), d in all_restock.items():
        if sum(d['sizes'].values()) <= 0: continue
        if actual_date and actual_date.date() >= cutoff:
            restock_recent.add(skc)

    unmatched = {skc: qty for skc, qty in neg_skcs.items() if skc not in restock_recent}
    logger.info("无翻单负库存: %d条", len(unmatched))

    # ── 6. 生意参谋评分 ────────────────────────────────────────────────
    biz_data = load_business_data(paths['biz_dir'], pats['biz_advisor'])
    scores   = score_and_advise(unmatched, biz_data)

    return results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched, scores, \
           paths['output_dir'], sheet2_data, wh_entity_total, product_table

# ── Excel 输出 ───────────────────────────────────────────────────────────────
def to_excel(results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched, scores,
             output_dir, sheet2_data, wh_entity_total, product_table=None):
    from datetime import date
    today_str = date.today().strftime('%Y%m%d')
    out_path  = f"{output_dir}窗口期到货_负库存_{today_str}.xlsx"
    os.makedirs(output_dir, exist_ok=True)

    wb = openpyxl.Workbook()

    hf   = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    hfont = Font(color='FFFFFF', bold=True, size=10)
    tb   = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),   bottom=Side(style='thin'))

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
        entity   = wh_entity_total.get(skc, 0)

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
        for ci in range(1, len(headers1) + 1):
            cell = ws1.cell(rn, ci)
            cell.border = tb
            cell.alignment = Alignment(horizontal='center', vertical='center')
            # 背景色：负数→红色提示，正常→绿色
            fill_color = 'FCE4D6' if wh_total < 0 else 'DAEFCE'
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw1 = [10, 14, 10, 10, 6, 6, 6, 6, 6, 6, 10, 6, 6, 6, 6, 6, 6, 12, 20, 8]
    for ci, w in enumerate(cw1, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w

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

    today_disp = date.today().strftime('%Y-%m-%d')
    for d2 in sheet2_data:
        skc        = d2['skc']
        cur_sizes  = d2['cur_sizes']
        prev_sizes = d2['prev_sizes']
        entity     = wh_entity_total.get(skc, 0)
        备注       = d2.get('备注', '')
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
            备注,
        ]
        ws2.append(row)
        rn = ws2.max_row
        # 颜色区分：自动铺货=绿色，老款同步=橙色，新款待上架=灰色
        fill_color = {'自动铺货': 'DAEFCE', '开启商品同步': 'FFF2CC', '新款待上架': 'D9D9D9'}.get(备注, 'DAEFCE')
        for ci in range(1, len(headers2) + 1):
            cell = ws2.cell(rn, ci)
            cell.border = tb
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw2 = [10, 14, 10, 10, 10, 10, 6, 6, 6, 6, 6, 6, 6, 10, 6, 6, 6, 6, 6, 6, 6, 12]
    for ci, w in enumerate(cw2, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w

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

    wb.save(out_path)
    logger.info("已保存: %s", out_path)
    return out_path

# ── 入口 ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='负库存翻单到货分析')
    parser.add_argument('--skc', '--款号', help='只查询指定SKC（支持模糊匹配）')
    args = parser.parse_args()

    results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched, scores, \
        output_dir, sheet2_data, wh_entity_total, product_table = analyze()

    if args.skc:
        results  = [r for r in results  if args.skc in r[0]]
        unmatched = {k: v for k, v in unmatched.items() if args.skc in k}

    if results or unmatched:
        path = to_excel(results, neg_skcs, wh_neg_size, wh_colors_txt, unmatched,
                        scores, output_dir, sheet2_data, wh_entity_total, product_table)
        logger.info("完成: 窗口期到货%d条 | 库存回补%d条 | 无翻单%d条",
                    len(results), len(sheet2_data), len(unmatched))
    else:
        logger.info("无数据")
