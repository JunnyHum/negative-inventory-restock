#!/usr/bin/env python3
"""
负库存翻单到货分析
- 每次运行自动匹配最新数据源，路径通过 config.json 配置
- 仓库/SG表/NBA表/生意参谋：均取最新文件
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

# NOTE: 配置文件与技能定义放在一起，修改路径只需编辑 config.json
CONFIG_PATH = Path(__file__).parent.parent / 'config.json'

DEFAULT_CONFIG = {
    'paths': {
        'warehouse_dir': '~/Desktop/淘宝库存监控/仓库商品库存表/',
        'biz_dir': '~/Desktop/活动报名8折商品/',
        'restock_dir': '~/Desktop/淘宝库存监控/生产部补单表/',
        'output_dir': '~/Desktop/淘宝库存监控/归档/',
    },
    'file_patterns': {
        'warehouse': '实体库存_黄之俊',
        'sg_restock': 'SG品牌进度表*.xlsx',
        'nba_restock': 'SG.NBA电商专供订单进度*.xlsx',
        'biz_advisor': '*生意参谋*.xls*',
    }
}

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)


def load_config() -> dict:
    """
    加载配置文件，不存在则自动生成默认配置
    所有路径自动展开 ~ 为用户主目录，并校验是否存在
    """
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(
            json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2),
            encoding='utf-8'
        )
        logger.info("已生成默认配置: %s", CONFIG_PATH)

    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = json.load(f)

    paths = config.get('paths', {})
    for key, path in paths.items():
        expanded = os.path.expanduser(path)
        if not os.path.exists(expanded):
            raise FileNotFoundError(
                f"配置路径不存在: {key} = {path}\n"
                f"  展开后: {expanded}\n"
                f"  请检查 config.json: {CONFIG_PATH}"
            )
        paths[key] = expanded

    return config


SIZE_KEYS = ['XS', 'S', 'M', 'L', 'XL', '2XL', '均码']
SIZE_PAT = re.compile(r'(S|M|L|XL|XXL|均码)')
SIZE_NUM_PAT = re.compile(r'(\d+)/\d+([A-Z])\)$')  # (165/80S) -> S
SIZE_MEASURE_PAT = re.compile(r'\((\d+)/\d+[A-Z]\)$')  # (165/84A) -> 均码

def get_latest_file(directory, pattern='实体库存_黄之俊'):
    files = sorted(glob.glob(directory + '*.xlsx'), key=os.path.getmtime, reverse=True)
    for f in files:
        if pattern in f:
            return f
    return files[-1] if files else None

def parse_spec(spec, spec_name):
    """
    从仓库规格编码和规格名称解析SKC和尺码
    规格编码: WE0202137904 (10位=款号8位+颜色2位+尺码2位)
    规格名称: WE0202137904[79]黑色S(165/84A) 或 WE0000010020[00]透明均码
    返回: (skc, size_name, color_name)
    """
    if not spec or len(spec) < 10:
        return None, None, None

    skc = spec[:8] + spec[8:10]  # 款号+颜色（10位）

    # 尺码从规格编码末位提取：04/14=S, 05/15=M, 06/16=L, 07/17=XL, 08/18=2XL, 00/20=均码
    size_last = spec[11] if len(spec) >= 12 else ''
    size_name_map = {'3': 'XS', '4': 'S', '5': 'M', '6': 'L', '7': 'XL', '8': '2XL', '0': '均码'}
    size_name = size_name_map.get(size_last, None)

    # 颜色从规格名称解析
    color_name = None
    if spec_name:
        # 提取[xx]之后的部分
        after_bracket = re.split(r'\[\d+\]', spec_name)
        after_bracket = after_bracket[-1] if len(after_bracket) > 1 else spec_name
        # 去掉末尾的(身高)部分
        color_raw = re.sub(r'\([^)]+\)$', '', after_bracket).strip()
        # 去掉末尾的尺码字母得到颜色名
        color_name = re.sub(r'(S|M|L|XL|XXL|均码)$', '', color_raw).strip()

    return skc, size_name, color_name

def excel_serial_to_date(serial):
    if serial is None: return None
    if isinstance(serial, datetime): return serial
    try:
        serial = int(float(serial))
        if serial >= 60:
            serial -= 1
        return datetime(1899, 12, 31) + timedelta(days=serial)
    except (ValueError, TypeError, OverflowError):
        return None

def get_date_value(val):
    if val is None: return None
    if isinstance(val, datetime): return val
    if isinstance(val, (int, float)):
        return excel_serial_to_date(val)
    s = str(val).strip()
    if not s or s in ['-', 'None', '']: return None
    current_year = datetime.now().year
    m = re.match(r'^(\d+)/(\d+)-(\d+)$', s)
    if m:
        return datetime(current_year, int(m.group(1)), int(m.group(2)))
    m = re.match(r'^(\d+)/(\d+)$', s)
    if m:
        return datetime(current_year, int(m.group(1)), int(m.group(2)))
    return None

def load_business_data(biz_dir, pattern):
    files = sorted(glob.glob(biz_dir + pattern), key=os.path.getmtime, reverse=True)
    if not files:
        logger.info("未找到生意参谋数据")
        return {}
    try:
        ext = os.path.splitext(files[0])[1].lower()
        prod_data = defaultdict(lambda: {'visitors':0,'pv':0,'cart':0,'orders':0,'pay':0})
        if ext == '.xls':
            import xlrd
            wb = xlrd.open_workbook(files[0])
            ws = wb.sheet_by_index(0)
            headers = ws.row_values(4)
            for row_idx in range(5, ws.nrows):
                row = dict(zip(headers, ws.row_values(row_idx)))
                code = row.get('货号','')
                if not code: continue
                def to_int(v):
                    if v in ['','-',None]: return 0
                    return int(str(v).replace(',',''))
                prod_data[code]['visitors'] += to_int(row.get('商品访客数'))
                prod_data[code]['pv'] += to_int(row.get('商品浏览量'))
                prod_data[code]['cart'] += to_int(row.get('商品加购件数'))
                prod_data[code]['orders'] += to_int(row.get('下单件数'))
                prod_data[code]['pay'] += to_int(row.get('支付件数'))
        else:
            wb = openpyxl.load_workbook(files[0], data_only=True)
            ws = wb.active
            headers = [str(c.value) if c.value else '' for c in ws[4]]
            for row in ws.iter_rows(min_row=5, values_only=True):
                rd = dict(zip(headers, row))
                code = str(rd.get('货号','')).strip()
                if not code: continue
                def to_int(v):
                    if v is None or v == '' or v == '-': return 0
                    return int(str(v).replace(',',''))
                prod_data[code]['visitors'] += to_int(rd.get('商品访客数'))
                prod_data[code]['pv'] += to_int(rd.get('商品浏览量'))
                prod_data[code]['cart'] += to_int(rd.get('商品加购件数'))
                prod_data[code]['orders'] += to_int(rd.get('下单件数'))
                prod_data[code]['pay'] += to_int(rd.get('支付件数'))
        logger.info("生意参谋款号: %d", len(prod_data))
        return prod_data
    except Exception as e:
        logger.error("生意参谋读取失败: %s", e)
        return {}

def score_and_advise(unmatched_skcs, biz_data):
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
        pay = d.get('pay', 0)
        cart = d.get('cart', 0)
        if visitors == 0 and pay == 0:
            advice = '❌ 不建议（无验证）'
            score = 0
        else:
            sv = visitors / max_v * 100 if max_v else 0
            sp = pay / max_p * 100 if max_p else 0
            sc = cart / max_c * 100 if max_c else 0
            score = sv * 0.3 + sp * 0.5 + sc * 0.2
            if score > 50 and pay >= 3:
                advice = '🔥 建议翻单'
            elif score >= 30 and pay >= 1:
                advice = '⚠️ 考虑翻单'
            elif pay == 0:
                advice = '❌ 不建议（0销量）'
            else:
                advice = '⚠️ 考虑翻单'
        results[code] = {'score': round(score, 1), 'pay': pay, 'visitors': visitors,
                         'cart': cart, 'advice': advice}
    return results

def analyze():
    config = load_config()
    paths = config['paths']
    patterns = config['file_patterns']

    warehouse_dir = paths['warehouse_dir']
    biz_dir = paths['biz_dir']
    restock_dir = paths['restock_dir']
    output_dir = paths['output_dir']

    # 动态窗口期：今天前10天 ~ 今天后10天
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    window_start = today - timedelta(days=10)
    window_end = today + timedelta(days=10)
    logger.info("窗口期: %s ~ %s", window_start.strftime('%m/%d'), window_end.strftime('%m/%d'))

    # 1. Load warehouse negative SKCs with size breakdown
    wh_file = get_latest_file(warehouse_dir, patterns['warehouse'])
    logger.info("仓库文件: %s", wh_file)
    wb_wh = openpyxl.load_workbook(wh_file, data_only=True)
    ws_wh = wb_wh.active

    wh_neg_total = defaultdict(float)      # skc -> 总可销
    wh_neg_size = defaultdict(lambda: {k:0.0 for k in SIZE_KEYS})  # skc -> {尺码: 可销}
    wh_colors = {}  # skc -> 颜色名（从规格名称解析）

    for row in ws_wh.iter_rows(min_row=2):
        rd = [c.value for c in row]
        code = str(rd[2]).strip() if rd[2] else ''
        spec = str(rd[7]).strip() if rd[7] else ''
        spec_name = str(rd[8]).strip() if rd[8] else ''
        kexiao = float(rd[19]) if rd[19] else 0.0

        if not code or len(spec) < 10:
            continue

        skc, size_name, color_name = parse_spec(spec, spec_name)
        if not skc:
            continue

        if color_name:
            wh_colors[skc] = color_name

        # 先汇总所有kexiao（不逐行过滤），用sum合并两个仓库的数据
        wh_neg_total[skc] += kexiao
        if size_name:
            wh_neg_size[skc][size_name] += kexiao

    # 判断每个款号（8位）是否为均码款
    product_is_junma = {}  # 款号 -> True/False
    for skc in wh_neg_total.keys():
        code = skc[:8]
        if code not in product_is_junma:
            # 检查该款号下是否有均码规格的库存
            has_junma = any(
                s.startswith(code) and wh_neg_size.get(s, {}).get('均码', 0) > 0
                for s in wh_neg_size.keys()
            )
            product_is_junma[code] = has_junma

    # 再筛选：按款号类型判断是否进入列表
    # 均码款：只看均码列；分尺码款：看S/M/L/XL/2XL列
    neg_skcs = {}
    for skc, qty in wh_neg_total.items():
        code = skc[:8]
        is_junma = product_is_junma.get(code, False)
        sizes = wh_neg_size.get(skc, {})
        if is_junma:
            # 均码款：均码<10
            if sizes.get('均码', 0) < 10:
                neg_skcs[skc] = qty
        else:
            # 分尺码款：任一尺码<10
            if any(sizes.get(s, 0) < 10 for s in ['S','M','L','XL','2XL']):
                neg_skcs[skc] = qty
    logger.info("可销<10的SKC: %d", len(neg_skcs))

    # 加载前一天仓库数据（用于Sheet3库存回补对比）
    prev_wh_neg_total = defaultdict(float)
    prev_wh_neg_size = defaultdict(lambda: {k:0.0 for k in SIZE_KEYS})
    all_warehouse_files = sorted(glob.glob(warehouse_dir + '*.xlsx'), key=os.path.getmtime, reverse=True)
    prev_wh_files = [f for f in all_warehouse_files if patterns['warehouse'] in f]
    if len(prev_wh_files) >= 2:
        prev_wh_file = prev_wh_files[1]  # 第二新 = 前一天
        logger.info("前天仓库文件: %s", prev_wh_file)
        wb_prev = openpyxl.load_workbook(prev_wh_file, data_only=True)
        ws_prev = wb_prev.active
        for row in ws_prev.iter_rows(min_row=2):
            rd = [c.value for c in row]
            code = str(rd[2]).strip() if rd[2] else ''
            spec = str(rd[7]).strip() if rd[7] else ''
            kexiao = float(rd[19]) if rd[19] else 0.0
            if not code or len(spec) < 10:
                continue
            skc, size_name, _ = parse_spec(spec, '')
            if not skc:
                continue
            prev_wh_neg_total[skc] += kexiao
            if size_name:
                prev_wh_neg_size[skc][size_name] += kexiao

    # Sheet3 库存回补：前一天表里可销<=0，后一天可销>0 的SKC
    # 注意：款号必须在前一天表格里出现过才算，没出现过则跳过（无法对比）
    sheet3_data = []
    for skc, cur_total in wh_neg_total.items():
        if skc not in prev_wh_neg_total:
            continue  # 前一天表格里没有此款号，无法对比
        prev_total = prev_wh_neg_total.get(skc, 0)
        if prev_total <= 0 and cur_total > 0:
            prev_sizes = prev_wh_neg_size.get(skc, {k:0.0 for k in SIZE_KEYS})
            cur_sizes = wh_neg_size.get(skc, {k:0.0 for k in SIZE_KEYS})
            sheet3_data.append({
                'skc': skc,
                'color': wh_colors.get(skc, ''),
                'cur_total': cur_total,
                'prev_total': prev_total,
                'prev_sizes': prev_sizes,
                'cur_sizes': cur_sizes,
            })
    logger.info("库存回补（Sheet3）: %d条", len(sheet3_data))

    # NOTE: 延迟到运行时查找文件，避免 import 时因文件缺失崩溃
    sg_files = sorted(glob.glob(restock_dir + patterns['sg_restock']), key=os.path.getmtime, reverse=True)
    nba_files = sorted(glob.glob(restock_dir + patterns['nba_restock']), key=os.path.getmtime, reverse=True)
    if not sg_files:
        raise FileNotFoundError(f"未找到SG品牌进度表: {restock_dir}{patterns['sg_restock']}")
    nba_file = nba_files[0] if nba_files else restock_dir + patterns['nba_restock'].replace('*', '')
    restock_files = [sg_files[0], nba_file]

    # 加载所有翻单记录（不限制可销数量）
    # key改为 (skc, actual_date)，不同到货日期分开记录
    all_restock = {}  # {(skc, date): {'sizes':..., 'color':..., 'status':..., 'factory':...}}
    skc_has_any_restock = set()  # 用于Sheet2排除：只要有任一翻单记录就算

    # 2a. SG品牌进度表（加载所有翻单记录）
    wb1 = openpyxl.load_workbook(restock_files[0], data_only=True)
    ws1 = wb1['2026年春夏-SG电商进度']
    for row in ws1.iter_rows(min_row=2):
        rd = [c.value for c in row]
        pc = str(rd[2]).strip() if rd[2] else ''
        if rd[18] is not None and rd[18] != '': continue  # 已出货
        color = str(rd[5]) if rd[5] else ''
        mc = re.search(r'\[(\d+)\]', color)
        cc = mc.group(1) if mc else '?'
        skc = pc + cc
        actual_date = get_date_value(rd[21])
        restock_key = (skc, actual_date)
        # SG表：cols 10-14=S/M/L/XL/2XL，无均码列
        if restock_key not in all_restock:
            all_restock[restock_key] = {'sizes': {k:0 for k in SIZE_KEYS}, 'color': color, 'status':'', 'factory':''}
        for i, ci in enumerate([10,11,12,13,14]):
            key = ['S','M','L','XL','2XL'][i]
            all_restock[restock_key]['sizes'][key] += rd[ci] if isinstance(rd[ci],(int,float)) else 0
        if rd[23]: all_restock[restock_key]['status'] = str(rd[23])[:30]
        if rd[6]: all_restock[restock_key]['factory'] = str(rd[6])
        if color: all_restock[restock_key]['color'] = color
        if skc: skc_has_any_restock.add(skc)

    # 2b. SG.NBA订单进度（加载所有翻单记录）
    wb2 = openpyxl.load_workbook(restock_files[1], data_only=True)
    for sheet in wb2.sheetnames:
        ws2 = wb2[sheet]
        header_row = 3 if sheet in ['SG-SS26','NBA联名'] else 2

        # NOTE: 校验表头，防止生产部改表结构后静默出错
        header_cells = [str(c.value).strip() if c.value else '' for c in ws2[header_row]]
        expected_keywords = ['款号', '颜色']
        header_text = ''.join(header_cells)
        if not any(kw in header_text for kw in expected_keywords):
            logger.warning("Sheet [%s] 表头不匹配，跳过（表头: %s）", sheet, header_cells[:10])
            continue

        col_code = 2
        col_ship = 17
        col_color = 5 if header_row == 2 else 8
        col_sizes = [10,11,12,13,14,15]
        col_status = 23
        col_factory = 6
        col_arrival = 21

        for row in ws2.iter_rows(min_row=header_row+1):
            rd = [c.value for c in row]
            pc = str(rd[col_code]).strip() if rd[col_code] else ''
            if rd[col_ship] is not None and rd[col_ship] != '': continue
            color = str(rd[col_color]) if rd[col_color] else ''
            mc = re.search(r'\[(\d+)\]', color)
            cc = mc.group(1) if mc else '?'
            skc = pc + cc
            actual_date = get_date_value(rd[col_arrival])
            restock_key = (skc, actual_date)
            if restock_key not in all_restock:
                all_restock[restock_key] = {'sizes': {k:0 for k in SIZE_KEYS}, 'color': color, 'status':'', 'factory':''}
            for i, ci in enumerate(col_sizes):
                key = ['S','M','L','XL','2XL','均码'][i]
                all_restock[restock_key]['sizes'][key] += rd[ci] if isinstance(rd[ci],(int,float)) else 0
            if rd[col_status]: all_restock[restock_key]['status'] = str(rd[col_status])[:30]
            if rd[col_factory]: all_restock[restock_key]['factory'] = str(rd[col_factory])
            if color: all_restock[restock_key]['color'] = color
            if skc: skc_has_any_restock.add(skc)

    logger.info("翻单表中的SKC总数: %d", len(all_restock))

    # 3. Filter by window 04/12~05/02 using actual dates
    results = []
    for (skc, actual_date), d in all_restock.items():
        if skc not in neg_skcs: continue  # 只看可销<10的SKC
        total = sum(d['sizes'].values())
        if total <= 0: continue
        if actual_date is None: continue
        if not isinstance(actual_date, datetime):
            try:
                actual_date = datetime.strptime(actual_date, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
        if window_start <= actual_date <= window_end:
            d2 = dict(d)
            d2['delivery'] = actual_date.strftime('%Y-%m-%d')
            # 优先用仓库解析的颜色，没有则用生产部颜色
            d2['color'] = wh_colors.get(skc, d.get('color', ''))
            results.append((skc, actual_date, d2))

    results.sort(key=lambda x: (x[2].get('delivery', ''), neg_skcs.get(x[0], 0)))
    logger.info("窗口期到货: %d条", len(results))

    # 4. Unmatched negative SKCs - 排除窗口期（今天前10天）起有翻单记录的SKC
    cutoff = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=7)
    restock_from_apr15 = set()
    for (skc, actual_date), d in all_restock.items():
        if sum(d['sizes'].values()) <= 0:
            continue
        if actual_date and actual_date >= cutoff:
            restock_from_apr15.add(skc)
    logger.info("%s起有翻单的SKC: %d条", cutoff.strftime('%m/%d'), len(restock_from_apr15))
    unmatched = {skc: qty for skc, qty in neg_skcs.items() if skc not in restock_from_apr15}
    logger.info("无翻单负库存: %d条", len(unmatched))

    # 5. Business data scoring
    biz_data = load_business_data(biz_dir, patterns['biz_advisor'])
    scores = score_and_advise(unmatched, biz_data)

    return results, neg_skcs, wh_neg_size, wh_colors, unmatched, scores, output_dir, sheet3_data

def to_excel(results, neg_skcs, wh_neg_size, wh_colors, unmatched, scores, output_dir, sheet3_data):
    from datetime import date
    today = date.today().strftime('%Y%m%d')
    out_path = f"{output_dir}窗口期到货_负库存_{today}.xlsx"

    wb = openpyxl.Workbook()

    hf = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    hfont = Font(color='FFFFFF', bold=True, size=10)
    tb = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))

    # Sheet 1: 窗口期到货
    ws1 = wb.active
    ws1.title = '窗口期到货'
    headers = ['款号','SKC','颜色','仓库可销','XS','S','M','L','XL','2XL','均码',
               '翻单待入库','XS','S','M','L','XL','2XL','均码','实际到仓时间','生产状态','工厂']
    ws1.append(headers)
    for ci, h in enumerate(headers, 1):
        cell = ws1.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb

    for skc, actual_date, d in results:
        restock_total = sum(d['sizes'].values())
        wh_total = neg_skcs.get(skc, 0)
        wh_sizes = wh_neg_size.get(skc, {k:0 for k in SIZE_KEYS})
        wh_color = wh_colors.get(skc, d.get('color', ''))

        row = [
            skc[:8], skc, wh_color,
            round(wh_total, 0),
            0,  # XS（仓库无XS，固定0）
            round(wh_sizes.get('S',0),0), round(wh_sizes.get('M',0),0),
            round(wh_sizes.get('L',0),0), round(wh_sizes.get('XL',0),0),
            round(wh_sizes.get('2XL',0),0), round(wh_sizes.get('均码',0),0),
            round(restock_total, 0),
            0,  # XS（SG/NBA无XS，固定0）
            round(d['sizes'].get('S',0),0), round(d['sizes'].get('M',0),0),
            round(d['sizes'].get('L',0),0), round(d['sizes'].get('XL',0),0),
            round(d['sizes'].get('2XL',0),0), round(d['sizes'].get('均码',0),0),
            d.get('delivery',''), d.get('status',''), d.get('factory','')
        ]
        ws1.append(row)
        rn = ws1.max_row
        for ci in range(1, 23):
            cell = ws1.cell(rn, ci)
            cell.border = tb; cell.alignment = Alignment(horizontal='center', vertical='center')
            fill_color = {'2026-04-15':'FCE4D6','2026-04-20':'FFF2CC'}.get(d['delivery'], 'DAEFCE')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    cw1 = [10,14,10,10,6,6,6,6,6,6,6,10,6,6,6,6,6,6,12,20,8]
    for ci, w in enumerate(cw1, 1): ws1.column_dimensions[get_column_letter(ci)].width = w

    # Sheet 2: 无翻单需评分
    ws2 = wb.create_sheet('无翻单需评分')
    headers2 = ['款号','SKC','仓库可销','综合评分','7天访客','7天支付','7天加购','翻单建议']
    ws2.append(headers2)
    for ci, h in enumerate(headers2, 1):
        cell = ws2.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb
    for skc, qty in sorted(unmatched.items(), key=lambda x: neg_skcs[x[0]]):
        code = skc[:8]
        sc = scores.get(code, {})
        score = sc.get('score')
        advice = sc.get('advice', '⚪ 无数据')
        fill_color = 'F4CCCC' if '❌' in advice else ('FFF2CC' if '⚠️' in advice else 'DAEFCE')
        row = [code, skc, round(qty,0),
               str(score) if score is not None else '无数据',
               sc.get('visitors', 0), sc.get('pay', 0), sc.get('cart', 0), advice]
        ws2.append(row)
        rn = ws2.max_row
        for ci in range(1, 9):
            cell = ws2.cell(rn, ci)
            cell.border = tb; cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
    cw2 = [10,14,10,10,10,10,10,16]
    for ci, w in enumerate(cw2, 1): ws2.column_dimensions[get_column_letter(ci)].width = w

    # Sheet 3: 库存回补明细
    ws3 = wb.create_sheet('库存回补')
    headers3 = ['款号','SKC','颜色','当前可销',
                '前表日期','前表可销','XS前','S前','M前','L前','XL前','2XL前','均码前',
                '后表日期','后表可销','XS后','S后','M后','L后','XL后','2XL后','均码后']
    ws3.append(headers3)
    for ci, h in enumerate(headers3, 1):
        cell = ws3.cell(1, ci)
        cell.fill = hf; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = tb

    from datetime import date as d
    today_str = d.today().strftime('%Y-%m-%d')
    prev_date_str = '前表'

    for d3 in sheet3_data:
        skc = d3['skc']
        cur_sizes = d3['cur_sizes']
        prev_sizes = d3['prev_sizes']
        row = [
            skc[:8], skc, d3['color'],
            round(d3['cur_total'], 0),
            prev_date_str, round(d3['prev_total'], 0),
            round(prev_sizes.get('XS',0),0), round(prev_sizes.get('S',0),0),
            round(prev_sizes.get('M',0),0), round(prev_sizes.get('L',0),0),
            round(prev_sizes.get('XL',0),0), round(prev_sizes.get('2XL',0),0),
            round(prev_sizes.get('均码',0),0),
            today_str, round(d3['cur_total'], 0),
            round(cur_sizes.get('XS',0),0), round(cur_sizes.get('S',0),0),
            round(cur_sizes.get('M',0),0), round(cur_sizes.get('L',0),0),
            round(cur_sizes.get('XL',0),0), round(cur_sizes.get('2XL',0),0),
            round(cur_sizes.get('均码',0),0),
        ]
        ws3.append(row)
        rn = ws3.max_row
        for ci in range(1, 22):
            cell = ws3.cell(rn, ci)
            cell.border = tb; cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color='DAEFCE', end_color='DAEFCE', fill_type='solid')

    cw3 = [10,14,10,10,10,10,6,6,6,6,6,6,6,10,10,6,6,6,6,6,6,6]
    for ci, w in enumerate(cw3, 1): ws3.column_dimensions[get_column_letter(ci)].width = w

    wb.save(out_path)
    logger.info("已保存: %s", out_path)
    return out_path

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='负库存翻单到货分析')
    parser.add_argument('--skc', '--款号', help='只查询指定SKC（支持模糊匹配），如 WE12026208')
    args = parser.parse_args()

    results, neg_skcs, wh_neg_size, wh_colors, unmatched, scores, output_dir, sheet3_data = analyze()

    # 按SKC过滤（可选）
    if args.skc:
        results = [r for r in results if args.skc in r[0]]  # r[0]=SKC
        unmatched = {k: v for k, v in unmatched.items() if args.skc in k}

    if results or unmatched:
        path = to_excel(results, neg_skcs, wh_neg_size, wh_colors, unmatched, scores, output_dir, sheet3_data)
        logger.info("完成: 窗口期到货%d条 | 库存回补%d条 | 无翻单%d条", len(results), len(sheet3_data), len(unmatched))
    else:
        logger.info("无数据")
