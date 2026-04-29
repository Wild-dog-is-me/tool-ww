# -*- coding: utf-8 -*-
from flask import Flask, request, render_template, send_file, session
import os
import sys as _sys
import uuid
import xlrd
import xlwt
import xlutils.copy
import random
import logging
from logging.handlers import RotatingFileHandler
import datetime as dt
import json

if getattr(_sys, 'frozen', False):
    EXE_DIR = os.path.dirname(_sys.executable)
else:
    EXE_DIR = os.path.dirname(os.path.abspath(__file__))

def _load_config():
    config_path = os.path.join(EXE_DIR, 'config.json')
    default_config = {'work_dir': 'data', 'port': 5000, 'host': '127.0.0.1'}
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                raw = f.read()
            import re
            cleaned = re.sub(r'//.*', '', raw)
            cfg = json.loads(cleaned)
            default_config.update(cfg)
        except Exception:
            pass
    return default_config

CONFIG = _load_config()

_work_dir = CONFIG.get('work_dir', 'data')
if os.path.isabs(_work_dir):
    WORK_DIR = _work_dir
else:
    WORK_DIR = os.path.join(EXE_DIR, _work_dir)

def app_path(*parts):
    return os.path.join(WORK_DIR, *parts)

for d in ['log', 'uploads', 'images']:
    os.makedirs(app_path(d), exist_ok=True)

logger = logging.getLogger('recipe_app')
logger.setLevel(logging.DEBUG)

log_file = app_path('log', 'app.log')
fh = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=5, encoding='utf-8')
fh.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
fh.setFormatter(formatter)
logger.addHandler(fh)

logger.info("=" * 60)
logger.info("应用启动")

if getattr(_sys, 'frozen', False):
    _template_dir = os.path.join(_sys._MEIPASS, 'templates')
else:
    _template_dir = os.path.join(EXE_DIR, 'templates')

app = Flask(__name__, template_folder=_template_dir)
app.secret_key = 'supersecretkey'

# 周数到行号的映射表
WEEK_ROW_MAP = {
    1: 1,
    2: 20,
    3: 39,
    4: 57,
    5: 76,
    6: 94,
    7: 113,
    8: 131,
    9: 150,
    10: 168,
    11: 187,
    12: 205,
    13: 224,
    14: 242,
    15: 261,
    16: 279,
    17: 298,
    18: 316,
    19: 335,
    20: 353
}

# 护理方法列表
CARE_METHODS = [
    '垫汗巾',
    '喝水休息',
    '擦汗',
    '适当休息',
    '调整活动强度',
    '补充水分'
]

session_data = {}

# 日期到周几的计算
def get_weekday(year, month, day):
    """计算给定日期是周几"""
    try:
        date = dt.date(year, month, day)
        weekday = date.weekday()
        weekday_names = ['周一', '周二', '周三', '周四', '周五']
        if 0 <= weekday <= 4:
            result = weekday_names[weekday]
            logger.debug(f"[get_weekday] {year}/{month}/{day} -> {result}")
            return result
        logger.warning(f"[get_weekday] {year}/{month}/{day} -> 周末，返回None")
        return None
    except Exception as e:
        logger.error(f"[get_weekday] {year}/{month}/{day} 计算失败: {e}")
        return None

# 百度云OCR识别菜谱图片
def ocr_recipe_image(image_path):
    """使用百度云OCR识别菜谱图片"""
    logger.info(f"[OCR] 开始识别: {image_path}")
    logger.info(f"[OCR] 文件存在: {os.path.exists(image_path)}, 大小: {os.path.getsize(image_path) if os.path.exists(image_path) else 'N/A'}")
    try:
        from aip import AipOcr

        APP_ID = '123067277'
        API_KEY = 'Jg0jsopGOqNVca2ilUWFgLYa'
        SECRET_KEY = '88HbqnzUILByg7c7Ohafk7LJhYbIZokz'

        client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
        logger.debug(f"[OCR] AipOcr客户端初始化成功")

        def get_file_content(filePath):
            with open(filePath, 'rb') as fp:
                return fp.read()

        image = get_file_content(image_path)
        logger.debug(f"[OCR] 图片读取成功, 字节数: {len(image)}")

        result = client.basicAccurate(image)
        logger.debug(f"[OCR] API返回keys: {list(result.keys()) if result else 'None'}")

        if 'words_result' in result:
            text = ''
            for item in result['words_result']:
                text += item['words'] + ' '
            final_text = text.strip()
            logger.info(f"[OCR] 识别成功, 文本长度: {len(final_text)}, 行数: {result.get('words_result_num', 'N/A')}")
            logger.info(f"[OCR] 识别文本: {final_text}")
            return final_text
        else:
            err_msg = result.get('error_msg', '未知错误')
            err_code = result.get('error_code', 'N/A')
            logger.error(f"[OCR] API返回错误: code={err_code}, msg={err_msg}")
            logger.error(f"[OCR] 完整返回: {result}")
            return ''
    except Exception as e:
        logger.error(f"[OCR] 异常: {e}")
        import traceback
        logger.error(f"[OCR] 堆栈:\n{traceback.format_exc()}")
        return ''

# 从空格分隔的文本中按列索引提取值
# OCR文本每行格式: "标签 v1 v2 v3 v4 v5" (5列=周一到周五)
# 或双行格式: "标签 n1 n2 n3 n4 n5 q1 q2 q3 q4 q5" (名称+数量)
def pick_column_values(section_text, col_idx, num_cols=5):
    """从OCR文本片段中按列索引提取值
    参数：
        section_text: 包含标签和值的文本片段
        col_idx: 目标列索引 (0-4 对应 周一-周五)
        num_cols: 总列数，默认5（周一至周五）
    返回：
        提取到的值列表（可能1个或2个元素，对应单行或双行格式）
    """
    parts = section_text.split()
    if len(parts) <= 1:
        return []
    values = parts[1:]
    total_values = len(values)
    if total_values < num_cols:
        return []
    num_rows = total_values // num_cols if total_values % num_cols == 0 else 1
    result = []
    for row in range(num_rows):
        idx = row * num_cols + col_idx
        if idx < len(values):
            val = values[idx]
            val = val.replace(' ', '')  # 去除OCR可能插入的空格
            result.append(val)
    return result

def extract_recipe_data(ocr_text, weekday):
    """从OCR文本中提取菜谱数据
    采用标签分段 + 模式匹配的混合策略，兼容不同OCR输出布局:
    - Week2: 早点 牛奶×5 饼干×5 米饭×5 主食 稻米×5 ...
    - Week5: 牛奶×5 早点 饼干×5 主食 米饭×5 稻米×5 ...
    """
    import re
    recipe = {'早点': '', '午餐': '', '午点': '', '蔬菜': ''}

    try:
        logger.info(f"[EXTRACT] weekday={weekday}, ocr_text长度={len(ocr_text)}")
        if not ocr_text:
            logger.warning("[EXTRACT] OCR文本为空")
            return recipe

        weekday_index = {'周一': 0, '周二': 1, '周三': 2, '周四': 3, '周五': 4}
        if weekday not in weekday_index:
            logger.warning(f"[EXTRACT] weekday '{weekday}' 不在工作日范围")
            return recipe

        col = weekday_index[weekday]
        logger.info(f"[EXTRACT] 目标列: {weekday} (col={col})")

        tokens = ocr_text.split()
        logger.info(f"[EXTRACT] 总token数: {len(tokens)}")

        LABELS = ['早点', '主食', '荤菜', '素菜', '汤', '水果', '午点']
        label_map = {}
        for i, tok in enumerate(tokens):
            if tok in LABELS and tok not in label_map:
                label_map[tok] = i

        if not label_map:
            logger.error(f"[EXTRACT] 未找到菜谱标签! tokens前20: {tokens[:20]}")
            return recipe

        logger.info(f"[EXTRACT] 标签位置: {[(l, p) for l, p in label_map.items()]}")

        def get_section(label_name):
            if label_name not in label_map:
                return []
            pos = label_map[label_name]
            next_pos = min([p for l, p in label_map.items() if p > pos], default=len(tokens))
            return tokens[pos + 1:next_pos]

        def at_col(items, default=''):
            return items[col] if 0 <= col < len(items) else default

        breakfast_parts = []
        lunch_parts = []
        snack_parts = []

        # === 早点: 全局扫描牛奶/饼干（兼容牛奶在标签前或标签后的情况） ===
        milk_pat = re.compile(r'牛奶\d+克')
        cookie_pat = re.compile(r'饼干\d+克')
        all_milk = [t for t in tokens if milk_pat.match(t)]
        all_cookie = [t for t in tokens if cookie_pat.match(t)]
        logger.debug(f"[EXTRACT] 牛奶: {all_milk}, 饼干: {all_cookie}")

        milk = at_col(all_milk) if len(all_milk) >= 5 else ''
        cookie = at_col(all_cookie) if len(all_cookie) >= 5 else ''
        if milk:
            breakfast_parts.append(milk)
        if cookie:
            breakfast_parts.append(cookie)
        if breakfast_parts:
            recipe['早点'] = '、'.join(breakfast_parts)
            logger.info(f"[EXTRACT] {weekday}早点: {recipe['早点']}")

        # === 主食 ===
        bean_sec = get_section('早点')
        staple_sec = get_section('主食')
        rice_pat = re.compile(r'.*米饭$')
        rice_amount_pat = re.compile(r'.*稻米.*\d+克')

        rice_types_from_bean = [t for t in bean_sec if rice_pat.match(t)]
        rice_types_from_staple = [t for t in staple_sec if rice_pat.match(t)]
        all_rice_types = rice_types_from_bean if len(rice_types_from_bean) >= 5 else rice_types_from_staple
        rice_type = at_col(all_rice_types) if len(all_rice_types) >= 5 else ''
        logger.debug(f"[EXTRACT] 米饭类型: {all_rice_types} -> col{col}={repr(rice_type)}")

        all_rice_amounts = [t for t in staple_sec if rice_amount_pat.match(t)]
        rice_amount = at_col(all_rice_amounts) if len(all_rice_amounts) >= 5 else ''
        logger.debug(f"[EXTRACT] 稻米量: {all_rice_amounts} -> col{col}={repr(rice_amount)}")

        if rice_type and rice_amount:
            lunch_parts.append(f'{rice_type}{rice_amount}')
            logger.info(f"[EXTRACT] {weekday}主食: {rice_type}{rice_amount}")

        # === 荤菜 ===
        meat_sec = get_section('荤菜')
        if len(meat_sec) >= 10:
            meat_name = at_col(meat_sec[:5])
            meat_amount = at_col(meat_sec[5:10])
            if meat_name and meat_amount:
                lunch_parts.append(f'{meat_name}{meat_amount}')
                logger.info(f"[EXTRACT] {weekday}荤菜: {meat_name}{meat_amount}")

        # === 素菜 ===
        veg_sec = get_section('素菜')
        if len(veg_sec) >= 5:
            veg_name = at_col(veg_sec[:5])
            recipe['蔬菜'] = veg_name
        if len(veg_sec) >= 10:
            veg_amount = at_col(veg_sec[5:10])
            if veg_name and veg_amount:
                lunch_parts.append(f'{veg_name}{veg_amount}')
                logger.info(f"[EXTRACT] {weekday}素菜: {veg_name}{veg_amount}")

        # === 汤: 按模式分离名称(含"汤")和数量(克模式) ===
        soup_sec = get_section('汤')
        soup_name_pat = re.compile(r'汤$')
        amount_pat = re.compile(r'\d+克(\d+克)*$|每人\d+个')

        soup_names = [t for t in soup_sec if soup_name_pat.search(t)]
        soup_amounts_all = [t for t in soup_sec if amount_pat.match(t) and not soup_name_pat.search(t)]
        soup_leftover = [t for t in soup_sec if t not in soup_names and t not in soup_amounts_all]
        logger.debug(f"[EXTRACT] 汤names: {soup_names}, amounts: {soup_amounts_all}, leftover: {soup_leftover}")

        if len(soup_names) >= 5 and len(soup_amounts_all) >= 5:
            # 修复OCR错位: 若某个amount出现在最后一个name之前，旋转对齐
            last_name_pos = max(soup_sec.index(n) for n in soup_names if n in soup_sec)
            shift = sum(1 for a in soup_amounts_all if soup_sec.index(a) < last_name_pos)
            if shift > 0:
                soup_amounts_all = soup_amounts_all[shift:] + soup_amounts_all[:shift]
                logger.debug(f"[EXTRACT] 汤amounts旋转{shift}位: {soup_amounts_all}")
            soup_name = at_col(soup_names)
            soup_amount = at_col(soup_amounts_all)
            if soup_name and soup_amount:
                lunch_parts.append(f'{soup_name}{soup_amount}')
                logger.info(f"[EXTRACT] {weekday}汤: {soup_name}{soup_amount}")

        if lunch_parts:
            recipe['午餐'] = '、'.join(lunch_parts)
            logger.info(f"[EXTRACT] {weekday}午餐: {recipe['午餐']}")

        # === 水果: 合并汤段溢出token + 水果段，按模式匹配 ===
        fruit_sec_raw = soup_leftover + get_section('水果')
        # 把 粥/南瓜/馒头/红薯/炮 类型的午点token移到午点段
        snack_keywords = re.compile(r'粥|南瓜|馒头|红薯|炮')
        fruit_sec = []
        extra_snack = []
        for t in fruit_sec_raw:
            if snack_keywords.search(t):
                extra_snack.append(t)
            else:
                fruit_sec.append(t)
        logger.debug(f"[EXTRACT] 水果段: {fruit_sec}, 溢出到午点: {extra_snack}")

        fruit_combined_pat = re.compile(r'.+\d+克$')
        fruit_combined = [t for t in fruit_sec if fruit_combined_pat.match(t) and '汤' not in t]
        logger.debug(f"[EXTRACT] 水果combined: {fruit_combined}")

        if len(fruit_combined) >= 5:
            fruit = at_col(fruit_combined)
        else:
            fruit_amounts = [t for t in fruit_sec if amount_pat.match(t) and '汤' not in t]
            fruit_names = [t for t in fruit_sec if not amount_pat.match(t)]
            logger.debug(f"[EXTRACT] 水果names: {fruit_names}, amounts: {fruit_amounts}")
            if len(fruit_names) >= 5:
                f_name = at_col(fruit_names)
                f_amount = at_col(fruit_amounts) if col < len(fruit_amounts) else ''
                fruit = f'{f_name}{f_amount}' if f_amount else f_name
            elif len(fruit_sec) >= 5:
                fruit = at_col(fruit_sec)
            else:
                fruit = ''

        if fruit:
            snack_parts.append(fruit)
            logger.info(f"[EXTRACT] {weekday}水果: {fruit}")

        # === 午点: 合并水果段溢出token + 午点段，按模式分离 ===
        snack_sec = extra_snack + get_section('午点')
        snack_amounts = [t for t in snack_sec if amount_pat.match(t)]
        snack_names_raw = [t for t in snack_sec if t not in snack_amounts]

        logger.debug(f"[EXTRACT] 午点sec: {snack_sec}")
        logger.debug(f"[EXTRACT] 午点names: {snack_names_raw}, amounts: {snack_amounts}")

        if len(snack_names_raw) >= 5 and len(snack_amounts) >= 5:
            s_name = at_col(snack_names_raw)
            s_amount = at_col(snack_amounts)
        else:
            s_name = at_col(snack_sec[:5]) if len(snack_sec) >= 5 else ''
            s_amount = at_col(snack_sec[5:10]) if len(snack_sec) >= 10 else ''

        if s_name and s_amount:
            snack_parts.append(f'{s_name}{s_amount}')
            logger.info(f"[EXTRACT] {weekday}午点: {s_name}{s_amount}")

        if snack_parts:
            recipe['午点'] = '、'.join(snack_parts)
            logger.info(f"[EXTRACT] {weekday}午点: {recipe['午点']}")

        logger.info(f"[EXTRACT] 最终recipe: {recipe}")
        logger.info(f"[EXTRACT] all_values={all(recipe.values())}")

    except Exception as e:
        logger.error(f"[EXTRACT] 异常: {e}")
        import traceback
        logger.error(f"[EXTRACT] 堆栈:\n{traceback.format_exc()}")

    return recipe

# 数字转中文
def num_to_chinese(num):
    chinese_nums = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
    if num <= 10:
        return chinese_nums[num]
    elif num < 20:
        return '十' + chinese_nums[num-10]
    elif num < 100:
        return chinese_nums[num//10] + '十' + (chinese_nums[num%10] if num%10 !=0 else '')
    return str(num)

# 创建宋体12号带边框样式
def create_style():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = '宋体'
    font.height = 12 * 20
    style.font = font
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    return style

# 创建居中样式
def create_center_style():
    style = create_style()
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment
    return style

# 创建红色宋体12号带边框样式
def create_red_style():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = '宋体'
    font.height = 12 * 20
    font.colour_index = 2
    style.font = font
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    return style

# 填充单条记录
def fill_record(source_path, target_path, week, month, day, recorder_name, recipe_data=None):
    # 确保参数类型正确
    week = int(week)
    month = int(month)
    day = int(day)
    
    wb = xlrd.open_workbook(source_path, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    
    # 获取目标行
    if week not in WEEK_ROW_MAP:
        raise ValueError(f"周数 {week} 超出范围（1-20周）")
    target_row = WEEK_ROW_MAP[week]
    
    # 检查目标行是否存在
    if target_row >= sheet.nrows:
        raise ValueError(f"第{week}周的目标行 {target_row} 超出文件范围")
    
    wb_new = xlutils.copy.copy(wb)
    sheet_new = wb_new.get_sheet(0)
    
    # 填充周数和记录人（居中）
    chinese_week = num_to_chinese(week)
    new_text = f"第   {chinese_week}  周   {month}  月  {day}   日                    记录人：{recorder_name}           "
    sheet_new.write(target_row, 0, new_text, create_center_style())
    
    # 填充食谱信息（进食食物种类、进餐量下方：早点/午餐/午点）
    if recipe_data:
        # 填充早点（在目标行+13的位置）
        if target_row + 13 < sheet.nrows:
            breakfast = recipe_data.get('早点', '')
            if breakfast:
                new_text = f"早点：{breakfast}"
                sheet_new.write(target_row + 13, 0, new_text, create_style())
        
        # 填充午餐（在目标行+14的位置）
        if target_row + 14 < sheet.nrows:
            lunch = recipe_data.get('午餐', '')
            if lunch:
                new_text = f"午餐：{lunch}"
                sheet_new.write(target_row + 14, 0, new_text, create_style())
        
        # 填充午点（在目标行+15的位置）
        if target_row + 15 < sheet.nrows:
            snack = recipe_data.get('午点', '')
            if snack:
                new_text = f"午点：{snack}"
                sheet_new.write(target_row + 15, 0, new_text, create_style())

        # 填充2.进食情况（喜欢程度、进食快慢、挑食情况）
        if target_row + 16 < sheet.nrows:
            vegetable = recipe_data.get('蔬菜', '')
            if vegetable:
                PREFIX = '2.进食情况（喜欢程度、进食快慢、挑食情况）:'
                original_text = sheet.cell(target_row + 16, 0).value or ''
                if PREFIX in original_text:
                    original_text = original_text.split(PREFIX)[0] + PREFIX
                eating_text = f"{original_text}喜欢、快、不喜欢吃{vegetable}"
                sheet_new.write(target_row + 16, 0, eating_text, create_red_style())
                logger.info(f"[FILL] 进食情况: {eating_text}")
    
    # 填充运动强度（加大打勾）
    if target_row + 4 < sheet.nrows:
        original_text = sheet.cell(target_row + 4, 0).value
        if original_text and '运动强度：' in original_text:
            # 只处理当前周的运动强度，不影响其他周
            new_text = original_text
            # 先移除当前行的所有勾
            new_text = new_text.replace('□✔', '□')
            # 只在加大后面打勾
            if '加大 □' in new_text:
                new_text = new_text.replace('加大 □', '加大 □✔')
            sheet_new.write(target_row + 4, 0, new_text, create_style())
    
    # 填充运动密度（随机增多1-2次）
    if target_row + 5 < sheet.nrows:
        original_text = sheet.cell(target_row + 5, 0).value
        if original_text and '运动密度：' in original_text:
            increase_times = random.randint(1, 2)
            new_text = original_text.replace('增多   次', f'增多 {increase_times} 次')
            sheet_new.write(target_row + 5, 0, new_text, create_style())
    
    # 填充护理方法（随机选择）
    if target_row + 6 < sheet.nrows:
        care_method = random.choice(CARE_METHODS)
        sheet_new.write(target_row + 6, 0, care_method, create_style())
    
    # 填充心理状况（只勾选喜欢）
    import re
    if target_row + 8 < sheet.nrows:
        original_text = sheet.cell(target_row + 8, 0).value
        if original_text and '心理状况：' in original_text:
            # 只处理当前周的心理状况，不影响其他周
            new_text = original_text
            # 先移除当前行的所有勾
            new_text = new_text.replace('□✔', '□')
            # 使用正则表达式精确匹配独立的"喜欢"选项
            new_text = re.sub(r'(?<!勉强)喜欢 □', '喜欢 □✔', new_text)
            # 确保不匹配"不喜欢"中的"喜欢"
            new_text = re.sub(r'不喜欢 □✔', '不喜欢 □', new_text)
            sheet_new.write(target_row + 8, 0, new_text, create_style())
    
    # 填充生理状况（正常或汗多打勾）
    if target_row + 9 < sheet.nrows:
        original_text = sheet.cell(target_row + 9, 0).value
        if original_text and '生理状况：' in original_text:
            # 只处理当前周的生理状况，不影响其他周
            new_text = original_text
            # 先移除当前行的所有勾
            new_text = new_text.replace('□✔', '□')
            if random.random() < 0.5:
                # 正常
                if '正常 □' in new_text:
                    new_text = new_text.replace('正常 □', '正常 □✔')
            else:
                # 汗多
                if '汗多 □' in new_text:
                    new_text = new_text.replace('汗多 □', '汗多 □✔')
            sheet_new.write(target_row + 9, 0, new_text, create_style())
    
    # 填充总体评价（全部打勾）
    if target_row + 11 < sheet.nrows:
        original_text = sheet.cell(target_row + 11, 0).value
        if original_text and '运动强度是否合适' in original_text:
            # 只处理当前周的总体评价，不影响其他周
            new_text = original_text
            # 先移除当前行的所有勾
            new_text = new_text.replace('（   ✔   ）', '（      ）')
            # 确保只替换3个空括号
            count = 0
            while '（      ）' in new_text and count < 3:
                new_text = new_text.replace('（      ）', '（   ✔   ）', 1)
                count += 1
            sheet_new.write(target_row + 11, 0, new_text, create_style())
    
    wb_new.save(target_path)

@app.route('/')
def index():
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    session_id = session['session_id']
    
    if session_id not in session_data:
        session_data[session_id] = {
            'uploaded_file': None,
            'records': [],
            'recorder_name': ''
        }
    
    return render_template('index.html')

@app.route('/api/students')
def get_students():
    students = CONFIG.get('students', [])
    return app.response_class(json.dumps(students, ensure_ascii=False), mimetype='application/json')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    
    if file:
        session_id = session['session_id']
        filename = f"{session_id}_uploaded.xls"
        file_path = app_path('uploads', filename)
        
        file.save(file_path)
        session_data[session_id]['uploaded_file'] = file_path
        session_data[session_id]['original_filename'] = file.filename
        return 'File uploaded successfully'

@app.route('/add_record', methods=['POST'])
def add_record():
    session_id = session['session_id']
    week = int(request.form['week'])
    month = int(request.form['month'])
    day = int(request.form['day'])
    recorder_name = request.form['recorder_name']

    logger.info(f"[ADD_RECORD] session={session_id[:8]}..., week={week}, month={month}, day={day}, recorder={recorder_name}")

    images = request.files.getlist('images[]')
    image_paths = []
    logger.info(f"[ADD_RECORD] 接收到 {len(images)} 张图片")

    if images:
        for i, image in enumerate(images):
            if image and image.filename:
                image_filename = f"{session_id}_recipe_{week}_{month}_{day}_{i}.jpg"
                image_path = app_path('images', image_filename)
                logger.info(f"[ADD_RECORD] 保存图片[{i}]: filename={image.filename}, path={image_path}")
                image.save(image_path)
                image_paths.append(image_path)
                logger.info(f"[ADD_RECORD] 图片保存完成, 文件大小: {os.path.getsize(image_path)}")
            else:
                logger.warning(f"[ADD_RECORD] 图片[{i}]无效: image={image is not None}, filename={image.filename if image else 'N/A'}")
    else:
        logger.info("[ADD_RECORD] 未上传任何图片")

    session_data[session_id]['recorder_name'] = recorder_name
    session_data[session_id]['records'].append({
        'week': week,
        'month': month,
        'day': day,
        'images': image_paths
    })

    logger.info(f"[ADD_RECORD] 记录添加成功, 当前共{len(session_data[session_id]['records'])}条记录")
    logger.info(f"[ADD_RECORD] 本轮图片路径: {image_paths}")

    return 'Record added successfully'

@app.route('/fill_excel', methods=['POST'])
def fill_excel():
    session_id = session['session_id']
    data = session_data[session_id]

    logger.info("=" * 60)
    logger.info(f"[FILL] 开始填充, session={session_id[:8]}...")
    logger.info(f"[FILL] 记录数: {len(data['records'])}, 记录人: {data['recorder_name']}")

    if not data['uploaded_file']:
        logger.error("[FILL] 未上传Excel文件")
        return 'Please upload a file first'

    if not data['records']:
        logger.error("[FILL] 没有记录可填充")
        return 'Please add at least one record'

    recorder_name = request.form.get('recorder_name', '').strip() or data['recorder_name']
    target_file = f"{session_id}_filled.xls"
    target_path = app_path('uploads', target_file)

    import shutil
    shutil.copy(data['uploaded_file'], target_path)
    logger.info(f"[FILL] 复制模板: {data['uploaded_file']} -> {target_path}")

    DEFAULT_RECIPE = {
        '早点': '牛奶120克、饼干10克',
        '午餐': '米饭稻米60克、鸡腿每人一只、炒西兰花80克、冬瓜海带虾皮汤60克10克2克',
        '午点': '苹果50克、南瓜粥80克10克',
        '蔬菜': '西兰花'
    }

    for idx, record in enumerate(data['records']):
        week = record['week']
        month = record['month']
        day = record['day']

        logger.info(f"")
        logger.info(f"[FILL] ===== 记录[{idx}]: 第{week}周 {month}月{day}日 =====")
        logger.info(f"[FILL] record完整内容: {record}")

        recipe_data = None

        images = record.get('images', [])
        logger.info(f"[FILL] record.images: {images} (数量={len(images)})")

        if not images:
            logger.warning(f"[FILL] 记录[{idx}] 无图片 -> 将使用默认食谱")
        else:
            image_path = images[0]
            logger.info(f"[FILL] 使用图片: {image_path}")
            logger.info(f"[FILL] 图片文件存在: {os.path.exists(image_path)}")

            if not os.path.exists(image_path):
                logger.error(f"[FILL] 图片文件不存在! path={image_path}")
            else:
                ocr_text = ocr_recipe_image(image_path)
                logger.info(f"[FILL] OCR完成: 成功={bool(ocr_text)}, 文本长度={len(ocr_text)}")

                if not ocr_text:
                    logger.warning(f"[FILL] OCR识别失败/返回空 -> 将使用默认食谱")
                else:
                    current_year = dt.datetime.now().year
                    weekday = get_weekday(current_year, month, day)
                    logger.info(f"[FILL] 日期计算: {current_year}/{month}/{day} -> weekday={weekday}")

                    if not weekday:
                        logger.warning(f"[FILL] 非工作日 ({current_year}/{month}/{day} 是周末) -> 将使用默认食谱")
                    else:
                        recipe_data = extract_recipe_data(ocr_text, weekday)
                        logger.info(f"[FILL] extract_recipe_data返回: {recipe_data}")
                        logger.info(f"[FILL] all(recipe_data.values()): {all(recipe_data.values())}")
                        logger.info(f"[FILL] 各字段非空: 早点={bool(recipe_data.get('早点'))}, 午餐={bool(recipe_data.get('午餐'))}, 午点={bool(recipe_data.get('午点'))}")

        if not recipe_data or not all(recipe_data.values()):
            logger.warning(f"[FILL] 触发兜底逻辑: not recipe_data={not recipe_data}, not all_values={not all(recipe_data.values()) if recipe_data else 'N/A'}")
            recipe_data = DEFAULT_RECIPE.copy()
            logger.info(f"[FILL] 使用默认食谱: {recipe_data}")
        else:
            logger.info(f"[FILL] 使用OCR食谱: {recipe_data}")

        logger.info(f"[FILL] 最终填充数据: {recipe_data}")
        fill_record(target_path, target_path, week, month, day, recorder_name, recipe_data)
        logger.info(f"[FILL] 记录[{idx}] 填充完成")

    download_filename = data.get('original_filename', 'filled_excel.xls')
    logger.info(f"[FILL] 全部填充完成, 下载文件名: {download_filename}")
    logger.info("=" * 60)
    return send_file(target_path, as_attachment=True, download_name=download_filename)

if __name__ == '__main__':
    import webbrowser, threading
    _host = CONFIG.get('host', '127.0.0.1')
    _port = CONFIG.get('port', 5000)
    def open_browser():
        webbrowser.open(f'http://{_host}:{_port}')
    threading.Timer(1.0, open_browser).start()
    app.run(host=_host, port=_port, debug=False)