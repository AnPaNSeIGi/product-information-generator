# =============================================================================
# 乐天市场商品上架自动化系统（合并增强版 - 修复标题变体生成）
# 功能：图片识别 → AI 生成文案 → HTML 组装 → Excel 输出
# 合并策略：代码一框架 + 代码二 AI 引擎
# =============================================================================
# ---------- 成本计算 ----------
# 是否启用成本追踪
ENABLE_COST_TRACKING = True

# 模型单价（单位：元/千 token）
# 请根据阿里云 DashScope 官网最新价格填写：
# https://help.aliyun.com/zh/dashscope/developer-reference/api-details
INPUT_TOKEN_PRICE = 0.008   # qwen-vl-max 输入价格（示例，请核实）
OUTPUT_TOKEN_PRICE = 0.016  # qwen-vl-max 输出价格（示例，请核实）

import os
import time
import base64
import re
import logging
import uuid
import random
import json
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from natsort import natsorted

# Excel 处理
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# 数据处理
import pandas as pd

# AI 调用（统一使用 Dashscope 原生 SDK）
import dashscope
from dashscope import MultiModalConversation

# =============================================================================
# ================= 动态路径（自动定位项目根目录） =================
# =============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)

# =============================================================================
# ================= 用户配置区（请根据实际情况修改） =================
# =============================================================================

# ---------- 必填 ----------
# 阿里云 DashScope API Key（获取：https://dashscope.console.aliyun.com/）
DASHSCOPE_API_KEY = "sk-xxx"   # 请替换为你的真实 Key

# 源目录：存放商品子文件夹（每个子文件夹放一个商品的图片）
MAIN_DIR = os.path.join(PROJECT_ROOT, "data", "source")

# 输出目录：Excel和报告将保存在这里
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "data", "output")

# 关键词库路径（如果不需要生成标题，设为 None）
KEYWORD_LIB_PATH = os.path.join(PROJECT_ROOT, "data", "keyword_library.json")

# 日志文件存放路径（留空则自动生成到 OUTPUT_DIR）
LOG_DIR = None      # 若设为 None，则日志保存在 OUTPUT_DIR 下
# ---------- 店铺配置 ----------
# 请根据你的乐天店铺信息修改以下内容
SHOP_NAME = "X店"                      # 店铺名称
SHOP_WELCOME = "楽天市場店へようこそ！ご愛顧いたきありがとうございます"  # 欢迎语
SHOP_CODE = "X店"                     # 店铺代码（仅用于显示）
SHOP_LINK_PREFIX = "https://image.rakuten.co.jp/your_shop/cabinet/xxxxx/"  # 图片链接前缀(xxxxx中可为多层文件夹)

# =============================================================================
# ================= 配置区域 =================
# =============================================================================

# 模型配置（代码二使用 qwen-vl-max，支持多图片输入）
MODEL_NAME = "qwen-vl-max"

# 图片限制
MAX_IMAGES_FOR_AI = 20  # 传给 AI 的最大图片数量
MAX_VARIATIONS = 5      # 文案变体上限（超过则循环使用）

# 速率限制（防 API 限流）
API_DELAY_SECONDS = 2   # 每个商品处理间隔秒数

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()   # 只保留控制台输出
    ]
)
logger = logging.getLogger(__name__)

# =============================================================================
# ================= HTML 模板（代码一完整版） =================
# 包含注意事项配置区（仅简单替换文字即可）
# 包含品质保证配置区
# =============================================================================

HTML_TEMPLATE = """<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td align="center" bgcolor="#e6f7ff" bordercolor="#bae7ff">
<font size="2">{shop_welcome}</font>
</td>
</tr>

<tr>
<td>
<table width="100%" border="1" cellpadding="8" cellspacing="0" bordercolor="#e0e0e0">
<tr>
<td bgcolor="#1890ff">
<font color="#ffffff"><b>● 商品特徴</b></font>
</td>
</tr>
<tr>
<td>
{product_features}
</td>
</tr>
</table>
</td>
</tr>

<tr>
<td>
<table width="100%" border="1" cellpadding="8" cellspacing="0" bordercolor="#e0e0e0">
<tr>
<td bgcolor="#52c41a">
<font color="#ffffff"><b>● 商品仕様</b></font>
</td>
</tr>
<tr>
<td>
{product_specs}
</td>
</tr>
</table>
</td>
</tr>

<tr>
<td>
<table width="100%" border="1" cellpadding="8" cellspacing="0" bordercolor="#e0e0e0">
<tr>
<td bgcolor="#faad14">
<font color="#ffffff"><b>● 注意事項</b></font>
</td>
</tr>
<tr>
<td>
<b>※配送について</b><br>
当店はいま配送日時指定対応できません。設置しても対応しませんので、予めご了承ください。<br><br>
<b>※使用時の破損</b><br>
使用時の破損や事故等につきましては責任を負いかねます。<br><br>
<b>※箱の状態</b><br>
海外輸入製品のため、箱が潰れる可能性がありますが、お気になる方はご注意ください。<br><br>
<b>※ラッピング</b><br>
ラッピングは対応しておりません。<br><br>
<b>※仕様変更</b><br>
予告なく仕様変更になる場合がございます。<br><br>
<b>※言語表記</b><br>
日本語表記はございません。パッケージ言語は中国語もしくは英語です。<br><br>
<b>※輸送ダメージ</b><br>
輸入品の為、輸送の際に生じるキズ・汚れ・箱潰れがある場合がございますが新品です。<br><br>
<b>※サイズ誤差</b><br>
サイズ・重量は、多少の誤差が生じる場合がございます。<br><br>
<b>※イメージ違い</b><br>
あくまで画像はイメージですので、商品改良の為パッケージや一部仕様が少し変更になる場合もございます。<br><br>
<b>※色の違い</b><br>
お使いのブラウザや設定により、画像と実際の商品との若干の色の違いが生じる場合がございます。<br><br>
<b>※送料</b><br>
基本は全国一律発送無料と対応させていただきますが、沖縄・へお届け場合は 2000 円の送料をご請求させいただきます。<br><br>
<b>※評価について</b><br>
こちらの対応と商品について何かご不満がありましたら、「悪い評価」を付ける前に一度当方とご連絡ください。出品者の誠意を持って最後まで対応いたします。
</td>
</tr>
</table>
</td>
</tr>

<tr>
<td>
<table width="100%" border="1" cellpadding="8" cellspacing="0" bordercolor="#e0e0e0">
<tr>
<td bgcolor="#722ed1">
<font color="#ffffff"><b>● 品質保証</b></font>
</td>
</tr>
<tr>
<td>
<b>※新品保証</b><br>
出品する商品は全て新品未使用です。<br><br>
<b>※検品について</b><br>
販売している商品は全て新品未使用です、倉庫から出荷前に商品検査必要ですので、商品箱を開封する場合もございます、予めご了承くださいいませ。<br><br>
<b>※初期不良対応</b><br>
初期不良の場合は到着後 1 週間以内にご連絡ください。<br><br>
<b>※交換・返金</b><br>
初期不良の場合は写真やビデオをご提供下さい、確認後無料で新品交換もしくは御返金致しますのでご連絡下さい。<br><br>
<font color="#fa541c"><b>※ご注意</b></font><br>
写真やビデオをご提供出来ない場合は技術者と確認できません、返品返金対応出来ない可能性が御座います、予めご了承ください。
</td>
</tr>
</table>
</td>
</tr>

<tr>
<td>
<table width="100%" border="1" cellpadding="8" cellspacing="0" bordercolor="#e0e0e0">
<tr>
<td bgcolor="#8c8c8c">
<font color="#ffffff"><b>● 関連キーワード</b></font>
</td>
</tr>
<tr>
<td>
{product_keywords}
</td>
</tr>
</table>
</td>
</tr>

<tr>
<td align="center">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td bgcolor="#f9f9f9" bordercolor="#e8e8e8">
<font size="2"><b>ご購入を検討いただきありがとうございます。<br>より良いサービスを提供できるよう努めて参ります。</b></font>
</td>
</tr>
</table>
</td>
</tr>
</table>"""


# SEO 同义词库（用于标题变体生成）- 扩展版
# 可以丰富标题质量产出
SEO_SYNONYMS = {
    "高品質": ["高質感", "プレミアム", "上質", "高級感ある", "優れた品質の", "極上の", "洗練された"],
    "便利": ["使いやすい", "便利な", "手軽な", "シンプルな", "効率的な", "快適な", "スムーズな"],
    "実用的": ["実用性抜群", "日常使い", "実用重視", "機能性", "多目的", "汎用性", "活用度"],
    "耐久性": ["長持ち", "耐久性抜群", "頑丈", "長期間使用可能", "丈夫な", "堅牢な", "長寿命"],
    "おすすめ": ["人気", "注目", "定番", "定番商品", "話題の", "評判の", "支持の"],
    "送料無料": ["送料込", "無料配送", "配送料無料", "送料無料対応", "送料サービス", "配送無料"],
    "新商品": ["新作", "新登場", "新入荷", "NEW", "最新", "新リリース", "新規"],
    "限定": ["特別", "スペシャル", "特別仕様", "限定モデル", "特別版", "希少", "レア"],
    "快適": ["快適な", "スムーズな", "心地よい", "ストレスフリー", "快調な", "良好な"],
    "安心": ["安全", "信頼", "安心して", "確かな", "信頼できる", "確実な", "安定した"],
    "軽量": ["軽い", "軽量化", "スリム", "コンパクト", "持ち運びやすい", "携帯性"],
    "大容量": ["多容量", "たっぷり", "十分な容量", "広々", "余裕の容量", "大型"],
    "シンプル": ["簡易", "単純", "ミニマル", "すっきり", "無駄のない", "清潔感のある"],
    "おしゃれ": ["スタイリッシュ", "モダン", "洗練", "センスのいい", "魅力的な", "美しい"],
    "丈夫": ["頑丈", "堅牢", "強固", "しっかり", "耐久性", "長持ちする", "壊れにくい"]
}

# 标题结构组件（用于排列组合）
TITLE_COMPONENTS = {
    "开头": ["【送料無料】", "【高品質】", "【新登場】", "【限定】", "【人気】", ""],
    "结尾": [" 送料無料", " 高品質", " 新作", " 限定品", " おすすめ", " 人気商品", ""]
}

# 禁止品牌词列表（防止侵权）
FORBIDDEN_BRANDS = [
    "Nike", "Adidas", "Apple", "Samsung", "Sony", "Panasonic", "Toyota", 
    "Honda", "Canon", "Nikon", "LG", "Philips", "Dyson", "Xiaomi", "Huawei",
    "無印良品", "ニトリ", "IKEA", "ZOZOTOWN", "Amazon", "楽天", "ヤフー",
    "ユニクロ", "GU", "ZARA", "H&M"
]

# =============================================================================
# ================= 工具函数 =================
# =============================================================================

def count_fullwidth_chars(text):
    """计算全角字符数"""
    count = 0
    for char in text:
        if ord(char) > 126:
            count += 1
        else:
            count += 0.5
    return int(count)

def sanitize_text_for_html(text):
    """
    清理文本，移除特殊符号，适配 HTML 和乐天规则
    保留<br>标签，只转义其他 HTML 标签
    """
    if not text:
        return ""
    
    # 先保护<br>标签，替换为临时占位符
    text = text.replace("<br>", "{{BR_TAG}}")
    text = text.replace("<br/>", "{{BR_TAG}}")
    text = text.replace("<br />", "{{BR_TAG}}")
    
    # 移除特殊符号
    special_chars = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩',
                     '❶', '❷', '❸', '❹', '❺', '❻', '❼', '❽', '❾', '❿',
                     '㎜', '㎝', '㎞', '㎎', '㎏', '㎐', '㎑', '㎒', '㎓', '㎔',
                     '★', '●', '○', '▲', '△', '▼', '▽', '◆', '◇', '■', '□']
    for char in special_chars:
        text = text.replace(char, "")
    
    # 移除 Markdown 格式
    text = text.replace("**", "")
    text = text.replace("##", "")
    text = text.replace("###", "")
    text = text.replace("`", "")
    
    # HTML 转义（转义其他标签）
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    
    # 恢复<br>标签
    text = text.replace("{{BR_TAG}}", "<br>")
    
    return text.strip()

def check_forbidden_brands(text):
    """
    检查文本中是否包含禁止的品牌词
    返回：(是否包含品牌词，发现的品牌列表)
    """
    found_brands = []
    for brand in FORBIDDEN_BRANDS:
        if brand.lower() in text.lower() or brand in text:
            found_brands.append(brand)
    return len(found_brands) > 0, found_brands

def remove_forbidden_brands(text):
    """
    移除文本中的禁止品牌词
    """
    cleaned_text = text
    for brand in FORBIDDEN_BRANDS:
        cleaned_text = re.sub(re.escape(brand), "", cleaned_text, flags=re.IGNORECASE)
    return cleaned_text

def generate_title_variations(base_title, variation_count):
    """
    【修正】生成标题变体（SEO 关键词排列组合/同义词重写）
    确保每个变体真正不同，不出现版本号字眼
    """
    variations = [base_title]
    
    for i in range(1, variation_count):
        varied_title = base_title
        
        # 策略 1：多词同义词替换（至少替换 2-4 个词）
        replace_count = random.randint(2, min(4, len(SEO_SYNONYMS)))
        selected_keys = random.sample(list(SEO_SYNONYMS.keys()), replace_count)
        
        for key in selected_keys:
            if key in varied_title:
                synonyms = SEO_SYNONYMS[key]
                # 选择与之前不同的同义词
                available_synonyms = [s for s in synonyms if s not in varied_title]
                if available_synonyms:
                    replacement = random.choice(available_synonyms)
                    varied_title = varied_title.replace(key, replacement, 1)
        
        # 策略 2：调整开头/结尾组件
        if i % 2 == 0:  # 偶数变体调整开头
            for prefix in TITLE_COMPONENTS["开头"]:
                if prefix and not varied_title.startswith(prefix):
                    varied_title = prefix + varied_title
                    break
        else:  # 奇数变体调整结尾
            for suffix in TITLE_COMPONENTS["结尾"]:
                if suffix and not varied_title.endswith(suffix):
                    varied_title = varied_title + suffix
                    break
        
        # 策略 3：调整关键词顺序（如果有多个关键词短语）
        if i >= 3:
            # 尝试重新排列部分短语
            parts = varied_title.split(" ")
            if len(parts) > 3:
                # 交换中间部分
                mid_idx = len(parts) // 2
                parts[mid_idx-1], parts[mid_idx] = parts[mid_idx], parts[mid_idx-1]
                varied_title = " ".join(parts)
        
        # 确保不重复且不含版本号
        if varied_title not in variations and "パターン" not in varied_title and "バリエーション" not in varied_title:
            variations.append(varied_title)
        else:
            # 如果重复，继续尝试不同的替换组合
            for attempt in range(5):
                varied_title = base_title
                # 随机选择更多词替换
                replace_count = random.randint(3, min(5, len(SEO_SYNONYMS)))
                selected_keys = random.sample(list(SEO_SYNONYMS.keys()), replace_count)
                
                for key in selected_keys:
                    if key in varied_title:
                        synonyms = SEO_SYNONYMS[key]
                        replacement = random.choice(synonyms)
                        varied_title = varied_title.replace(key, replacement, 1)
                
                if varied_title not in variations and "パターン" not in varied_title and "バリエーション" not in varied_title:
                    variations.append(varied_title)
                    break
            else:
                # 实在无法生成不同变体，添加功能性描述
                functional_additions = [
                    " 機能性抜群", " 実用性重視", " 日常に便利", " 使い勝手良好",
                    " 品質確実", " 性能安定", " 満足度高", " コスパ優秀"
                ]
                addition = functional_additions[i % len(functional_additions)]
                varied_title = base_title + addition
                variations.append(varied_title)
    
    return variations

def generate_subtitle_variations(base_subtitle, variation_count):
    """
    【修正】生成副标题变体
    确保每个变体真正不同，不出现版本号字眼
    """
    variations = [base_subtitle]
    
    for i in range(1, variation_count):
        varied_subtitle = base_subtitle
        
        # 策略 1：多词同义词替换（至少替换 1-3 个词）
        replace_count = random.randint(1, min(3, len(SEO_SYNONYMS)))
        selected_keys = random.sample(list(SEO_SYNONYMS.keys()), replace_count)
        
        for key in selected_keys:
            if key in varied_subtitle:
                synonyms = SEO_SYNONYMS[key]
                available_synonyms = [s for s in synonyms if s not in varied_subtitle]
                if available_synonyms:
                    replacement = random.choice(available_synonyms)
                    varied_subtitle = varied_subtitle.replace(key, replacement, 1)
        
        # 策略 2：调整描述角度
        angle_additions = [
            " 毎日使える", " 長く愛用", " 手軽に活用", " 快適に使用",
            " 安心して利用", " 効率アップ", " 時間節約", " スペース有効"
        ]
        if i % 2 == 0 and len(varied_subtitle) < 80:
            varied_subtitle = varied_subtitle + angle_additions[i % len(angle_additions)]
        
        # 确保不重复且不含版本号
        if varied_subtitle not in variations and "パターン" not in varied_subtitle and "バリエーション" not in varied_subtitle:
            variations.append(varied_subtitle)
        else:
            # 如果重复，继续尝试
            for attempt in range(5):
                varied_subtitle = base_subtitle
                replace_count = random.randint(2, min(4, len(SEO_SYNONYMS)))
                selected_keys = random.sample(list(SEO_SYNONYMS.keys()), replace_count)
                
                for key in selected_keys:
                    if key in varied_subtitle:
                        synonyms = SEO_SYNONYMS[key]
                        replacement = random.choice(synonyms)
                        varied_subtitle = varied_subtitle.replace(key, replacement, 1)
                
                if varied_subtitle not in variations and "パターン" not in varied_subtitle and "バリエーション" not in varied_subtitle:
                    variations.append(varied_subtitle)
                    break
            else:
                # 添加功能性描述
                functional_additions = [
                    " 機能充実", " 実用設計", " 便利仕様", " 品質重視"
                ]
                addition = functional_additions[i % len(functional_additions)]
                varied_subtitle = base_subtitle + addition
                variations.append(varied_subtitle)
    
    return variations

def adjust_title_length(title, min_length=120, max_length=126):
    """
    调整标题字数到指定范围
    """
    current_length = count_fullwidth_chars(title)
    
    # 通用补充关键词（有效内容）
    generic_keywords = [
        "高品質", "新商品", "送料無料", "便利", "実用的", 
        "使いやすい", "耐久性", "おすすめ", "人気", "定番",
        "日常使い", "長持ち", "快適", "安心", "シンプル",
        "効率的", "多目的", "実用性抜群", "機能性", "軽量",
        "コンパクト", "持ち運び", "外出先", "オフィス", "家庭用",
        "一人暮らし", "ファミリー", "贈り物", "プレゼント", "ギフト"
    ]
    
    # 字数不足时补充
    while current_length < min_length:
        for keyword in generic_keywords:
            if current_length >= min_length:
                break
            if keyword not in title:
                title = f"{title} {keyword}"
                current_length = count_fullwidth_chars(title)
    
    # 字数过多时删减
    if current_length > max_length:
        # 找到合适的截断点
        truncated = title[:max_length]
        last_space = max(truncated.rfind(" "), truncated.rfind(" "))
        if last_space > min_length - 10:
            title = truncated[:last_space]
        else:
            title = truncated
    
    return title

def parse_image_filename(filename):
    """
    代码一图片解析逻辑：解析 x-y-z.jpg 命名规则
    y=1 为主图，y≠1 为详情图
    返回：(是否为主图，y 值，z 值)
    """
    name = os.path.splitext(filename)[0]
    parts = name.split('-')
    if len(parts) >= 3:
        try:
            y_value = int(parts[-2])  # y 值（倒数第二个）
            z_value = int(parts[-1])  # z 值（最后一个）
            is_main = (y_value == 1)  # y=1 为主图
            return is_main, y_value, z_value
        except ValueError:
            pass
    # 无法解析则视为主图
    return True, 0, 0

def get_images_in_folder(folder_path):
    """获取文件夹内所有图片，按自然排序"""
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
    all_files = []
    try:
        for f in os.listdir(folder_path):
            if f.lower().endswith(tuple(image_extensions)):
                all_files.append(f)
    except Exception as e:
        logger.error(f"无法读取文件夹 {folder_path}: {str(e)}")
        return []
    return natsorted(all_files)

def encode_image_to_base64(image_path):
    """
    按照代码二的方式编码图片（包含前缀）
    """
    try:
        with open(image_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode('utf-8')
            return f"data:image/jpeg;base64,{encoded}"
    except Exception as e:
        logger.error(f"图片编码失败：{image_path} - {str(e)}")
        return None

def build_ai_prompt(variation_count, total_images, shop_name):
    """
    增强版 Prompt，确保标题质量和字数
    """
    return f"""
你是一位专业的乐天市场（Rakuten Ichiba）商品文案专家。
我提供了 **{total_images} 张同一商品的不同角度/信息图片**，请综合分析所有图片内容，生成一份完整的商品详情页文案。

**所有输出内容必须全程使用【日语】，不要输出任何中文解释。**

### 核心要求：

1. **商品タイトル（正标题）**：
   - **字数严格：120-126 全角文字以内**（必须达到 120 字以上）
   - **有效内容比例：90% 以上**（避免无意义填充词）
   - 構造：[核心売点] + [製品カテゴリー] + [適用人群] + [主要機能] + [デザイン特徴] + [使用シーン] + [仕様情報] + [同義語補充]
   - キーワードを前に配置し、主要検索語をカバー
   - **禁止事项**：
     - 不得使用品牌名称（如 Nike、Apple、Sony、無印良品等）
     - 不得虚假宣传（如"100%""絶対""全て"等夸大词汇）
     - 不得侵犯商标权

2. **商品サブタイトル（副标题）**：
   - 80-87 全角文字以内
   - 包含内容：データ化/定量化された売りポイント、感覚的な記述、核心製品名、1-2 個の重要な補足キーワード
   - 差別化された利点を強調し、簡潔で力強い表現
   - **禁止事项**：同标题

3. **商品仕様表**：
   - 仅生成 1 份（综合所有图片信息）
   - 根据所有图片内的参数汇总、去重、整合
   - **重要**：如果所有图片中都没有提到某个参数，直接删除该栏，不要留空，不要写"不明"
   - 格式使用 `<br>` 换行，不要使用 Markdown 代码块
   - 范例格式：
     製品名：...<br>
     製品タイプ：...<br>
     サイズ：...<br>
     カラー：...<br>
     材質：...<br>
     重量：...<br>
     **重要**：如果所有图片中都没有提到某个参数，直接删除该栏，不要留空，不要写"不明"
     **重要**：如果所有图片中都没有提到某个参数，直接删除该栏，不要留空，不要写"不明"
     **重要**：如果所有图片中都没有提到某个参数，直接删除该栏，不要留空，不要写"不明"
     **重要**：如果所有图片中都没有提到某个参数，直接删除该栏，不要留空，不要写"不明"
     **重要**：如果所有图片中都没有提到某个参数，直接删除该栏，不要留空，不要写"不明"
     例如：若没有提到具体的"原産国"，"原産国：...<br>"这一栏应该直接删除，这是绝对的规则，不能违背

4. **商品特徴（介绍文案）**：
   - 生成 **{variation_count} 个不同的变体版本**
   - 每个变体侧重不同的卖点或角度
   - 每个变体约 80-100 字
   - 变体之间用 `###VARIATION###` 分隔
   - 每一句话（以句号 `。` 结尾）后面必须立即加上 `<br>` 换行符
   - 例如：これは商品です。<br>とても便利です。<br>

5. **関連キーワード**：
   - 生成 10-15 个高密度搜索关键词
   - 用全角空格分隔
   - 涵盖人群、场景、节日等维度
   - 総文字数：300-500 字

6. **Bluetooth 機能チェック**：
   - 检查商品是否含有蓝牙功能
   - 输出"あり"或"なし"

### 输出格式（严格遵守）：

【TITLE】
(标题内容，必须 120-126 全角字)

【SUBTITLE】
(副标题内容，80-87 全角字)

【SPECS】
(规格表内容，使用<br>换行)

【FEATURES】
(介绍文案变体 1)
###VARIATION###
(介绍文案变体 2)
###VARIATION###
...

【KEYWORDS】
(关键词 1 关键词 2 关键词 3 ...)

【BLUETOOTH】
(あり/なし)

### 禁止事项：
- 禁止使用特殊符号：① ② ❶ ㎜ ㎝ ㎏ ★ ● ○ ▲ △ ◆ ■ □ 等
- 禁止使用 Markdown 格式（如 **bold**）
- 禁止输出任何中文解释
- 禁止留空或写"不明"
- 换行符必须使用<br>，不能使用\\n
- **禁止使用任何品牌名称、商标词**
- **禁止虚假宣传、夸大其词**

商品店铺：{shop_name}
"""

def call_qwen_vl(image_paths, variation_count, shop_name):
    """
    调用多模态模型，返回 (响应文本, usage字典)
    usage = {'input_tokens': int, 'output_tokens': int} 或 None
    """
    dashscope.api_key = DASHSCOPE_API_KEY
    
    try:
        content_list = []
        for img_path in image_paths[:MAX_IMAGES_FOR_AI]:
            base64_img = encode_image_to_base64(img_path)
            if base64_img:
                content_list.append({"image": base64_img})
        
        prompt_text = build_ai_prompt(variation_count, len(image_paths), shop_name)
        content_list.append({"text": prompt_text})
        
        messages = [{"role": "user", "content": content_list}]
        response = MultiModalConversation.call(model=MODEL_NAME, messages=messages)
        
        if response.status_code == 200:
            # 提取 usage
            usage = None
            if hasattr(response, 'usage'):
                usage = {
                    'input_tokens': response.usage.input_tokens,
                    'output_tokens': response.usage.output_tokens
                }
            elif hasattr(response, 'output') and hasattr(response.output, 'usage'):
                usage = {
                    'input_tokens': response.output.usage.input_tokens,
                    'output_tokens': response.output.usage.output_tokens
                }
            text = response.output.choices[0].message.content[0]['text']
            return text, usage
        else:
            logger.error(f"API 调用失败：{response.code} - {response.message}")
            return None, None
            
    except Exception as e:
        logger.error(f"AI 调用异常：{str(e)}")
        return None, None

def parse_ai_response(response_text):
    """解析 AI 返回的文本，提取各部分内容"""
    result = {
        'title': '',
        'subtitle': '',
        'specs': '',
        'features': [],
        'keywords': '',
        'bluetooth': 'なし'
    }
    
    if not response_text:
        logger.warning("AI 响应为空")
        return result
    
    logger.info(f"AI 响应长度：{len(response_text)} 字符")
    logger.info(f"AI 响应预览：{response_text[:500]}...")
    
    try:
        # 提取标题
        title_match = re.search(r'【TITLE】\s*\n(.*?)(?=\n【|$)', response_text, re.DOTALL)
        if title_match:
            result['title'] = title_match.group(1).strip()
            logger.info(f"提取到标题：{result['title'][:50]}...")
            logger.info(f"标题字数：{count_fullwidth_chars(result['title'])} 字")
        else:
            logger.warning("未提取到标题")
        
        # 提取副标题
        subtitle_match = re.search(r'【SUBTITLE】\s*\n(.*?)(?=\n【|$)', response_text, re.DOTALL)
        if subtitle_match:
            result['subtitle'] = subtitle_match.group(1).strip()
            logger.info(f"提取到副标题：{result['subtitle'][:50]}...")
            logger.info(f"副标题字数：{count_fullwidth_chars(result['subtitle'])} 字")
        else:
            logger.warning("未提取到副标题")
        
        # 提取规格
        specs_match = re.search(r'【SPECS】\s*\n(.*?)(?=\n【|$)', response_text, re.DOTALL)
        if specs_match:
            result['specs'] = specs_match.group(1).strip()
            logger.info(f"提取到规格：{result['specs'][:100]}...")
        else:
            logger.warning("未提取到规格")
        
        # 提取特征变体
        features_match = re.search(r'【FEATURES】\s*\n(.*?)(?=\n【KEYWORDS】|$)', response_text, re.DOTALL)
        if features_match:
            features_text = features_match.group(1).strip()
            # 按分隔符拆分变体
            variations = re.split(r'###VARIATION###', features_text)
            result['features'] = [v.strip() for v in variations if v.strip()]
            logger.info(f"提取到 {len(result['features'])} 个特征变体")
        else:
            logger.warning("未提取到特征")
        
        # 提取关键词
        keywords_match = re.search(r'【KEYWORDS】\s*\n(.*?)(?=\n【|$)', response_text, re.DOTALL)
        if keywords_match:
            result['keywords'] = keywords_match.group(1).strip()
            logger.info(f"提取到关键词：{result['keywords'][:100]}...")
        else:
            logger.warning("未提取到关键词")
        
        # 提取蓝牙检查
        bt_match = re.search(r'【BLUETOOTH】\s*\n(.*?)(?=\n|$)', response_text, re.DOTALL)
        if bt_match:
            bt_text = bt_match.group(1).strip()
            result['bluetooth'] = 'あり' if 'あり' in bt_text else 'なし'
            logger.info(f"提取到蓝牙检查：{result['bluetooth']}")
        else:
            logger.warning("未提取到蓝牙检查")
            
    except Exception as e:
        logger.error(f"解析 AI 响应失败：{str(e)}")
    
    return result

def generate_html(shop_config, specs, features, keywords, bluetooth):
    """生成完整 HTML 详情页代码"""
    # 清理文本
    specs_clean = sanitize_text_for_html(specs)
    features_clean = sanitize_text_for_html(features)
    keywords_clean = sanitize_text_for_html(keywords)
    
    # 蓝牙注意事项
    bt_note = ""
    if bluetooth == 'あり':
        bt_note = (
            "<b>※技適マークについて</b><br>"
            "本商品は、電波法令で定められている技術基準に適合していることを証明する技適マークが貼付されていない無線機器であり、日本国内で使用する場合は、"
            "電波法違反になるおそれがございます。ご使用の際には、十分ご注意いただきますようお願いいたします。<br><br>"
        )
    
    # 填充模板
    html = HTML_TEMPLATE.format(
        shop_welcome=shop_config['welcome'],
        product_features=features_clean,
        product_specs=specs_clean,
        product_keywords=keywords_clean
    )
    
    # 插入蓝牙注意事项（在注意事项板块开头）
    if bt_note:
        html = html.replace("<b>※配送について</b>", bt_note + "<b>※配送について</b>")
    
    return html

def generate_image_links(folder_path, link_prefix, shop_name_code):
    """
    简化版：按自然排序，第一张为主图，其余为详情图
    返回：主图文件名列表（只有一项），详情图链接列表，行数据列表（只有一行）
    """
    folder_name = os.path.basename(folder_path)
    
    # 支持的图片扩展名
    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
    
    # 获取所有图片文件
    image_files = []
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            file_ext = os.path.splitext(file.lower())[1]
            if file_ext in image_extensions:
                image_files.append(file)
    
    if not image_files:
        logger.warning(f"在文件夹 {folder_name} 中未找到图片文件")
        return [], [], []
    
    # 自然排序
    image_files = natsorted(image_files)
    
    # 第一张为主图
    main_image_files = [image_files[0]]
    detail_files = image_files[1:]  # 其余为详情图
    
    # 生成链接
    detail_links = [f"{link_prefix}{f}" for f in detail_files]
    all_links = [f"{link_prefix}{f}" for f in image_files]
    
    # 生成行数据（只有一行）
    row_num = 2  # 从第2行开始（第1行是表头）
    main_file = main_image_files[0]
    main_link = f"{link_prefix}{main_file}"
    
    # 手机端和电脑端HTML框架（包含所有图片）
    mobile_html = ""
    pc_html = ""
    for link in all_links:
        mobile_html += f'<img src="{link}" width="100%"><br><br>\n'
        pc_html += f'<img src="{link}" width="1000"><br><br>\n'
    
    row_data = [{
        'row_num': row_num,
        'mobile_html': mobile_html,
        'pc_html': pc_html,
        'main_link': main_link,
        'main_file': main_file,
        'y_val': 0  # 不再使用
    }]
    
    logger.info(f"=== 图片链接处理完成：{folder_name} - 主图 1 张，详情图 {len(detail_files)} 张 ===")
    
    return main_image_files, detail_links, row_data
def load_keyword_library(file_path):
    """从本地 JSON 文件加载关键词库，返回 keyword_library 部分"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        keyword_lib = full_data.get("keyword_library", {})
        if not keyword_lib:
            logger.warning("JSON 文件中未找到 'keyword_library' 键")
        return keyword_lib
    except Exception as e:
        logger.error(f"读取关键词库失败: {e}")
        return {}

def safe_word_list(words, category):
    """安全地将输入转换为列表，用于后续拼接"""
    if words is None:
        return []
    if isinstance(words, list):
        seen = set()
        deduped = []
        for w in words:
            if w not in seen:
                seen.add(w)
                deduped.append(w)
        return deduped[:50]
    if isinstance(words, str):
        return [words]
    if isinstance(words, dict):
        logger.warning(f"分类 '{category}' 的值是字典类型，将提取所有字典值作为关键词。")
        flattened = []
        for v in words.values():
            if isinstance(v, list):
                flattened.extend(v)
            elif isinstance(v, str):
                flattened.append(v)
        return flattened[:50]
    return [str(words)][:50]

def detect_category_from_folder(folder_name):
    """
    根据文件夹名判断商品品类。
    返回品类标识符（如 'fan_cooling', 'cleaning_vacuum'），若无法判断则返回 None。
    """
    folder_lower = folder_name.lower()
    category_keywords = {
        'fan_cooling': ['fan', '风扇', '扇風', '冷却', 'cooling', '冷風', 'ハンディファン', '扇風機'],
        'cleaning_vacuum': ['clean', 'vacuum', '吸尘', '掃除', 'クリーナー', '掃除機'],
        'light': ['ランプ','E27'],
        'statue': ['仏像','観音']
    }
    for cat, keywords in category_keywords.items():
        for kw in keywords:
            if kw in folder_lower:
                return cat
    return None

def build_title_prompt(product_params, keyword_lib, category=None):
    params_text = "\n".join([f"- {k}: {v}" for k, v in product_params.items()])

    allowed_categories = [
        "core_product_names", "synonyms_variations", "specs_params",
        "features_benefits", "usage_scenarios", "marketing_emotional", "subtitle_hooks"
    ]

    lib_text = ""

    # 如果指定了品类且存在，添加该品类的词
    if category and category in keyword_lib:
        for cat in allowed_categories:
            if cat in keyword_lib[category]:
                words = keyword_lib[category][cat]
                safe_words = safe_word_list(words, cat)
                if safe_words:
                    lib_text += f"\n{cat}（{category}）: {', '.join(safe_words)}"

    # 添加通用词（general 分类）
    if 'general' in keyword_lib:
        for cat in allowed_categories:
            if cat in keyword_lib['general']:
                words = keyword_lib['general'][cat]
                safe_words = safe_word_list(words, cat)
                if safe_words:
                    lib_text += f"\n{cat}（一般）: {', '.join(safe_words)}"

    # 如果仍未收集到任何词（例如品类不存在或 general 缺失），回退到扁平结构（兼容旧版）
    if not lib_text:
        for cat in allowed_categories:
            if cat in keyword_lib:
                words = keyword_lib[cat]
                safe_words = safe_word_list(words, cat)
                if safe_words:
                    lib_text += f"\n{cat}: {', '.join(safe_words)}"

    prompt = f"""
你是一名拥有 10 年经验的日本乐天 (Rakuten Japan) SEO 运营专家，擅长撰写高点击、高搜索权重的商品标题。

请根据以下【商品参数】和【关键词库】，生成一组乐天商品的正标题和副标题。商品属于{category if category else '未知'}品类，请严格基于商品参数选择相关词汇，禁止引入无关品类词汇。

# 商品参数
{params_text}

# 关键词库（从中选取最适合的词汇）
{lib_text}

# 标题生成规则
1. **正标题 (main_title)**:
   - 长度控制在 100-126 字符之间（日文全角）。
   - 必须包含【core_product_names】中的至少 2 个词。
   - 必须包含【specs_params】和【features_benefits】中与商品匹配的词。
   - 关键词之间用半角空格分隔。
   - 顺序逻辑：核心词 + 规格 + 功能 + 场景 + 营销词。
   - 严禁出现违禁词，确保日语语法通顺（名词罗列即可）。

2. **副标题 (sub_title)**:
   - 长度控制在 70-87 字符之间。
   - 必须包含【subtitle_hooks】中的一个标签（如【軽量モデル】）。
   - 侧重用户利益点 (Benefit) 和促销信息 (如 送料無料)。
   - 语气亲切，吸引点击。

3. **匹配逻辑**:
   - 从词库中优先选择与商品参数完全匹配的关键词。
   - 如果商品是风扇类，必须包含“扇風機”或“ハンディファン”。
   - 年份词（如 2025, 2026）根据当前时间或商品属性自动选择最新的。

# 输出格式
请仅输出以下 JSON 格式，不要包含其他解释：
{{
  "main_title": "生成的正标题字符串",
  "sub_title": "生成的副标题字符串",
  "used_keywords": ["列出用到的核心关键词"]
}}
"""
    return prompt

def call_qwen_vl_for_text(prompt):
    """纯文本调用，返回 (响应文本, usage字典)"""
    dashscope.api_key = DASHSCOPE_API_KEY
    try:
        messages = [{"role": "user", "content": [{"text": prompt}]}]
        response = MultiModalConversation.call(model=MODEL_NAME, messages=messages)
        if response.status_code == 200:
            usage = None
            if hasattr(response, 'usage'):
                usage = {
                    'input_tokens': response.usage.input_tokens,
                    'output_tokens': response.usage.output_tokens
                }
            elif hasattr(response, 'output') and hasattr(response.output, 'usage'):
                usage = {
                    'input_tokens': response.output.usage.input_tokens,
                    'output_tokens': response.output.usage.output_tokens
                }
            text = response.output.choices[0].message.content[0]['text']
            return text, usage
        else:
            logger.error(f"API 调用失败：{response.code} - {response.message}")
            return None, None
    except Exception as e:
        logger.error(f"API 调用异常：{str(e)}")
        return None, None
    
def generate_title_from_params(product_params, keyword_lib, category=None):
    if not keyword_lib:
        logger.error("关键词库为空，无法生成标题")
        return None, None

    prompt = build_title_prompt(product_params, keyword_lib, category)
    logger.info(f"标题生成提示词构建完成（品类：{category}），开始调用模型...")
    response_text, usage = call_qwen_vl_for_text(prompt)
    if not response_text:
        return None, None

    try:
        start = response_text.find('{')
        end = response_text.rfind('}') + 1
        if start != -1 and end > start:
            json_str = response_text[start:end]
            result = json.loads(json_str)
            if 'main_title' in result and 'sub_title' in result:
                return result, usage
            else:
                logger.error("返回的 JSON 缺少必要字段")
                return None, None
        else:
            logger.error("模型返回的内容中没有找到有效的 JSON")
            logger.debug(f"原始返回内容: {response_text}")
            return None, None
    except Exception as e:
        logger.error(f"解析 JSON 失败: {e}")
        return None, None

def extract_params_from_specs(specs_text):
    """
    从规格文本（如 '製品名：...<br>サイズ：...'）中提取键值对，返回字典
    键名为日文原始字段名，值为对应的内容
    """
    params = {}
    if not specs_text:
        return params
    lines = specs_text.split('<br>')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        for sep in ['：', ':']:
            if sep in line:
                key, value = line.split(sep, 1)
                key = key.strip()
                value = value.strip()
                if key and value:
                    params[key] = value
                break
        # 无冒号的行忽略
    return params    

# =============================================================================
# ================= 核心处理函数 =================
# =============================================================================

def process_single_product(folder_path, shop_config, output_dir, index, total, workbook, default_ws, keyword_lib):
    """处理单个商品文件夹（合并逻辑）"""
    folder_name = os.path.basename(folder_path)
    total_tokens = {'input': 0, 'output': 0}
    
    print("=" * 60)
    print(f"📦 处理商品 [{index}/{total}]: {folder_name}")
    print("=" * 60)
    
    start_time = time.time()
    
    # 获取图片列表
    image_files = get_images_in_folder(folder_path)
    
    if not image_files:
        print(f"⚠️ 警告：文件夹 {folder_name} 中没有找到图片文件，跳过。")
        logger.warning(f"跳过商品 {folder_name}: 无图片")
        return {'status': 'skipped', 'reason': 'no_images', 'folder': folder_name}
    
    print(f"🔍 找到 {len(image_files)} 张图片：")
    for f in image_files[:5]:
        print(f"   - {f}")
    if len(image_files) > 5:
        print(f"   ... 还有 {len(image_files) - 5} 张")
    
    # 代码一图片链接生成逻辑
    main_images, detail_links, link_row_data = generate_image_links(folder_path, shop_config['link_prefix'], shop_config['code'])
    
    # 变体数量固定为 MAX_VARIATIONS（至少1）
    if len(main_images) == 0:
        variation_count = 1
    else:
        variation_count = MAX_VARIATIONS  # 例如5
    print(f"📝 将生成 {variation_count} 个文案变体")
    
    print(f"📷 识别到主图：{len(main_images)} 张，详情图：{len(detail_links)} 张")
    print(f"📝 将生成 {variation_count} 个文案变体")
    
    # 准备 AI 输入图片路径（所有图片，限制 20 张）
    image_paths = [os.path.join(folder_path, f) for f in image_files[:MAX_IMAGES_FOR_AI]]
    print(f"🤖 将发送 {len(image_paths)} 张图片给 AI 分析")
    
    # 调用 AI（代码二方式）
    print("🔄 正在调用 AI 生成内容...")
    ai_response, usage_main = call_qwen_vl(image_paths, variation_count, shop_config['name'])
    if usage_main:
        total_tokens['input'] += usage_main.get('input_tokens', 0)
        total_tokens['output'] += usage_main.get('output_tokens', 0)
    
    if not ai_response:
        print(f"❌ AI 调用失败，跳过商品 {folder_name}")
        logger.error(f"AI 调用失败：{folder_name}")
        return {'status': 'failed', 'reason': 'ai_error', 'folder': folder_name}
    
    # 解析 AI 响应
    ai_data = parse_ai_response(ai_response)
    # ========== 新增：基于关键词库和规格参数生成正副标题 ==========
    # 从规格文本提取参数
    # 从规格文本提取参数
    product_params = extract_params_from_specs(ai_data['specs'])

    # 根据文件夹名判断品类
    folder_name = os.path.basename(folder_path)
    category = detect_category_from_folder(folder_name)
    if category is None:
        logger.warning(f"无法从文件夹名 '{folder_name}' 判断品类，使用默认品类 'fan_cooling'")
        category = 'fan_cooling'  # 可改为更合适的默认值

    # 使用新方法生成正副标题
    base_title = None
    base_subtitle = None
    if keyword_lib and product_params:
        title_result, usage_title = generate_title_from_params(product_params, keyword_lib, category)
        if usage_title:   # 在此处累加
            total_tokens['input'] += usage_title.get('input_tokens', 0)
            total_tokens['output'] += usage_title.get('output_tokens', 0)
        if title_result:
            base_title = title_result['main_title']
            base_subtitle = title_result['sub_title']
            logger.info(f"新标题生成成功（品类：{category}）：{base_title[:50]}...")
        else:
            logger.warning("标题生成失败，回退使用原 AI 标题")
            base_title = ai_data['title']
            base_subtitle = ai_data['subtitle']
    else:
        logger.warning("关键词库为空或未提取到参数，使用原 AI 标题")
        base_title = ai_data['title']
        base_subtitle = ai_data['subtitle']
    # ========== 结束新增 ==========
    
    # 检查是否有内容
    if not ai_data['title'] and not ai_data['specs']:
        print(f"❌ AI 返回内容为空，跳过商品 {folder_name}")
        logger.error(f"AI 返回内容为空：{folder_name}")
        return {'status': 'failed', 'reason': 'empty_response', 'folder': folder_name}
    
    # 检查并移除品牌词（对新的 base_title/base_subtitle）
    has_brand, found_brands = check_forbidden_brands(base_title)
    if has_brand:
        print(f"⚠️ 警告：标题中包含品牌词 {found_brands}，已自动移除")
        logger.warning(f"标题包含品牌词：{found_brands}")
        base_title = remove_forbidden_brands(base_title)
    
    has_brand, found_brands = check_forbidden_brands(base_subtitle)
    if has_brand:
        print(f"⚠️ 警告：副标题中包含品牌词 {found_brands}，已自动移除")
        logger.warning(f"副标题包含品牌词：{found_brands}")
        base_subtitle = remove_forbidden_brands(base_subtitle)
    
    # 调整标题字数到 120-126 全角字
    original_title_length = count_fullwidth_chars(base_title)
    base_title = adjust_title_length(base_title, min_length=120, max_length=126)
    adjusted_title_length = count_fullwidth_chars(base_title)
    print(f"📏 标题字数调整：{original_title_length} → {adjusted_title_length} 字")
    logger.info(f"标题字数调整：{original_title_length} → {adjusted_title_length} 字")
    
    # 副标题字数简单处理（不低于70，不超过87）
    original_sub_len = count_fullwidth_chars(base_subtitle)
    if count_fullwidth_chars(base_subtitle) > 87:
        # 简单截断，保留完整词语
        while count_fullwidth_chars(base_subtitle) > 87 and ' ' in base_subtitle:
            base_subtitle = base_subtitle.rsplit(' ', 1)[0]
    if count_fullwidth_chars(base_subtitle) < 70:
        base_subtitle = base_subtitle + " 人気商品"  # 补充
    adjusted_sub_len = count_fullwidth_chars(base_subtitle)
    print(f"📏 副标题字数调整：{original_sub_len} → {adjusted_sub_len} 字")
    logger.info(f"副标题字数调整：{original_sub_len} → {adjusted_sub_len} 字")
    
    # 生成标题变体（基于新基础标题）
    title_variations = generate_title_variations(base_title, len(link_row_data))
    subtitle_variations = generate_subtitle_variations(base_subtitle, len(link_row_data))
    print(f"📝 已生成 {len(title_variations)} 个标题变体")
    
    # 打印变体预览（用于调试）
    for i, (t, s) in enumerate(zip(title_variations, subtitle_variations), 1):
        print(f"  变体{i}: 标题={t[:30]}... | 副标题={s[:30]}...")
    
    elapsed_time = time.time() - start_time
    print(f"✅ AI 生成完成，耗时 {elapsed_time:.2f} 秒")
    
    # 生成 HTML（每个变体一个 HTML）
    html_rows = []
    
    # 确保有足够的变体文案
    features_list = ai_data['features']
    if len(features_list) == 0:
        features_list = [ai_data['title'] or "商品特徴"]
    
    # 循环使用变体（如果主图数量超过变体数量）
    total_rows = len(link_row_data) if link_row_data else 1
    
    for i in range(total_rows):
        # 选择变体文案（循环）
        feature_text = features_list[i % len(features_list)]
        
        # 选择对应的标题变体
        title_text = title_variations[i % len(title_variations)]
        subtitle_text = subtitle_variations[i % len(subtitle_variations)]
        
        # 生成 HTML
        html_content = generate_html(shop_config, ai_data['specs'], feature_text, ai_data['keywords'], ai_data['bluetooth'])
        
        # 获取对应行的图片 HTML
        row_info = link_row_data[i] if i < len(link_row_data) else link_row_data[0]
        
        html_rows.append({
            'title': title_text,
            'subtitle': subtitle_text,
            'sku': f"{folder_name}_{i+1}",
            'price': '',
            'html': html_content,
            'mobile_html': row_info['mobile_html'],
            'pc_html': row_info['pc_html'],
            'main_link': row_info['main_link'],
            'category_id': '',
            'category_name': '',
            'specs': ai_data['specs'],
            'features': feature_text,
            'bluetooth': ai_data['bluetooth'],
            'keywords': ai_data['keywords'],
            'detail_links': detail_links,
            'row_num': row_info['row_num']
        })
    
    # 写入 Excel
    try:
        # 创建工作表
        sheet_name = "".join([c for c in folder_name if c.isalnum() or c in '._- ']).strip()[:31]
        
        # 检查是否已存在同名工作表
        existing_sheets = [ws.title for ws in workbook.worksheets]
        if sheet_name in existing_sheets:
            sheet_name = f"{sheet_name}_{uuid.uuid4().hex[:5]}"
        
        if index == 0:
            worksheet = default_ws
            worksheet.title = sheet_name
        else:
            worksheet = workbook.create_sheet(title=sheet_name)
        
        # 设置表头（代码一布局）
        worksheet['A1'] = "标题"
        worksheet['B1'] = "副标题"
        worksheet['C1'] = "sku"
        worksheet['D1'] = "价格"
        worksheet['E1'] = "代码"
        worksheet['F1'] = "手机图"
        worksheet['G1'] = "电脑图"
        worksheet['H1'] = "种类 ID"
        worksheet['I1'] = "店内种类"
        
        # 设置表头样式
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
            cell = f'{col}1'
            if cell in worksheet:
                worksheet[cell].font = Font(bold=True)
                worksheet[cell].alignment = Alignment(horizontal='center')
        
        # 填充数据行
        for row_info in html_rows:
            row_num = row_info['row_num']
            worksheet[f'A{row_num}'] = row_info['title']
            worksheet[f'B{row_num}'] = row_info['subtitle']
            worksheet[f'C{row_num}'] = row_info['sku']
            worksheet[f'D{row_num}'] = row_info['price']
            worksheet[f'E{row_num}'] = row_info['html']
            worksheet[f'F{row_num}'] = row_info['mobile_html']
            worksheet[f'G{row_num}'] = row_info['pc_html']
            worksheet[f'H{row_num}'] = row_info['main_link']
            worksheet[f'I{row_num}'] = row_info['category_name']
            worksheet[f'K{row_num}'] = row_info['specs']
            worksheet[f'L{row_num}'] = row_info['features']
            worksheet[f'M{row_num}'] = "有" if row_info['bluetooth'] == 'あり' else "无"
            worksheet[f'N{row_num}'] = row_info['keywords']
        
        # J 列：详情图链接（从 J2 开始）
        worksheet[f'J1'] = f"{shop_config['code']}侧图链接"
        worksheet[f'J1'].font = Font(bold=True)
        worksheet[f'J1'].alignment = Alignment(horizontal='center')
        for i, link in enumerate(detail_links, start=2):
            worksheet[f'J{i}'] = link
        
        # 调整列宽
        worksheet.column_dimensions['A'].width = 50
        worksheet.column_dimensions['B'].width = 50
        worksheet.column_dimensions['E'].width = 80
        worksheet.column_dimensions['F'].width = 50
        worksheet.column_dimensions['G'].width = 50
        worksheet.column_dimensions['H'].width = 50
        worksheet.column_dimensions['J'].width = 50
        worksheet.column_dimensions['K'].width = 50
        worksheet.column_dimensions['L'].width = 50
        worksheet.column_dimensions['M'].width = 20
        worksheet.column_dimensions['N'].width = 80
        
        print(f"📄 Excel 工作表已创建：{sheet_name}")
        logger.info(f"Excel 写入成功：{sheet_name}")
        
    except Exception as e:
        print(f"❌ Excel 写入失败：{str(e)}")
        logger.error(f"Excel 写入失败：{str(e)}")
        return {'status': 'failed', 'reason': 'excel_error', 'folder': folder_name, 'error': str(e)}
    
    
    return {
        'status': 'success',
        'folder': folder_name,
        'output_file': "N/A",  # 不再生成报告文件
        'image_count': len(image_files),
        'variations': variation_count,
        'tokens': total_tokens
    }

# =============================================================================
# ================= 主函数 =================
# =============================================================================


def main():
    """主入口函数（代码一逻辑）"""
    print("=" * 60)
    print("🚀 乐天市场商品上架自动化系统（合并增强版 - 修复标题变体生成）")
    print("=" * 60)
    
    # 检查 API Key
    if DASHSCOPE_API_KEY in (None, "", "sk-你的API密钥", "YOUR_DASHSCOPE_API_KEY_HERE"):
        print("❌ 错误：请在配置区填入有效的 DASHSCOPE_API_KEY。")
        return
    
    print(f"📂 主目录：{MAIN_DIR}")
    print(f"💾 输出目录：{OUTPUT_DIR}")
    print(f"🤖 模型：{MODEL_NAME}")
    print("=" * 60)
    
    # 检查主目录是否存在
    if not os.path.exists(MAIN_DIR):
        print(f"❌ 错误：目录 {MAIN_DIR} 不存在。")
        print("💡 提示：请确认目录路径是否正确。")
        return
    
    # 确保输出目录存在
    if not os.path.exists(OUTPUT_DIR):
        try:
            os.makedirs(OUTPUT_DIR)
            print(f"✅ 已创建输出目录：{OUTPUT_DIR}")
        except Exception as e:
            print(f"❌ 无法创建输出目录：{str(e)}")
            return
    
    # 店铺配置（从用户配置区读取）
    shop_config = {
        "name": SHOP_NAME,
        "welcome": SHOP_WELCOME,
        "code": SHOP_CODE,
        "link_prefix": SHOP_LINK_PREFIX
    }
    print(f"✅ 已选择店铺：{shop_config['name']}")
    logger.info(f"店铺选择：{shop_config['name']}")
    logger.info(f"链接前缀：{shop_config['link_prefix']}")
    logger.info(f"店铺代码：{shop_config['code']}")
    
    # 加载关键词库
    keyword_lib = load_keyword_library(KEYWORD_LIB_PATH)
    if keyword_lib:
        logger.info(f"关键词库加载成功，包含分类：{list(keyword_lib.keys())}")
    else:
        logger.warning("关键词库加载失败，将使用原 AI 生成标题")

    # 扫描商品文件夹（代码一路径）
    print(f"\n🔍 正在扫描商品文件夹...")
    
    product_folders = []
    try:
        all_items = os.listdir(MAIN_DIR)
    except Exception as e:
        print(f"❌ 无法读取目录 {MAIN_DIR}: {str(e)}")
        return
    
    for item in all_items:
        item_path = os.path.join(MAIN_DIR, item)
        if not os.path.isdir(item_path):
            continue
        if item.startswith('.') or item in ['@eaDir', 'System Volume Information', '$RECYCLE.BIN']:
            continue
        product_folders.append({
            'name': item,
            'path': item_path
        })
    
    # 自然排序
    product_folders = sorted(product_folders, key=lambda x: x['name'])
    product_count = len(product_folders)
    
    if product_count == 0:
        print("❌ 错误：未找到任何商品子文件夹。")
        print(f"💡 提示：请在 {MAIN_DIR} 下创建子文件夹，每个文件夹放一个商品的图片。")
        return
    
    logger.info(f"找到 {len(product_folders)} 个子文件夹")
    print(f"\n✅ 识别到【商品数量】: {product_count} 个")
    
    print("\n📋 商品文件夹列表：")
    for i, folder in enumerate(product_folders, 1):
        print(f"   {i}. {folder['name']}")
    
    
    # 创建 Excel 工作簿
    workbook = Workbook()
    default_ws = workbook.active
    default_ws.title = "首页"
    
    # 统计结果
    results = {'success': 0, 'skipped': 0, 'failed': 0, 'details': []}
    global_tokens = {'input': 0, 'output': 0}   # 新增
    
    # 逐个处理商品
    start_total_time = time.time()
    
    for i, folder_info in enumerate(product_folders, 1):
        result = process_single_product(
            folder_info['path'], 
            shop_config, 
            OUTPUT_DIR, 
            i, 
            product_count,
            workbook,
            default_ws,
            keyword_lib   # 新增参数
        )
        results['details'].append(result)
        
        if result['status'] == 'success':
            results['success'] += 1
        elif result['status'] == 'skipped':
            results['skipped'] += 1
        else:
            results['failed'] += 1

        # 累加 token
        if result.get('tokens'):
            global_tokens['input'] += result['tokens']['input']
            global_tokens['output'] += result['tokens']['output']    
        
        # 速率限制：每个商品之间等待
        if i < product_count:
            print(f"⏳ 等待 {API_DELAY_SECONDS} 秒后处理下一个商品...")
            time.sleep(API_DELAY_SECONDS)
    
    total_elapsed = time.time() - start_total_time
    
    # 输出总结报告
    print("=" * 60)
    print("📊 批量处理完成报告")
    print("=" * 60)
    print(f"总商品数：{product_count}")
    print(f"✅ 成功：{results['success']}")
    print(f"⚠️ 跳过：{results['skipped']} (无图片)")
    print(f"❌ 失败：{results['failed']}")
    print(f"总耗时：{total_elapsed:.2f} 秒")
    
    # 保存 Excel（代码一路径：桌面）
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = os.path.join(OUTPUT_DIR, f"整合输出_{timestamp}.xlsx")
        workbook.save(excel_path)
        print(f"📄 Excel 文件已保存至：{excel_path}")
        logger.info(f"Excel 保存成功：{excel_path}")
    except Exception as e:
        print(f"❌ Excel 保存失败：{str(e)}")
        logger.error(f"Excel 保存失败：{str(e)}")
    
    # 显示 Excel 布局信息
    print("=" * 60)
    print("📋 Excel 工作表布局:")
    print(" - A 列：正标题（每行不同变体）")
    print(" - B 列：副标题（每行不同变体）")
    print(" - C 列：sku")
    print(" - D 列：价格")
    print(" - E 列：AI 生成的 HTML 内容")
    print(" - F 列：手机端 HTML 框架")
    print(" - G 列：电脑端 HTML 框架")
    print(" - H 列：主图链接")
    print(" - I 列：店内种类")
    print(f" - J1: {shop_config['code']}侧图链接")
    print(" - J 列从 J2 开始：详情图链接")
    print(" - K 列：商品规格（AI 提取）")
    print(" - L 列：商品介绍（AI 生成）")
    print(" - M 列：蓝牙功能检查结果")
    print(" - N 列：乐天关键词")
    print("=" * 60)
    
    logger.info("处理完成")
# 输出成本统计
    if ENABLE_COST_TRACKING:
        input_cost = global_tokens['input'] / 1000 * INPUT_TOKEN_PRICE
        output_cost = global_tokens['output'] / 1000 * OUTPUT_TOKEN_PRICE
        total_cost = input_cost + output_cost
        print("\n" + "=" * 60)
        print("💰 本次运行成本统计")
        print("=" * 60)
        print(f"输入 Token 总数：{global_tokens['input']}")
        print(f"输出 Token 总数：{global_tokens['output']}")
        print(f"输入成本：¥{input_cost:.4f}")
        print(f"输出成本：¥{output_cost:.4f}")
        print(f"总成本：¥{total_cost:.4f}")
    print(f"🎉 结果已保存到：{OUTPUT_DIR}")
    print(f"共处理了 {product_count} 个子文件夹")

if __name__ == "__main__":
    main()
