# 乐天市场商品上架自动化工具

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

> 批量处理商品图片，调用 AI 自动生成日文商品文案，一键输出乐天上架用 Excel。

---

## 📖 简介

本工具专为 **乐天市场（Rakuten Ichiba）卖家** 设计，包含两大核心功能：

- **图片预处理**（`src.py`）：批量重命名、规范尺寸、添加水印
- **AI 文案生成**（`pig.py`）：识别图片内容，自动生成标题、副标题、规格、介绍文、关键词等完整详情页文案

两个工具可独立使用，也可串联使用，帮你从繁琐的商品上架工作中解放出来。

> 本工具使用阿里云 DashScope（通义千问 VL）多模态模型，需自行申请 API Key。

---

## ✨ 主要特性

- ✅ 图片自动重命名（时间戳+序号，避免重复）
- ✅ 统一缩放至指定宽度（默认 800px）
- ✅ 水印添加（支持透明度、缩放比例调整）
- ✅ 自动备份原始图片
- ✅ AI 识别多张图片，生成完整日文商品详情页文案
- ✅ 自动标题变体生成（SEO 优化）
- ✅ 品牌词自动过滤（防止侵权）
- ✅ 蓝牙功能检测并添加合规提示
- ✅ 输出标准 Excel 文件，包含 HTML 详情页代码
- ✅ **Token 成本追踪**（输入/输出 Token 数及费用统计）
- ✅ 完全配置化，无需修改代码逻辑

---

## 📁 目录结构
pig/
├── data/
│ ├── source/ # 【用户放置】原始商品图片文件夹
│ │ ├── 1.电风扇/ # 子文件夹命名格式：数字.名称（如 1.风扇）
│ │ ├── 2.加湿器/
│ │ └── ...
│ ├── output/ # 【自动生成】Excel 输出目录
│ ├── backup/ # 【自动生成】图片预处理时的备份
│ └── watermark.png # 【用户放置】水印图片（可选，PNG 透明背景）
├── src/
│ ├── pig.py # AI 文案生成主程序
│ └── src.py # 图片预处理工具
├── requirements.txt # Python 依赖清单
└── README.md # 本文件

---

## 🚀 快速开始

### 1. 安装依赖

确保已安装 Python 3.8 或更高版本，然后执行：

```bash
pip install -r requirements.txt
```

---

本工具使用阿里云 DashScope（通义千问 VL 模型）生成文案，你需要：

访问 DashScope 控制台

注册/登录阿里云账号

创建 API Key（注意保存，只显示一次）

### 将 API Key 填入 src/pig.py 的配置区

参数——说明——示例
DASHSCOPE_API_KEY——阿里云 DashScope API Key（必填）——"sk-xxxxx"
MAIN_DIR——商品图片源目录——自动生成，一般无需修改
OUTPUT_DIR——输出目录——自动生成，一般无需修改
KEYWORD_LIB_PATH——关键词库路径（不需要设为 None）——None 或 "路径"
SHOP_NAME——店铺名称——"弊店"
SHOP_WELCOME——店铺欢迎语——"楽天市場店へようこそ！..."
SHOP_CODE——店铺代码（仅显示用）——"1 店"
SHOP_LINK_PREFIX——图片链接前缀（关键）——"https://image.rakuten.co.jp/your_shop/cabinet/mg/mg2/"
ENABLE_COST_TRACKING——是否启用成本统计——True / False
INPUT_TOKEN_PRICE——模型输入价格（元/千 Token）——0.008
OUTPUT_TOKEN_PRICE——模型输出价格（元/千 Token）——0.016

### src.py 配置区（src/src.py 顶部）

参数——说明——示例
SOURCE_DIR——图片源目录——自动生成，一般无需修改
BACKUP_DIR——备份目录（设为 None 则跳过）——自动生成或 None
WATERMARK_CONFIGS——水印配置（空列表则不加）——见下文
RENAME_PREFIX——重命名后的文件名前缀（已废弃）——可忽略
TARGET_WIDTH——缩放目标宽度（像素）——800
DELETE_ORIGINAL——是否删除原图（现已覆盖，此参数无效）——True / False

### src.py中水印配置示例
WATERMARK_CONFIGS = [
    {
        "path": os.path.join(PROJECT_ROOT, "data", "watermark.png"),
        "scale": 0.3,          # 水印宽度占图片宽度的比例
        "opacity": 0.7         # 透明度（0.0~1.0）
    },
]

## 📝 使用步骤

### 步骤一：图片预处理（重命名、缩放、加水印）
运行图片预处理工具，自动完成：备份 → 重命名 → 缩放 → 加水印（可选）。

cd pig
python src/src.py

#### 执行效果：
1. data/source/ 下的图片会被重命名为 20260814_153045_001.jpg、20260814_153045_002.jpg 等格式（时间戳+序号）

2. 图片缩放至统一宽度（默认 800px）

3. 如果有水印配置，直接覆盖原图（带水印）

4. 原始图片自动备份到 data/backup/

### 步骤二：AI 生成文案
运行 AI 文案生成程序：python src/pig.py

#### 执行过程：
1. 扫描 data/source/ 下的所有商品子文件夹

2. 对每个商品，将图片发送给 AI 分析（按自然排序，第一张作为主图）

3. AI 返回：标题、副标题、规格表、介绍文案、关键词、蓝牙检测

4. 生成完整的 HTML 详情页代码

5. 输出到 Excel 文件（data/output/整合输出_时间戳.xlsx）

※执行前请先在代码内，根据店铺的实际情况，配置好详情文案中【注意事项】、【品质保证板块的文案】。

#### 输出结果：
1. Excel 文件：包含标题、副标题、SKU、HTML 代码、手机/电脑端图片框架等

2. 控制台输出处理进度和成本统计（无额外 txt 文件）

## 📊 Excel 输出说明
列——内容
A——正标题（SEO 优化，120-126 全角字符）
B——副标题（80-87 全角字符）
C——SKU（自动生成）
D——价格（预留，可手动填写）
E——完整的商品详情 HTML 代码
F——手机端图片 HTML 框架
G——电脑端图片 HTML 框架
H——主图链接
J——详情图链接列表
K——商品规格（AI 提取）
L——商品介绍（AI 生成）
M——蓝牙功能检查（有/无）
N——乐天关键词

## 🛠️ 常见问题

### Q1: 提示 "目录不存在" 怎么办？
确认 data/source/ 文件夹存在且包含商品子文件夹

检查是否在项目根目录下运行（cd pig）

### Q2: API 调用失败？
确认 API Key 正确且账户有余额

检查网络连接是否正常

可降低 MAX_IMAGES_FOR_AI 值（默认 20 张）减少调用成本

### Q3: 水印显示异常？
确认水印图片为 PNG 格式且包含透明通道

调整 opacity 参数（建议 0.5~0.8）

调整 scale 参数控制水印大小

### Q4: 标题字数不够或超长？
工具会自动调整标题到 120-126 字符

如仍不理想，可在 pig.py 中修改 adjust_title_length() 的参数

### Q5: 如何修改图片链接前缀？
修改 pig.py 中的 SHOP_LINK_PREFIX 变量，替换为你的乐天图片存储路径

### Q6: 不想生成多个标题变体？
修改 pig.py 中的 MAX_VARIATIONS 变量（默认 5），设为 1 即可只生成一个

### Q7: 成本统计中的价格不准确？
请根据阿里云 DashScope 官方最新价格修改 INPUT_TOKEN_PRICE 和 OUTPUT_TOKEN_PRICE

## 📄 许可证
本项目采用 MIT 许可证，可自由使用、修改和分发。详见 LICENSE 文件。

## 🤝 贡献
欢迎提交 Issue 和 Pull Request。

## 📮 联系方式
如有问题，请通过 GitHub Issues 反馈。
