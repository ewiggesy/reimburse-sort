# reimburse-sort（报销发票整理助手）

自动整理报销发票的 Kimi Skill + 独立 Python 脚本：

- **出差识别**：从车票/机票/会议通知提取出差起止时间、目的地/省份、出差人（支持多人）

- **分类**：识别发票类别（交通 / 会议 / 图书 / 顺丰快递 / 其他），会议内的火车票/机票跟随会议计入会议总额
- **统计**：自动提取 PDF 发票金额，按类别、按会议分组统计（会务费 / 住宿费 / 交通费 / 其他费用细分），发票号码自动去重
- **整理**：按类别 + 原子文件夹输出到新目录，生成 `报销统计.md` 和 `报销明细.csv`
- **纸质票**：`纸质发票.csv` 手工登记或 `add` 命令连续录入，金额自动并入合计
- **打印**：发票内容顺时针旋转 90° 后保持原始尺寸打印，纸张方向按内容自动选择；会议通知等文档保持竖版，照片/截图原样直印；静默无对话框

## 用法

```bash
# 扫描统计（不改动文件）
python scripts/reimburse.py scan "<发票文件夹>"

# 按类别整理 + 生成报告（复制，不删除原件）
python scripts/reimburse.py organize "<发票文件夹>"

# 登记纸质发票（无电子版）
python scripts/reimburse.py add "<发票文件夹>" --category 会议 --group 8月-杭州会议 --etype 会务费 --amount 200

# 只输金额算总和
python scripts/reimburse.py total 120 45.5 200

# 打印（--landscape = 仅发票内容顺时针旋转90°并保持原始尺寸，其他文档/照片直印）
python scripts/reimburse.py print --landscape "<分类整理输出目录>"
```

## 依赖

Python 3.9+，依赖：`pdfplumber`、`pypdfium2`、`Pillow`（打印/扫描时需要）。

## 作为 Kimi Skill 使用

将整个 `reimburse-sort/` 目录放入 Kimi 的 skills 目录（或在 Kimi Work 中安装），之后直接把装发票的文件夹发给 Kimi，说"整理一下"即可，无需其他说明。

## 口径说明

- 会议文件夹内的火车票/机票**跟随会议**计入会议总额（不单独归入交通）
- 只统计能提取到金额的正式发票；会议通知、付款凭证图片、zip/OFD 副本不计入金额
- 同一张发票的 pdf/zip/ofd 多份副本按发票号码去重，只计一次
