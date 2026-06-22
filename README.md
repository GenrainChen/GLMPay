# GLMPay

智谱 AI(GLM)开放平台**费用明细分析工具**：读取平台导出的账单 Excel，生成一套「Quantum Terminal HUD」风格的动态可视化 HTML 报告——黑色基调、科技感、动态且不卡顿。

## 功能特性

- 📊 解析智谱 AI 开放平台导出的费用明细 `.xlsx`
- 💰 结合 GLM 系列模型定价表(`glm_pricing.csv`)核算成本
- 🎨 输出酷炫的 HUD 风格 HTML 报告(动态效果 + 高性能渲染)

## 目录结构

```
.
├── analyze_bill.py          # 主程序:解析账单并生成报告
├── report_template.html     # 报告 HTML 模板
├── report_style.css         # 报告样式(Quantum Terminal HUD 风格)
├── glm_pricing.csv          # GLM 系列模型定价参考表
└── requirements.txt         # Python 依赖
```

> 账单 `.xlsx` 与生成的 `bill_report.html` 为本地数据，已通过 `.gitignore` 排除，不上传至仓库。

## 快速开始

```bash
# 1. 安装依赖
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 2. 放入智谱 AI 平台导出的费用明细 xlsx 到项目根目录

# 3. 生成报告
python analyze_bill.py
```

运行后在项目根目录生成 `bill_report.html`，用浏览器打开即可查看。

## 依赖

- `pandas>=2.0`
- `openpyxl>=3.1`

## License

MIT
