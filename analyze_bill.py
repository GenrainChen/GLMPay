#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
智谱AI 费用明细分析脚本
自动扫描目录下所有「智谱AI开放平台费用明细」开头的 xlsx 文件，
合并为一份统一的 HTML 报告。

用法:
  python analyze_bill.py                                 # 自动扫描
  python analyze_bill.py 文件1.xlsx 文件2.xlsx           # 指定文件
"""

import math
import re
import sys
import warnings
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import pandas as pd
from dateutil.relativedelta import relativedelta

# ── 配置 ──────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent
PRICING_FILE = SCRIPT_DIR / "glm_pricing.csv"
TEMPLATE_FILE = SCRIPT_DIR / "report_template.html"
STYLE_FILE = SCRIPT_DIR / "report_style.css"

# ── 列名常量 ──────────────────────────────────────────
COL_MODEL = "模型编码（推理专用）"
COL_PACKAGE = "Tokens资源包名称"
COL_DATE = "账期(自然日)"
COL_PRICE_TYPE = "价格类型"
COL_AMOUNT = "用量"
COL_UNIT_PRICE = "单价"
COL_COST = "理论费用"
COL_PAID = "已付款金额"
COL_REQUESTS = "请求次数 (仅API)"
COL_CUSTOMER = "客户id"
COL_PRODUCT_TYPE = "产品类型"
COL_PRODUCT_NAME = "模型产品名称"
SUBSCRIPTION_TYPE = "订阅套餐"

# ── 预编译正则 ────────────────────────────────────────
_RE_MONTH = re.compile(r"(\d{4})-(\d{2})")
_RE_MONTH_SHORT = re.compile(r"\d{4}-\d{2}")

MODEL_NAMES: dict[str, str] = {
    "glm-5.2": "GLM-5.2",
    "glm-5.1": "GLM-5.1",
    "glm-5-turbo": "GLM-5-Turbo",
    "glm-5": "GLM-5",
    "glm-4.7": "GLM-4.7",
    "glm-4.6v": "GLM-4.6V",
    "glm-4.5-air": "GLM-4.5-Air",
    "search-prime-claude": "Search Prime",
    "web-reader": "Web Reader",
    "zread": "ZRead",
}

PRICE_TYPE_NAMES: dict[str, str] = {
    "输入": "输入",
    "输出": "输出",
    "缓存命中": "缓存命中",
    "不区分输入输出": "调用次数",
}

PACKAGE_COLORS: dict[str, str] = {
    "GLM Coding Max V2 - 季": "#6366f1",
    "GLM Coding Max V2 - 月": "#818cf8",
    "GLM Coding Pro V2 - 季": "#8b5cf6",
    "联网搜索/读取 - Max - 包季计划": "#06b6d4",
    "联网搜索/读取 - Pro - 包季计划": "#0ea5e9",
    "【实名认证】500万GLM-4.7体验包": "#10b981",
    "【新用户专享】200万通用模型推理资源包": "#f59e0b",
    "1000万GLM-4.7资源包": "#22c55e",
}

# 套餐产品名（订阅套餐行的「模型产品名称」）展示名与配色
PRODUCT_NAMES: dict[str, str] = {
    "GLM Coding Max": "GLM Coding Max",
    "GLM Coding Pro": "GLM Coding Pro",
}
PRODUCT_COLORS: dict[str, str] = {
    "GLM Coding Max": "#6366f1",
    "GLM Coding Pro": "#8b5cf6",
}

MODEL_COLORS: dict[str, str] = {
    "glm-5.2": "#22d3ee",
    "glm-5.1": "#a78bfa",
    "glm-5-turbo": "#8b5cf6",
    "glm-4.7": "#06b6d4",
    "glm-4.5-air": "#10b981",
    "glm-4.6v": "#f59e0b",
    "search-prime-claude": "#ef4444",
    "web-reader": "#ec4899",
    "zread": "#84cc16",
    "glm-5": "#f97316",
}

MONTH_COLORS = ["#6366f1", "#06b6d4", "#10b981", "#f59e0b", "#ef4444", "#ec4899"]

# ── 套餐订阅模型（成本摊销）──────────────────────────
# 官方月费(元/月); 折扣按月数:月1.0 / 季0.9 / 年0.8
PLAN_MONTHLY = {"Lite": 49, "Pro": 149, "Max": 469}
PLAN_DISCOUNT = {1: 1.0, 3: 0.9, 12: 0.8}
# 订阅清单:(账号标签, 档位, 月数, 生效日期)
#   升级时旧套餐在新套餐生效日终止;剩余价值折算已计入新套餐公允价值
#   公允价值 = PLAN_MONTHLY[档位] × 月数 × PLAN_DISCOUNT[月数]
SUBSCRIPTIONS = [
    # (账号, 档位, 月数, 生效日期, 实付)
    ("账号A", "Pro", 3, "2026-05-14", 402.30),   # 5/20 升级 Max,提前终止
    ("账号A", "Max", 3, "2026-05-20", 886.10),   # 升级,公允=Max季价(含Pro折算)
    ("账号B", "Max", 1, "2026-06-13", 469.00),   # 独立账号
]

# SVG 饼图参数
PIE_SVG_RADIUS = 18
PIE_SVG_CIRCUMFERENCE = 2 * math.pi * PIE_SVG_RADIUS


# ── 工具函数 ──────────────────────────────────────────

def fmt(n: float | int) -> str:
    """格式化数值：大数千分位，小数保留适当精度"""
    if isinstance(n, float):
        if n >= 1:
            return f"{n:,.2f}"
        elif n >= 0.01:
            return f"{n:.4f}"
        else:
            return f"{n:.6f}"
    return f"{n:,}"


def fmt_tok(n: int | float) -> str:
    """格式化 token 数量为可读形式（K/M）"""
    if n >= 1_000_000:
        return f"{n / 1_000_000:,.1f}M"
    elif n >= 1_000:
        return f"{n / 1_000:,.1f}K"
    return str(int(n))


def detect_month(filepath: str | Path) -> str:
    """从文件名提取月份标签，如 '2026年5月'"""
    name = Path(filepath).name
    m = _RE_MONTH.search(name)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月"
    return "未知月份"


def detect_month_short(filepath: str | Path) -> str:
    """从文件名提取短月份标识，如 '2026-05'"""
    name = Path(filepath).name
    m = _RE_MONTH_SHORT.search(name)
    return m.group(0) if m else "unknown"


def render(template_str: str, **kwargs: Any) -> str:
    """将模板中的 {{key}} 占位符替换为对应值"""
    for key, val in kwargs.items():
        template_str = template_str.replace("{{" + key + "}}", str(val))
    return template_str


# ── 数据加载 ──────────────────────────────────────────

def load_single(filepath: str | Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """加载单个 xlsx，返回 (用量行, 订阅套餐行) 两个 DataFrame。

    - 用量行：Tokens资源包名称含「GLM Coding」，承载理论费用/token/请求/缓存
    - 订阅套餐行：产品类型==「订阅套餐」且模型产品名称含「GLM Coding」，承载套餐实付
    """
    warnings.filterwarnings("ignore")
    raw = pd.read_excel(filepath)
    month_label = detect_month(filepath)
    month_short = detect_month_short(filepath)

    # ── 用量行 ──
    usage = raw[raw[COL_PACKAGE].fillna("").astype(str).str.contains("GLM Coding", na=False)].copy()
    if not usage.empty:
        usage[COL_MODEL] = usage[COL_MODEL].fillna("unknown")
        usage[COL_DATE] = usage[COL_DATE].astype(str)
        usage[COL_COST] = usage[COL_AMOUNT] * usage[COL_UNIT_PRICE] / 1000
        usage["月份标签"] = month_label
        usage["月份"] = month_short
        usage["日期排序"] = pd.to_datetime(usage[COL_DATE])
        usage = usage.sort_values("日期排序")

    # ── 订阅套餐行 ──
    plan = raw[raw[COL_PRODUCT_TYPE].fillna("").astype(str).str.contains(SUBSCRIPTION_TYPE, na=False)].copy()
    plan = plan[plan[COL_PRODUCT_NAME].fillna("").astype(str).str.contains("GLM Coding", na=False)].copy()
    if not plan.empty:
        plan[COL_DATE] = plan[COL_DATE].astype(str)
        plan["月份"] = plan[COL_DATE].str[:7]
        plan["日期排序"] = pd.to_datetime(plan[COL_DATE])
        plan = plan.sort_values("日期排序")

    return usage, plan


def load_all(files: list[Path]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """加载并合并所有费用文件，返回 (用量行, 订阅套餐行)"""
    usages, plans = [], []
    for f in files:
        u, p = load_single(f)
        usages.append(u)
        plans.append(p)
        cust = u[COL_CUSTOMER].nunique() if (not u.empty and COL_CUSTOMER in u.columns) else 0
        print(f"  ✓ {detect_month(f)}: 用量 {len(u)} 条 · {u[COL_COST].sum():.2f} 元"
              f" · 套餐 {len(p)} 笔 · {p[COL_PAID].sum():.2f} 元 · 账号 {cust}")

    usages_ne = [u for u in usages if not u.empty]
    plans_ne = [p for p in plans if not p.empty]
    usage_df = (pd.concat(usages_ne, ignore_index=True).sort_values("日期排序")
                if usages_ne else pd.DataFrame())
    plan_df = (pd.concat(plans_ne, ignore_index=True).sort_values("日期排序")
               if plans_ne else pd.DataFrame())
    return usage_df, plan_df


# ── 图表构建 ──────────────────────────────────────────

def build_trend_30d(
    by_date_30: pd.DataFrame,
    max_val: float,
    months: list[str],
) -> str:
    """构建最近 30 日趋势图:柱子(最新月紫青渐变)+ 标签两层"""
    latest = months[-1] if months else ""

    bars = ""
    for _, row in by_date_30.iterrows():
        d = row[COL_DATE]
        cost = row[COL_COST]
        empty = cost < 0.001
        h = (cost / max_val * 100) if (max_val > 0 and not empty) else 0
        if empty:
            cls = "trend-bar empty"
        elif d[:7] == latest:
            cls = "trend-bar jun"
        else:
            cls = "trend-bar"
        if not empty and max_val > 0 and cost >= max_val:
            cls += " peak"
        tip = f' data-tip="{d[-5:]}: {fmt(cost)}元"' if not empty else ""
        bars += f'<div class="{cls}" style="height:{h:.1f}%"{tip}></div>\n'

    labels = ""
    n = len(by_date_30)
    for i, (_, row) in enumerate(by_date_30.iterrows()):
        show = (i % 3 == 0) or (i == n - 1)
        d = row[COL_DATE]
        labels += f'<div class="trend-lbl">{d[-5:] if show else ""}</div>\n'

    nonzero = by_date_30[by_date_30[COL_COST] > 0.001][COL_COST]
    avg_val = float(nonzero.mean()) if len(nonzero) else 0.0
    avg_pct = (avg_val / max_val * 100) if max_val > 0 else 0.0
    avg_line = (f'<div class="trend-avg" style="bottom:{avg_pct:.1f}%"><span>均 {fmt(avg_val)}元</span></div>'
                if avg_val > 0 else "")
    return f'<div class="trend">{bars}{avg_line}</div><div class="trend-labels">{labels}</div>'


def build_sparkline(values: list[float], w: int = 100, h: int = 24) -> str:
    """构建 mini sparkline SVG(归一化面积+折线,用于 KPI 卡片底部 30 天趋势)"""
    vals = [max(0.0, float(v)) for v in values]
    mx = max(vals) if vals else 0.0
    n = len(vals)
    if mx <= 0 or n < 2:
        return ""
    pts = []
    for i, v in enumerate(vals):
        x = (i / (n - 1)) * w
        y = h - (v / mx) * (h - 3) - 1.5
        pts.append(f"{x:.1f},{y:.1f}")
    line = " ".join(pts)
    area = f"0,{h} " + line + f" {w},{h}"
    return (f'<svg class="spark-svg" viewBox="0 0 {w} {h}" preserveAspectRatio="none">'
            f'<polyline class="spark-area" points="{area}"/>'
            f'<polyline class="spark-line" points="{line}" pathLength="1"/></svg>')


def build_rate_trend(rate_map: dict[str, float], all_days: list[str]) -> str:
    """构建百分比小柱图（如每日缓存命中率，0~100%）"""
    bars = ""
    for d in all_days:
        v = rate_map.get(d, 0.0)
        cls = "trend-bar jun" if v > 0 else "trend-bar empty"
        tip = f' data-tip="{d[-5:]}: {v:.0f}%"' if v > 0 else ""
        bars += f'<div class="{cls}" style="height:{v:.1f}%"{tip}></div>\n'
    labels = ""
    n = len(all_days)
    for i, d in enumerate(all_days):
        show = (i % 5 == 0) or (i == n - 1)
        labels += f'<div class="trend-lbl">{d[-5:] if show else ""}</div>\n'
    return f'<div class="trend trend-mini">{bars}</div><div class="trend-labels">{labels}</div>'


def _sector_path(cx: float, cy: float, r: float, start_pct: float, end_pct: float) -> str:
    """生成饼图扇区 SVG path(从顶部起,顺时针;start/end_pct 为 0-100 占比)"""
    a1 = math.radians(start_pct / 100 * 360 - 90)
    a2 = math.radians(end_pct / 100 * 360 - 90)
    x1, y1 = cx + r * math.cos(a1), cy + r * math.sin(a1)
    x2, y2 = cx + r * math.cos(a2), cy + r * math.sin(a2)
    large = 1 if (end_pct - start_pct) > 50 else 0
    return f"M {cx:.2f} {cy:.2f} L {x1:.2f} {y1:.2f} A {r:.2f} {r:.2f} 0 {large} 1 {x2:.2f} {y2:.2f} Z"


def build_pie(
    by_model: pd.DataFrame, total_cost: float
) -> tuple[str, str]:
    """构建饼图 SVG(可交互扇区) 和图例 HTML"""
    total_pie = total_cost if total_cost > 0 else 1
    cx, cy, r = 100.0, 100.0, 80.0
    angle = 0.0
    sectors: list[str] = []
    legend = ""
    for _, row in by_model.iterrows():
        m = row[COL_MODEL]
        c = MODEL_COLORS.get(m, "#94a3b8")
        name = MODEL_NAMES.get(m, m)
        pct = row[COL_COST] / total_pie * 100
        nxt = angle + pct
        mid_a = math.radians((angle + nxt) / 2 / 100 * 360 - 90)
        ex, ey = math.cos(mid_a) * 6, math.sin(mid_a) * 6
        sectors.append(
            f'<path class="pie-sector" d="{_sector_path(cx, cy, r, angle, nxt)}" '
            f'fill="{c}" data-name="{name}" data-pct="{pct:.1f}" '
            f'data-cost="{fmt(row[COL_COST])}元" style="--ex:{ex:.1f}px;--ey:{ey:.1f}px"></path>'
        )
        angle = nxt
        if row[COL_COST] >= 0.005:
            legend += f'''<div class="legend-item" data-name="{name}">
            <span class="legend-dot" style="background:{c}"></span>
            <span class="legend-name">{name}</span>
            <span class="legend-pct">{pct:.1f}%</span>
            <span class="legend-val">{fmt(row[COL_COST])}元</span>
            <div class="legend-bar"><i style="width:{pct:.1f}%;background:{c}"></i></div>
        </div>'''
    svg = (
        '<svg class="pie-svg" viewBox="0 0 200 200" role="img" '
        f'aria-label="模型费用占比饼图">{" ".join(sectors)}</svg>'
    )
    return svg, legend


# ── 子报告构建 ────────────────────────────────────────

def build_date_rows(
    df: pd.DataFrame,
    by_date: pd.DataFrame,
    months: list[str],
    max_daily_val: float,
) -> str:
    """构建每日费用明细表格行"""
    # 预构建日期→月份索引映射，避免 O(n*m) 嵌套
    date_month_idx: dict[str, tuple[int, str]] = {}
    for mi, m in enumerate(months):
        for d in df[df["月份"] == m][COL_DATE].unique():
            date_month_idx[d] = (mi, m)

    rows = ""
    for _, row in by_date.iterrows():
        d = row[COL_DATE]
        cost = row[COL_COST]
        pct = cost / max_daily_val * 100 if max_daily_val > 0 else 0

        m_tag = ""
        if d in date_month_idx:
            mi, m = date_month_idx[d]
            mc = MONTH_COLORS[mi % len(MONTH_COLORS)]
            m_label = m[-2:] + "月"
            m_tag = f'<span class="month-tag" style="background:{mc}20; color:{mc}">{m_label}</span> '

        rows += f'''<tr>
            <td class="td-date">{m_tag}{d}</td>
            <td class="td-cost">{fmt(cost)}元</td>
            <td class="td-tokens">{fmt_tok(row[COL_AMOUNT])}</td>
            <td class="td-reqs">{fmt(row[COL_REQUESTS])}</td>
            <td class="td-bar"><div class="td-bar-track"><div class="inline-bar" style="width:{pct:.0f}%"></div></div></td>
        </tr>'''
    return rows


def build_pricing_rows(detail: pd.DataFrame) -> str:
    """构建模型单价对照表行"""
    model_total_usage = detail.groupby(COL_MODEL)[COL_AMOUNT].sum()
    detail_sorted = detail.copy()
    detail_sorted["_model_total"] = detail_sorted[COL_MODEL].map(model_total_usage)
    detail_sorted = detail_sorted.sort_values(["_model_total", COL_AMOUNT], ascending=[False, False])

    rows = ""
    for _, row in detail_sorted.iterrows():
        m = row[COL_MODEL]
        name = MODEL_NAMES.get(m, m)
        pt = PRICE_TYPE_NAMES.get(row[COL_PRICE_TYPE], row[COL_PRICE_TYPE])
        c = MODEL_COLORS.get(m, "#94a3b8")
        rows += f'''<tr>
            <td><span class="model-tag" style="background:{c}20; color:{c}">{name}</span></td>
            <td>{pt}</td>
            <td>{fmt(row[COL_AMOUNT])} tokens</td>
            <td>{row[COL_UNIT_PRICE] * 1000:.2f}元</td>
            <td class="td-cost">{fmt(row[COL_COST])}元</td>
        </tr>'''
    return rows


# ── 新增分析模块 ──────────────────────────────────────

def build_cost_saving(
    theory_cost: float,
    today: date,
    plan_df: pd.DataFrame,
) -> str:
    """模块1：成本节省 —— 同期按量理论费用 vs 套餐按已用天数摊销。

    套餐为预付,按「已用天数/周期天数」摊销公允价值;季度=3日历月;
    升级时旧套餐在新套餐生效日终止,剩余价值折算已计入新套餐公允价值。
    对比口径为同期(最早订阅日 ~ 今天)。
    """
    # ── 按账号展开订阅,处理升级终止与摊销 ──
    by_acct: dict[str, list[dict]] = {}
    for acct, tier, months, start_str, paid in SUBSCRIPTIONS:
        start = date.fromisoformat(start_str)
        end = start + relativedelta(months=months)
        fair = PLAN_MONTHLY[tier] * months * PLAN_DISCOUNT[months]
        by_acct.setdefault(acct, []).append(
            {"acct": acct, "tier": tier, "months": months,
             "start": start, "end": end, "fair": fair, "paid": paid})

    sub_rows_html = ""
    total_amort = 0.0
    total_paid = 0.0
    for acct in sorted(by_acct):
        subs = sorted(by_acct[acct], key=lambda s: s["start"])
        for i, s in enumerate(subs):
            # 升级终止:同账号下一条生效日早于到期则提前终止
            term = min(s["end"], subs[i + 1]["start"]) if i + 1 < len(subs) else s["end"]
            used_days = max(0, (min(term, today) - s["start"]).days)
            total_days = (s["end"] - s["start"]).days
            amort = s["fair"] * used_days / total_days if total_days else 0.0
            progress = used_days / total_days * 100 if total_days else 0.0
            total_amort += amort
            total_paid += s["paid"]

            pname = {1: "月", 3: "季", 12: "年"}.get(s["months"], str(s["months"]))
            c = PRODUCT_COLORS.get(f"GLM Coding {s['tier']}", "#94a3b8")
            ended = ' <span class="cmp-mono">(升级终止)</span>' if term < s["end"] else ""
            sub_rows_html += f'''<tr>
                <td>{s['acct']}</td>
                <td><span class="model-tag" style="background:{c}20;color:{c}">{s['tier']}·{pname}</span></td>
                <td class="td-date">{s['start']}~{s['end']}{ended}</td>
                <td>{used_days}/{total_days}天</td>
                <td class="td-bar"><div class="td-bar-track"><div class="inline-bar" style="width:{progress:.0f}%"></div></div></td>
                <td class="td-cost">¥{amort:.2f}</td>
                <td class="td-cost">¥{s['paid']:.2f}</td>
            </tr>'''

    saved = theory_cost - total_amort
    amort_rate = (total_amort / theory_cost * 100) if theory_cost > 0 else 0
    plan_paid_sum = plan_df[COL_PAID].sum() if not plan_df.empty else 0.0

    return f'''<div class="saving-wrap reveal">
        <div class="saving-bar">
            <div class="saving-cell">
                <div class="saving-label">同期按量理论费用</div>
                <div class="saving-val">¥<span class="counter" data-target="{theory_cost:.2f}" data-decimals="2">0</span></div>
                <div class="saving-sub">用量×单价 累加</div>
            </div>
            <div class="saving-arrow">→</div>
            <div class="saving-cell">
                <div class="saving-label">套餐已摊销</div>
                <div class="saving-val">¥<span class="counter" data-target="{total_amort:.2f}" data-decimals="2">0</span></div>
                <div class="saving-track"><div class="saving-fill" style="width:{amort_rate:.1f}%"></div></div>
                <div class="saving-sub">摊销占理论 {amort_rate:.1f}%</div>
            </div>
            <div class="saving-arrow">→</div>
            <div class="saving-cell saving-win">
                <div class="saving-label">同期省下</div>
                <div class="saving-val"><span class="grad">¥<span class="counter" data-target="{saved:.2f}" data-decimals="2">0</span></span></div>
                <div class="saving-sub">套餐预付杠杆</div>
            </div>
        </div>
        <div class="table-card">
            <table><thead><tr><th>账号</th><th>套餐</th><th>周期</th><th>已用/总天数</th><th>进度</th><th>摊销成本</th><th>实付</th></tr></thead>
            <tbody>{sub_rows_html}</tbody></table>
        </div>
        <div class="saving-note">口径:套餐为预付,摊销 = 公允价值 × 已用天数/周期天数;季度按 3 个日历月;升级折算已计入公允价值。实付合计 ¥{total_paid:.2f}(账单校验 ¥{plan_paid_sum:.2f}),对比口径为同期 [最早订阅 ~ {today}]。</div>
    </div>'''


def build_cache_saving(
    df: pd.DataFrame,
    detail: pd.DataFrame,
    all_days: list[str],
) -> str:
    """模块2：缓存节省 —— 缓存命中省了多少钱 + 各模型折扣 + 命中率趋势"""
    # ── 各模型缓存节省 ──
    rows_data: list[tuple] = []
    total_save = 0.0
    for m in detail[COL_MODEL].unique():
        sub = detail[detail[COL_MODEL] == m]
        in_r = sub[sub[COL_PRICE_TYPE] == "输入"]
        ca_r = sub[sub[COL_PRICE_TYPE] == "缓存命中"]
        if in_r.empty or ca_r.empty:
            continue
        p_in = float(in_r[COL_UNIT_PRICE].iloc[0])
        p_ca = float(ca_r[COL_UNIT_PRICE].iloc[0])
        q_ca = float(ca_r[COL_AMOUNT].iloc[0])
        save = q_ca * (p_in - p_ca) / 1000
        total_save += save
        rows_data.append((m, p_in, p_ca, q_ca, save))
    rows_data.sort(key=lambda x: -x[4])

    cache_rows = ""
    for m, p_in, p_ca, q_ca, save in rows_data:
        name = MODEL_NAMES.get(m, m)
        c = MODEL_COLORS.get(m, "#94a3b8")
        disc = (1 - p_ca / p_in) * 100 if p_in > 0 else 0
        cache_rows += f'''<tr>
            <td><span class="model-tag" style="background:{c}20;color:{c}">{name}</span></td>
            <td>{p_in * 1000:.2f}</td>
            <td>{p_ca * 1000:.2f}</td>
            <td><span class="hl-pos">-{disc:.0f}%</span></td>
            <td>{fmt_tok(q_ca)}</td>
            <td class="td-cost">¥{fmt(save)}</td>
        </tr>'''

    return f'''<div class="section-title reveal">⚡ 缓存节省分析（近 30 天）</div>
    <div class="cache-wrap reveal">
        <div class="cache-hero panel">
            <div class="cache-label">缓存命中为你省下</div>
            <div class="cache-val"><span class="grad">¥<span class="counter" data-target="{total_save:.2f}" data-decimals="2">0</span></span></div>
            <div class="cache-sub">缓存单价仅为输入价的 1/5 ~ 1/4</div>
        </div>
        <div class="table-card"><table>
            <thead><tr><th>模型</th><th>输入价(元/百万)</th><th>缓存价</th><th>折扣</th><th>缓存用量</th><th>节省</th></tr></thead>
            <tbody>{cache_rows}</tbody>
        </table></div>
    </div>'''


def build_model_efficiency(df: pd.DataFrame) -> str:
    """模块3：模型性价比排行 —— 单次请求 token / 单次成本（请求去重）"""
    # 同天同模型 输入/输出/缓存 三条请求次数相同，取 first 避免重复 3 倍
    pdm = df.groupby([COL_DATE, COL_MODEL]).agg(
        请求=(COL_REQUESTS, "first"), tokens=(COL_AMOUNT, "sum"), cost=(COL_COST, "sum"))
    eff = pdm.groupby(COL_MODEL).sum()
    eff["单次token"] = eff["tokens"] / eff["请求"]
    eff["单次成本"] = eff["cost"] / eff["请求"]
    eff = eff.sort_values("单次成本")

    rows = ""
    for m, r in eff.iterrows():
        name = MODEL_NAMES.get(m, m)
        c = MODEL_COLORS.get(m, "#94a3b8")
        rows += f'''<tr>
            <td><span class="model-tag" style="background:{c}20;color:{c}">{name}</span></td>
            <td class="td-cost">¥{fmt(r['cost'])}</td>
            <td>{fmt_tok(r['tokens'])}</td>
            <td>{fmt_tok(r['单次token'])}</td>
            <td class="td-cost">¥{r['单次成本']:.3f}</td>
        </tr>'''
    return rows


def build_monthly_compare(usage_full: pd.DataFrame) -> str:
    """模块4：月度环比 —— 最近两个月 日均(费用/token/请求)对比,消除当月未满的天数差异"""
    pdm = usage_full.groupby([COL_DATE, COL_MODEL]).agg(
        请求=(COL_REQUESTS, "first"), tokens=(COL_AMOUNT, "sum"), cost=(COL_COST, "sum")).reset_index()
    pdm["月份"] = pdm[COL_DATE].str[:7]
    monthly = pdm.groupby("月份").agg(
        费用=("cost", "sum"), tokens=("tokens", "sum"), 请求=("请求", "sum"),
        天数=(COL_DATE, "nunique"))
    months = sorted(monthly.index.tolist())
    if len(months) < 2:
        return ""
    prev_m, cur_m = months[-2], months[-1]
    p, c = monthly.loc[prev_m], monthly.loc[cur_m]
    # 日均 = 总量 / 有数据天数(消除当月未满的天数差异,公平环比)
    def daily(row, col):
        return row[col] / row["天数"] if row["天数"] else 0.0
    pv = {k: daily(p, k) for k in ("费用", "tokens", "请求")}
    cv = {k: daily(c, k) for k in ("费用", "tokens", "请求")}

    def growth(pv_: float, cv_: float) -> float:
        return ((cv_ - pv_) / pv_ * 100) if pv_ else 0.0

    def card(label: str, pval: float, cval: float, fmt_fn) -> str:
        g = growth(pval, cval)
        up = cval >= pval
        arrow, cls = ("↑", "hl-pos") if up else ("↓", "hl-neg")
        return f'''<div class="cmp-card">
            <div class="cmp-label">{label} · 日均</div>
            <div class="cmp-from">{fmt_fn(pval)} <span class="cmp-mono">{prev_m}·{int(p["天数"])}天</span></div>
            <div class="cmp-arrow {cls}">{arrow} {abs(g):.0f}%</div>
            <div class="cmp-to">{fmt_fn(cval)} <span class="cmp-mono">{cur_m}·{int(c["天数"])}天</span></div>
        </div>'''

    cards = "".join([
        card("费用", pv["费用"], cv["费用"], lambda x: f"¥{fmt(x)}/天"),
        card("Token", pv["tokens"], cv["tokens"], lambda x: fmt_tok(x) + "/天"),
        card("请求", pv["请求"], cv["请求"], lambda x: f"{int(x):,}/天"),
    ])
    return f'''<div class="section-title reveal">📅 月度环比 · {prev_m} → {cur_m}（日均,当月进行中）</div>
    <div class="cmp-grid reveal">{cards}</div>'''


# ── 主报告 ────────────────────────────────────────────

def build_report(usage_full: pd.DataFrame, plan_df: pd.DataFrame) -> str:
    """从合并后的 DataFrame 构建完整的 HTML 报告"""
    # ── 30 天滚动窗口(排除今日,当天数据不完整)──
    now_cn = datetime.now(timezone(timedelta(hours=8)))
    window_end = (now_cn - timedelta(days=1)).strftime("%Y-%m-%d")
    window_start = (now_cn - timedelta(days=30)).strftime("%Y-%m-%d")
    df = usage_full[usage_full[COL_DATE].between(window_start, window_end)].copy()

    # ── 全局汇总（窗口内用量）──
    total_cost = df[COL_COST].sum()
    total_paid = df[COL_PAID].sum()

    input_mask = df[COL_PRICE_TYPE] == "输入"
    output_mask = df[COL_PRICE_TYPE] == "输出"
    cache_mask = df[COL_PRICE_TYPE] == "缓存命中"

    total_input = df.loc[input_mask, COL_AMOUNT].sum()
    total_output = df.loc[output_mask, COL_AMOUNT].sum()
    total_cache = df.loc[cache_mask, COL_AMOUNT].sum()
    input_cost = df.loc[input_mask, COL_COST].sum()
    output_cost = df.loc[output_mask, COL_COST].sum()
    cache_cost = df.loc[cache_mask, COL_COST].sum()

    total_reqs = df[COL_REQUESTS].sum()
    dates = sorted(df[COL_DATE].unique())
    n_days = len(dates)
    n_models = df[COL_MODEL].nunique()
    n_accounts = df[COL_CUSTOMER].nunique() if COL_CUSTOMER in df.columns else 1
    avg_daily = total_cost / n_days if n_days > 0 else 0
    avg_daily_tok = (total_input + total_output + total_cache) / n_days if n_days > 0 else 0

    # ── 按模型汇总 ──
    by_model = (df.groupby(COL_MODEL)
        .agg({COL_COST: "sum", COL_AMOUNT: "sum", COL_REQUESTS: "sum"})
        .sort_values(COL_COST, ascending=False).reset_index())

    # ── 按日期汇总 ──
    by_date = (df.groupby(COL_DATE)
        .agg({COL_COST: "sum", COL_AMOUNT: "sum", COL_REQUESTS: "sum"})
        .sort_index().reset_index())

    # ── 模型×价格类型 ──
    detail = (df.groupby([COL_MODEL, COL_PRICE_TYPE])
        .agg({COL_AMOUNT: "sum", COL_UNIT_PRICE: "first", COL_COST: "sum", COL_REQUESTS: "sum"})
        .reset_index())

    # ── 月份列表 ──
    months = sorted(df["月份"].unique())

    # ── 最近 30 天趋势（窗口内补 0）──
    all_30_days = pd.date_range(start=window_start, end=window_end, freq="D").strftime("%Y-%m-%d").tolist()

    cost_map = dict(zip(by_date[COL_DATE], by_date[COL_COST]))
    reqs_map = dict(zip(by_date[COL_DATE], by_date[COL_REQUESTS]))
    tok_map = dict(zip(by_date[COL_DATE], by_date[COL_AMOUNT]))
    by_date_30 = pd.DataFrame({
        COL_DATE: all_30_days,
        COL_COST: [cost_map.get(d, 0.0) for d in all_30_days],
        COL_REQUESTS: [reqs_map.get(d, 0) for d in all_30_days],
        COL_AMOUNT: [tok_map.get(d, 0) for d in all_30_days],
    })

    max_daily_val = by_date_30[COL_COST].max() if len(by_date_30) > 0 else 1
    cache_rate = total_cache / (total_input + total_cache) * 100 if (total_input + total_cache) > 0 else 0
    io_ratio = total_input / total_output if total_output > 0 else 0
    cache_ring_offset = PIE_SVG_CIRCUMFERENCE * (1 - cache_rate / 100)

    bar_html = build_trend_30d(by_date_30, max_daily_val, months)
    hero_spark = build_sparkline(by_date_30[COL_COST].tolist())
    req_spark = build_sparkline([float(x) for x in by_date_30[COL_REQUESTS].tolist()])
    _cache_daily = df[df[COL_PRICE_TYPE] == "缓存命中"].groupby(COL_DATE)[COL_COST].sum()
    cache_spark = build_sparkline([float(_cache_daily.get(d, 0.0)) for d in all_30_days])
    _io_daily = df[df[COL_PRICE_TYPE].isin(["输入", "输出"])].groupby(COL_DATE)[COL_COST].sum()
    io_spark = build_sparkline([float(_io_daily.get(d, 0.0)) for d in all_30_days])
    pie_grad, pie_legend = build_pie(by_model, total_cost)

    # ── 子表格构建 ──
    date_rows = build_date_rows(df, by_date, months, max_daily_val)
    pricing_rows = build_pricing_rows(detail)

    month_labels = ", ".join(sorted(df["月份标签"].unique()))

    # ── 新增模块 ──
    theory_cost_full = usage_full[COL_COST].sum()
    cost_saving_html = build_cost_saving(theory_cost_full, now_cn.date(), plan_df)
    cache_saving_html = build_cache_saving(df, detail, all_30_days)
    model_efficiency_rows = build_model_efficiency(df)
    monthly_compare_html = build_monthly_compare(usage_full)

    # ── 加载模板并渲染 ──
    css = STYLE_FILE.read_text(encoding="utf-8")
    template = TEMPLATE_FILE.read_text(encoding="utf-8")
    gen_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 新增派生量(占位符参数)
    cache_cost_share = f"{cache_cost / total_cost * 100:.0f}%" if total_cost > 0 else "0%"
    io_total_tok = fmt_tok(total_input + total_output)
    avg_daily_reqs = f"{total_reqs // n_days:,}" if n_days > 0 else "0"

    return render(template,
        css=css,
        month_labels=month_labels,
        date_start=dates[0],
        date_end=dates[-1],
        n_days=n_days,
        n_models=n_models,
        n_months=len(months),
        n_accounts=n_accounts,
        window_start=window_start,
        window_end=window_end,
        cache_cost_share=cache_cost_share,
        io_total_tok=io_total_tok,
        avg_daily_reqs=avg_daily_reqs,
        input_cost=fmt(input_cost) + "元",
        input_cost_raw=input_cost,
        total_input_tok=fmt_tok(total_input),
        output_cost=fmt(output_cost) + "元",
        output_cost_raw=output_cost,
        total_output_tok=fmt_tok(total_output),
        cache_cost=fmt(cache_cost) + "元",
        cache_cost_raw=cache_cost,
        total_cache_tok=fmt_tok(total_cache),
        total_cost=fmt(total_cost) + "元",
        total_cost_raw=total_cost,
        avg_daily=fmt(avg_daily),
        avg_daily_raw=avg_daily,
        total_reqs=f"{total_reqs:,}",
        bar_html=bar_html,
        hero_spark=hero_spark,
        cache_spark=cache_spark,
        io_spark=io_spark,
        req_spark=req_spark,
        pie_grad=pie_grad,
        total_cost_fmt=fmt(total_cost),
        pie_legend=pie_legend,
        date_rows=date_rows,
        pricing_rows=pricing_rows,
        avg_daily_tok=fmt_tok(avg_daily_tok),
        cache_rate=f"{cache_rate:.1f}%",
        cache_ring_offset=f"{cache_ring_offset:.1f}",
        io_ratio=f"{io_ratio:.1f}:1",
        gen_time=gen_time,
        cost_saving_html=cost_saving_html,
        cache_saving_html=cache_saving_html,
        model_efficiency_rows=model_efficiency_rows,
        monthly_compare_html=monthly_compare_html,
    )


# ── 入口 ──────────────────────────────────────────────

def main() -> None:
    print("📊 智谱AI 费用分析工具（统一报告）")
    print("=" * 40)

    files = sorted(SCRIPT_DIR.glob("智谱AI开放平台费用明细*.xlsx"))
    if len(sys.argv) > 1:
        files = [Path(a) for a in sys.argv[1:] if a.endswith(".xlsx")]

    if not files:
        print("❌ 未找到任何 xlsx 费用明细文件")
        sys.exit(1)

    print(f"✓ 找到 {len(files)} 个账单文件:")
    usage_df, plan_df = load_all(files)

    total = usage_df[COL_COST].sum()
    paid = plan_df[COL_PAID].sum() if not plan_df.empty else 0.0
    print(f"\n  总计: 用量 {len(usage_df)} 条 · 同期按量理论 {total:.2f} 元"
          f" · 套餐实付(累计) {paid:.2f} 元")
    print(f"  💡 成本节省按「已用天数摊销」计算,详见报告")

    html = build_report(usage_df, plan_df)
    output = SCRIPT_DIR / "bill_report.html"
    output.write_text(html, encoding="utf-8")
    print(f"\n✅ 统一报告已生成: {output}")

    try:
        import webbrowser
        webbrowser.open(output.as_uri())
        print("🌐 已在浏览器中打开")
    except Exception:
        print(f"💡 请手动打开: {output}")


if __name__ == "__main__":
    main()
