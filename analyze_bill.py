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
TEMPLATE_FILE = SCRIPT_DIR / "report_template.html"
STYLE_FILE = SCRIPT_DIR / "report_style.css"

# ── 列名常量 ──────────────────────────────────────────
COL_MODEL = "模型编码（推理专用）"
COL_PACKAGE = "Tokens资源包名称"
COL_DATE = "账期(自然日)"
COL_PRICE_TYPE = "价格类型"
COL_AMOUNT = "用量"
COL_UNIT_PRICE = "单价"
COL_COST = "理论费用"  # 账单无此列;为「用量×单价/1000」重算的按量目录理论费用列名
COL_PAID = "已付款金额"
COL_REQUESTS = "请求次数 (仅API)"
COL_CUSTOMER = "客户id"
COL_PRODUCT_TYPE = "产品类型"
COL_PRODUCT_NAME = "模型产品名称"
SUBSCRIPTION_TYPE = "订阅套餐"
# 用量行/套餐行依赖的账单原始列(缺失则报错;COL_COST 为重算列不在此列)
REQUIRED_COLS = [
    COL_PACKAGE, COL_MODEL, COL_DATE, COL_PRICE_TYPE,
    COL_AMOUNT, COL_UNIT_PRICE, COL_REQUESTS, COL_CUSTOMER,
    COL_PRODUCT_TYPE, COL_PRODUCT_NAME, COL_PAID,
]

# ── 预编译正则 ────────────────────────────────────────
_RE_MONTH = re.compile(r"(\d{4})-(\d{2})")

MODEL_NAMES: dict[str, str] = {
    "glm-5.2": "GLM-5.2",
    "glm-5.1": "GLM-5.1",
    "glm-5-turbo": "GLM-5-Turbo",
    "glm-5": "GLM-5",
    "glm-4.7": "GLM-4.7",
    "glm-4.6v": "GLM-4.6V",
    "glm-4.5-air": "GLM-4.5-Air",
}

PRICE_TYPE_NAMES: dict[str, str] = {
    "输入": "输入",
    "输出": "输出",
    "缓存命中": "缓存命中",
    "不区分输入输出": "调用次数",
}

# 套餐 tier 配色(build_cost_saving 用)
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
    ("账号A", "Max", 3, "2026-05-20", 886.10),   # 升级;fair=Max季价1266.30,paid886.10已抵扣Pro剩余价值
    ("账号B", "Max", 1, "2026-06-13", 469.00),   # 独立账号
]



# ── 工具函数 ──────────────────────────────────────────

def fmt(n: float | int) -> str:
    """格式化数值:大数千分位,小数保留适当精度;NaN/inf→—,0 与负数安全"""
    if isinstance(n, float):
        if math.isnan(n) or math.isinf(n):
            return "—"
        if n == 0:
            return "0"
        a = abs(n)
        if a >= 1:
            return f"{n:,.2f}"
        if a >= 0.01:
            return f"{n:.4f}"
        return f"{n:.6f}"
    return f"{n:,}"


def fmt_tok(n: int | float) -> str:
    """格式化 token 数量为可读形式(K/M);NaN/inf→—"""
    if isinstance(n, float) and (math.isnan(n) or math.isinf(n)):
        return "—"
    if n >= 1_000_000:
        return f"{n / 1_000_000:,.1f}M"
    if n >= 1_000:
        return f"{n / 1_000:,.1f}K"
    return str(int(n))


def detect_month(filepath: str | Path) -> str:
    """从文件名提取月份标签，如 '2026年5月'"""
    name = Path(filepath).name
    m = _RE_MONTH.search(name)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月"
    return "未知月份"


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
    # 列存在性校验(账单表头若变更,给出清晰提示而非晦涩的 KeyError)
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        raise ValueError(
            f"账单缺少必需列: {missing}\n"
            f"实际列名: {list(raw.columns)}\n"
            f"→ 可能智谱平台导出表头已变更,请核对 analyze_bill.py 顶部的列名常量。"
        )

    # ── 用量行 ──
    usage = raw[raw[COL_PACKAGE].fillna("").astype(str).str.contains("GLM Coding", na=False)].copy()
    if not usage.empty:
        usage[COL_MODEL] = usage[COL_MODEL].fillna("unknown")
        usage[COL_DATE] = usage[COL_DATE].astype(str)
        usage[COL_COST] = usage[COL_AMOUNT] * usage[COL_UNIT_PRICE] / 1000
        # 月份基于行日期(与 plan 一致),避免跨月文件时文件名月份≠行日期
        usage["月份"] = usage[COL_DATE].str[:7]
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
    """加载并合并所有费用文件，返回 (用量行, 订阅套餐行)

    单个文件损坏(读取异常)或表头不符(列校验失败)时跳过并警告,不影响其他文件。
    """
    usages, plans = [], []
    for f in files:
        try:
            u, p = load_single(f)
        except Exception as e:
            print(f"  ⚠ 跳过无法解析的文件 {f.name}: {type(e).__name__}: {e}")
            continue
        usages.append(u)
        plans.append(p)
        cust = u[COL_CUSTOMER].nunique() if (not u.empty and COL_CUSTOMER in u.columns) else 0
        u_cost = u[COL_COST].sum() if (not u.empty and COL_COST in u.columns) else 0.0
        p_paid = p[COL_PAID].sum() if (not p.empty and COL_PAID in p.columns) else 0.0
        print(f"  ✓ {detect_month(f)}: 用量 {len(u)} 条 · {u_cost:.2f} 元"
              f" · 套餐 {len(p)} 笔 · {p_paid:.2f} 元 · 账号 {cust}")

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
    n_sectors = len(by_model)
    for i, (_, row) in enumerate(by_model.iterrows()):
        m = row[COL_MODEL]
        c = MODEL_COLORS.get(m, "#94a3b8")
        name = MODEL_NAMES.get(m, m)
        pct = row[COL_COST] / total_pie * 100
        # 末扇区强制闭合到 100%,消除浮点累积导致的缝隙
        nxt = 100.0 if i == n_sectors - 1 else angle + pct
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
    window_start: str,
    window_end: str,
    plan_df: pd.DataFrame,
) -> str:
    """模块1:成本节省 —— 近30天窗口内 按量理论费用 vs 套餐摊销。

    与报告主体(近30天)口径一致:theory 与 amort 都限定在30天窗口。
    套餐摊销 = 公允价值 × (套餐周期 ∩ 窗口)天数 / 周期总天数;
    完全在窗口外的订阅(如已终止的旧套餐)不计入近30天成本。
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
    win_start_d = date.fromisoformat(window_start)
    win_end_d = date.fromisoformat(window_end)
    for acct in sorted(by_acct):
        subs = sorted(by_acct[acct], key=lambda s: s["start"])
        for i, s in enumerate(subs):
            # 升级终止:同账号下一条生效日早于到期则提前终止
            term = min(s["end"], subs[i + 1]["start"]) if i + 1 < len(subs) else s["end"]
            total_days = (s["end"] - s["start"]).days
            # 近30天窗口内已用 = 套餐周期 ∩ 窗口(升级终止用 term)
            used_days = max(0, (min(term, win_end_d) - max(s["start"], win_start_d)).days)
            if used_days <= 0:
                continue  # 完全在窗口外(如已终止的旧套餐),不计入近30天成本
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
        <div class="saving-note">口径:近30天窗口 [{window_start} ~ {window_end}],与报告主体一致。摊销 = 公允价值 × (套餐周期∩窗口)天数 / 周期总天数;窗口外已终止套餐(如账号A Pro·季)不计入。窗口内活跃订阅实付 ¥{total_paid:.2f}(账单全部实付 ¥{plan_paid_sum:.2f})。</div>
    </div>'''


def build_cache_saving(
    df: pd.DataFrame,
    detail: pd.DataFrame,
    all_days: list[str],
) -> str:
    """模块2：缓存节省 —— 加权均价口径(自动处理同模型多档单价) + 命中率趋势"""
    # ── 各模型缓存节省(加权均价) ──
    # 缓存节省 = 缓存用量 × 输入加权均价 − 缓存实际费用
    # 加权均价 = 各档总费用 / 各档总量(元/token);避免旧版 iloc[0] 只取第一档单价,
    # 同模型多上下文档时旧版节省偏差可达 10%+。
    rows_data: list[tuple] = []
    total_save = 0.0
    for m in detail[COL_MODEL].unique():
        sub = detail[detail[COL_MODEL] == m]
        in_r = sub[sub[COL_PRICE_TYPE] == "输入"]
        ca_r = sub[sub[COL_PRICE_TYPE] == "缓存命中"]
        if in_r.empty or ca_r.empty:
            continue
        in_cost = float(in_r[COL_COST].sum())
        in_qty = float(in_r[COL_AMOUNT].sum())
        ca_cost = float(ca_r[COL_COST].sum())
        ca_qty = float(ca_r[COL_AMOUNT].sum())
        if in_qty <= 0 or ca_qty <= 0:
            continue
        in_avg = in_cost / in_qty            # 元/token(加权)
        ca_avg = ca_cost / ca_qty
        save = ca_qty * in_avg - ca_cost
        disc = (1 - ca_avg / in_avg) * 100
        total_save += save
        # 单价显示转 元/百万tokens
        rows_data.append((m, in_avg * 1_000_000, ca_avg * 1_000_000, ca_qty, save, disc))
    rows_data.sort(key=lambda x: -x[4])

    cache_rows = ""
    for m, p_in, p_ca, q_ca, save, disc in rows_data:
        name = MODEL_NAMES.get(m, m)
        c = MODEL_COLORS.get(m, "#94a3b8")
        cache_rows += f'''<tr>
            <td><span class="model-tag" style="background:{c}20;color:{c}">{name}</span></td>
            <td>{p_in:.2f}</td>
            <td>{p_ca:.2f}</td>
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
    """模块3：模型性价比排行 —— 单次请求 token / 单次成本

    请求次数取「同天同模型 输入行之和」:每次 API 调用必产生输入,且输入唯一落入
    某个上下文长度档(如 GLM-5.1 的 [0,32K)/[32K+)),故输入请求次数之和 = 真实调用数。
    旧版 groupby+first 只取到某档请求次数,严重低估(实测偏低约 2 倍,单次成本虚高)。
    """
    # 请求:输入行之和;tokens/cost:全价格类型之和
    req = (df[df[COL_PRICE_TYPE] == "输入"]
           .groupby([COL_DATE, COL_MODEL])[COL_REQUESTS].sum())
    tok_cost = df.groupby([COL_DATE, COL_MODEL]).agg(
        tokens=(COL_AMOUNT, "sum"), cost=(COL_COST, "sum"))
    eff = tok_cost.join(req.rename("请求"), how="left").fillna(0)
    eff = eff.groupby(COL_MODEL).sum()
    eff_req = eff["请求"]
    safe_req = eff_req.replace(0, float("nan"))
    eff["单次token"] = eff["tokens"] / safe_req
    eff["单次成本"] = eff["cost"] / safe_req
    eff = eff.sort_values("单次成本", na_position="last")

    rows = ""
    for m, r in eff.iterrows():
        name = MODEL_NAMES.get(m, m)
        c = MODEL_COLORS.get(m, "#94a3b8")
        has_req = r["请求"] > 0
        tok_per = fmt_tok(r["单次token"]) if has_req else "—"
        cost_per = f"¥{r['单次成本']:.3f}" if has_req else "—"
        rows += f'''<tr>
            <td><span class="model-tag" style="background:{c}20;color:{c}">{name}</span></td>
            <td class="td-cost">¥{fmt(r['cost'])}</td>
            <td>{fmt_tok(r['tokens'])}</td>
            <td>{tok_per}</td>
            <td class="td-cost">{cost_per}</td>
        </tr>'''
    return rows


def build_biweekly_compare(
    df: pd.DataFrame, window_start: str, window_end: str
) -> str:
    """模块4:双周环比 —— 近两周 vs 前两周(都在30天窗口内,完整周期公平对比)。

    报告主体为近30天,旧版月度环比会跨月并用到窗口外数据、且当月进行中不公平;
    改为双周(各14天)环比,完全落在窗口内且周期等长。
    """
    we = date.fromisoformat(window_end)
    this_start = (we - timedelta(days=13)).isoformat()   # 近14天 [we-13, we]
    prev_start = (we - timedelta(days=27)).isoformat()   # 前14天 [we-27, we-14]
    prev_end = (we - timedelta(days=14)).isoformat()

    def agg(sub: pd.DataFrame) -> dict:
        if sub.empty:
            return {"费用": 0.0, "tokens": 0.0, "请求": 0, "天数": 0}
        return {
            "费用": float(sub[COL_COST].sum()),
            "tokens": float(sub[COL_AMOUNT].sum()),
            "请求": int(sub.loc[sub[COL_PRICE_TYPE] == "输入", COL_REQUESTS].sum()),
            "天数": int(sub[COL_DATE].nunique()),
        }

    p = agg(df[df[COL_DATE].between(prev_start, prev_end)])
    c = agg(df[df[COL_DATE].between(this_start, window_end)])
    if p["天数"] == 0 or c["天数"] == 0:
        return ""   # 某段无数据,无法环比

    def daily(row, col):
        return row[col] / row["天数"] if row["天数"] else 0.0

    def growth(pv_: float, cv_: float) -> float:
        return ((cv_ - pv_) / pv_ * 100) if pv_ else 0.0

    def card(label: str, pval: float, cval: float, fmt_fn) -> str:
        g = growth(pval, cval)
        up = cval >= pval
        arrow, cls = ("↑", "hl-pos") if up else ("↓", "hl-neg")
        return f'''<div class="cmp-card">
            <div class="cmp-label">{label} · 日均</div>
            <div class="cmp-from">{fmt_fn(pval)} <span class="cmp-mono">{prev_start[-5:]}~{prev_end[-5:]}·{p["天数"]}天</span></div>
            <div class="cmp-arrow {cls}">{arrow} {abs(g):.0f}%</div>
            <div class="cmp-to">{fmt_fn(cval)} <span class="cmp-mono">{this_start[-5:]}~{window_end[-5:]}·{c["天数"]}天</span></div>
        </div>'''

    pv = {k: daily(p, k) for k in ("费用", "tokens", "请求")}
    cv = {k: daily(c, k) for k in ("费用", "tokens", "请求")}
    cards = "".join([
        card("费用", pv["费用"], cv["费用"], lambda x: f"¥{fmt(x)}/天"),
        card("Token", pv["tokens"], cv["tokens"], lambda x: fmt_tok(x) + "/天"),
        card("请求", pv["请求"], cv["请求"], lambda x: f"{round(x):,}/天"),
    ])
    return f'''<div class="section-title reveal">📅 双周环比（近两周 vs 前两周,日均）</div>
    <div class="cmp-grid reveal">{cards}</div>'''


# ── 主报告 ────────────────────────────────────────────

def build_report(usage_full: pd.DataFrame, plan_df: pd.DataFrame) -> str:
    """从合并后的 DataFrame 构建完整的 HTML 报告"""
    # ── 30 天滚动窗口(排除今日,当天数据不完整)──
    now_cn = datetime.now(timezone(timedelta(hours=8)))
    window_end = (now_cn - timedelta(days=1)).strftime("%Y-%m-%d")
    window_start = (now_cn - timedelta(days=30)).strftime("%Y-%m-%d")
    df = usage_full[usage_full[COL_DATE].between(window_start, window_end)].copy()
    if df.empty and not usage_full.empty:
        # 近30天窗口内无数据(数据已超出窗口,如一个月后再跑) → 回退到全量,避免空报告崩溃
        df = usage_full.copy()
        window_start = df[COL_DATE].min()
        window_end = df[COL_DATE].max()

    # ── 全局汇总（窗口内用量）──
    total_cost = df[COL_COST].sum()

    input_mask = df[COL_PRICE_TYPE] == "输入"
    output_mask = df[COL_PRICE_TYPE] == "输出"
    cache_mask = df[COL_PRICE_TYPE] == "缓存命中"

    total_input = df.loc[input_mask, COL_AMOUNT].sum()
    total_output = df.loc[output_mask, COL_AMOUNT].sum()
    total_cache = df.loc[cache_mask, COL_AMOUNT].sum()
    input_cost = df.loc[input_mask, COL_COST].sum()
    output_cost = df.loc[output_mask, COL_COST].sum()
    cache_cost = df.loc[cache_mask, COL_COST].sum()

    # 请求次数按输入行去重(=真实 API 调用数)
    total_reqs = int(df.loc[df[COL_PRICE_TYPE] == "输入", COL_REQUESTS].sum())
    dates = sorted(df[COL_DATE].unique())
    n_days = len(dates)
    n_models = df[COL_MODEL].nunique()
    n_accounts = df[COL_CUSTOMER].nunique() if COL_CUSTOMER in df.columns else 1
    avg_daily = total_cost / n_days if n_days > 0 else 0

    # ── 按模型汇总(请求按输入行去重,避免输入/输出/缓存三行重复 3 倍) ──
    by_model = (df.groupby(COL_MODEL)
        .agg({COL_COST: "sum", COL_AMOUNT: "sum"})
        .sort_values(COL_COST, ascending=False).reset_index())
    _req_by_model = df[df[COL_PRICE_TYPE] == "输入"].groupby(COL_MODEL)[COL_REQUESTS].sum()
    by_model[COL_REQUESTS] = by_model[COL_MODEL].map(_req_by_model).fillna(0).astype(int)
    # 占比 <10% 的模型合并为"其它"(饼图扇区过多时不清晰)
    if total_cost > 0 and len(by_model) > 1:
        _pct = by_model[COL_COST] / total_cost
        _small = by_model[_pct < 0.10]
        if not _small.empty:
            _other = pd.DataFrame([{
                COL_MODEL: "其它", COL_COST: _small[COL_COST].sum(),
                COL_AMOUNT: _small[COL_AMOUNT].sum(),
                COL_REQUESTS: int(_small[COL_REQUESTS].sum()),
            }])
            by_model = pd.concat([by_model[_pct >= 0.10], _other], ignore_index=True)

    # ── 按日期汇总(请求按输入行去重) ──
    by_date = (df.groupby(COL_DATE)
        .agg({COL_COST: "sum", COL_AMOUNT: "sum"}).sort_index().reset_index())
    _req_by_date = df[df[COL_PRICE_TYPE] == "输入"].groupby(COL_DATE)[COL_REQUESTS].sum()
    by_date[COL_REQUESTS] = by_date[COL_DATE].map(_req_by_date).fillna(0).astype(int)

    # ── 模型×价格类型×单价档(同模型多上下文档各自成行,单价准确) ──
    detail = (df.groupby([COL_MODEL, COL_PRICE_TYPE, COL_UNIT_PRICE])
        .agg({COL_AMOUNT: "sum", COL_COST: "sum", COL_REQUESTS: "sum"})
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

    # ── 新增模块 ──
    # 口径:近30天窗口(theory 与 amort 都限定窗口,与报告主体一致)
    theory_cost_full = df[COL_COST].sum()
    cost_saving_html = build_cost_saving(theory_cost_full, window_start, window_end, plan_df)
    cache_saving_html = build_cache_saving(df, detail, all_30_days)
    model_efficiency_rows = build_model_efficiency(df)
    monthly_compare_html = build_biweekly_compare(df, window_start, window_end)

    # ── 加载模板并渲染 ──
    css = STYLE_FILE.read_text(encoding="utf-8")
    template = TEMPLATE_FILE.read_text(encoding="utf-8")
    gen_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 新增派生量(占位符参数)
    cache_cost_share = f"{cache_cost / total_cost * 100:.0f}%" if total_cost > 0 else "0%"
    io_total_tok = fmt_tok(total_input + total_output)
    avg_daily_reqs = f"{round(total_reqs / n_days):,}" if n_days > 0 else "0"

    return render(template,
        css=css,
        n_days=n_days,
        n_models=n_models,
        n_accounts=n_accounts,
        window_start=window_start,
        window_end=window_end,
        cache_cost_share=cache_cost_share,
        io_total_tok=io_total_tok,
        avg_daily_reqs=avg_daily_reqs,
        input_cost_raw=input_cost,
        output_cost_raw=output_cost,
        cache_cost_raw=cache_cost,
        total_cost_raw=total_cost,
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
        cache_rate=f"{cache_rate:.1f}%",
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

    if usage_df.empty:
        print("❌ 未在账单中找到任何 GLM Coding 用量数据")
        print("   (账单可能只含套餐行,或资源包名称不含 'GLM Coding')")
        sys.exit(1)

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
