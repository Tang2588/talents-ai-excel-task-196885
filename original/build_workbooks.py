from __future__ import annotations

import math
from calendar import monthrange
from datetime import datetime
from pathlib import Path

import xlsxwriter


OUT_DIR = Path(__file__).resolve().parent
TASK_PATH = OUT_DIR / "股票组合风险分析_任务文件.xlsx"
ANSWER_PATH = OUT_DIR / "股票组合风险分析_参考答案.xlsx"

STOCKS = [
    ("S001", "华曜科技", "科技", 0.10),
    ("S002", "云帆软件", "科技", 0.09),
    ("S003", "江海银行", "金融", 0.10),
    ("S004", "安信证券", "金融", 0.08),
    ("S005", "远岭医药", "医药", 0.09),
    ("S006", "康泽生物", "医药", 0.07),
    ("S007", "东辰消费", "消费", 0.09),
    ("S008", "新禾食品", "消费", 0.08),
    ("S009", "瀚源能源", "能源", 0.08),
    ("S010", "启明电力", "能源", 0.07),
    ("S011", "联港工业", "工业", 0.08),
    ("S012", "锐虎机械", "工业", 0.07),
]

SCENARIOS = {
    "温和回撤": {"科技": -0.08, "金融": -0.05, "医药": -0.03, "消费": -0.04, "能源": -0.06, "工业": -0.05},
    "流动性冲击": {"科技": -0.18, "金融": -0.22, "医药": -0.10, "消费": -0.14, "能源": -0.17, "工业": -0.19},
    "增长反弹": {"科技": 0.16, "金融": 0.07, "医药": 0.09, "消费": 0.11, "能源": 0.05, "工业": 0.10},
}


def month_ends(start_year: int, start_month: int, count: int) -> list[datetime]:
    dates = []
    year, month = start_year, start_month
    for _ in range(count):
        dates.append(datetime(year, month, monthrange(year, month)[1]))
        month += 1
        if month == 13:
            month = 1
            year += 1
    return dates


DATES = month_ends(2024, 1, 24)


def build_market_data():
    market = []
    benchmark = 1000.0
    benchmark_prices = []
    stock_prices = {code: [] for code, _, _, _ in STOCKS}
    stock_dividends = {code: [] for code, _, _, _ in STOCKS}

    for month_idx, date in enumerate(DATES):
        market_return = 0.006 + 0.027 * math.sin((month_idx + 1) * 0.77) - 0.011 * math.cos((month_idx + 2) * 0.41)
        if month_idx in (7, 15):
            market_return -= 0.055
        if month_idx in (10, 20):
            market_return += 0.045
        if month_idx > 0:
            benchmark *= 1 + market_return
        benchmark_prices.append(round(benchmark, 2))

        for stock_idx, (code, name, sector, _) in enumerate(STOCKS):
            if month_idx == 0:
                price = 18.0 + stock_idx * 3.15
            else:
                beta = 0.72 + (stock_idx % 6) * 0.12
                alpha = 0.0025 + (stock_idx % 4) * 0.0012
                idio = 0.025 * math.sin((month_idx + 1) * (stock_idx + 2) * 0.37)
                cyclical = 0.012 * math.cos((month_idx + stock_idx + 3) * 0.53)
                monthly_return = alpha + beta * market_return + idio + cyclical
                price = stock_prices[code][-1] * (1 + monthly_return)
            dividend = 0.0
            if month_idx in (5, 17):
                dividend = round((0.08 + stock_idx * 0.012), 3)
            stock_prices[code].append(round(price, 2))
            stock_dividends[code].append(dividend)
            market.append((date, code, name, sector, round(price, 2), dividend, benchmark_prices[-1]))

    return market, benchmark_prices, stock_prices, stock_dividends


MARKET_ROWS, BENCHMARK_PRICES, STOCK_PRICES, STOCK_DIVIDENDS = build_market_data()


def calculate_metrics():
    stock_returns = {}
    for code, _, _, _ in STOCKS:
        prices = STOCK_PRICES[code]
        dividends = STOCK_DIVIDENDS[code]
        stock_returns[code] = [
            (prices[i] + dividends[i]) / prices[i - 1] - 1 for i in range(1, len(prices))
        ]

    benchmark_returns = [
        BENCHMARK_PRICES[i] / BENCHMARK_PRICES[i - 1] - 1 for i in range(1, len(BENCHMARK_PRICES))
    ]
    weights = [stock[3] for stock in STOCKS]
    portfolio_returns = [
        sum(weights[j] * stock_returns[STOCKS[j][0]][i] for j in range(len(STOCKS)))
        for i in range(len(benchmark_returns))
    ]

    def average(values):
        return sum(values) / len(values)

    def stdev(values):
        mean = average(values)
        return math.sqrt(sum((value - mean) ** 2 for value in values) / (len(values) - 1))

    def slope(y_values, x_values):
        x_mean, y_mean = average(x_values), average(y_values)
        numerator = sum((x - x_mean) * (y - y_mean) for x, y in zip(x_values, y_values))
        denominator = sum((x - x_mean) ** 2 for x in x_values)
        return numerator / denominator

    rf = 0.022
    benchmark_annual = (BENCHMARK_PRICES[-1] / BENCHMARK_PRICES[0]) ** (12 / 23) - 1
    metrics = {}
    for code, _, _, weight in STOCKS:
        prices = STOCK_PRICES[code]
        returns = stock_returns[code]
        cumulative = math.prod(1 + value for value in returns) - 1
        annual_return = (1 + cumulative) ** (12 / 23) - 1
        annual_vol = stdev(returns) * math.sqrt(12)
        beta = slope(returns, benchmark_returns)
        capm = rf + beta * (benchmark_annual - rf)
        alpha = annual_return - capm
        negative = [value for value in returns if value < 0]
        downside = math.sqrt(sum(value * value for value in negative) / len(negative)) * math.sqrt(12)
        metrics[code] = {
            "latest": prices[-1],
            "initial": prices[0],
            "cumulative": cumulative,
            "annual_return": annual_return,
            "annual_vol": annual_vol,
            "beta": beta,
            "capm": capm,
            "alpha": alpha,
            "downside": downside,
            "weight": weight,
            "weighted_return": weight * annual_return,
            "weighted_beta": weight * beta,
            "var_contribution": weight * 10_000_000 * 1.645 * annual_vol / math.sqrt(12),
        }

    portfolio_nav = []
    benchmark_nav = []
    running_portfolio = 1.0
    running_benchmark = 1.0
    drawdowns = []
    running_peak = 1.0
    for portfolio_return, benchmark_return in zip(portfolio_returns, benchmark_returns):
        running_portfolio *= 1 + portfolio_return
        running_benchmark *= 1 + benchmark_return
        running_peak = max(running_peak, running_portfolio)
        portfolio_nav.append(running_portfolio)
        benchmark_nav.append(running_benchmark)
        drawdowns.append(running_portfolio / running_peak - 1)

    portfolio_annual = portfolio_nav[-1] ** (12 / 23) - 1
    portfolio_vol = stdev(portfolio_returns) * math.sqrt(12)
    portfolio_beta = sum(item[3] * metrics[item[0]]["beta"] for item in STOCKS)
    portfolio_var = 10_000_000 * 1.645 * stdev(portfolio_returns)
    sharpe = (portfolio_annual - rf) / portfolio_vol
    max_drawdown = min(drawdowns)

    stress = {}
    for scenario, shocks in SCENARIOS.items():
        stress[scenario] = []
        for code, name, sector, weight in STOCKS:
            shock = shocks[sector]
            loss = 10_000_000 * weight * shock
            stress[scenario].append((code, name, sector, weight, shock, 10_000_000 * weight, loss))

    return {
        "stock_returns": stock_returns,
        "benchmark_returns": benchmark_returns,
        "portfolio_returns": portfolio_returns,
        "portfolio_nav": portfolio_nav,
        "benchmark_nav": benchmark_nav,
        "drawdowns": drawdowns,
        "metrics": metrics,
        "benchmark_annual": benchmark_annual,
        "portfolio_annual": portfolio_annual,
        "portfolio_vol": portfolio_vol,
        "portfolio_beta": portfolio_beta,
        "portfolio_var": portfolio_var,
        "sharpe": sharpe,
        "max_drawdown": max_drawdown,
        "stress": stress,
    }


RESULTS = calculate_metrics()


def workbook_formats(workbook):
    return {
        "title": workbook.add_format({"bold": True, "font_size": 16, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"}),
        "section": workbook.add_format({"bold": True, "font_size": 11, "font_color": "#FFFFFF", "bg_color": "#2F75B5", "align": "left", "valign": "vcenter"}),
        "header": workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#4472C4", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True}),
        "subheader": workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True}),
        "label": workbook.add_format({"bold": True, "bg_color": "#E2F0D9", "border": 1}),
        "text": workbook.add_format({"border": 1, "valign": "vcenter"}),
        "text_center": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"}),
        "input": workbook.add_format({"border": 1, "bg_color": "#FFF2CC", "font_color": "#7F6000"}),
        "date": workbook.add_format({"border": 1, "num_format": "yyyy-mm-dd", "align": "center"}),
        "money": workbook.add_format({"border": 1, "num_format": "¥#,##0;[Red]-¥#,##0"}),
        "number": workbook.add_format({"border": 1, "num_format": "0.00"}),
        "percent": workbook.add_format({"border": 1, "num_format": "0.00%;[Red]-0.00%"}),
        "percent1": workbook.add_format({"border": 1, "num_format": "0.0%;[Red]-0.0%"}),
        "formula": workbook.add_format({"border": 1, "bg_color": "#EAF2F8", "num_format": "0.00"}),
        "formula_pct": workbook.add_format({"border": 1, "bg_color": "#EAF2F8", "num_format": "0.00%;[Red]-0.00%"}),
        "formula_money": workbook.add_format({"border": 1, "bg_color": "#EAF2F8", "num_format": "¥#,##0;[Red]-¥#,##0"}),
        "note": workbook.add_format({"font_color": "#666666", "italic": True, "text_wrap": True, "valign": "top"}),
        "kpi_label": workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#5B9BD5", "border": 1, "align": "center", "valign": "vcenter"}),
        "kpi_value": workbook.add_format({"bold": True, "font_size": 14, "bg_color": "#DDEBF7", "border": 1, "align": "center", "valign": "vcenter", "num_format": "0.00%;[Red]-0.00%"}),
    }


def configure_sheet(ws, freeze=(1, 0), zoom=90):
    ws.hide_gridlines(2)
    ws.freeze_panes(*freeze)
    ws.set_zoom(zoom)


def write_raw_data(workbook, formats):
    ws = workbook.add_worksheet("原始行情")
    configure_sheet(ws, freeze=(2, 0), zoom=85)
    ws.merge_range("A1:G1", "模拟股票月度行情（数据已完整提供，无需外部检索）", formats["title"])
    headers = ["日期", "股票代码", "股票名称", "行业", "收盘价", "当月现金分红", "基准指数收盘"]
    ws.write_row("A2", headers, formats["header"])
    for row_idx, row in enumerate(MARKET_ROWS, start=2):
        date, code, name, sector, close, dividend, benchmark = row
        ws.write_datetime(row_idx, 0, date, formats["date"])
        ws.write(row_idx, 1, code, formats["text_center"])
        ws.write(row_idx, 2, name, formats["text"])
        ws.write(row_idx, 3, sector, formats["text_center"])
        ws.write_number(row_idx, 4, close, formats["number"])
        ws.write_number(row_idx, 5, dividend, formats["number"])
        ws.write_number(row_idx, 6, benchmark, formats["number"])
    ws.autofilter(1, 0, len(MARKET_ROWS) + 1, 6)
    ws.set_column("A:A", 12)
    ws.set_column("B:B", 11)
    ws.set_column("C:C", 14)
    ws.set_column("D:D", 10)
    ws.set_column("E:G", 15)


def write_parameters(workbook, formats):
    ws = workbook.add_worksheet("参数设置")
    configure_sheet(ws, freeze=(2, 0))
    ws.merge_range("A1:E1", "股票组合分析参数", formats["title"])
    ws.write_row("A2", ["参数", "数值", "说明"], formats["header"])
    parameters = [
        ("组合规模", 10_000_000, "用于仓位、市值与风险金额计算"),
        ("无风险年利率", 0.022, "CAPM 与夏普比率使用"),
        ("单股权重上限", 0.15, "任何单只股票不得超过该上限"),
        ("行业权重上限", 0.35, "同一行业合计权重不得超过该上限"),
        ("VaR 单尾置信系数", 1.645, "95% 单尾正态近似"),
        ("年化月份数", 12, "月度数据年化系数"),
    ]
    for row_idx, (label, value, note) in enumerate(parameters, start=2):
        ws.write(row_idx, 0, label, formats["label"])
        if row_idx in (3, 4):
            ws.write_number(row_idx, 1, value, formats["percent"])
        elif row_idx == 2:
            ws.write_number(row_idx, 1, value, formats["money"])
        else:
            ws.write_number(row_idx, 1, value, formats["number"])
        ws.write(row_idx, 2, note, formats["text"])

    ws.merge_range("D2:E2", "目标组合权重", formats["section"])
    ws.write_row("D3", ["股票代码", "目标权重"], formats["header"])
    for row_idx, (code, _, _, weight) in enumerate(STOCKS, start=3):
        ws.write(row_idx, 3, code, formats["text_center"])
        ws.write_number(row_idx, 4, weight, formats["percent"])
    ws.write(15, 3, "合计", formats["label"])
    ws.write_formula(15, 4, "=SUM(E4:E15)", formats["formula_pct"], 1.0)

    ws.merge_range("A20:G20", "压力情景：行业价格冲击", formats["section"])
    sectors = ["科技", "金融", "医药", "消费", "能源", "工业"]
    ws.write(20, 0, "行业", formats["header"])
    for col_idx, scenario in enumerate(SCENARIOS, start=1):
        ws.write(20, col_idx, scenario, formats["header"])
    for row_offset, sector in enumerate(sectors, start=21):
        ws.write(row_offset, 0, sector, formats["text_center"])
        for col_idx, scenario in enumerate(SCENARIOS, start=1):
            ws.write_number(row_offset, col_idx, SCENARIOS[scenario][sector], formats["percent"])

    ws.set_column("A:A", 18)
    ws.set_column("B:B", 16)
    ws.set_column("C:C", 38)
    ws.set_column("D:D", 13)
    ws.set_column("E:E", 13)
    ws.set_column("F:G", 13)


def write_analysis_template(workbook, formats, answer: bool):
    ws = workbook.add_worksheet("组合分析")
    configure_sheet(ws, freeze=(5, 0), zoom=80)
    ws.merge_range("A1:Q1", "股票组合收益与风险分析", formats["title"])
    ws.merge_range("A2:Q2", "蓝色单元格为公式结果；所有收益率、风险指标与贡献值均应保持公式联动。", formats["note"])
    headers = ["股票代码", "股票名称", "行业", "最新价", "期初价", "累计收益率", "年化收益率", "年化波动率", "Beta", "CAPM期望收益", "Alpha", "仅基于负月收益的年化下行波动率", "目标权重", "加权收益贡献", "加权Beta贡献", "月度VaR贡献", "配置判断"]
    ws.write_row("A5", headers, formats["header"])

    for stock_idx, (code, name, sector, weight) in enumerate(STOCKS):
        row = 5 + stock_idx
        excel_row = row + 1
        ws.write(row, 0, code, formats["text_center"])
        ws.write(row, 1, name, formats["text"])
        ws.write(row, 2, sector, formats["text_center"])
        if answer:
            metrics = RESULTS["metrics"][code]
            raw_first = 3 + stock_idx
            raw_last = 3 + (len(DATES) - 1) * len(STOCKS) + stock_idx
            return_col = xlsxwriter.utility.xl_col_to_name(1 + stock_idx)
            ws.write_formula(row, 3, f"='原始行情'!E{raw_last}", formats["formula"], metrics["latest"])
            ws.write_formula(row, 4, f"='原始行情'!E{raw_first}", formats["formula"], metrics["initial"])
            ws.write_formula(row, 5, f"=PRODUCT(1+'月度收益'!{return_col}$2:{return_col}$24)-1", formats["formula_pct"], metrics["cumulative"])
            ws.write_formula(row, 6, f"=(1+F{excel_row})^(12/23)-1", formats["formula_pct"], metrics["annual_return"])
            ws.write_formula(row, 7, f"=STDEV.S('月度收益'!{return_col}$2:{return_col}$24)*SQRT(12)", formats["formula_pct"], metrics["annual_vol"])
            ws.write_formula(row, 8, f"=SLOPE('月度收益'!{return_col}$2:{return_col}$24,'月度收益'!$N$2:$N$24)", formats["formula"], metrics["beta"])
            ws.write_formula(row, 9, f"='参数设置'!$B$4+I{excel_row}*('组合分析'!$B$28-'参数设置'!$B$4)", formats["formula_pct"], metrics["capm"])
            ws.write_formula(row, 10, f"=G{excel_row}-J{excel_row}", formats["formula_pct"], metrics["alpha"])
            ws.write_formula(row, 11, f'=SQRT(SUMPRODUCT((\'月度收益\'!{return_col}$2:{return_col}$24<0)*\'月度收益\'!{return_col}$2:{return_col}$24*\'月度收益\'!{return_col}$2:{return_col}$24)/COUNTIF(\'月度收益\'!{return_col}$2:{return_col}$24,"<0"))*SQRT(12)', formats["formula_pct"], metrics["downside"])
            ws.write_formula(row, 12, f"='参数设置'!E{4 + stock_idx}", formats["formula_pct"], weight)
            ws.write_formula(row, 13, f"=M{excel_row}*G{excel_row}", formats["formula_pct"], metrics["weighted_return"])
            ws.write_formula(row, 14, f"=M{excel_row}*I{excel_row}", formats["formula"], metrics["weighted_beta"])
            ws.write_formula(row, 15, f"=M{excel_row}*'参数设置'!$B$3*'参数设置'!$B$7*H{excel_row}/SQRT(12)", formats["formula_money"], metrics["var_contribution"])
            judgment = "增配观察" if metrics["alpha"] > 0.015 and metrics["beta"] < 1.15 else "维持" if metrics["alpha"] > -0.01 else "降低暴露"
            ws.write_formula(row, 16, f'=IF(AND(K{excel_row}>1.5%,I{excel_row}<1.15),"增配观察",IF(K{excel_row}>-1%,"维持","降低暴露"))', formats["text_center"], judgment)
        else:
            for col in range(3, 17):
                ws.write_blank(row, col, None, formats["input"])

    ws.write("A20", "组合汇总", formats["section"])
    summary = [
        ("权重合计", "=SUM(M6:M17)", 1.0, "percent"),
        ("组合年化收益率", "='月度收益'!P24^(12/23)-1", RESULTS["portfolio_annual"], "percent"),
        ("组合年化波动率", "=STDEV.S('月度收益'!O2:O24)*SQRT(12)", RESULTS["portfolio_vol"], "percent"),
        ("组合Beta", "=SUM(O6:O17)", RESULTS["portfolio_beta"], "number"),
        ("95%单月VaR", "='参数设置'!B3*'参数设置'!B7*STDEV.S('月度收益'!O2:O24)", RESULTS["portfolio_var"], "money"),
        ("夏普比率", "=(B22-'参数设置'!B4)/B23", RESULTS["sharpe"], "number"),
        ("最大回撤", "=MIN('月度收益'!R2:R24)", RESULTS["max_drawdown"], "percent"),
        ("基准年化收益率", "=('原始行情'!G279/'原始行情'!G3)^(12/23)-1", RESULTS["benchmark_annual"], "percent"),
    ]
    for idx, (label, formula, value, fmt_key) in enumerate(summary, start=20):
        ws.write(idx, 0, label, formats["label"])
        if answer:
            target_fmt = formats["formula_pct"] if fmt_key == "percent" else formats["formula_money"] if fmt_key == "money" else formats["formula"]
            ws.write_formula(idx, 1, formula, target_fmt, value)
        else:
            ws.write_blank(idx, 1, None, formats["input"])

    ws.conditional_format("K6:K17", {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
    ws.conditional_format("P6:P17", {"type": "data_bar", "bar_color": "#5B9BD5"})
    ws.set_column("A:A", 11)
    ws.set_column("B:B", 14)
    ws.set_column("C:C", 9)
    ws.set_column("D:P", 14)
    ws.set_column("Q:Q", 13)
    ws.set_row(0, 26)


def write_monthly_returns(workbook, formats):
    ws = workbook.add_worksheet("月度收益")
    configure_sheet(ws, freeze=(1, 1), zoom=78)
    headers = ["月份"] + [stock[0] for stock in STOCKS] + ["基准月收益率", "组合月收益率", "组合累计净值", "基准累计净值", "组合回撤"]
    ws.write_row("A1", headers, formats["header"])
    for month_idx in range(1, len(DATES)):
        row = month_idx
        ws.write_datetime(row, 0, DATES[month_idx], formats["date"])
        for stock_idx, (code, _, _, _) in enumerate(STOCKS):
            current_raw = 3 + month_idx * len(STOCKS) + stock_idx
            previous_raw = 3 + (month_idx - 1) * len(STOCKS) + stock_idx
            value = RESULTS["stock_returns"][code][month_idx - 1]
            ws.write_formula(row, 1 + stock_idx, f"=('原始行情'!E{current_raw}+'原始行情'!F{current_raw})/'原始行情'!E{previous_raw}-1", formats["formula_pct"], value)
        current_benchmark_raw = 3 + month_idx * len(STOCKS)
        previous_benchmark_raw = 3 + (month_idx - 1) * len(STOCKS)
        benchmark_return = RESULTS["benchmark_returns"][month_idx - 1]
        portfolio_return = RESULTS["portfolio_returns"][month_idx - 1]
        portfolio_nav = RESULTS["portfolio_nav"][month_idx - 1]
        benchmark_nav = RESULTS["benchmark_nav"][month_idx - 1]
        drawdown = RESULTS["drawdowns"][month_idx - 1]
        ws.write_formula(row, 13, f"='原始行情'!G{current_benchmark_raw}/'原始行情'!G{previous_benchmark_raw}-1", formats["formula_pct"], benchmark_return)
        ws.write_formula(row, 14, f"=SUMPRODUCT(B{row + 1}:M{row + 1},'参数设置'!$E$4:$E$15)", formats["formula_pct"], portfolio_return)
        nav_formula = f"=1+O{row + 1}" if row == 1 else f"=P{row}*(1+O{row + 1})"
        bench_nav_formula = f"=1+N{row + 1}" if row == 1 else f"=Q{row}*(1+N{row + 1})"
        ws.write_formula(row, 15, nav_formula, formats["formula"], portfolio_nav)
        ws.write_formula(row, 16, bench_nav_formula, formats["formula"], benchmark_nav)
        ws.write_formula(row, 17, f"=P{row + 1}/MAX($P$2:P{row + 1})-1", formats["formula_pct"], drawdown)
    ws.set_column("A:A", 12)
    ws.set_column("B:R", 13)
    ws.conditional_format("R2:R24", {"type": "data_bar", "bar_color": "#C00000", "bar_negative_color": "#C00000"})


def write_stress_template(workbook, formats, answer: bool):
    ws = workbook.add_worksheet("压力测试")
    configure_sheet(ws, freeze=(4, 0), zoom=85)
    ws.merge_range("A1:G1", "股票组合情景压力测试", formats["title"])
    ws.merge_range("A2:G2", "按行业冲击映射至每只股票，计算持仓市值和情景损益；负值代表亏损。", formats["note"])
    ws.write_row("A4", ["情景", "股票代码", "行业", "目标权重", "行业冲击", "持仓市值", "情景损益"], formats["header"])
    row = 4
    for scenario_idx, scenario in enumerate(SCENARIOS):
        for stock_idx, (code, _, sector, weight) in enumerate(STOCKS):
            excel_row = row + 1
            ws.write(row, 0, scenario, formats["text_center"])
            ws.write(row, 1, code, formats["text_center"])
            ws.write(row, 2, sector, formats["text_center"])
            if answer:
                record = RESULTS["stress"][scenario][stock_idx]
                ws.write_formula(row, 3, f"='参数设置'!E{4 + stock_idx}", formats["formula_pct"], weight)
                scenario_col = xlsxwriter.utility.xl_col_to_name(1 + scenario_idx)
                sector_param_row = 22 + ["科技", "金融", "医药", "消费", "能源", "工业"].index(sector)
                ws.write_formula(row, 4, f"='参数设置'!{scenario_col}{sector_param_row}", formats["formula_pct"], record[4])
                ws.write_formula(row, 5, f"=D{excel_row}*'参数设置'!$B$3", formats["formula_money"], record[5])
                ws.write_formula(row, 6, f"=E{excel_row}*F{excel_row}", formats["formula_money"], record[6])
            else:
                for col in range(3, 7):
                    ws.write_blank(row, col, None, formats["input"])
            row += 1

    ws.write_row("J3", ["情景", "组合损益", "损益率", "风险判断"], formats["header"])
    for idx, scenario in enumerate(SCENARIOS):
        row_idx = 3 + idx
        start_row = 5 + idx * len(STOCKS)
        end_row = start_row + len(STOCKS) - 1
        total_loss = sum(item[6] for item in RESULTS["stress"][scenario])
        ws.write(row_idx, 9, scenario, formats["text_center"])
        if answer:
            ws.write_formula(row_idx, 10, f"=SUM(G{start_row}:G{end_row})", formats["formula_money"], total_loss)
            ws.write_formula(row_idx, 11, f"=K{row_idx + 1}/'参数设置'!$B$3", formats["formula_pct"], total_loss / 10_000_000)
            label = "高风险" if total_loss / 10_000_000 <= -0.15 else "中风险" if total_loss < 0 else "正向情景"
            ws.write_formula(row_idx, 12, f'=IF(L{row_idx + 1}<=-15%,"高风险",IF(L{row_idx + 1}<0,"中风险","正向情景"))', formats["text_center"], label)
        else:
            ws.write_blank(row_idx, 10, None, formats["input"])
            ws.write_blank(row_idx, 11, None, formats["input"])
            ws.write_blank(row_idx, 12, None, formats["input"])

    ws.conditional_format("G5:G40", {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
    ws.set_column("A:A", 15)
    ws.set_column("B:C", 12)
    ws.set_column("D:E", 13)
    ws.set_column("F:G", 16)
    ws.set_column("H:I", 3)
    ws.set_column("J:J", 15)
    ws.set_column("K:L", 15)
    ws.set_column("M:M", 13)


def write_dashboard(workbook, formats, answer: bool):
    ws = workbook.add_worksheet("投资仪表盘")
    configure_sheet(ws, freeze=(0, 0), zoom=85)
    ws.merge_range("A1:L1", "股票组合投资仪表盘", formats["title"])
    kpis = [
        ("组合年化收益率", "='组合分析'!B22", RESULTS["portfolio_annual"]),
        ("组合年化波动率", "='组合分析'!B23", RESULTS["portfolio_vol"]),
        ("95%单月VaR率", "='组合分析'!B25/'参数设置'!B3", RESULTS["portfolio_var"] / 10_000_000),
        ("最大回撤", "='组合分析'!B27", RESULTS["max_drawdown"]),
    ]
    for idx, (label, formula, value) in enumerate(kpis):
        start_col = idx * 3
        ws.merge_range(2, start_col, 2, start_col + 1, label, formats["kpi_label"])
        if answer:
            ws.merge_range(3, start_col, 4, start_col + 1, "", formats["kpi_value"])
            ws.write_formula(3, start_col, formula, formats["kpi_value"], value)
        else:
            ws.merge_range(3, start_col, 4, start_col + 1, "", formats["input"])

    ws.write("A7", "风险贡献最高的 5 只股票", formats["section"])
    ws.write_row("A8", ["排名", "股票代码", "股票名称", "目标权重", "月度VaR贡献"], formats["header"])
    ranked = sorted(STOCKS, key=lambda item: RESULTS["metrics"][item[0]]["var_contribution"], reverse=True)[:5]
    for idx, stock in enumerate(ranked, start=1):
        row = 7 + idx
        source_row = 6 + [item[0] for item in STOCKS].index(stock[0])
        ws.write_number(row, 0, idx, formats["text_center"])
        if answer:
            ws.write_formula(row, 1, f"='组合分析'!A{source_row}", formats["text_center"], stock[0])
            ws.write_formula(row, 2, f"='组合分析'!B{source_row}", formats["text"], stock[1])
            ws.write_formula(row, 3, f"='组合分析'!M{source_row}", formats["formula_pct"], stock[3])
            ws.write_formula(row, 4, f"='组合分析'!P{source_row}", formats["formula_money"], RESULTS["metrics"][stock[0]]["var_contribution"])
        else:
            for col in range(1, 5):
                ws.write_blank(row, col, None, formats["input"])

    if answer:
        line_chart = workbook.add_chart({"type": "line"})
        line_chart.add_series({"name": "组合累计净值", "categories": "='月度收益'!$A$2:$A$24", "values": "='月度收益'!$P$2:$P$24", "line": {"color": "#4472C4", "width": 2.25}})
        line_chart.add_series({"name": "基准累计净值", "categories": "='月度收益'!$A$2:$A$24", "values": "='月度收益'!$Q$2:$Q$24", "line": {"color": "#70AD47", "width": 2.0}})
        line_chart.set_title({"name": "组合与基准累计净值"})
        line_chart.set_x_axis({"name": "月份", "date_axis": True, "num_format": "yyyy-mm"})
        line_chart.set_y_axis({"name": "累计净值", "major_gridlines": {"visible": False}, "num_format": "0.00"})
        line_chart.set_legend({"position": "bottom"})
        line_chart.set_style(10)
        ws.insert_chart("G7", line_chart, {"x_scale": 1.15, "y_scale": 1.1})

        column_chart = workbook.add_chart({"type": "column"})
        column_chart.add_series({"name": "情景组合损益", "categories": "='压力测试'!$J$4:$J$6", "values": "='压力测试'!$K$4:$K$6", "fill": {"color": "#ED7D31"}, "border": {"none": True}, "data_labels": {"value": True, "num_format": "¥0.0,,\"百万\""}})
        column_chart.set_title({"name": "压力情景组合损益"})
        column_chart.set_x_axis({"name": "情景"})
        column_chart.set_y_axis({"name": "损益金额（元）", "major_gridlines": {"visible": False}, "num_format": "¥0.0,,\"百万\""})
        column_chart.set_legend({"none": True})
        column_chart.set_style(10)
        ws.insert_chart("G20", column_chart, {"x_scale": 1.15, "y_scale": 1.05})

    ws.set_column("A:A", 9)
    ws.set_column("B:B", 12)
    ws.set_column("C:C", 14)
    ws.set_column("D:E", 15)
    ws.set_column("F:F", 3)
    ws.set_column("G:L", 13)
    ws.set_row(0, 28)


def build_workbook(path: Path, answer: bool):
    workbook = xlsxwriter.Workbook(path)
    workbook.set_properties({
        "title": "股票组合风险分析",
        "subject": "TalentsAI Excel 任务",
        "author": "TalentsAI Expert",
        "comments": "全部数据均为模拟数据，仅用于 Excel 分析任务。",
    })
    formats = workbook_formats(workbook)
    write_raw_data(workbook, formats)
    write_parameters(workbook, formats)
    write_analysis_template(workbook, formats, answer=answer)
    if answer:
        write_monthly_returns(workbook, formats)
    write_stress_template(workbook, formats, answer=answer)
    write_dashboard(workbook, formats, answer=answer)
    workbook.close()


def main():
    build_workbook(TASK_PATH, answer=False)
    build_workbook(ANSWER_PATH, answer=True)
    print(TASK_PATH)
    print(ANSWER_PATH)


if __name__ == "__main__":
    main()
