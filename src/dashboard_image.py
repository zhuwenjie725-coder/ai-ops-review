"""将本地分析结果导出为适合作品集截图的静态 PNG。"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager, ticker
from matplotlib.patches import FancyBboxPatch


def _font():
    font_path = Path("C:/Windows/Fonts/msyh.ttc")
    if font_path.is_file():
        font = font_manager.FontProperties(fname=str(font_path))
        plt.rcParams["font.family"] = font.get_name()
        return font
    return None


def _compact(value, money=False):
    if pd.isna(value):
        return "—"
    prefix = "¥" if money else ""
    if abs(value) >= 10000:
        return "{}{:,.1f}万".format(prefix, value / 10000)
    return "{}{:,.0f}".format(prefix, value)


def save_dashboard_image(frame, values, labels, output_path):
    """生成 2560×1440 图片；数值与 Excel 看板使用同一份计算结果。"""
    font = _font()
    plt.rcParams["axes.unicode_minus"] = False
    figure = plt.figure(figsize=(16, 9), dpi=160, facecolor="#F3F6FA")
    title_style = {"fontproperties": font} if font else {}
    figure.text(0.045, 0.955, "达人投放数据自动化复盘｜V1.5", fontsize=24, weight="bold", color="#20364B", va="top", **title_style)
    figure.text(0.955, 0.958, "模拟测试数据，仅用于功能演示", fontsize=12, color="#9A6700", ha="right", va="top", **title_style)
    figure.text(0.045, 0.915, "本项目数据不代表任何公司真实经营数据", fontsize=10, color="#617588", va="top", **title_style)

    cards = [
        ("总消耗", _compact(values["总投放消耗"], money=True)),
        ("总播放", _compact(values["总播放量"])),
        ("总线索", _compact(values["总线索量"])),
        ("总GMV", _compact(values["总GMV"], money=True)),
        ("整体ROI", "{:.2f}".format(values["整体ROI"]) if pd.notna(values["整体ROI"]) else "—"),
        ("整体CPL", "¥{:,.2f}".format(values["整体CPL"]) if pd.notna(values["整体CPL"]) else "—"),
    ]
    for index, (label, value) in enumerate(cards):
        x = 0.045 + index * 0.153
        patch = FancyBboxPatch((x, 0.755), 0.141, 0.13, boxstyle="round,pad=0.006,rounding_size=0.012", transform=figure.transFigure, linewidth=0, facecolor="white")
        figure.patches.append(patch)
        figure.text(x + 0.012, 0.855, label, fontsize=11, color="#617588", va="top", **title_style)
        figure.text(x + 0.012, 0.815, value, fontsize=21, weight="bold", color="#20364B", va="top", **title_style)

    layout = [(0.045, 0.15, 0.26, 0.54), (0.348, 0.15, 0.27, 0.54), (0.661, 0.15, 0.295, 0.54)]
    axes = []
    for index, (x, y, width, height) in enumerate(layout):
        figure.patches.append(FancyBboxPatch((x, y), width, height, boxstyle="round,pad=0.005,rounding_size=0.012", transform=figure.transFigure, linewidth=0, facecolor="white", zorder=-1))
        left = 0.072 if index == 1 else (0.06 if index == 2 else 0.035)
        ax = figure.add_axes((x + left, y + 0.075, width - left - 0.023, height - 0.135), facecolor="white")
        ax.set_zorder(2)
        axes.append(ax)
    group_ax, roi_ax, scatter_ax = axes

    groups = ["建议复投", "继续观察", "建议淘汰"]
    counts = [values[name + "数量"] for name in groups]
    bars = group_ax.bar(groups, counts, color=["#438F83", "#E7A64D", "#8D9BAA"], width=0.58)
    group_ax.set_title("达人分层数量", fontsize=14, weight="bold", color="#20364B", pad=18, **title_style)
    group_ax.set_ylim(0, max(counts + [1]) * 1.25)
    group_ax.yaxis.set_major_locator(ticker.MaxNLocator(nbins=5, integer=True))
    group_ax.grid(axis="y", color="#E5EBF0", linewidth=0.8)
    group_ax.set_axisbelow(True)
    for bar, count in zip(bars, counts):
        group_ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(counts + [1]) * 0.03, str(count), ha="center", va="bottom", fontsize=11, color="#20364B")

    complete = frame[["粉丝数", "播放量", "点赞数", "评论数", "转发数", "投放消耗", "线索量", "成交量", "GMV"]].notna().all(axis=1)
    complete &= ~frame["数据异常说明"].str.contains("达人名称缺失", regex=False)
    top = frame.loc[complete & frame["ROI"].notna(), ["达人名称", "ROI"]].nlargest(10, "ROI")
    roi_ax.barh(list(top["达人名称"])[::-1], list(top["ROI"])[::-1], color="#5B86A6", height=0.68)
    roi_ax.set_title("ROI Top10 达人", fontsize=14, weight="bold", color="#20364B", pad=18, **title_style)
    roi_ax.set_xlim(left=0)
    roi_ax.xaxis.set_major_locator(ticker.MaxNLocator(nbins=5))
    roi_ax.xaxis.set_major_formatter(ticker.FormatStrFormatter("%.1f"))
    roi_ax.grid(axis="x", color="#E5EBF0", linewidth=0.8)
    roi_ax.set_axisbelow(True)
    roi_ax.set_xlabel("ROI", fontsize=10, color="#617588", **title_style)

    scatter_ax.set_title("投放消耗 vs GMV", fontsize=14, weight="bold", color="#20364B", pad=18, **title_style)
    colors = {"建议复投": "#438F83", "继续观察": "#E7A64D", "建议淘汰": "#8D9BAA"}
    for group in groups:
        selected = [index for index, item in enumerate(labels) if item[0] == group and pd.notna(frame.iloc[index]["投放消耗"]) and pd.notna(frame.iloc[index]["GMV"])]
        if selected:
            scatter_ax.scatter(frame.iloc[selected]["投放消耗"], frame.iloc[selected]["GMV"], s=22, alpha=0.72, color=colors[group], label=group)
    scatter_ax.set_xlim(left=0)
    scatter_ax.set_ylim(bottom=0)
    scatter_ax.xaxis.set_major_locator(ticker.MaxNLocator(nbins=4))
    scatter_ax.yaxis.set_major_locator(ticker.MaxNLocator(nbins=5))
    scatter_ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda value, _: "{:,.0f}".format(value / 1000)))
    scatter_ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda value, _: "{:,.0f}".format(value / 1000)))
    scatter_ax.set_xlabel("投放消耗（千元）", fontsize=10, color="#617588", **title_style)
    scatter_ax.set_ylabel("GMV（千元）", fontsize=10, color="#617588", **title_style)
    scatter_ax.grid(color="#E5EBF0", linewidth=0.8)
    scatter_ax.set_axisbelow(True)
    scatter_ax.legend(loc="upper left", fontsize=8, frameon=False)

    for ax in axes:
        ax.tick_params(axis="both", labelsize=9, colors="#617588", length=0)
        for spine in ax.spines.values():
            spine.set_visible(False)
        if font:
            for tick in ax.get_xticklabels() + ax.get_yticklabels():
                tick.set_fontproperties(font)
    figure.text(0.045, 0.075, "整体指标：关键字段齐全的记录 {} / {}；缺失和异常不作0。分层阈值仅用于个人项目演示。".format(values["有效汇总记录数"], len(frame)), fontsize=10, color="#617588", **title_style)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(str(output_path), dpi=160, facecolor=figure.get_facecolor())
    plt.close(figure)
