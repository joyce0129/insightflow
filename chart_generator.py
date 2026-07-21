import os
import re
import matplotlib.pyplot as plt
from wordcloud import WordCloud


def _clean_title(text, max_len=16):
    """移除 emoji 與換行，截斷過長文字（供折線圖 callout 標籤使用）"""
    cleaned = ''.join(c for c in str(text) if ord(c) <= 0xFFFF)
    cleaned = re.sub(r'[\x00-\x1f]', ' ', cleaned)
    cleaned = ' '.join(cleaned.split())
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len] + '…'
    return cleaned


def build_trend_callouts(trend, articles, limit=4):
    """從熱門文章挑出互動數最高的幾篇，依日期排序後轉為折線圖 callout 事件"""
    trend_dates = {item["date"] for item in trend["daily"]}
    pool = articles[:10] if articles else []
    top = sorted(pool, key=lambda a: a.get("engagement", 0), reverse=True)[:limit]
    top.sort(key=lambda a: a.get("time", ""))

    events = []
    for a in top:
        date_part = str(a.get("time", "")).split(" ")[0]
        parts = date_part.split("-")
        trend_date = "/".join(parts) if len(parts) == 3 else date_part
        if trend_date not in trend_dates:
            continue
        md = f"{parts[1]}/{parts[2]}" if len(parts) == 3 else date_part
        events.append({
            "date": trend_date,
            "text": f"{md}｜{a.get('source', '')}\n{_clean_title(a.get('title', ''), 16)}"
        })
    return events


def generate_volume_trend_chart(trend, brand_name="品牌", out_dir="output/charts", callouts=None):
    import os
    import matplotlib.pyplot as plt

    os.makedirs(out_dir, exist_ok=True)

    # 中文字體設定
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    dates = [item["date"] for item in trend["daily"]]
    volumes = [item["volume"] for item in trend["daily"]]

    # 找出聲量高峰
    peak_index = volumes.index(max(volumes))
    peak_date = dates[peak_index]
    peak_volume = volumes[peak_index]

    fig, ax = plt.subplots(figsize=(11, 5.5))

    ax.plot(
        dates,
        volumes,
        marker="o",
        linewidth=2,
        color="#4C7A9E",
        zorder=3
    )

    ax.set_title(f"{brand_name} 聲量趨勢")
    ax.set_xlabel("日期")
    ax.set_ylabel("聲量")
    ax.tick_params(axis="x", rotation=45)
    ax.grid(True, linestyle="--", alpha=0.5)

    # 標註每個點的數字
    for x, y in zip(dates, volumes):
        ax.text(
            x,
            y,
            str(y),
            ha="center",
            va="bottom",
            fontsize=8
        )

    callout_dates = {c["date"] for c in callouts} if callouts else set()

    # 事件 callout（白底黑框 + 折線指標，仿範本樣式）
    if callouts:
        ax.set_ylim(top=max(volumes) * 1.4)
        stagger = [95, 145, 95, 145, 120]
        for i, ev in enumerate(callouts):
            if ev["date"] not in dates:
                continue
            idx = dates.index(ev["date"])
            y = volumes[idx]
            jitter = 3 if i % 2 == 0 else -3
            ax.annotate(
                ev["text"],
                xy=(ev["date"], y),
                xytext=(jitter, stagger[i % len(stagger)]),
                textcoords="offset points",
                ha="center",
                fontsize=8,
                zorder=4,
                bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="#333333", lw=1),
                arrowprops=dict(
                    arrowstyle="-", color="#333333",
                    connectionstyle="angle,angleA=0,angleB=90,rad=5"
                )
            )
    else:
        ax.set_ylim(top=max(volumes) * 1.15)

    # 標註高峰（僅在無 callout 時顯示；有 callout 時峰值已由逐點數字標示，避免重複視覺雜訊）
    if not callouts and peak_date not in callout_dates:
        ax.annotate(
            f"高峰：{peak_volume} 筆",
            xy=(peak_date, peak_volume),
            xytext=(0, 30),
            textcoords="offset points",
            arrowprops=dict(arrowstyle="->"),
            fontsize=10,
            ha="center"
        )

    fig.tight_layout()

    fig.savefig(
        f"{out_dir}/volume_trend.png",
        dpi=150,
        bbox_inches="tight"
    )

    plt.close(fig)


def generate_sentiment_pie_chart(sentiment, brand_name="品牌", out_dir="output/charts"):
    import os

    os.makedirs(out_dir, exist_ok=True)

    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    labels = ["正面", "中立", "負面"]

    values = [
        sentiment["positive"],
        sentiment["neutral"],
        sentiment["negative"]
    ]

    colors = ["#3B7DD8", "#7A7D82", "#C0392B"]

    fig, ax = plt.subplots(figsize=(6, 6))

    ax.pie(
        values,
        labels=labels,
        autopct="%1.1f%%",
        colors=colors,
        wedgeprops=dict(width=0.4, edgecolor="white")
    )

    pn_value = sentiment.get("pn_value")
    ax.text(
        0, 0.08,
        "P/N值",
        ha="center", va="center",
        fontsize=13, color="#5B6470"
    )
    ax.text(
        0, -0.1,
        str(pn_value) if pn_value is not None else "—",
        ha="center", va="center",
        fontsize=26, fontweight="bold", color="#1A233B"
    )

    ax.set_title(f"{brand_name} 情緒分布")

    fig.savefig(f"{out_dir}/sentiment_pie.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

def generate_keyword_bar_chart(keywords, out_dir="output/charts"):

    import os
    import matplotlib.pyplot as plt

    os.makedirs(out_dir, exist_ok=True)

    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    labels = [
        item["keyword"]
        for item in keywords[:10]
    ]

    values = [
        item["count"]
        for item in keywords[:10]
    ]

    plt.figure(figsize=(10, 6))

    plt.barh(
        labels[::-1],
        values[::-1]
    )

    plt.title("熱門關鍵字 TOP10")
    plt.xlabel("出現次數")

    plt.tight_layout()

    plt.savefig(
        f"{out_dir}/keyword_bar.png",
        dpi=150,
        bbox_inches="tight"
    )

    plt.close()

def generate_keyword_cloud(keywords, out_dir="output/charts"):

    import os

    os.makedirs(out_dir, exist_ok=True)

    freq = {}

    for item in keywords[:30]:
        freq[item["keyword"]] = item["count"]

    wc = WordCloud(
        width=1200,
        height=600,
        background_color="white",

        # Windows 中文字型
        font_path=r"C:\Windows\Fonts\msjh.ttc"
    )

    wc.generate_from_frequencies(freq)

    wc.to_file(f"{out_dir}/keyword_cloud.png")


COMPETITOR_COLORS = ["#E67E22", "#27AE60", "#8E44AD"]


def generate_competitor_volume_chart(brand_name, trend, competitor_data, out_dir="output/charts"):
    """主品牌 vs 競品聲量折線疊圖"""
    os.makedirs(out_dir, exist_ok=True)

    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(11, 5.5))

    main_dates = [item["date"] for item in trend["daily"]]
    main_volumes = [item["volume"] for item in trend["daily"]]
    ax.plot(main_dates, main_volumes, marker="o", linewidth=2.5,
            color="#CC3232", label=brand_name, zorder=3)

    for i, comp in enumerate(competitor_data):
        c_dates = [item["date"] for item in comp["trend"]["daily"]]
        c_volumes = [item["volume"] for item in comp["trend"]["daily"]]
        color = COMPETITOR_COLORS[i % len(COMPETITOR_COLORS)]
        ax.plot(c_dates, c_volumes, marker="o", markersize=3, linewidth=1.5,
                color=color, label=comp["name"], alpha=0.85)

    ax.set_title(f"{brand_name} vs 競品 聲量趨勢比較")
    ax.set_xlabel("日期")
    ax.set_ylabel("聲量")
    ax.tick_params(axis="x", rotation=45)
    ax.grid(True, linestyle="--", alpha=0.5)
    ax.legend(loc="upper left")

    fig.tight_layout()
    fig.savefig(f"{out_dir}/competitor_volume.png", dpi=150, bbox_inches="tight")
    plt.close(fig)


def generate_competitor_sentiment_chart(brand_name, sentiment, competitor_data, out_dir="output/charts"):
    """主品牌 vs 競品 情緒占比 100% 堆疊橫條圖"""
    os.makedirs(out_dir, exist_ok=True)

    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    brands = [brand_name] + [c["name"] for c in competitor_data]
    sentiments = [sentiment] + [c["sentiment"] for c in competitor_data]

    pos = []
    neu = []
    neg = []
    for s in sentiments:
        total = s["positive"] + s["neutral"] + s["negative"]
        if total == 0:
            pos.append(0)
            neu.append(0)
            neg.append(0)
        else:
            pos.append(s["positive"] / total * 100)
            neu.append(s["neutral"] / total * 100)
            neg.append(s["negative"] / total * 100)

    fig, ax = plt.subplots(figsize=(11, 0.7 * len(brands) + 1.5))

    y = range(len(brands))
    ax.barh(y, pos, color="#3B7DD8", label="正面")
    ax.barh(y, neu, left=pos, color="#7A7D82", label="中立")
    left_neg = [p + n for p, n in zip(pos, neu)]
    ax.barh(y, neg, left=left_neg, color="#C0392B", label="負面")

    for i in range(len(brands)):
        ax.text(pos[i] / 2, i, f"{pos[i]:.1f}%", ha="center", va="center",
                fontsize=9, color="white")
        ax.text(left_neg[i] + neg[i] / 2, i, f"{neg[i]:.1f}%", ha="center",
                va="center", fontsize=9, color="white")

    ax.set_yticks(list(y))
    ax.set_yticklabels(brands)
    ax.invert_yaxis()
    ax.set_xlim(0, 100)
    ax.set_xlabel("占比（%）")
    ax.set_title(f"{brand_name} vs 競品 情緒占比比較")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=3)

    fig.tight_layout()
    fig.savefig(f"{out_dir}/competitor_sentiment.png", dpi=150, bbox_inches="tight")
    plt.close(fig)