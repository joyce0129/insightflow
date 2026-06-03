import os
import matplotlib.pyplot as plt
from wordcloud import WordCloud

def generate_volume_trend_chart(trend):
    import os
    import matplotlib.pyplot as plt

    os.makedirs("output/charts", exist_ok=True)

    # 中文字體設定
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    dates = [item["date"] for item in trend["daily"]]
    volumes = [item["volume"] for item in trend["daily"]]

    # 找出聲量高峰
    peak_index = volumes.index(max(volumes))
    peak_date = dates[peak_index]
    peak_volume = volumes[peak_index]

    plt.figure(figsize=(10, 5))

    plt.plot(
        dates,
        volumes,
        marker="o",
        linewidth=2
    )

    plt.title("iRent 聲量趨勢")
    plt.xlabel("日期")
    plt.ylabel("聲量")
    plt.xticks(rotation=45)
    plt.grid(True, linestyle="--", alpha=0.5)

    # 標註每個點的數字
    for x, y in zip(dates, volumes):
        plt.text(
            x,
            y,
            str(y),
            ha="center",
            va="bottom",
            fontsize=9
        )

    # 標註高峰
    plt.annotate(
        f"高峰：{peak_volume} 筆",
        xy=(peak_date, peak_volume),
        xytext=(peak_index - 1, peak_volume + 200),
        arrowprops=dict(arrowstyle="->"),
        fontsize=10
    )

    plt.tight_layout()

    plt.savefig(
        "output/charts/volume_trend.png",
        dpi=150,
        bbox_inches="tight"
    )

    plt.close()


def generate_sentiment_pie_chart(sentiment):

    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    labels = ["正面", "中立", "負面"]

    values = [
        sentiment["positive"],
        sentiment["neutral"],
        sentiment["negative"]
    ]

    plt.figure(figsize=(6, 6))

    plt.pie(
        values,
        labels=labels,
        autopct="%1.1f%%"
    )

    plt.title("iRent 情緒分布")

    plt.savefig("output/charts/sentiment_pie.png")
    plt.close()

def generate_keyword_bar_chart(keywords):

    import os
    import matplotlib.pyplot as plt

    os.makedirs("output/charts", exist_ok=True)

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
        "output/charts/keyword_bar.png",
        dpi=150,
        bbox_inches="tight"
    )

    plt.close()

def generate_keyword_cloud(keywords):

    import os

    os.makedirs(
        "output/charts",
        exist_ok=True
    )

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

    wc.to_file(
        "output/charts/keyword_cloud.png"
    )