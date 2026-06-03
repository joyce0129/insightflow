import os
from datetime import datetime, timedelta

from local_ai_analyzer import analyze_with_ollama
from ppt_generator import generate_ppt
from keypo_fetcher import fetch_keypo_data

from keypo_loader import (
    load_json,
    parse_sentiment,
    parse_hot_keywords,
    parse_hotrank,
    parse_volume_trend,
    parse_kol_channels
)

from report_generator import generate_summary

from chart_generator import (
    generate_volume_trend_chart,
    generate_sentiment_pie_chart,
    generate_keyword_bar_chart,
    generate_keyword_cloud
)


def main():

    print("=" * 50)
    print("InsightFlow AI 輿情分析系統")
    print("=" * 50)

    keyword = input(
        "\n請輸入品牌名稱："
    ).strip()

    if not keyword:
        print("❌ 品牌名稱不可空白")
        return

    # 最近 7 天
    today = datetime.today()

    end_date = today.strftime("%Y-%m-%d")

    start_date = (
        today - timedelta(days=6)
    ).strftime("%Y-%m-%d")

    print(f"\n開始分析：{keyword}")
    print(f"分析期間：{start_date} ~ {end_date}")

    # =========================
    # 抓取 KEYPO 資料
    # =========================

    fetch_keypo_data(
        keyword,
        start_date,
        end_date
    )

    # =========================
    # 載入資料
    # =========================

    sentiment = parse_sentiment(
        load_json("data/sentidist.json")
    )

    keywords = parse_hot_keywords(
        load_json("data/hotkw.json")
    )

    articles = parse_hotrank(
        load_json("data/hotrank.json")
    )

    trend = parse_volume_trend(
        load_json("data/freqdist.json")
    )

    kols = parse_kol_channels(
    load_json("data/sprdtrnd.json")
    )

    # =========================
    # 測試 KOL 資料
    # =========================

    print("\n")
    print("=" * 60)
    print("網路關鍵領袖 TOP10")
    print("=" * 60)

    for kol in kols[:10]:
        print(
            f"{kol['rank']}. "
            f"[{kol['platform']}] "
            f"{kol['channel']} "
            f"({kol['volume']})"
        )


    # =========================
    # 產生 PPT 日期
    # =========================

    dates = [item["date"] for item in trend["daily"]]

    report_start_date = dates[0].replace("/", "")
    report_end_date = dates[-1].replace("/", "")

    # =========================
    # 產生摘要
    # =========================

    report = generate_summary(
        sentiment,
        keywords,
        articles,
        trend
    )

    # =========================
    # 圖表
    # =========================

    generate_volume_trend_chart(trend)
    generate_sentiment_pie_chart(sentiment)
    generate_keyword_bar_chart(keywords)
    generate_keyword_cloud(keywords)
    print("\n")
    print(report)

    # =========================
    # AI 洞察
    # =========================

    print("\n")
    print("=" * 60)
    print("AI 洞察")
    print("=" * 60)

    ai_report = analyze_with_ollama(report)

    print(ai_report)

    # =========================
    # 輸出資料夾
    # =========================

    os.makedirs("output", exist_ok=True)

    with open(
        "output/weekly_report.md",
        "w",
        encoding="utf-8"
    ) as f:
        f.write(report)

    with open(
        "output/ai_insight.md",
        "w",
        encoding="utf-8"
    ) as f:
        f.write(ai_report)

    # =========================
    # 產生 PPT
    # =========================

    ppt_path = generate_ppt(
    keyword,
    report,
    ai_report,
    keywords,
    kols,
    report_start_date,
    report_end_date
    )

    print("\n✅ 已輸出報告、圖表與 PPT")
    print(f"✅ PPT 路徑：{ppt_path}")


if __name__ == "__main__":
    main()