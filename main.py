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
from detailed_report_generator import generate_detailed_report

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
        print("品牌名稱不可空白")
        return

    # 查詢字串（可自訂複雜 KEYPO 布林語法，預設同品牌名稱）
    print("\n查詢關鍵字（直接 Enter 使用品牌名稱，可輸入複雜布林語法）")
    print(f"  預設：{keyword}")
    print(f"  範例：(((CREED|克雷德)&(香水|淡香精|香味|氣息)))")
    query_input = input("  查詢字串：").strip()
    if query_input:
        # 開頭是 & 時自動補上品牌名稱
        if query_input.startswith("&"):
            search_query = f"{keyword}{query_input}"
        else:
            search_query = query_input
    else:
        search_query = keyword

    # 觀測期間輸入
    today = datetime.today()
    default_start = (today - timedelta(days=6)).strftime("%Y-%m-%d")
    default_end   = today.strftime("%Y-%m-%d")

    print(f"\n觀測期間（格式：YYYY-MM-DD，直接 Enter 使用預設近 7 天）")

    start_input = input(
        f"  開始日期 [{default_start}]："
    ).strip()

    end_input = input(
        f"  結束日期 [{default_end}]："
    ).strip()

    start_date = start_input if start_input else default_start
    end_date   = end_input   if end_input   else default_end

    # 格式驗證
    try:
        datetime.strptime(start_date, "%Y-%m-%d")
        datetime.strptime(end_date,   "%Y-%m-%d")
    except ValueError:
        print("日期格式錯誤，請使用 YYYY-MM-DD")
        return

    if start_date > end_date:
        print("開始日期不可晚於結束日期")
        return

    # 是否排除抽獎
    lottery_input = input(
        "\n是否排除抽獎類貼文？[Y/n]："
    ).strip().lower()
    exclude_lottery = lottery_input != "n"

    # 自訂雜訊排除
    print("\n自訂雜訊排除（直接 Enter 跳過）")
    print("  範例：!(代購|直播|販售文|二手衣出清)")
    custom_exclusion = input("  輸入排除字串：").strip()

    print(f"\n開始分析：{keyword}")
    print(f"查詢字串：{search_query}")
    print(f"分析期間：{start_date} ~ {end_date}")
    print(f"排除抽獎：{'是' if exclude_lottery else '否'}")
    if custom_exclusion:
        print(f"自訂排除：{custom_exclusion}")

    # =========================
    # 抓取 KEYPO 資料
    # =========================

    fetch_keypo_data(
        search_query,
        start_date,
        end_date,
        exclude_lottery=exclude_lottery,
        custom_exclusion=custom_exclusion
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

    out_dir = f"output/{keyword}"
    charts_dir = f"{out_dir}/charts"

    generate_volume_trend_chart(trend, keyword, out_dir=charts_dir)
    generate_sentiment_pie_chart(sentiment, keyword, out_dir=charts_dir)
    generate_keyword_bar_chart(keywords, out_dir=charts_dir)
    generate_keyword_cloud(keywords, out_dir=charts_dir)
    print("\n")
    print(report)

    # =========================
    # AI 洞察
    # =========================

    print("\n")
    print("=" * 60)
    print("AI 洞察")
    print("=" * 60)

    ai_report = analyze_with_ollama(
        report,
        keywords=keywords,
        articles=articles,
        kols=kols
    )

    print(ai_report)

    # =========================
    # 輸出資料夾
    # =========================

    os.makedirs(out_dir, exist_ok=True)

    with open(
        f"{out_dir}/{keyword}_{report_start_date}_{report_end_date}_summary.md",
        "w",
        encoding="utf-8"
    ) as f:
        f.write(report)

    with open(
        f"{out_dir}/{keyword}_{report_start_date}_{report_end_date}_ai_insight.md",
        "w",
        encoding="utf-8"
    ) as f:
        f.write(ai_report)

    # =========================
    # 產生詳細分析報告
    # =========================

    detailed_report = generate_detailed_report(
        keyword,
        sentiment,
        keywords,
        articles,
        trend,
        kols,
        ai_report,
        report_start_date,
        report_end_date
    )

    detailed_path = (
        f"{out_dir}/{keyword}_{report_start_date}"
        f"_{report_end_date}_詳細分析報告.md"
    )

    with open(detailed_path, "w", encoding="utf-8") as f:
        f.write(detailed_report)

    # =========================
    # 產生 PPT
    # =========================

    ppt_path = generate_ppt(
        keyword,
        report,
        ai_report,
        keywords,
        kols,
        articles,
        sentiment,
        trend,
        report_start_date,
        report_end_date,
        out_dir=out_dir
    )

    print("\n✅ 已輸出報告、圖表與 PPT")
    print(f"✅ PPT 路徑：{ppt_path}")


if __name__ == "__main__":
    main()