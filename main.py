import os
from datetime import datetime, timedelta

from local_ai_analyzer import (
    analyze_with_ai,
    narrative_for_articles,
    narrative_for_keywords,
    narrative_for_kol,
    narrative_for_competitor,
    narrative_for_trend
)
import ai_backend
from ppt_generator import generate_ppt
from keypo_fetcher import fetch_keypo_data, fetch_volume_only, fetch_endpoints

from keypo_loader import (
    load_json,
    parse_sentiment,
    parse_hot_keywords,
    parse_hotrank,
    parse_volume_trend,
    parse_kol_channels,
    parse_kol_from_articles,
    parse_source_distribution
)

from report_generator import generate_summary

from chart_generator import (
    generate_volume_trend_chart,
    generate_sentiment_pie_chart,
    generate_source_pie_chart,
    generate_keyword_bar_chart,
    generate_keyword_cloud,
    build_trend_callouts,
    generate_competitor_volume_chart,
    generate_competitor_sentiment_chart
)

from period_history import (
    load_previous_period,
    load_same_period_last_year,
    same_period_last_year_dates,
    save_current_period,
    compute_growth_rate
)


def _safe_caption(prefix, narrative_fn, *args):
    """組出「模板數字段 + AI 敘事段」的表格摘要句；AI 呼叫失敗時優雅退化成純模板句"""
    try:
        narrative = narrative_fn(*args)
    except Exception as e:
        print(f"敘事摘要產生失敗，僅顯示基本數據：{e}")
        narrative = ""
    return f"{prefix}{narrative}"


def main():

    print("=" * 50)
    print("InsightFlow AI 輿情分析系統")
    print(f"AI 後端：{ai_backend.backend_name()}"
          "（可用環境變數 AI_BACKEND 切換 ollama / claude / openai）")
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

    # 競品分析（選填）
    print("\n競品分析（選填，直接 Enter 跳過，最多 3 個，逗號分隔）")
    print("  範例：新光三越,微風廣場,SOGO百貨")
    competitor_input = input("  競品品牌：").strip()
    competitors = (
        [c.strip() for c in competitor_input.split(",") if c.strip()][:3]
        if competitor_input else []
    )

    # 各競品自訂雜訊排除（選填；每個品牌雜訊不同，各自輸入）
    competitor_exclusions = {}
    if competitors:
        print("\n各競品自訂雜訊排除（直接 Enter 跳過該品牌）")
        print("  範例：!(代購|直播|販售文|二手衣出清)")
        for comp in competitors:
            excl = input(f"  {comp} 排除字串：").strip()
            if excl:
                competitor_exclusions[comp] = excl

    # 產業／背景（供「分析架構與腳本」AI 推斷議題；直接 Enter 由 AI 依品牌推斷）
    print("\n產業／背景（選填，直接 Enter 跳過；有助分析架構腳本的議題更精準）")
    print("  範例：美妝保養／醫美級保養品；或 台北地標（觀景台、購物中心、信義商圈）")
    industry_hint = input("  產業／背景：").strip()

    print(f"\n開始分析：{keyword}")
    print(f"查詢字串：{search_query}")
    print(f"分析期間：{start_date} ~ {end_date}")
    print(f"排除抽獎：{'是' if exclude_lottery else '否'}")
    if custom_exclusion:
        print(f"自訂排除：{custom_exclusion}")
    if competitors:
        print(f"競品分析：{'、'.join(competitors)}")
        for comp in competitors:
            if comp in competitor_exclusions:
                print(f"  {comp} 自訂排除：{competitor_exclusions[comp]}")

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

    # =========================
    # 上期比較 ／ 去年同期比較（簡化版：僅比較總聲量）
    # =========================

    prev_period = load_previous_period(keyword)
    mom_rate = compute_growth_rate(trend["total"], prev_period)

    yoy_period = load_same_period_last_year(keyword, start_date, end_date)
    if yoy_period is None:
        yoy_start, yoy_end = same_period_last_year_dates(start_date, end_date)
        print(f"\n查無去年同期（{yoy_start} ~ {yoy_end}）本地快照，"
              f"向 KEYPO 額外查詢總聲量以供比較…")
        try:
            yoy_data_path = f"data/_yoy_freqdist.json"
            fetch_volume_only(
                search_query, yoy_start, yoy_end, yoy_data_path,
                exclude_lottery=exclude_lottery,
                custom_exclusion=custom_exclusion
            )
            yoy_trend = parse_volume_trend(load_json(yoy_data_path))
            yoy_period = {
                "start_date": yoy_start,
                "end_date": yoy_end,
                "total_volume": yoy_trend["total"]
            }
            # 快取起來，下次分析同品牌時可直接沿用，不必重複查詢
            save_current_period(keyword, yoy_start, yoy_end, yoy_trend["total"])
        except Exception as e:
            print(f"去年同期資料查詢失敗，略過年增率比較：{e}")
            yoy_period = None

    yoy_rate = compute_growth_rate(trend["total"], yoy_period)

    growth_info = {
        "mom": (
            {"rate": mom_rate, "prev_total": prev_period["total_volume"]}
            if mom_rate is not None else None
        ),
        "yoy": (
            {"rate": yoy_rate, "prev_total": yoy_period["total_volume"]}
            if yoy_rate is not None else None
        ),
    }

    kols = parse_kol_channels(
    load_json("data/sprdtrnd.json")
    )

    # 關鍵領袖多維指標（發表數／回文／按讚／互動總計）——依熱門文章實際發文者聚合，
    # 對照 KEYPO 官方月報的四維 KOL 表；報告與 PPT 的 KOL 頁改用此資料。
    kol_metrics = parse_kol_from_articles(articles)

    # 來源類型分布（社群／新聞／討論區／部落格）——供 PPT「來源分布」頁與圓餅
    source_dist = parse_source_distribution(articles)

    # 聲量趨勢事件敘事——把前三大高峰對應到具體事件（Ollama，失敗自動退化）
    trend_narrative = narrative_for_trend(trend, articles, source_dist, keyword)

    # =========================
    # 競品分析（選填）
    # =========================

    competitor_data = []
    for comp in competitors:
        print(f"\n抓取競品資料：{comp}")
        try:
            comp_dir = f"data/_competitors/{comp}"
            fetch_endpoints(
                comp, start_date, end_date, comp_dir,
                ["freqdist", "sentidist", "hotrank"],
                exclude_lottery=exclude_lottery,
                custom_exclusion=competitor_exclusions.get(comp, "")
            )
            competitor_data.append({
                "name": comp,
                "trend": parse_volume_trend(load_json(f"{comp_dir}/freqdist.json")),
                "sentiment": parse_sentiment(load_json(f"{comp_dir}/sentidist.json")),
                "articles": parse_hotrank(load_json(f"{comp_dir}/hotrank.json")),
            })
        except Exception as e:
            print(f"競品「{comp}」資料抓取失敗，略過：{e}")

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

    trend_callouts = build_trend_callouts(trend, articles)
    generate_volume_trend_chart(
        trend, keyword, out_dir=charts_dir, callouts=trend_callouts
    )
    generate_sentiment_pie_chart(sentiment, keyword, out_dir=charts_dir)
    generate_source_pie_chart(source_dist, keyword, out_dir=charts_dir)
    generate_keyword_bar_chart(keywords, out_dir=charts_dir)
    generate_keyword_cloud(keywords, out_dir=charts_dir)

    if competitor_data:
        generate_competitor_volume_chart(keyword, trend, competitor_data, out_dir=charts_dir)
        generate_competitor_sentiment_chart(keyword, sentiment, competitor_data, out_dir=charts_dir)

    print("\n")
    print(report)

    # =========================
    # AI 洞察
    # =========================

    print("\n")
    print("=" * 60)
    print("AI 洞察")
    print("=" * 60)

    ai_report = analyze_with_ai(
        report,
        keywords=keywords,
        articles=articles,
        kols=kols
    )

    print(ai_report)

    # =========================
    # 表格敘事摘要（AI 輔助，數字段落由程式組出、敘事段落交給 Ollama）
    # =========================

    period_text = f"{start_date}~{end_date}"

    table_captions = {
        "keywords": _safe_caption(
            f"{keyword} {period_text}，熱門關鍵字聲量共 {sum(k['count'] for k in keywords[:20]):,} 筆。",
            narrative_for_keywords, keywords, articles
        ),
        "articles": _safe_caption(
            f"{keyword} {period_text}，熱門文章互動數合計 "
            f"{sum(a['engagement'] for a in articles[:10]):,} 次。",
            narrative_for_articles, articles
        ),
        "kol": _safe_caption(
            f"{keyword} {period_text}，關鍵領袖互動總計 "
            f"{sum(k['engagement'] for k in kol_metrics[:10]):,} 次。",
            narrative_for_kol, kol_metrics, articles
        ),
        "competitors": {
            c["name"]: _safe_caption(
                f"{c['name']} {period_text}，聲量為 {c['trend']['total']:,} 筆。",
                narrative_for_competitor, c["name"], c["articles"]
            )
            for c in competitor_data
        }
    }

    # =========================
    # 輸出資料夾
    # =========================
    # 註：精簡摘要（summary）與 AI 洞察（ai_report）已整併進詳細報告 Word——
    # §1/§3/§8 涵蓋摘要數據、§9 完整收錄 AI 六節洞察，故不再另存 summary.md / ai_insight.md。

    os.makedirs(out_dir, exist_ok=True)

    # =========================
    # 產生詳細分析報告（Word 版，真表格，供正式交付）
    # =========================
    try:
        from detailed_report_docx import generate_detailed_report_docx
        detailed_docx = (
            f"{out_dir}/{keyword}_{report_start_date}"
            f"_{report_end_date}_詳細分析報告.docx"
        )
        generate_detailed_report_docx(
            keyword, sentiment, keywords, articles, trend, kol_metrics,
            ai_report, report_start_date, report_end_date, detailed_docx,
            growth_info=growth_info, competitor_data=competitor_data,
            table_captions=table_captions, source_dist=source_dist,
            trend_narrative=trend_narrative,
        )
        print(f"✅ 詳細報告 Word：{detailed_docx}")
    except Exception as e:
        print(f"詳細報告 Word 版產生失敗，已略過：{e}")

    # =========================
    # 產生 PPT
    # =========================

    ppt_path = generate_ppt(
        keyword,
        report,
        ai_report,
        keywords,
        kol_metrics,
        articles,
        sentiment,
        trend,
        report_start_date,
        report_end_date,
        out_dir=out_dir,
        growth_info=growth_info,
        competitor_data=competitor_data,
        table_captions=table_captions,
        source_dist=source_dist,
        trend_narrative=trend_narrative
    )

    # =========================
    # 市場議題分析規劃（前瞻文件，AI 依品牌自動擬定議題，Word）
    # 定位：規劃「接下來要分析哪些市場議題」，與本期單品牌數據報告（PPT）不同方向。
    # 與報告同步產出，存到同一個 output/{品牌}/ 資料夾；失敗不影響前面產物。
    # =========================

    framework_path = None
    try:
        from framework_generator import build_config_with_ai, generate_framework
        print("\n產生市場議題分析規劃（AI 前瞻擬定議題／矩陣／腳本）…")
        fw_config = build_config_with_ai(
            keyword,
            industry=industry_hint,
            competitors=competitors,
            period=f"{start_date} – {end_date}",
        )
        framework_path = f"{out_dir}/市場議題分析規劃_{keyword}.docx"
        generate_framework(fw_config, framework_path)
    except Exception as e:
        print(f"市場議題分析規劃產生失敗，已略過（不影響報告）：{e}")

    # =========================
    # 儲存本期快照（供下次比較用）
    # =========================

    save_current_period(keyword, start_date, end_date, trend["total"])

    print("\n✅ 已輸出報告、圖表與 PPT")
    print(f"✅ PPT 路徑：{ppt_path}")
    if framework_path:
        print(f"✅ 市場議題分析規劃：{framework_path}")


if __name__ == "__main__":
    main()