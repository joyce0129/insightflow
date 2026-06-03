def generate_summary(
    sentiment,
    keywords,
    articles,
    trend
):

    report = []

    report.append("=== 輿情重點摘要 ===")
    report.append("")

    report.append(
        f"總聲量：{trend['total']} 筆"
    )

    report.append(
        f"聲量高峰：{trend['peak_date']} "
        f"({trend['peak_volume']} 筆)"
    )

    report.append("")

    report.append(
        f"正面聲量 {sentiment['positive']} 筆 "
        f"({sentiment['positive_percent']})"
    )

    report.append(
        f"中立聲量 {sentiment['neutral']} 筆 "
        f"({sentiment['neutral_percent']})"
    )

    report.append(
        f"負面聲量 {sentiment['negative']} 筆 "
        f"({sentiment['negative_percent']})"
    )

    report.append(
        f"P/N值：{sentiment['pn_value']}"
    )

    report.append("")
    report.append("=== 熱門關鍵字 ===")

    for kw in keywords[:5]:
        report.append(
            f"{kw['rank']}. "
            f"{kw['keyword']} "
            f"({kw['count']})"
        )

    report.append("")
    report.append("=== 熱門文章 ===")

    for article in articles[:5]:
        report.append(
            f"{article['rank']}. "
            f"[{article['source']}] "
            f"{article['title'][:50]}"
        )

    report.append("")
    report.append("=== 風險觀察 ===")

    if sentiment["negative"] > 100:
        report.append(
            "本期負面聲量超過100筆，"
            "建議持續監控社群討論。"
        )
    else:
        report.append(
            "本期負面聲量較低。"
        )

    return "\n".join(report)