"""
detailed_report_generator.py — 自動產生詳細輿情分析報告（Markdown）
涵蓋 §1–§11，§7 核心議題以熱門文章自動填入框架，供分析師補充。
§12 競品聲量分析僅在呼叫時提供 competitor_data 才會輸出（選填章節）。
"""
import re
from datetime import datetime

from keypo_loader import classify_article, match_keywords_to_articles


# ── 工具函式 ─────────────────────────────────────────

def _clean(text, max_len=0):
    """移除 emoji、控制字元並截斷"""
    text = ''.join(c for c in str(text) if ord(c) <= 0xFFFF)
    text = re.sub(r'[\x00-\x1f]', ' ', text)
    text = ' '.join(text.split())
    if max_len and len(text) > max_len:
        text = text[:max_len] + '…'
    return text


def _fmt_date(s):
    """20260708 → 2026/07/08"""
    return f"{s[:4]}/{s[4:6]}/{s[6:]}"


def _parse_ai(text):
    """解析 AI 報告各節（# 標題）"""
    sections = {}
    key, lines = None, []
    for line in text.split('\n'):
        if line.startswith('# '):
            if key:
                sections[key] = '\n'.join(lines).strip()
            key, lines = line[2:].strip(), []
        else:
            lines.append(line)
    if key:
        sections[key] = '\n'.join(lines).strip()
    return sections


def _strip_md(text):
    """移除 Markdown 粗體/斜體符號"""
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)
    text = re.sub(r'\*([^*]+)\*', r'\1', text)
    return text


def _safe_float(s):
    try:
        return float(str(s).replace('%', ''))
    except Exception:
        return 0.0


def _growth_lines(growth_info):
    """把 mom/yoy 成長率轉成報告用的條列文字（無資料時回傳空清單）"""
    if not growth_info:
        return []
    lines = []
    mom = growth_info.get("mom")
    yoy = growth_info.get("yoy")
    if mom:
        word = "成長" if mom["rate"] >= 0 else "下滑"
        lines.append(
            f"本期聲量較上期（{mom['prev_total']:,}筆）{word} "
            f"{abs(mom['rate'])}%。"
        )
    if yoy:
        word = "成長" if yoy["rate"] >= 0 else "下滑"
        lines.append(
            f"本期聲量較去年同期（{yoy['prev_total']:,}筆）{word} "
            f"{abs(yoy['rate'])}%。"
        )
    return lines


# ══════════════════════════════════════════════════════
# 主函式
# ══════════════════════════════════════════════════════

AI_CAPTION_NOTE = "（以上事件摘要由 AI 輔助生成，請核對後使用，比照 §0 原則 5）"


def generate_detailed_report(
    brand_name, sentiment, keywords, articles,
    trend, kols, ai_report, start_date, end_date,
    growth_info=None, competitor_data=None, table_captions=None,
    source_dist=None, trend_narrative=None
):
    table_captions = table_captions or {}
    today  = datetime.today().strftime("%Y/%m/%d")
    disp_s = _fmt_date(start_date)
    disp_e = _fmt_date(end_date)
    period = f"{disp_s} – {disp_e}"
    days   = len(trend["daily"])
    avg_vol = trend["total"] // days if days else 0

    ai_secs = _parse_ai(ai_report)

    L = []
    h  = L.append
    br = lambda: L.append("")
    def div(n=82): L.append("─" * n)
    def eq(n=82):  L.append("=" * n)

    # ══ Header ════════════════════════════════════════
    eq()
    h(f" {brand_name} 品牌網路輿情分析報告")
    h(f" 觀測期間：{period}（共 {days} 天）")
    h(f" 資料來源：KEYPO 大數據關鍵引擎")
    h(f" 產出日期：{today}")
    h(f" 分析工具：InsightFlow AI 輿情分析系統")
    eq()
    br()

    # ══ §1 執行摘要 ═══════════════════════════════════
    div()
    h("§1  執行摘要")
    div()
    br()

    ai_summary = _strip_md(ai_secs.get("執行摘要", "")).strip()
    if ai_summary:
        h(ai_summary)
        br()

    h("關鍵結論：")
    h(f"  ▪ 本週總聲量 {trend['total']:,}筆，平均每日 {avg_vol:,}筆。")
    for line in _growth_lines(growth_info):
        h(f"  ▪ {line}")
    h(f"  ▪ 聲量高峰：{trend['peak_date']}（{trend['peak_volume']:,}筆）。")
    h(f"  ▪ 情緒結構：正面 {sentiment['positive_percent']}、"
      f"中立 {sentiment['neutral_percent']}、"
      f"負面 {sentiment['negative_percent']}；P/N 值 {sentiment['pn_value']}。")

    if len(keywords) >= 3:
        top3 = "、".join(
            f"{k['keyword']}（{k['count']}）" for k in keywords[:3]
        )
        h(f"  ▪ 熱門關鍵字 TOP3：{top3}。")

    neg = sentiment["negative"]
    if neg > 100:
        h(f"  ▪ 負面聲量 {neg:,}筆（{sentiment['negative_percent']}），"
          f"超過 100 筆，建議持續監控社群討論。")
    else:
        h(f"  ▪ 負面聲量 {neg:,}筆（{sentiment['negative_percent']}），維持低水位。")

    if kols:
        h(f"  ▪ 最具影響力 KOL："
          f"{kols[0]['platform']} {kols[0]['channel']}"
          f"（互動 {kols[0].get('engagement', kols[0].get('volume', 0)):,}）。")
    br()

    # ══ §2 聲量趨勢 ═══════════════════════════════════
    div()
    h("§2  聲量趨勢分析")
    div()
    br()

    h(f"{'日期':<14}  {'聲量（筆）':>10}  趨勢備註")
    h("─" * 55)
    for item in trend["daily"]:
        mark = "  ★高峰" if item["date"] == trend["peak_date"] else ""
        h(f"{item['date']:<14}  {item['volume']:>10,}{mark}")
    h("─" * 55)
    h(f"{'週累積':<14}  {trend['total']:>10,}")
    if growth_info:
        mom = growth_info.get("mom")
        yoy = growth_info.get("yoy")
        if mom:
            h(f"{'上期':<14}  {mom['prev_total']:>10,}"
              f"（本期成長率 {'+' if mom['rate'] >= 0 else ''}{mom['rate']}%）")
        if yoy:
            h(f"{'去年同期':<14}  {yoy['prev_total']:>10,}"
              f"（年增率 {'+' if yoy['rate'] >= 0 else ''}{yoy['rate']}%）")
    br()

    h("聲量觀察：")
    if trend_narrative:
        # 事件驅動敘事（總聲量＋來源結構＋各高峰對應具體事件）
        for line in trend_narrative:
            h(f"  ▪ {line}")
    else:
        h(f"  ▪ 高峰日 {trend['peak_date']} 聲量達 {trend['peak_volume']:,}筆，"
          f"建議對照熱門文章確認事件驅動因素。")
        h(f"  ▪ 週平均每日 {avg_vol:,}筆。")
    for line in _growth_lines(growth_info):
        h(f"  ▪ {line}")

    last_vol = trend["daily"][-1]["volume"]
    if last_vol < avg_vol * 0.5:
        h(f"  ▪ 末日聲量（{last_vol:,}筆）低於日均，"
          f"若為截止當天資料（非完整 24 小時）屬正常現象。")
    br()

    if source_dist and source_dist.get("total"):
        h("來源結構（依熱門文章來源類型，結構參考值）：")
        parts = [
            f"{it['type']} {it['count']}篇（{it['percent']}%）"
            for it in source_dist["items"] if it["count"]
        ]
        h("  ▪ " + "、".join(parts) + "。")
        br()

    # ══ §3 情緒分布 ═══════════════════════════════════
    div()
    h("§3  情緒分布分析")
    div()
    br()

    h(f"{'情緒':<6}  {'聲量（筆）':>10}  {'占比':>8}")
    h("─" * 32)
    h(f"{'正面':<6}  {sentiment['positive']:>10,}  {sentiment['positive_percent']:>8}")
    h(f"{'中立':<6}  {sentiment['neutral']:>10,}  {sentiment['neutral_percent']:>8}")
    h(f"{'負面':<6}  {sentiment['negative']:>10,}  {sentiment['negative_percent']:>8}")
    h("─" * 32)
    h(f"{'合計':<6}  {trend['total']:>10,}  {'100.00%':>8}  "
      f"P/N 值：{sentiment['pn_value']}")
    br()

    h("情緒解讀：")
    br()

    h(f"  【正面（{sentiment['positive_percent']}）】")
    pos_text = _strip_md(ai_secs.get("正面議題", "")).strip()
    if pos_text:
        for line in pos_text.split('\n')[:4]:
            line = line.strip()
            if line:
                h(f"    {line}")
    else:
        h(f"    本期正面聲量佔比 {sentiment['positive_percent']}，"
          f"主要來自品牌正向討論與產品好評。")
    br()

    h(f"  【中立（{sentiment['neutral_percent']}）】")
    h(f"    中立聲量佔比最高（{sentiment['neutral_percent']}），"
      f"以資訊轉發型、選購詢問型討論為主，情緒中性。")
    br()

    h(f"  【負面（{sentiment['negative_percent']}）】")
    neg_text = _strip_md(ai_secs.get("負面議題", "")).strip()
    if neg_text:
        for line in neg_text.split('\n')[:4]:
            line = line.strip()
            if line:
                h(f"    {line}")
    else:
        h(f"    本期負面聲量佔比 {sentiment['negative_percent']}，"
          f"請對照熱門文章確認主要負面來源。")
    br()

    pn = _safe_float(sentiment['pn_value'])
    h(f"  P/N 值 {sentiment['pn_value']} 評估：")
    if pn >= 2.0:
        h(f"    P/N 值高於 2.0，整體輿情正向，品牌口碑良好。")
    elif pn >= 1.5:
        h(f"    P/N 值介於 1.5–2.0，整體輿情偏正面，品牌基本盤穩固。")
    elif pn >= 1.0:
        h(f"    P/N 值介於 1.0–1.5，正面略優於負面，仍有提升空間。")
    else:
        h(f"    P/N 值低於 1.0，負面聲量超越正面，建議加強正向議題操作。")
    br()

    # ══ §4 熱門關鍵字 ═════════════════════════════════
    div()
    h("§4  熱門關鍵字分析（TOP 20）")
    div()
    br()

    h(f"{'排名':>4}  {'關鍵字':<18}  {'聲量（筆）':>10}")
    h("─" * 38)
    for kw in keywords[:20]:
        h(f"{kw['rank']:>4}  {kw['keyword']:<18}  {kw['count']:>10,}")
    br()

    # 關鍵字對應事件（把關鍵字還原成代表性文章）
    kw_events = match_keywords_to_articles(keywords, articles, top_n=8)
    if kw_events:
        h("關鍵字對應事件（TOP 8，取互動最高之命中文章）：")
        for ke in kw_events:
            art = ke.get("article")
            if art:
                event = f"[{art['source']}] {_clean(art.get('title', ''), 44)}"
            else:
                event = "（本期熱門文章未見直接對應）"
            h(f"  ▪ {ke['keyword']}（命中 {ke['match_count']} 篇）：{event}")
        br()

    if table_captions.get("keywords"):
        h(table_captions["keywords"])
        h(AI_CAPTION_NOTE)
        br()

    # ══ §5 熱門文章 ═══════════════════════════════════
    div()
    h("§5  熱門文章分析（TOP 10）")
    div()
    br()

    h(f"{'排名':>4}  {'平台':<10}  {'分類':<8}  {'互動':>6}  {'按讚':>8}  {'頻道':<16}  標題")
    h("─" * 100)
    for art in articles[:10]:
        title = _clean(art["title"], max_len=28)
        ch    = _clean(art["channel"], max_len=14)
        cls   = classify_article(art, brand_name)
        h(f"{art['rank']:>4}  {art['source']:<10}  {cls:<8}  "
          f"{art['engagement']:>6,}  {art['likes']:>8,}  "
          f"{ch:<16}  {title}")
    br()

    if table_captions.get("articles"):
        h(table_captions["articles"])
        h(AI_CAPTION_NOTE)
        br()

    h("熱門文章洞察：")
    if articles:
        top = articles[0]
        h(f"  ▪ 互動數最高：[{top['source']}] {_clean(top['channel'], 14)}"
          f"（互動 {top['engagement']:,}、按讚 {top['likes']:,}）。")

    top_like = max(articles[:10], key=lambda x: x["likes"]) if articles else None
    if top_like and articles and top_like["rank"] != articles[0]["rank"]:
        h(f"  ▪ 按讚數最高：[{top_like['source']}] "
          f"{_clean(top_like['channel'], 14)}（按讚 {top_like['likes']:,}）。")
    br()

    # ══ §6 KOL 分析 ═══════════════════════════════════
    div()
    h("§6  網路關鍵領袖（KOL）分析 — TOP 10")
    div()
    br()

    h(f"{'排名':>4}  {'平台':<10}  {'頻道名稱':<20}  "
      f"{'發表':>5}  {'回文':>7}  {'按讚':>9}  {'互動總計':>10}")
    h("─" * 82)
    for kol in kols[:10]:
        ch = _clean(kol["channel"], max_len=18)
        h(f"{kol['rank']:>4}  {kol['platform']:<10}  {ch:<20}  "
          f"{kol.get('posts', 0):>5,}  {kol.get('replies', 0):>7,}  "
          f"{kol.get('likes', 0):>9,}  {kol.get('engagement', 0):>10,}")
    br()

    if table_captions.get("kol"):
        h(table_captions["kol"])
        h(AI_CAPTION_NOTE)
        br()

    h("KOL 傳播分析：")
    if kols:
        h(f"  ▪ 最具影響力 KOL：{kols[0]['platform']} {kols[0]['channel']}"
          f"（互動總計 {kols[0].get('engagement', 0):,}、"
          f"按讚 {kols[0].get('likes', 0):,}），為本週聲量主要擴散來源。")

    platforms = {}
    for kol in kols[:10]:
        p = kol["platform"]
        platforms[p] = platforms.get(p, 0) + 1
    if platforms:
        top_p = max(platforms, key=platforms.get)
        h(f"  ▪ KOL 平台分布：{top_p} 頻道最多（{platforms[top_p]} 個），"
          f"為本週主要擴散平台。")
    br()

    # ══ §7 核心議題深度解讀 ═══════════════════════════
    div()
    h("§7  核心議題深度解讀")
    div()
    br()
    h("  ★ 以下依熱門文章 TOP3 自動框架，請分析師補充事件背景與留言走向。")
    br()

    for i, art in enumerate(articles[:3], 1):
        title = _clean(art["title"], max_len=55)
        ch    = _clean(art["channel"], max_len=20)
        h(f"【議題 {i}】{title}")
        h(f"  平台：{art['source']}　頻道：{ch}　時間：{art['time']}")
        h(f"  互動：{art['engagement']:,}　按讚：{art['likes']:,}　"
          f"分享：{art['replies']:,}")
        br()
        h(f"  事件背景：（請分析師補充）")
        h(f"  輿論走向：")
        h(f"    ▪ 正面：（請補充）")
        h(f"    ▪ 負面：（請補充）")
        h(f"  風險/機會面：（請補充）")
        br()

    # ══ §8 風險觀察 ═══════════════════════════════════
    div()
    h("§8  風險觀察與預警")
    div()
    br()

    h(f"{'風險項目':<18}  {'等級':<6}  說明")
    h("─" * 72)

    neg_vol = sentiment["negative"]
    if neg_vol > 300:
        h(f"{'負面聲量偏高':<18}  {'高':<6}  "
          f"負面聲量 {neg_vol:,}筆，超過 300 筆，建議立即查明主要負面來源。")
    elif neg_vol > 100:
        h(f"{'負面聲量偏高':<18}  {'中':<6}  "
          f"負面聲量 {neg_vol:,}筆，超過 100 筆，建議持續監控社群討論。")
    else:
        h(f"{'負面聲量':<18}  {'低':<6}  "
          f"負面聲量 {neg_vol:,}筆，維持低水位。")

    ai_risk = _strip_md(ai_secs.get("風險觀察", "")).strip()
    if ai_risk:
        for line in ai_risk.split('\n'):
            line = line.strip()
            if line:
                h(f"{'AI 風險提示':<18}  {'中':<6}  {line[:55]}")
    br()

    # ══ §9 AI 輿情洞察 ════════════════════════════════
    div()
    h("§9  AI 輿情洞察（Ollama / qwen2.5:7b）")
    div()
    br()
    h(ai_report.strip())
    br()

    # ══ §10 綜合建議行動 ══════════════════════════════
    div()
    h("§10  綜合建議行動")
    div()
    br()

    act_text = _strip_md(ai_secs.get("建議行動", "")).strip()
    if act_text:
        h("【AI 建議行動】")
        h(act_text)
        br()

    h("【分析師補充建議】")
    h("  短期（本週~下週）：")
    h("    1. （請補充）")
    h("    2. （請補充）")
    h("  中期（1 個月內）：")
    h("    1. （請補充）")
    h("    2. （請補充）")
    br()

    # ══ §11 研究方法 ══════════════════════════════════
    div()
    h("§11  研究方法與資料說明")
    div()
    br()

    h(f"資料來源    KEYPO 大數據關鍵引擎")
    h(f"查詢關鍵字  {brand_name}")
    h(f"觀測期間    {period}（共 {days} 天）")
    h(f"總處理聲量  {trend['total']:,} 筆")
    br()
    h("分析模組：")
    h("  ▪ 情緒分析   — KEYPO sentidist")
    h("  ▪ 關鍵字分析 — KEYPO hotkw")
    h("  ▪ 熱門文章   — KEYPO hotrank")
    h("  ▪ KOL 分析   — KEYPO sprdtrnd")
    h("  ▪ 聲量趨勢   — KEYPO freqdist")
    h("  ▪ AI 洞察    — Ollama qwen2.5:7b（本地模型）")
    br()
    h("報告產出工具：InsightFlow AI 輿情分析系統")
    br()

    # ══ §12 競品聲量分析（僅在有競品資料時輸出）══════════
    if competitor_data:
        div()
        h("§12  競品聲量分析")
        div()
        br()

        ranked = sorted(
            [{
                "name": brand_name, "total": trend["total"],
                "pn": sentiment["pn_value"],
                "pos": sentiment["positive_percent"],
                "neg": sentiment["negative_percent"],
            }] + [{
                "name": c["name"], "total": c["trend"]["total"],
                "pn": c["sentiment"]["pn_value"],
                "pos": c["sentiment"]["positive_percent"],
                "neg": c["sentiment"]["negative_percent"],
            } for c in competitor_data],
            key=lambda x: x["total"], reverse=True
        )

        h(f"{'品牌':<14}  {'總聲量（筆）':>10}  {'P/N值':>8}  {'正面占比':>10}  {'負面占比':>10}")
        h("─" * 62)
        for r in ranked:
            mark = "  ★本品牌" if r["name"] == brand_name else ""
            h(f"{r['name']:<14}  {r['total']:>10,}  {r['pn']:>8}  "
              f"{r['pos']:>10}  {r['neg']:>10}{mark}")
        br()

        competitor_captions = table_captions.get("competitors", {})
        if competitor_captions:
            for c in competitor_data:
                caption = competitor_captions.get(c["name"])
                if caption:
                    h(caption)
            h(AI_CAPTION_NOTE)
            br()

        h("比較觀察：")
        top = ranked[0]
        h(f"  ▪ 本期聲量最高為「{top['name']}」（{top['total']:,}筆）。")
        if top["name"] != brand_name:
            main_rank = next(i for i, r in enumerate(ranked, 1) if r["name"] == brand_name)
            main_total = next(r["total"] for r in ranked if r["name"] == brand_name)
            if top["total"]:
                ratio = round(top["total"] / main_total, 2) if main_total else None
                ratio_text = f"，約為本品牌的 {ratio} 倍" if ratio else ""
                h(f"  ▪ {brand_name} 本期聲量排名第 {main_rank}{ratio_text}。")
        else:
            h(f"  ▪ {brand_name} 本期聲量領先所有列入比較的競品。")

        best_pn = max(ranked, key=lambda x: _safe_float(x["pn"]))
        h(f"  ▪ 本期 P/N 值最高為「{best_pn['name']}」（{best_pn['pn']}）。")
        br()

        for c in competitor_data:
            if not c["articles"]:
                continue
            h(f"{c['name']} 熱門文章 TOP3：")
            for art in c["articles"][:3]:
                title = _clean(art["title"], max_len=40)
                h(f"  {art['rank']}. [{art['source']}] {title}"
                  f"（互動 {art['engagement']:,}）")
            br()

    # ══ Footer ════════════════════════════════════════
    eq()
    h(" — 報告完 —")
    h(f" {brand_name} 品牌網路輿情分析報告  ／  {period}")
    h(" 本報告由 InsightFlow AI 自動產出，AI 洞察為模型輔助分析。")
    h(" 數值與事件描述以 KEYPO 原始資料為準；§7 核心議題請分析師人工補充。")
    eq()

    return "\n".join(L)
