import re

import ai_backend


def analyze_with_ai(report_text, keywords=None, articles=None, kols=None):

    # 熱門關鍵字 TOP10
    kw_lines = ""
    if keywords:
        kw_lines = "\n".join(
            f"  {kw['rank']}. {kw['keyword']}（{kw['count']} 次）"
            for kw in keywords[:10]
        )

    # 熱門文章 TOP5（含內文摘要，供模型引用具體細節）
    article_lines = _article_lines(articles, 5) if articles else ""

    # KOL TOP5
    kol_lines = ""
    if kols:
        kol_lines = "\n".join(
            f"  {k['rank']}. [{k['platform']}] {k['channel']}（{k['volume']} 筆）"
            for k in kols[:5]
        )

    prompt = f"""
你是一位資深品牌輿情分析顧問。

請根據以下輿情資料產出深度分析，洞察需具體引用數據。

--- 聲量與情緒摘要 ---
{report_text}

--- 熱門關鍵字 TOP10 ---
{kw_lines or "（無資料）"}

--- 熱門文章 TOP5 ---
{article_lines or "（無資料）"}

--- 關鍵傳播者 KOL TOP5 ---
{kol_lines or "（無資料）"}

請輸出以下六個章節，每節至少 2～3 句具體說明，禁止籠統帶過：

# 執行摘要

# 重點洞察
（請列出 3 點，每點需引用具體關鍵字或文章）

# 正面議題

# 負面議題

# 風險觀察

# 建議行動

使用繁體中文。
"""

    # 六節深度分析較長，給較大的輸出上限
    return ai_backend.generate(prompt, timeout=300, max_tokens=2000)


def _clean_title(text, max_len=40):
    text = ''.join(c for c in str(text) if ord(c) <= 0xFFFF)
    text = ' '.join(text.split())
    if len(text) > max_len:
        text = text[:max_len] + '…'
    return text


def _summary_snippet(art, max_len=110):
    """取文章內文摘要（summary），去除 <em> 標記與 emoji，供 AI 敘事萃取細節。
    無 summary 時退回標題。"""
    text = art.get("summary") or art.get("title") or ""
    text = re.sub(r'<[^>]+>', '', str(text))          # 去除 <em> 命中標記
    text = ''.join(c for c in text if ord(c) <= 0xFFFF)
    text = re.sub(r'[\x00-\x1f]', ' ', text)
    text = ' '.join(text.split())
    if len(text) > max_len:
        text = text[:max_len] + '…'
    return text


def _article_lines(articles, n=5, with_summary=True):
    """組出供 AI 參考的文章清單：標題 + 內文摘要片段。
    內文摘要讓模型能寫出具體的人事時地物，而非僅憑標題臆測。"""
    lines = []
    for a in (articles or [])[:n]:
        head = f"- [{a.get('source', '')}] {_clean_title(a.get('title', ''), 60)}"
        if with_summary:
            snip = _summary_snippet(a, 110)
            if snip and snip != _clean_title(a.get('title', ''), 60):
                head += f"\n    內文摘要：{snip}"
        lines.append(head)
    return "\n".join(lines)


def _ensure_period(text):
    """確保敘事以句號收尾（摘要重點後要句點）"""
    text = text.strip()
    if text and text[-1] not in "。！？!?…":
        text += "。"
    return text


def _clean_narrative(text, max_len=260):
    """移除 emoji、控制字元，確保輸出與 Microsoft JhengHei 字型相容並控制長度；
    未被截斷時補上句號結尾。"""
    text = ''.join(c for c in str(text) if ord(c) <= 0xFFFF)
    text = re.sub(r'[\x00-\x1f]', ' ', text)
    text = re.sub(r'^["「『]|["」』]$', '', text.strip())
    text = ' '.join(text.split())
    text = re.sub(r'。(?:\s*。)+', '。', text)   # 收斂重複句號「。。」
    if len(text) > max_len:
        return text[:max_len] + '…'
    return _ensure_period(text)


def _short_gen(prompt, timeout=90):
    """共用的短文字生成呼叫，供表格敘事摘要使用（與完整六節分析分開，
    避免互相拖慢；逾時/連線失敗時直接拋出例外，由呼叫端決定如何降級）。
    後端由 AI_BACKEND 決定（ollama / claude / openai）。"""
    return _clean_narrative(ai_backend.generate(prompt, timeout=timeout, max_tokens=600))


_NARRATIVE_RULES = (
    "請務必使用繁體中文（台灣用語習慣），絕對不可出現任何簡體字或簡體用語。\n"
    "請寫 2 到 4 句完整具體的敘事，每一句都必須是完整句子並以句號「。」結尾。"
    "內容需包含明確的人事時地物（是誰、在哪個平台/頻道或品牌、做了什麼具體行為、"
    "涉及哪個活動/品項/地點、引發什麼反應或後續發展），"
    "務必從文章標題與內文摘要中萃取具體細節（活動名稱、品項、地點、人物、爭議點、贈品等）"
    "寫成完整故事，嚴禁「有網友討論」「引發熱議」「反應熱烈」「備受關注」等籠統空泛的說法。"
    "不要提及任何聲量筆數、百分比或排名數字，不要加上「摘要：」等前綴，"
    "不要使用「摘要」「總結」等字眼，不要用引號包住整段文字，直接輸出敘事文字即可。"
)


def narrative_for_articles(articles):
    """依熱門文章標題產生本期熱門文章的事件敘事摘要"""
    if not articles:
        return ""
    lines = _article_lines(articles, 5)
    prompt = f"""你是一位品牌輿情分析師。以下是本期最熱門的討論文章（含標題與內文摘要）：

{lines}

請摘要這些文章反映的主要事件或話題。{_NARRATIVE_RULES}
"""
    return _short_gen(prompt)


def narrative_for_keywords(keywords, articles):
    """依熱門關鍵字與代表性文章產生關鍵字話題脈絡敘事"""
    if not keywords:
        return ""
    kw_line = "、".join(k["keyword"] for k in keywords[:5])
    article_lines = _article_lines(articles, 5)
    prompt = f"""你是一位品牌輿情分析師。本期熱門關鍵字包含：{kw_line}。
以下是相關的熱門討論文章（含標題與內文摘要），供參考話題脈絡：

{article_lines or "（無相關文章資料）"}

請說明這些關鍵字反映出的主要話題脈絡，可直接引用關鍵字本身。{_NARRATIVE_RULES}
"""
    return _short_gen(prompt)


def narrative_for_kol(kols, articles):
    """依關鍵領袖與代表性文章產生 KOL 影響力敘事"""
    if not kols:
        return ""
    kol_line = "、".join(f"{k['platform']} {k['channel']}" for k in kols[:3])
    article_lines = _article_lines(articles, 5)
    prompt = f"""你是一位品牌輿情分析師。本期主要的關鍵領袖（KOL）／頻道包含：{kol_line}。
以下是本期熱門討論文章（含標題與內文摘要），供參考話題脈絡：

{article_lines or "（無相關文章資料）"}

請說明這些關鍵領袖對本期輿情擴散的影響或帶動的話題。{_NARRATIVE_RULES}
"""
    return _short_gen(prompt)


def narrative_for_competitor(name, articles):
    """依競品的熱門文章標題產生該競品的事件敘事摘要"""
    if not articles:
        return ""
    lines = _article_lines(articles, 5)
    prompt = f"""你是一位品牌輿情分析師。以下是「{name}」本期最熱門的討論文章（含標題與內文摘要）：

{lines}

請摘要這些文章反映的主要事件或話題。{_NARRATIVE_RULES}
"""
    return _short_gen(prompt)


# ══════════════════════════════════════════════════════
# 聲量趨勢事件敘事（對照 KEYPO 官方月報：把每個聲量高峰對應到具體事件）
# ══════════════════════════════════════════════════════

_PEAK_WORDS = ["最高峰", "次高峰", "第三高峰", "第四高峰", "第五高峰"]


def _to_trend_date(time_str):
    """'2026-06-16 15:05' -> '2026/06/16'（對齊 freqdist 日期格式）"""
    d = str(time_str).split(" ")[0]
    parts = d.split("-")
    return "/".join(parts) if len(parts) == 3 else d


def _md(trend_date):
    """'2026/06/08' -> '6/08'"""
    parts = str(trend_date).split("/")
    if len(parts) == 3:
        return f"{int(parts[1])}/{parts[2]}"
    return str(trend_date)


def _peak_events(trend, articles, n=3):
    """取聲量前 n 高峰日，並各自配對當日互動最高的代表文章"""
    daily = sorted(trend.get("daily", []),
                   key=lambda x: x["volume"], reverse=True)
    events = []
    for p in daily[:n]:
        same_day = [
            a for a in articles
            if _to_trend_date(a.get("time", "")) == p["date"]
        ]
        best = (max(same_day, key=lambda a: a.get("engagement", 0))
                if same_day else None)
        events.append({"date": p["date"], "volume": p["volume"], "article": best})
    return events


def _source_summary_sentence(brand, trend, source_dist):
    """趨勢敘事第 1 句：總聲量 + 來源結構（deterministic）"""
    total = trend.get("total", 0)
    items = [it for it in (source_dist or {}).get("items", []) if it["count"]]
    if items:
        items = sorted(items, key=lambda x: x["count"], reverse=True)
        top = items[0]
        rest = "、".join(it["type"] for it in items[1:3])
        rest_txt = f"，其次為{rest}" if rest else ""
        return (f"本期{brand}總聲量共{total:,}筆，討論來源以{top['type']}為主"
                f"（佔熱門文章{top['percent']}%）{rest_txt}。")
    return f"本期{brand}總聲量共{total:,}筆。"


def _fallback_peak_sentence(ev, word):
    """Ollama 不可用時的具體退化句（仍帶日期、代表文章與聲量）"""
    md = _md(ev["date"])
    art = ev["article"]
    if art:
        return (f"{word}出現在{md}，代表討論為[{art.get('source', '')}]"
                f"{_clean_title(art.get('title', ''), 42)}，"
                f"當日聲量達{ev['volume']:,}筆。")
    return f"{word}出現在{md}，當日聲量達{ev['volume']:,}筆。"


def _clean_event_desc(text):
    """清理模型產出的事件描述：去 <em>、去模型可能重複的高峰詞/日期前綴、去尾句號
    （固定的「{高峰名稱}出現在{M/D}，」前綴由程式組，模型只負責事件內容，較穩健）"""
    text = re.sub(r'<[^>]+>', '', str(text)).strip()
    for w in _PEAK_WORDS:
        text = text.replace(w, '')
    # 去除模型自帶的「出現在.../於.../在M月D日」等前綴
    text = re.sub(r'^[，,。：:\s]*', '', text)
    text = re.sub(r'^(出現在|發生在|於|在)?\s*'
                  r'([0-9一二三四五六七八九十兩零〇]{1,3}[/月]'
                  r'[0-9一二三四五六七八九十兩零〇]{1,3}日?)?[，,：:\s]*',
                  '', text)
    text = _clean_narrative(text, 150)
    return text.rstrip('。').strip()


def _ollama_peak_events(events, timeout=120):
    """請 Ollama 為每個高峰寫『事件描述』（不含日期與高峰詞，前綴由程式組），
    以「高峰N：」標記對齊，回傳 dict{index: 事件描述}。"""
    blocks = []
    for i, ev in enumerate(events):
        art = ev["article"]
        if art:
            blocks.append(
                f"高峰{i + 1}（{_md(ev['date'])}）\n"
                f"  代表文章：[{art.get('source', '')}] "
                f"{_clean_title(art.get('title', ''), 80)}\n"
                f"  內文摘要：{_summary_snippet(art, 130)}\n"
                f"  互動 {art.get('engagement', 0):,}｜按讚 {art.get('likes', 0):,}"
                f"｜回文 {art.get('replies', 0):,}"
            )
        else:
            blocks.append(f"高峰{i + 1}（{_md(ev['date'])}）\n  （當日無明確代表文章）")
    joined = "\n\n".join(blocks)
    prompt = f"""你是品牌輿情分析師，正在撰寫「聲量趨勢分析」。以下是本期由高到低的聲量高峰，以及各高峰當日互動最高的代表文章：

{joined}

請針對每一個高峰，各寫「一句」描述『當天發生什麼具體事件』的敘事：
- 必須包含具體的人事時地物（誰、在哪個平台或品牌、做了什麼、關鍵細節如活動名稱、品項、贈品、爭議點）。
- 務必從內文摘要萃取具體細節，嚴禁「引發熱議」「反應熱烈」「備受關注」等籠統說法。
- 句尾以互動或留言規模收束（把提供的互動/回文數改寫成如「吸引近萬人留言」「回文逾五千筆」）。
- 不要寫日期、不要寫「最高峰／次高峰」等字眼、不要寫聲量筆數，只描述事件本身。
- 使用繁體中文（台灣用語），絕對不可出現簡體字。

輸出格式：每個高峰一行，以「高峰1：」「高峰2：」「高峰3：」開頭，冒號後直接接該事件的一句敘事。
"""
    raw = ai_backend.generate(prompt, timeout=timeout, max_tokens=800)

    result = {}
    for ln in raw.split("\n"):
        m = re.match(r'^\D{0,4}高峰\s*([0-9１-９])\s*[：:、.\)]\s*(.+)', ln.strip())
        if m:
            idx = int(m.group(1).translate(str.maketrans("１２３４５６７８９", "123456789")))
            desc = _clean_event_desc(m.group(2))
            if desc:
                result[idx] = desc
    return result


def narrative_for_trend(trend, articles, source_dist=None, brand="品牌", n=3):
    """聲量趨勢分析的條列敘事：第 1 點為總聲量＋來源結構（deterministic），
    第 2 點起為各聲量高峰「{高峰名稱}出現在{M/D}，{事件}。」，事件由 Ollama 產出
    （失敗時退化為具體模板句）。每句均以句號結尾。回傳 list[str]。"""
    bullets = [_source_summary_sentence(brand, trend, source_dist)]
    events = _peak_events(trend, articles, n)

    ai_events = {}
    try:
        ai_events = _ollama_peak_events(events)
    except Exception as e:
        print(f"趨勢事件敘事產生失敗，改用基本敘事：{e}")

    for i, ev in enumerate(events):
        word = _PEAK_WORDS[i] if i < len(_PEAK_WORDS) else f"第{i + 1}高峰"
        desc = ai_events.get(i + 1)
        if desc:
            bullets.append(_ensure_period(f"{word}出現在{_md(ev['date'])}，{desc}"))
        else:
            bullets.append(_fallback_peak_sentence(ev, word))
    return bullets