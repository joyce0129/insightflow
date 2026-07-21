import re
import httpx


def analyze_with_ollama(report_text, keywords=None, articles=None, kols=None):

    # 熱門關鍵字 TOP10
    kw_lines = ""
    if keywords:
        kw_lines = "\n".join(
            f"  {kw['rank']}. {kw['keyword']}（{kw['count']} 次）"
            for kw in keywords[:10]
        )

    # 熱門文章 TOP5
    article_lines = ""
    if articles:
        article_lines = "\n".join(
            f"  {a['rank']}. [{a['source']}] {a['title'][:60]}"
            for a in articles[:5]
        )

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

    response = httpx.post(
        "http://localhost:11434/api/generate",
        json={
            "model": "qwen2.5:7b",
            "prompt": prompt,
            "stream": False
        },
        timeout=300
    )

    response.raise_for_status()

    return response.json()["response"]


def _clean_title(text, max_len=40):
    text = ''.join(c for c in str(text) if ord(c) <= 0xFFFF)
    text = ' '.join(text.split())
    if len(text) > max_len:
        text = text[:max_len] + '…'
    return text


def _clean_narrative(text, max_len=260):
    """移除 emoji、控制字元，確保輸出與 Microsoft JhengHei 字型相容並控制長度"""
    text = ''.join(c for c in str(text) if ord(c) <= 0xFFFF)
    text = re.sub(r'[\x00-\x1f]', ' ', text)
    text = re.sub(r'^["「『]|["」』]$', '', text.strip())
    text = ' '.join(text.split())
    if len(text) > max_len:
        text = text[:max_len] + '…'
    return text


def _ollama_short(prompt, timeout=90):
    """共用的短文字生成呼叫，供表格敘事摘要使用（與完整六節分析分開，
    避免互相拖慢；逾時/連線失敗時直接拋出例外，由呼叫端決定如何降級）"""
    response = httpx.post(
        "http://localhost:11434/api/generate",
        json={
            "model": "qwen2.5:7b",
            "prompt": prompt,
            "stream": False
        },
        timeout=timeout
    )
    response.raise_for_status()
    return _clean_narrative(response.json()["response"])


_NARRATIVE_RULES = (
    "請務必使用繁體中文（台灣用語習慣），絕對不可出現任何簡體字或簡體用語。\n"
    "請寫 2 到 4 句完整具體的敘事，內容需包含明確的人事時地物"
    "（是誰、在哪個平台/頻道、做了什麼具體行為、引發什麼反應或後續發展），"
    "盡量從文章標題中萃取具體細節（涉及的品項、地點、行為、爭議點等）寫成完整故事，"
    "避免「有網友討論」「引發熱議」「反應熱烈」等籠統空泛的說法。"
    "不要提及任何聲量筆數、百分比或排名數字，不要加上「摘要：」等前綴，"
    "不要使用「摘要」「總結」等字眼，不要用引號包住整段文字，直接輸出敘事文字即可。"
)


def narrative_for_articles(articles):
    """依熱門文章標題產生本期熱門文章的事件敘事摘要"""
    if not articles:
        return ""
    lines = "\n".join(
        f"- [{a.get('source', '')}] {_clean_title(a.get('title', ''), 100)}"
        for a in articles[:5]
    )
    prompt = f"""你是一位品牌輿情分析師。以下是本期最熱門的討論文章標題：

{lines}

請摘要這些文章反映的主要事件或話題。{_NARRATIVE_RULES}
"""
    return _ollama_short(prompt)


def narrative_for_keywords(keywords, articles):
    """依熱門關鍵字與代表性文章產生關鍵字話題脈絡敘事"""
    if not keywords:
        return ""
    kw_line = "、".join(k["keyword"] for k in keywords[:5])
    article_lines = "\n".join(
        f"- [{a.get('source', '')}] {_clean_title(a.get('title', ''), 100)}"
        for a in (articles or [])[:5]
    )
    prompt = f"""你是一位品牌輿情分析師。本期熱門關鍵字包含：{kw_line}。
以下是相關的熱門討論文章標題，供參考話題脈絡：

{article_lines or "（無相關文章資料）"}

請說明這些關鍵字反映出的主要話題脈絡，可直接引用關鍵字本身。{_NARRATIVE_RULES}
"""
    return _ollama_short(prompt)


def narrative_for_kol(kols, articles):
    """依關鍵領袖與代表性文章產生 KOL 影響力敘事"""
    if not kols:
        return ""
    kol_line = "、".join(f"{k['platform']} {k['channel']}" for k in kols[:3])
    article_lines = "\n".join(
        f"- [{a.get('source', '')}] {_clean_title(a.get('title', ''), 100)}"
        for a in (articles or [])[:5]
    )
    prompt = f"""你是一位品牌輿情分析師。本期主要的關鍵領袖（KOL）／頻道包含：{kol_line}。
以下是本期熱門討論文章標題，供參考話題脈絡：

{article_lines or "（無相關文章資料）"}

請說明這些關鍵領袖對本期輿情擴散的影響或帶動的話題。{_NARRATIVE_RULES}
"""
    return _ollama_short(prompt)


def narrative_for_competitor(name, articles):
    """依競品的熱門文章標題產生該競品的事件敘事摘要"""
    if not articles:
        return ""
    lines = "\n".join(
        f"- [{a.get('source', '')}] {_clean_title(a.get('title', ''), 100)}"
        for a in articles[:5]
    )
    prompt = f"""你是一位品牌輿情分析師。以下是「{name}」本期最熱門的討論文章標題：

{lines}

請摘要這些文章反映的主要事件或話題。{_NARRATIVE_RULES}
"""
    return _ollama_short(prompt)