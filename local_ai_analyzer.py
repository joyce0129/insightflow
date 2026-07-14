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