import httpx


def analyze_with_ollama(report_text):

    prompt = f"""
你是一位資深品牌輿情分析顧問。

請根據以下輿情資料產出分析。

{report_text}

請輸出：

# 執行摘要

# 重點洞察
(3點)

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
        timeout=180
    )

    response.raise_for_status()

    return response.json()["response"]