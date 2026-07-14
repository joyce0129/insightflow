from anthropic import Anthropic
import os
from dotenv import load_dotenv

load_dotenv()

client = Anthropic(
    api_key=os.getenv("ANTHROPIC_API_KEY")
)

def analyze_report(report_text):

    prompt = f"""
你是一位資深品牌輿情分析顧問。

以下是KEYPO輿情資料：

{report_text}

請輸出：

1. 重點洞察
2. 正面議題
3. 負面議題
4. 風險觀察
5. 建議行動

使用繁體中文。
"""

    response = client.messages.create(
        model=os.getenv("CLAUDE_MODEL"),
        max_tokens=1500,
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    return response.content[0].text