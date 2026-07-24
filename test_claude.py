import os
from dotenv import load_dotenv
from anthropic import Anthropic

load_dotenv()

client = Anthropic(
    api_key=os.getenv("ANTHROPIC_API_KEY")
)

response = client.messages.create(
    model=os.getenv("CLAUDE_MODEL"),
    max_tokens=100,
    messages=[
        {
            "role": "user",
            "content": "請回答：測試成功"
        }
    ]
)

print(response.content[0].text)