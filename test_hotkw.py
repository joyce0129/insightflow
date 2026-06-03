from keypo_loader import load_json, parse_hot_keywords

data = load_json("data/hotkw.json")
keywords = parse_hot_keywords(data)

print("熱門關鍵字 TOP 10")
print("=" * 40)

for item in keywords[:10]:
    print(f"{item['rank']}. {item['keyword']}：{item['count']} 次")