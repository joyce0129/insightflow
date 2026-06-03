from keypo_loader import load_json, parse_hotrank

data = load_json("data/hotrank.json")
articles = parse_hotrank(data)

print("熱門文章 TOP 10")
print("=" * 60)

for item in articles[:10]:
    print(
        f"{item['rank']}. [{item['source']}] "
        f"{item['title'][:40]}"
    )