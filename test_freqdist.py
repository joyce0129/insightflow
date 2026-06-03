from keypo_loader import load_json, parse_volume_trend

data = load_json("data/freqdist.json")
trend = parse_volume_trend(data)

print("每日聲量趨勢")
print("=" * 40)
print(f"總聲量：{trend['total']} 筆")
print(f"聲量高峰：{trend['peak_date']}，{trend['peak_volume']} 筆")
print()

for item in trend["daily"]:
    print(f"{item['date']}：{item['volume']} 筆")