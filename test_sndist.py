from keypo_loader import load_json, parse_sn_dist

data = load_json("data/sndist.json")
sn = parse_sn_dist(data)

print("社群活躍度 S/N")
print("=" * 40)
print(f"社群/討論區/部落格聲量：{sn['social_volume']} 筆")
print(f"新聞聲量：{sn['news_volume']} 筆")
print(f"S/N 值：{sn['sn_value']}")