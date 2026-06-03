import json


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def parse_sentiment(data):
    result = {}

    for item in data["data"]:
        sentiment = item["senti"]
        result[sentiment] = {
            "total": item["t"],
            "percent": item["pc"],
            "daily": item["g"]
        }

    positive = result.get("正面", {}).get("total", 0)
    negative = result.get("負面", {}).get("total", 0)

    pn_value = round(positive / negative, 2) if negative else None

    return {
        "positive": positive,
        "neutral": result.get("中立", {}).get("total", 0),
        "negative": negative,
        "positive_percent": result.get("正面", {}).get("percent", "0%"),
        "neutral_percent": result.get("中立", {}).get("percent", "0%"),
        "negative_percent": result.get("負面", {}).get("percent", "0%"),
        "pn_value": pn_value,
        "dates": data["date"],
        "daily": result
    }

def parse_channel_rank(data):
    result = []

    for row in data["data"]:
        channel_index = row[0]
        volume = row[2]

        result.append({
            "rank": len(result) + 1,
            "channel": data["channels"][channel_index],
            "source": data["sprdtrnd_src"][channel_index],
            "name": data["sprdtrnd_ch"][channel_index],
            "volume": volume
        })

    return result

def parse_hot_keywords(data):
    result = []

    for index, item in enumerate(data["data"], start=1):
        result.append({
            "rank": index,
            "keyword": item["name"],
            "count": item["value"],
            "font_size": item["font_size"]
        })

    return result

def parse_hotrank(data):
    result = []

    for index, item in enumerate(data["data"], start=1):
        result.append({
            "rank": index,
            "title": item["title"],
            "source": item["src"],
            "channel": item["ch"],
            "author": item["author"],
            "time": item["time"],
            "url": item["url"],
            "engagement": item["cc"],
            "likes": item["lc"],
            "replies": item["rc"]
        })

    return result

def parse_sn_dist(data):
    result = {
        "social_volume": 0,
        "news_volume": 0,
        "sn_value": 0
    }

    for item in data["data"]:
        if item.get("nlist") == "nlist":
            result["social_volume"] = item["t"]
        elif item.get("slist") == "slist":
            result["news_volume"] = item["t"]
        elif item.get("sn") == "sn":
            result["sn_value"] = item["t"]

    return result

def parse_volume_trend(data):
    dates = data["date"]
    volumes = data["data"][0]["g"]

    result = []

    for index, date in enumerate(dates):
        result.append({
            "date": date,
            "volume": volumes[index]
        })

    peak = max(result, key=lambda x: x["volume"])

    return {
        "total": data["total"],
        "daily": result,
        "peak_date": peak["date"],
        "peak_volume": peak["volume"]
    }

def parse_kol_channels(data):

    result = []

    for idx, row in enumerate(data["data"]):

        channel_index = row[0]
        volume = row[2]

        result.append({
            "rank": idx + 1,
            "platform": data["sprdtrnd_src"][channel_index],
            "channel": data["sprdtrnd_ch"][channel_index],
            "volume": volume
        })

    return result