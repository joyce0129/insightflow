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
    totals = {}
    for row in data["data"]:
        channel_index = row[0]
        totals[channel_index] = totals.get(channel_index, 0) + row[2]

    ranked = sorted(totals.items(), key=lambda x: x[1], reverse=True)

    result = []
    for rank, (channel_index, volume) in enumerate(ranked, start=1):
        result.append({
            "rank": rank,
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
            "summary": item.get("summary", ""),
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


# ── 文章分類（heuristic）─────────────────────────────
# KEYPO 未提供文章角色分類欄位，依 src 與頻道名稱推斷，供報告「分類」欄使用。

_FORUM_SRC = {"PTT", "DCARD", "MOBILE01", "KOMICA", "BAHAMUT", "巴哈姆特"}

_NEWS_KEYWORDS = (
    "新聞", "日報", "晚報", "時報", "週刊", "周刊", "雜誌", "報導", "媒體",
    "Yahoo", "ETtoday", "ETFASHION", "TVBS", "中時", "聯合", "自由", "三立",
    "東森", "民視", "鏡", "風傳媒", "NOWnews", "中央社", "CNA", "LINE TODAY",
    "msn", "MSN", "財經", "通訊社", "新聞網", "新聞雲", "科技世代",
)


def classify_article(art, brand=""):
    """回傳文章分類：討論區 / 新聞媒體 / 官方 / 網友分享"""
    src = str(art.get("source", "")).upper()
    text = f"{art.get('channel', '')} {art.get('author', '')}"

    if src in _FORUM_SRC:
        return "討論區"
    if "NEWS" in src or "新聞" in src:
        return "新聞媒體"
    if any(kw.upper() in text.upper() for kw in _NEWS_KEYWORDS):
        return "新聞媒體"
    if brand and brand in text:
        return "官方"
    return "網友分享"


# ── 來源頻道類型分類（社群／新聞／討論區／部落格）────────────
# 對照 KEYPO 官方月報每個主題頁都有的「來源分布」圓餅。KEYPO 未在 hotrank
# 提供頻道類型欄位，故依 src 值推斷；此分布以熱門文章篇數計（非全站聲量加權），
# 屬結構參考值。

def classify_source_type(src):
    """回傳來源類型：討論區 / 部落格 / 新聞 / 社群"""
    s = str(src).upper()
    if any(k in s for k in (
        "PTT", "DCARD", "MOBILE01", "巴哈", "KOMICA", "論壇", "BBS",
        "卡提諾", "EYNY",
    )):
        return "討論區"
    if any(k in s for k in (
        "痞客邦", "PIXNET", "隨意窩", "部落格", "BLOG", "方格子",
        "MATTERS", "WORDPRESS",
    )):
        return "部落格"
    if any(k in s for k in (
        "新聞", "NEWS", "報", "週刊", "雜誌", "ETTODAY", "YAHOO", "TVBS",
        "中時", "聯合", "自由", "三立", "東森", "民視", "鏡", "中央社",
        "MSN", "通訊社", "NOWNEWS", "CNA",
    )):
        return "新聞"
    return "社群"


def parse_source_distribution(articles):
    """依熱門文章的來源類型統計分布（社群／新聞／討論區／部落格），
    回傳固定順序清單與總數，供「來源分布」圓餅使用。"""
    order = ["社群", "新聞", "討論區", "部落格"]
    counts = {t: 0 for t in order}
    for a in articles:
        counts[classify_source_type(a.get("source", ""))] += 1
    total = sum(counts.values())
    denom = total or 1
    items = [
        {"type": t, "count": counts[t],
         "percent": round(counts[t] / denom * 100, 1)}
        for t in order
    ]
    return {"total": total, "items": items}


def match_keywords_to_articles(keywords, articles, top_n=10):
    """為每個熱門關鍵字找出最具代表性的文章（標題或內文命中且互動最高），
    供『關鍵字對應事件』頁對照——把關鍵字還原成背後的具體事件。"""
    result = []
    for kw in keywords[:top_n]:
        term = str(kw["keyword"])
        matched = [
            a for a in articles
            if term in str(a.get("title", "")) or term in str(a.get("summary", ""))
        ]
        best = (max(matched, key=lambda a: a.get("engagement", 0))
                if matched else None)
        result.append({
            "keyword": term,
            "count": kw.get("count", 0),
            "match_count": len(matched),
            "article": best,
        })
    return result


def parse_kol_from_articles(articles, limit=10):
    """依熱門文章的實際發文者聚合關鍵領袖多維指標
    （發表數／回文數／按讚數／互動總計），依互動總計排序。

    對照 KEYPO 官方月報的「關鍵領袖 TOP5」四維表；sprdtrnd 僅有頻道聲量，
    無互動細項，故改由 hotrank 逐篇聚合而得。"""
    agg = {}
    for a in articles:
        platform = a.get("source", "")
        channel = a.get("author") or a.get("channel", "")
        if not channel:
            continue
        key = (platform, channel)
        entry = agg.setdefault(
            key, {"posts": 0, "likes": 0, "replies": 0, "cc": 0}
        )
        entry["posts"] += 1
        entry["likes"] += a.get("likes", 0) or 0
        entry["replies"] += a.get("replies", 0) or 0
        entry["cc"] += a.get("engagement", 0) or 0

    result = []
    ranked = sorted(
        agg.items(),
        key=lambda kv: kv[1]["likes"] + kv[1]["replies"] + kv[1]["cc"],
        reverse=True,
    )
    for rank, ((platform, channel), m) in enumerate(ranked[:limit], start=1):
        # 互動總計 = 按讚 + 回文 + 互動(cc)，對照官方月報的「互動數總計」grand total
        engagement = m["likes"] + m["replies"] + m["cc"]
        result.append({
            "rank": rank,
            "platform": platform,
            "channel": channel,
            "posts": m["posts"],
            "replies": m["replies"],
            "likes": m["likes"],
            "engagement": engagement,
            "volume": m["posts"],  # 相容舊欄位（部分呼叫端讀 volume）
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

    totals = {}
    for row in data["data"]:
        channel_index = row[0]
        totals[channel_index] = totals.get(channel_index, 0) + row[2]

    ranked = sorted(totals.items(), key=lambda x: x[1], reverse=True)

    result = []
    for rank, (channel_index, volume) in enumerate(ranked, start=1):
        result.append({
            "rank": rank,
            "platform": data["sprdtrnd_src"][channel_index],
            "channel": data["sprdtrnd_ch"][channel_index],
            "volume": volume
        })

    return result