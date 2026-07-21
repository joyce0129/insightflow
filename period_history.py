"""
period_history.py — 品牌期間快照，供 PPT 計算「本期聲量成長率」與
「去年同期比較」用。簡化版：只記錄總聲量與觀測期間，不做完整歷史序列分析，
但改用清單保存多筆歷史快照，供跨年比對搜尋。
"""
import json
import os
from datetime import datetime, timedelta

HISTORY_DIR = "output/_history"

DATE_FMT = "%Y-%m-%d"
YOY_TOLERANCE_DAYS = 20   # 去年同期比對容許誤差天數


def same_period_last_year_dates(start_date, end_date):
    """回傳與本期對應的去年同期日期（YYYY-MM-DD），處理 2/29 邊界"""
    s = datetime.strptime(start_date, DATE_FMT)
    e = datetime.strptime(end_date, DATE_FMT)

    try:
        ys = s.replace(year=s.year - 1)
    except ValueError:
        ys = s.replace(year=s.year - 1, day=28)
    try:
        ye = e.replace(year=e.year - 1)
    except ValueError:
        ye = e.replace(year=e.year - 1, day=28)

    return ys.strftime(DATE_FMT), ye.strftime(DATE_FMT)


def _history_path(brand):
    return f"{HISTORY_DIR}/{brand}.json"


def _load_history(brand):
    path = _history_path(brand)
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # 相容舊格式（單一 dict，而非清單）
    if isinstance(data, dict):
        return [data]
    return data


def load_previous_period(brand):
    """回傳最近一次分析的快照（供「本期 vs 上期」比較），無資料則回傳 None"""
    history = _load_history(brand)
    if not history:
        return None
    return history[-1]


def load_same_period_last_year(brand, start_date, end_date):
    """在歷史快照中尋找與「去年同期」最接近的一筆（容許 ±20 天誤差），無資料則回傳 None"""
    history = _load_history(brand)
    if not history:
        return None

    try:
        cur_start = datetime.strptime(start_date, DATE_FMT)
    except ValueError:
        return None

    target = cur_start - timedelta(days=365)

    best, best_diff = None, None
    for entry in history:
        try:
            entry_start = datetime.strptime(entry["start_date"], DATE_FMT)
        except (KeyError, ValueError):
            continue
        diff = abs((entry_start - target).days)
        if diff <= YOY_TOLERANCE_DAYS and (best_diff is None or diff < best_diff):
            best, best_diff = entry, diff

    return best


def save_current_period(brand, start_date, end_date, total_volume):
    """將本期快照附加進歷史清單（同一期間重複執行時覆蓋舊紀錄）"""
    os.makedirs(HISTORY_DIR, exist_ok=True)
    history = _load_history(brand)

    history = [
        e for e in history
        if not (e.get("start_date") == start_date and e.get("end_date") == end_date)
    ]
    history.append({
        "start_date": start_date,
        "end_date": end_date,
        "total_volume": total_volume,
        "saved_at": datetime.now().isoformat()
    })
    history.sort(key=lambda e: e["start_date"])

    with open(_history_path(brand), "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def compute_growth_rate(current_total, previous):
    if not previous or not previous.get("total_volume"):
        return None
    prev_total = previous["total_volume"]
    if prev_total == 0:
        return None
    return round((current_total - prev_total) / prev_total * 100, 2)
