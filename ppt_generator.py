"""
ppt_generator.py — InsightFlow PPT 產生器（乾淨版面）
10 頁結構：封面 / 執行摘要 / 聲量趨勢 / 情緒分布 /
           熱門關鍵字 / 關鍵字雲 / 熱門文章 / KOL / AI洞察 / 封底
"""
import os
import re
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# ── 色盤 ─────────────────────────────────────────────
C_NAVY   = RGBColor(0x1A, 0x23, 0x3B)   # 深藍  — 標題底色
C_TEAL   = RGBColor(0x00, 0x78, 0x8A)   # 青藍  — 強調色、子標
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)   # 白
C_GRAY   = RGBColor(0x5B, 0x64, 0x70)   # 灰    — 次要文字
C_LIGHT  = RGBColor(0xF4, 0xF6, 0xF8)   # 淺灰  — 卡片背景
C_ORANGE = RGBColor(0xF5, 0xA6, 0x23)   # 橘    — INSIGHT 側條
C_DARK   = RGBColor(0x33, 0x33, 0x33)   # 深灰  — 內文
C_POS    = RGBColor(0x3B, 0x7D, 0xD8)   # 藍    — 正面
C_NEG    = RGBColor(0xC0, 0x39, 0x2B)   # 紅    — 負面
C_NEU    = RGBColor(0x7A, 0x7D, 0x82)   # 灰    — 中立
C_LINE   = RGBColor(0xD5, 0xD9, 0xDF)   # 淡灰  — 分隔線
C_HDR    = RGBColor(0x20, 0x2A, 0x44)   # 頁首底色

FONT = "Microsoft JhengHei"

# ── 版面常數 ─────────────────────────────────────────
W      = Inches(13.33)   # 投影片寬
H      = Inches(7.5)     # 投影片高
HDR_H  = Inches(0.72)    # 頁首高
FTR_Y  = Inches(7.1)     # 頁尾 Y
ML     = Inches(0.5)     # 左邊距
CW     = Inches(12.33)   # 內容寬 (W - 2*ML)
CY     = Inches(0.88)    # 內容起始 Y


# ══════════════════════════════════════════════════════
# 基礎工具
# ══════════════════════════════════════════════════════

def _rect(slide, x, y, w, h, fill=None):
    """加入矩形色塊"""
    shp = slide.shapes.add_shape(1, x, y, w, h)
    shp.line.fill.background()
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    return shp


def _tb(slide, x, y, w, h, text,
        size=13, bold=False, color=C_DARK,
        align=PP_ALIGN.LEFT, wrap=True):
    """加入單段文字框"""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = str(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def _bullet_box(slide, x, y, w, h, items,
                size=13, spacing=6, color=C_DARK):
    """加入項目符號清單"""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        if not first:
            p.space_before = Pt(spacing)
        first = False
        r = p.add_run()
        r.text = str(item)
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.color.rgb = color


def _cell(cell, text, size=11, bold=False,
          color=C_DARK, bg=None, align=PP_ALIGN.LEFT):
    """設定表格儲存格"""
    if bg:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
    cell.text = str(text)
    tf = cell.text_frame
    tf.word_wrap = True
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.name = FONT
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = color


def _clean(text, max_len=40):
    """移除 emoji 與換行，截斷過長文字"""
    cleaned = ''.join(c for c in str(text) if ord(c) <= 0xFFFF)
    cleaned = re.sub(r'[\x00-\x1f]', ' ', cleaned)
    cleaned = ' '.join(cleaned.split())
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len] + '…'
    return cleaned


def _parse_ai(text):
    """解析 AI 報告各節（# 標題）"""
    sections = {}
    key, lines = None, []
    for line in text.split('\n'):
        if line.startswith('# '):
            if key:
                sections[key] = '\n'.join(lines).strip()
            key, lines = line[2:].strip(), []
        else:
            lines.append(line)
    if key:
        sections[key] = '\n'.join(lines).strip()
    return sections


def _strip_md(text):
    """移除 Markdown 粗體、清單符號"""
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)
    text = re.sub(r'\*([^*]+)\*', r'\1', text)
    return text


def format_date(s):
    return f"{s[:4]}/{s[4:6]}/{s[6:]}"


# ══════════════════════════════════════════════════════
# 共用版面元件
# ══════════════════════════════════════════════════════

def _header(slide, title, page_num, brand, period):
    """深色頁首 + 青色側條 + 頁尾"""
    # 頁首底色
    _rect(slide, 0, 0, W, HDR_H, fill=C_HDR)
    # 左側青色條
    _rect(slide, 0, 0, Inches(0.07), HDR_H, fill=C_TEAL)
    # 標題
    _tb(slide, ML + Inches(0.1), Inches(0.14),
        Inches(11.0), Inches(0.5),
        title, size=19, bold=True, color=C_WHITE, wrap=False)
    # 頁碼
    _tb(slide, Inches(12.5), Inches(0.2),
        Inches(0.7), Inches(0.38),
        f"{page_num:02d}",
        size=11, color=RGBColor(0x88, 0x99, 0xAA),
        align=PP_ALIGN.RIGHT, wrap=False)
    # 頁尾分隔線
    _rect(slide, 0, FTR_Y, W, Inches(0.02), fill=C_LINE)
    # 頁尾文字
    _tb(slide, ML, Inches(7.13), CW, Inches(0.28),
        f"{brand} 網路輿情分析報告  ／  {period}"
        f"  ／  資料來源：KEYPO 大數據關鍵引擎",
        size=9, color=C_GRAY)


def _data_card(slide, x, y, w, h, value, label, accent=C_TEAL):
    """資料卡片（大數字 + 標籤）"""
    _rect(slide, x, y, w, h, fill=C_LIGHT)
    _rect(slide, x, y, Inches(0.07), h, fill=accent)
    _tb(slide, x + Inches(0.18), y + Inches(0.1),
        w - Inches(0.25), Inches(0.7),
        value, size=24, bold=True, color=C_NAVY, wrap=False)
    _tb(slide, x + Inches(0.18), y + h - Inches(0.36),
        w - Inches(0.25), Inches(0.3),
        label, size=10, color=C_GRAY)


def _insight(slide, x, y, w, h, items):
    """INSIGHT 橘色側條灰底框"""
    _rect(slide, x, y, w, h, fill=C_LIGHT)
    _rect(slide, x, y, Inches(0.07), h, fill=C_ORANGE)
    _tb(slide, x + Inches(0.18), y + Inches(0.07),
        Inches(1.5), Inches(0.26),
        "INSIGHT", size=10, bold=True, color=C_TEAL)
    _bullet_box(slide,
                x + Inches(0.18), y + Inches(0.37),
                w - Inches(0.28), h - Inches(0.47),
                [f"● {i}" for i in items],
                size=12, spacing=5)


def _section_label(slide, x, y, text, color=C_NAVY):
    """小節標籤（▎ 字首）"""
    _tb(slide, x, y, Inches(6.0), Inches(0.32),
        f"▎ {text}", size=12, bold=True, color=color)


# ══════════════════════════════════════════════════════
# 各頁生成
# ══════════════════════════════════════════════════════

def _slide_cover(prs, brand, period):
    """封面"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 左側色條組合
    _rect(slide, 0, 0, Inches(0.45), H, fill=C_TEAL)
    _rect(slide, Inches(0.45), 0, Inches(0.07), H, fill=C_LIGHT)

    # 品牌名稱
    _tb(slide, Inches(0.85), Inches(2.0), Inches(11.5), Inches(1.15),
        brand, size=42, bold=True, color=C_NAVY)

    # 副標
    _tb(slide, Inches(0.85), Inches(3.25), Inches(10.0), Inches(0.65),
        "網路輿情分析報告", size=22, color=C_TEAL)

    # 分隔線
    _rect(slide, Inches(0.85), Inches(4.05), Inches(9.5), Inches(0.03),
          fill=C_LINE)

    # 期間
    _tb(slide, Inches(0.85), Inches(4.2), Inches(10.0), Inches(0.45),
        f"分析期間：{period}", size=14, color=C_GRAY)

    # 資料來源
    _tb(slide, Inches(0.85), Inches(4.75), Inches(10.0), Inches(0.38),
        "資料來源｜KEYPO 大數據關鍵引擎", size=12, color=C_GRAY)


def _slide_summary(prs, brand, period, trend, sentiment, keywords):
    """執行摘要：資料卡片 + 重點列表"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "執行摘要", 1, brand, period)

    # ── 資料卡片 (4 個) ──────────────────────────────
    card_w = Inches(2.88)
    card_h = Inches(1.28)
    gap    = Inches(0.19)
    cy     = Inches(1.02)

    cards = [
        (f"{trend['total']:,}筆",          "本週總聲量",  C_TEAL),
        (str(sentiment["pn_value"]),        "P/N 值",     C_POS),
        (sentiment["positive_percent"],     "正面聲量",    C_POS),
        (sentiment["negative_percent"],     "負面聲量",    C_NEG),
    ]
    for i, (val, lbl, acc) in enumerate(cards):
        _data_card(slide,
                   ML + i * (card_w + gap), cy,
                   card_w, card_h, val, lbl, accent=acc)

    # ── 分隔 ─────────────────────────────────────────
    _tb(slide, ML, Inches(2.44), Inches(4.0), Inches(0.3),
        "▎ 本週重點摘要", size=12, bold=True, color=C_NAVY)
    _rect(slide, ML, Inches(2.78), CW, Inches(0.02), fill=C_LINE)

    # ── 重點列表 ─────────────────────────────────────
    peak = trend["peak_date"]
    neg  = sentiment["negative"]
    kw0, kw1, kw2 = keywords[0], keywords[1], keywords[2]

    bullets = [
        f"▪  聲量高峰 {peak}（{trend['peak_volume']:,}筆），"
        f"本週累積 {trend['total']:,}筆，平均每日"
        f" {trend['total'] // len(trend['daily']):,}筆。",

        f"▪  情緒結構：正面 {sentiment['positive_percent']}、"
        f"中立 {sentiment['neutral_percent']}、"
        f"負面 {sentiment['negative_percent']}；"
        f"P/N 值 {sentiment['pn_value']}。",

        f"▪  熱門關鍵字 TOP3：{kw0['keyword']}（{kw0['count']}）、"
        f"{kw1['keyword']}（{kw1['count']}）、"
        f"{kw2['keyword']}（{kw2['count']}）。",

        f"▪  負面聲量 {neg:,}筆（{sentiment['negative_percent']}），"
        + ("超過 100 筆，建議持續監控社群討論。"
           if neg > 100 else "維持低水位。"),
    ]

    _bullet_box(slide, ML, Inches(2.9), CW, Inches(4.0),
                bullets, size=14, spacing=14)


def _slide_trend(prs, brand, period, trend, charts_dir):
    """聲量趨勢圖"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "聲量趨勢分析", 2, brand, period)

    slide.shapes.add_picture(
        f"{charts_dir}/volume_trend.png",
        ML, CY, width=CW, height=Inches(4.6))

    peak = trend["peak_date"]
    _insight(slide, ML, Inches(5.68), CW, Inches(1.27), [
        f"本週聲量高峰出現於 {peak}（{trend['peak_volume']:,}筆），"
        f"建議對照熱門文章確認事件驅動因素。",
        f"週累積 {trend['total']:,}筆，"
        f"平均每日 {trend['total'] // len(trend['daily']):,}筆。",
    ])


def _slide_sentiment(prs, brand, period, sentiment, charts_dir):
    """情緒分布：圓餅圖 + 數據卡"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "情緒分布分析", 3, brand, period)

    # 圓餅圖（左）
    slide.shapes.add_picture(
        f"{charts_dir}/sentiment_pie.png",
        ML, CY, width=Inches(6.5), height=Inches(5.2))

    # 數據區（右）
    rx = Inches(7.3)
    _tb(slide, rx, Inches(1.05), Inches(5.5), Inches(0.32),
        "情緒結構詳細數據", size=13, bold=True, color=C_NAVY)
    _rect(slide, rx, Inches(1.42), Inches(5.5), Inches(0.03), fill=C_LINE)

    rows = [
        ("正面", sentiment["positive"], sentiment["positive_percent"], C_POS),
        ("中立", sentiment["neutral"],  sentiment["neutral_percent"],  C_NEU),
        ("負面", sentiment["negative"], sentiment["negative_percent"], C_NEG),
    ]
    for i, (label, vol, pct, color) in enumerate(rows):
        sy = Inches(1.58) + i * Inches(1.45)
        _rect(slide, rx, sy + Inches(0.16), Inches(0.22), Inches(0.22), fill=color)
        _tb(slide, rx + Inches(0.38), sy,
            Inches(1.8), Inches(0.35), label, size=13, bold=True, color=C_DARK)
        _tb(slide, rx + Inches(0.38), sy + Inches(0.38),
            Inches(5.0), Inches(0.5),
            f"{vol:,} 筆", size=22, bold=True, color=color, wrap=False)
        _tb(slide, rx + Inches(0.38), sy + Inches(0.92),
            Inches(5.0), Inches(0.32), pct, size=13, color=C_GRAY)

    # P/N 值
    _rect(slide, rx, Inches(5.98), Inches(5.5), Inches(0.03), fill=C_LINE)
    _tb(slide, rx, Inches(6.07), Inches(5.5), Inches(0.38),
        f"P/N 值：{sentiment['pn_value']}",
        size=15, bold=True, color=C_TEAL)


def _slide_keywords(prs, brand, period, charts_dir):
    """熱門關鍵字橫條圖"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "熱門關鍵字 TOP10", 4, brand, period)

    slide.shapes.add_picture(
        f"{charts_dir}/keyword_bar.png",
        ML, CY, width=CW, height=Inches(6.08))


def _slide_wordcloud(prs, brand, period, charts_dir):
    """關鍵字雲"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "關鍵字雲", 5, brand, period)

    slide.shapes.add_picture(
        f"{charts_dir}/keyword_cloud.png",
        ML, CY, width=CW, height=Inches(5.7))


def _slide_articles(prs, brand, period, articles):
    """熱門文章 TOP5 表格"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "熱門文章 TOP5", 6, brand, period)

    n = min(len(articles), 5)
    tbl = slide.shapes.add_table(
        n + 1, 5,
        ML, CY,
        CW, Inches(5.97)
    ).table

    # 欄寬（合計 = CW = 12.33"）
    for i, cw in enumerate(
        [Inches(0.55), Inches(1.3), Inches(1.9), Inches(6.8), Inches(1.78)]
    ):
        tbl.columns[i].width = cw

    # 表頭
    for j, hdr in enumerate(["排名", "平台", "頻道", "標題", "互動（cc）"]):
        _cell(tbl.cell(0, j), hdr,
              size=11, bold=True, color=C_WHITE,
              bg=C_NAVY, align=PP_ALIGN.CENTER)

    # 資料列
    for i, art in enumerate(articles[:n], start=1):
        bg = C_LIGHT if i % 2 == 0 else C_WHITE
        _cell(tbl.cell(i, 0), str(art["rank"]),
              bg=bg, align=PP_ALIGN.CENTER)
        _cell(tbl.cell(i, 1), art["source"],
              bg=bg, align=PP_ALIGN.CENTER)
        _cell(tbl.cell(i, 2), _clean(art["channel"], 14),
              bg=bg)
        _cell(tbl.cell(i, 3), _clean(art["title"], 44),
              bg=bg)
        _cell(tbl.cell(i, 4), f"{art['engagement']:,}",
              bg=bg, align=PP_ALIGN.CENTER)


def _slide_kol(prs, brand, period, kols):
    """網路關鍵領袖 TOP10 表格"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "網路關鍵領袖 TOP10", 7, brand, period)

    n = min(len(kols), 10)
    tbl = slide.shapes.add_table(
        n + 1, 4,
        ML, CY,
        CW, Inches(5.97)
    ).table

    # 欄寬（合計 = 12.33"）
    for i, cw in enumerate(
        [Inches(0.7), Inches(2.0), Inches(7.9), Inches(1.73)]
    ):
        tbl.columns[i].width = cw

    for j, hdr in enumerate(["排名", "平台", "頻道名稱", "聲量（筆）"]):
        _cell(tbl.cell(0, j), hdr,
              size=11, bold=True, color=C_WHITE,
              bg=C_NAVY, align=PP_ALIGN.CENTER)

    for i, kol in enumerate(kols[:n], start=1):
        bg = C_LIGHT if i % 2 == 0 else C_WHITE
        _cell(tbl.cell(i, 0), str(kol["rank"]),
              bg=bg, align=PP_ALIGN.CENTER)
        _cell(tbl.cell(i, 1), kol["platform"],
              bg=bg, align=PP_ALIGN.CENTER)
        _cell(tbl.cell(i, 2), kol["channel"], bg=bg)
        _cell(tbl.cell(i, 3), str(kol["volume"]),
              bg=bg, align=PP_ALIGN.CENTER)


def _slide_ai(prs, brand, period, ai_report):
    """AI 輿情洞察（雙欄版面）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, "AI 輿情洞察", 8, brand, period)

    secs = _parse_ai(ai_report)

    col_w = Inches(5.9)
    gap   = Inches(0.43)
    lx    = ML
    rx    = ML + col_w + gap

    # ── 左欄 ─────────────────────────────────────────
    # 執行摘要
    _section_label(slide, lx, Inches(0.95), "執行摘要")
    _rect(slide, lx, Inches(1.28), col_w, Inches(1.62), fill=C_LIGHT)
    summary = _strip_md(secs.get("執行摘要", ""))[:230]
    _bullet_box(slide, lx + Inches(0.15), Inches(1.35),
                col_w - Inches(0.2), Inches(1.5),
                [summary], size=11)

    # 重點洞察
    _section_label(slide, lx, Inches(3.02), "重點洞察")
    raw_lines = [
        _strip_md(l.strip())
        for l in secs.get("重點洞察", "").split('\n')
        if l.strip() and not l.strip().startswith('(')
    ]
    insight_items = []
    for line in raw_lines[:4]:
        if len(line) > 78:
            line = line[:78] + '…'
        insight_items.append(f"● {line}")

    _bullet_box(slide, lx, Inches(3.38),
                col_w, Inches(3.52),
                insight_items, size=12, spacing=8)

    # ── 右欄 ─────────────────────────────────────────
    # 正面議題
    _section_label(slide, rx, Inches(0.95), "正面議題", color=C_POS)
    pos_lines = [
        _strip_md(l.strip())
        for l in secs.get("正面議題", "").split('\n')
        if l.strip()
    ][:3]
    _bullet_box(slide, rx, Inches(1.28),
                col_w, Inches(1.55),
                [f"● {l[:70]}" for l in pos_lines],
                size=12, spacing=6)

    _rect(slide, rx, Inches(2.9), col_w, Inches(0.02), fill=C_LINE)

    # 負面議題
    _section_label(slide, rx, Inches(2.97), "負面議題", color=C_NEG)
    neg_lines = [
        _strip_md(l.strip())
        for l in secs.get("負面議題", "").split('\n')
        if l.strip()
    ][:3]
    _bullet_box(slide, rx, Inches(3.32),
                col_w, Inches(1.42),
                [f"● {l[:70]}" for l in neg_lines],
                size=12, spacing=6)

    _rect(slide, rx, Inches(4.8), col_w, Inches(0.02), fill=C_LINE)

    # 建議行動
    _section_label(slide, rx, Inches(4.87), "建議行動", color=C_TEAL)
    act_lines = [
        _strip_md(l.strip())
        for l in secs.get("建議行動", "").split('\n')
        if l.strip()
    ][:4]
    _bullet_box(slide, rx, Inches(5.22),
                col_w, Inches(1.7),
                [f"● {l[:70]}" for l in act_lines],
                size=12, spacing=6)


def _slide_back(prs, brand, period):
    """封底"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, fill=C_HDR)
    _rect(slide, 0, 0, Inches(0.45), H, fill=C_TEAL)
    _rect(slide, Inches(0.45), 0, Inches(0.07), H,
          fill=RGBColor(0x00, 0x50, 0x60))

    _tb(slide, Inches(0.9), Inches(2.5), Inches(11.0), Inches(1.1),
        brand, size=38, bold=True, color=C_WHITE)
    _tb(slide, Inches(0.9), Inches(3.7), Inches(10.0), Inches(0.6),
        "網路輿情分析報告", size=20, color=C_TEAL)
    _rect(slide, Inches(0.9), Inches(4.42), Inches(8.5), Inches(0.03),
          fill=RGBColor(0x3A, 0x4A, 0x5E))
    _tb(slide, Inches(0.9), Inches(4.57), Inches(10.0), Inches(0.4),
        period, size=13, color=C_GRAY)
    _tb(slide, Inches(0.9), Inches(5.08), Inches(10.0), Inches(0.35),
        "資料來源｜KEYPO 大數據關鍵引擎", size=11, color=C_GRAY)
    _tb(slide, Inches(0.9), Inches(6.45), Inches(10.0), Inches(0.35),
        "— 報告完 —", size=12,
        color=RGBColor(0x00, 0xBB, 0xCC))


# ══════════════════════════════════════════════════════
# 主函式
# ══════════════════════════════════════════════════════

def generate_ppt(
    brand_name, report, ai_report,
    keywords, kols, articles,
    sentiment, trend,
    start_date, end_date,
    out_dir="output"
):
    charts_dir = f"{out_dir}/charts"
    os.makedirs(out_dir, exist_ok=True)

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    disp_s = format_date(start_date)
    disp_e = format_date(end_date)
    period = f"{disp_s} – {disp_e}"

    _slide_cover(prs, brand_name, period)
    _slide_summary(prs, brand_name, period, trend, sentiment, keywords)
    _slide_trend(prs, brand_name, period, trend, charts_dir)
    _slide_sentiment(prs, brand_name, period, sentiment, charts_dir)
    _slide_keywords(prs, brand_name, period, charts_dir)
    _slide_wordcloud(prs, brand_name, period, charts_dir)
    _slide_articles(prs, brand_name, period, articles)
    _slide_kol(prs, brand_name, period, kols)
    _slide_ai(prs, brand_name, period, ai_report)
    _slide_back(prs, brand_name, period)

    timestamp = datetime.now().strftime("%H%M%S")
    output_path = (
        f"{out_dir}/{brand_name}_{start_date}_{end_date}"
        f"_{timestamp}_Report.pptx"
    )
    prs.save(output_path)

    print(f"PPT 已儲存：{output_path}")
    if os.path.exists(output_path):
        os.startfile(os.path.abspath(output_path))

    return output_path
