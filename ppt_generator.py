from pptx import Presentation
from pptx.util import Inches, Pt
import os
from datetime import datetime


def add_title(slide, title):
    box = slide.shapes.add_textbox(
        Inches(0.5),
        Inches(0.3),
        Inches(12),
        Inches(0.6)
    )
    p = box.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True


def add_text(slide, text, x, y, w, h, font_size=16):
    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h)
    )
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    tf.paragraphs[0].font.size = Pt(font_size)


def format_date(date_str):
    """
    將 20260504 轉成 2026/05/04
    """
    return f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:]}"

def set_table_font(table, font_size=12):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)

def generate_ppt(
    brand_name,
    report,
    ai_report,
    keywords,
    kols,
    start_date,
    end_date
):
    os.makedirs("output", exist_ok=True)

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    display_start_date = format_date(start_date)
    display_end_date = format_date(end_date)

    # Slide 1 封面
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(
        slide,
        f"{brand_name} 網路輿情調查報告",
        1.0,
        2.4,
        11,
        0.8,
        36
    )
    add_text(
        slide,
        f"分析期間：{display_start_date} - {display_end_date}",
        1.0,
        3.3,
        10,
        0.5,
        20
    )

    # Slide 2 重點摘要
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "輿情重點摘要")
    add_text(
        slide,
        report,
        0.8,
        1.2,
        12,
        5.8,
        15
    )

    # Slide 3 聲量趨勢
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "聲量趨勢")
    slide.shapes.add_picture(
        "output/charts/volume_trend.png",
        Inches(1),
        Inches(1.2),
        width=Inches(11)
    )

    # Slide 4 情緒分布
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "情緒分布")
    slide.shapes.add_picture(
        "output/charts/sentiment_pie.png",
        Inches(3.5),
        Inches(1.1),
        width=Inches(6)
    )

    # Slide 5 熱門關鍵字
    slide = prs.slides.add_slide(
        prs.slide_layouts[6]
    )
    add_title(
        slide,
        "熱門關鍵字 TOP10"
    )
    slide.shapes.add_picture(
        "output/charts/keyword_bar.png",
        Inches(1),
        Inches(1.2),
        width=Inches(11)
    )

    # Slide 6 關鍵字雲

    slide = prs.slides.add_slide(
        prs.slide_layouts[6]
    )

    add_title(
        slide,
        "關鍵字雲"
    )

    slide.shapes.add_picture(
        "output/charts/keyword_cloud.png",
        Inches(0.8),
        Inches(1.0),
        width=Inches(11.5)
    )

    # Slide 7 網路關鍵領袖分析

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_title(
        slide,
        "網路關鍵領袖分析"
    )

    rows = min(len(kols), 10) + 1
    cols = 4

    table = slide.shapes.add_table(
        rows,
        cols,
        Inches(0.5),
        Inches(1.2),
        Inches(12),
        Inches(4.5)
    ).table

    # 設定欄寬
    table.columns[0].width = Inches(1.0)
    table.columns[1].width = Inches(2.0)
    table.columns[2].width = Inches(7.0)
    table.columns[3].width = Inches(2.0)

    # 表頭
    headers = ["排名", "平台", "頻道名稱", "聲量"]

    for col_index, header in enumerate(headers):
        table.cell(0, col_index).text = header

    # 表格資料
    for row_index, kol in enumerate(kols[:10], start=1):
        table.cell(row_index, 0).text = str(kol["rank"])
        table.cell(row_index, 1).text = str(kol["platform"])
        table.cell(row_index, 2).text = str(kol["channel"])
        table.cell(row_index, 3).text = str(kol["volume"])

    set_table_font(table, font_size=12)

    # Slide 8 AI洞察
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "AI 輿情洞察")
    add_text(
        slide,
        ai_report,
        0.8,
        1.1,
        12,
        6.0,
        14
    )


    timestamp = datetime.now().strftime("%H%M%S")

    output_path = (
        f"output/{brand_name}_{start_date}_{end_date}_{timestamp}_Report.pptx"
    )

    prs.save(output_path)

    print("PPT已儲存：", output_path)

    print("檔案存在嗎？", os.path.exists(output_path))

    if os.path.exists(output_path):
        os.startfile(os.path.abspath(output_path))

    return output_path