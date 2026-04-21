"""
Génération PPTX agrégé : N questionnaires → camemberts par item + synthèse.
"""

import io
from collections import Counter, defaultdict

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


COLOR_PRIMARY = RGBColor(0x1A, 0x56, 0xDB)
COLOR_GREEN = RGBColor(0x22, 0xC5, 0x5E)
COLOR_YELLOW = RGBColor(0xFA, 0xCC, 0x15)
COLOR_ORANGE = RGBColor(0xF9, 0x73, 0x16)
COLOR_RED = RGBColor(0xEF, 0x44, 0x44)
COLOR_GRAY = RGBColor(0x9C, 0xA3, 0xAF)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK = RGBColor(0x1F, 0x29, 0x37)

RESPONSE_ORDER = [
    "Tout à fait satisfait",
    "Satisfait",
    "Peu satisfait",
    "Pas du tout satisfait",
]

RESPONSE_COLORS = {
    "Tout à fait satisfait": COLOR_GREEN,
    "Satisfait": COLOR_YELLOW,
    "Peu satisfait": COLOR_ORANGE,
    "Pas du tout satisfait": COLOR_RED,
}

RESPONSE_SCORES = {
    "Tout à fait satisfait": 4,
    "Satisfait": 3,
    "Peu satisfait": 2,
    "Pas du tout satisfait": 1,
}


def _aggregate(all_data: list[dict]) -> dict:
    agg = defaultdict(lambda: defaultdict(Counter))
    cat_order = []
    item_order = defaultdict(list)

    for data in all_data:
        for cat in data["categories"]:
            cat_name = cat["name"]
            if cat_name not in cat_order:
                cat_order.append(cat_name)
            for item in cat["items"]:
                label = item["label"]
                if label not in item_order[cat_name]:
                    item_order[cat_name].append(label)
                resp = item["response"]
                if resp in RESPONSE_SCORES:
                    agg[cat_name][label][resp] += 1

    return {
        "cat_order": cat_order,
        "item_order": item_order,
        "counts": agg,
        "n": len(all_data),
    }


def _add_title_slide(prs, title, n_questionnaires):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = COLOR_PRIMARY

    txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = f"Synthèse de {n_questionnaires} questionnaire(s)"
    p2.font.size = Pt(20)
    p2.font.color.rgb = COLOR_WHITE
    p2.alignment = PP_ALIGN.CENTER


def _add_pie_chart(slide, left, top, width, height, label, counter):
    txBox = slide.shapes.add_textbox(left, top - Inches(0.35), width, Inches(0.35))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = label
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    p.alignment = PP_ALIGN.CENTER

    chart_data = CategoryChartData()
    categories = []
    values = []
    for r in RESPONSE_ORDER:
        count = counter.get(r, 0)
        if count > 0:
            categories.append(r)
            values.append(count)

    if not values:
        return

    chart_data.categories = categories
    chart_data.add_series("", values)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, left, top, width, height, chart_data
    )
    chart = chart_frame.chart
    chart.has_legend = False

    series = chart.series[0]
    for idx, cat_name in enumerate(categories):
        point = series.points[idx]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = RESPONSE_COLORS[cat_name]

    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_percentage = True
    data_labels.show_value = False
    data_labels.show_category_name = False
    data_labels.font.size = Pt(9)
    data_labels.font.color.rgb = COLOR_DARK


def _add_category_slide(prs, cat_name, items, counts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.4), Inches(0.6))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = cat_name
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    pie_w = Inches(2.8)
    pie_h = Inches(2.5)
    x_start = Inches(0.3)
    x_gap = Inches(3.2)
    y_start = Inches(1.4)
    y_gap = Inches(3.2)

    for i, label in enumerate(items):
        col = i % 3
        row = i // 3
        left = x_start + col * x_gap
        top = y_start + row * y_gap
        counter = counts.get(label, Counter())
        _add_pie_chart(slide, left, top, pie_w, pie_h, label, counter)


def _add_global_summary_slide(prs, agg):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    txBox = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(9.4), Inches(0.6))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Synthèse globale"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY

    global_counter = Counter()
    total = 0
    score_sum = 0
    cat_scores = {}

    for cat_name in agg["cat_order"]:
        cat_total = 0
        cat_score = 0
        for label in agg["item_order"][cat_name]:
            for resp, count in agg["counts"][cat_name][label].items():
                global_counter[resp] += count
                total += count
                if resp in RESPONSE_SCORES:
                    score_sum += RESPONSE_SCORES[resp] * count
                    cat_score += RESPONSE_SCORES[resp] * count
                    cat_total += count
        if cat_total:
            cat_scores[cat_name] = cat_score / cat_total

    score_avg = score_sum / total if total else 0

    color = COLOR_GREEN if score_avg >= 3 else (COLOR_ORANGE if score_avg >= 2 else COLOR_RED)
    txBox2 = slide.shapes.add_textbox(Inches(0.3), Inches(1.0), Inches(3.5), Inches(1.2))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = f"Score moyen : {score_avg:.1f} / 4"
    p2.font.size = Pt(24)
    p2.font.bold = True
    p2.font.color.rgb = color

    p3 = tf2.add_paragraph()
    p3.text = f"{agg['n']} questionnaire(s), {total} réponses"
    p3.font.size = Pt(13)
    p3.font.color.rgb = COLOR_DARK

    y = Inches(2.5)
    for cat_name in agg["cat_order"]:
        if cat_name in cat_scores:
            avg = cat_scores[cat_name]
            c = COLOR_GREEN if avg >= 3 else (COLOR_ORANGE if avg >= 2 else COLOR_RED)
            txBox3 = slide.shapes.add_textbox(Inches(0.3), y, Inches(3.5), Inches(0.35))
            tf3 = txBox3.text_frame
            p4 = tf3.paragraphs[0]
            p4.text = f"{cat_name} : {avg:.1f}/4"
            p4.font.size = Pt(12)
            p4.font.color.rgb = c
            y += Inches(0.4)

    chart_data = CategoryChartData()
    cats = [r for r in RESPONSE_ORDER if global_counter.get(r, 0) > 0]
    vals = [global_counter[r] for r in cats]
    chart_data.categories = cats
    chart_data.add_series("", vals)

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(4.5), Inches(1.0), Inches(5), Inches(4.5),
        chart_data,
    )
    chart = chart_frame.chart

    series = chart.series[0]
    for idx, cat_name in enumerate(cats):
        point = series.points[idx]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = RESPONSE_COLORS[cat_name]

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(10)

    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_percentage = True
    data_labels.show_value = True
    data_labels.show_category_name = False
    data_labels.font.size = Pt(10)


def _add_legend_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    txBox = slide.shapes.add_textbox(Inches(2.5), Inches(2), Inches(5), Inches(3))
    tf = txBox.text_frame
    tf.word_wrap = True
    for resp in RESPONSE_ORDER:
        p = tf.add_paragraph()
        p.text = f"  {resp}"
        p.font.size = Pt(16)
        p.font.color.rgb = RESPONSE_COLORS[resp]
        p.font.bold = True
        p.space_after = Pt(8)


def generate_pptx(all_data: list[dict], title: str = "Résultats de satisfaction") -> bytes:
    """Génère un PPTX agrégé et retourne les bytes."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    agg = _aggregate(all_data)

    _add_title_slide(prs, title, agg["n"])

    for cat_name in agg["cat_order"]:
        items = agg["item_order"][cat_name]
        counts = agg["counts"][cat_name]
        for chunk_start in range(0, len(items), 6):
            chunk = items[chunk_start:chunk_start + 6]
            _add_category_slide(prs, cat_name, chunk, counts)

    _add_global_summary_slide(prs, agg)
    _add_legend_slide(prs)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()
