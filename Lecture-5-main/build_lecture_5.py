#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Лекция 5: Формирование лазерного пятна — геометрическая и физическая оптика.
Курс «Физика лазеров для аддитивных (3D) технологий».
"""

import os
import re
from datetime import datetime

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, KeepTogether
from reportlab.lib.colors import darkgray, gray
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PIL import Image, ImageDraw, ImageFont


BASE = os.path.dirname(os.path.abspath(__file__))
PPTX_PATH = os.path.join(BASE, "Лекция_5_Геометрическая_оптика.pptx")
PDF_PATH = os.path.join(BASE, "Лекция_5_Раскадровка_для_лектора.pdf")
DOCX_PATH = os.path.join(BASE, "Лекция_5_Геометрическая оптика.docx")

LECTURE_TITLE_SHORT = "ФОРМИРОВАНИЕ ЛАЗЕРНОГО ПЯТНА"
LECTURE_TITLE_FULL = (
    "Формирование лазерного пятна: геометрическая и физическая оптика"
)


def _detect_font():
    windir = os.environ.get("WINDIR", r"C:\Windows")
    candidates = [
        os.path.join(windir, "Fonts", "arial.ttf"),
        os.path.join(windir, "Fonts", "segoeui.ttf"),
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Helvetica.ttc",
    ]
    for p in candidates:
        if os.path.exists(p):
            bold = (
                p.replace("arial.ttf", "arialbd.ttf")
                .replace("Arial.ttf", "Arial Bold.ttf")
                .replace("segoeui.ttf", "segoeuib.ttf")
            )
            if not os.path.exists(bold):
                bold = p
            return p, bold
    return None, None


def _detect_math_font():
    windir = os.environ.get("WINDIR", r"C:\Windows")
    for p in [
        os.path.join(windir, "Fonts", "timesi.ttf"),
        os.path.join(windir, "Fonts", "cambriai.ttf"),
        os.path.join(windir, "Fonts", "times.ttf"),
        "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
        "/Library/Fonts/Times New Roman.ttf",
    ]:
        if os.path.exists(p):
            return p
    return None


FONT_PATH, FONT_BOLD_PATH = _detect_font()
MATH_FONT_PATH = _detect_math_font()
if FONT_PATH:
    pdfmetrics.registerFont(TTFont("CyrFont", FONT_PATH))
    pdfmetrics.registerFont(TTFont("CyrFontBold", FONT_BOLD_PATH or FONT_PATH))
    PDF_FONT, PDF_FONT_BOLD = "CyrFont", "CyrFontBold"
else:
    PDF_FONT, PDF_FONT_BOLD = "Helvetica", "Helvetica-Bold"

if MATH_FONT_PATH:
    pdfmetrics.registerFont(TTFont("MathFont", MATH_FONT_PATH))
    PDF_MATH = "MathFont"
else:
    PDF_MATH = PDF_FONT


def _pil_font(size=14, bold=False):
    cands = []
    if bold and FONT_BOLD_PATH:
        cands.append(FONT_BOLD_PATH)
    cands += [
        FONT_PATH,
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf"),
    ]
    for p in cands:
        if p and os.path.exists(p):
            try:
                return ImageFont.truetype(p, size=size)
            except OSError:
                pass
    return ImageFont.load_default()


def _tw(d, text, font):
    text = _pil_safe(text)
    try:
        return int(d.textlength(text, font=font))
    except Exception:
        return len(text) * 7


def _draw_text(d, xy, text, fill, font):
    d.text(xy, _pil_safe(text), fill=fill, font=font)


def _draw_sup(d, x, y, base, sup, font, color="#333"):
    fs = _pil_font(max(6, getattr(font, "size", 12) - 5))
    d.text((x, y), base, fill=color, font=font)
    ox = x + _tw(d, base, font)
    d.text((ox, y - 4), sup, fill=color, font=fs)
    return ox + _tw(d, sup, fs)


def _pil_greek_font(size=12):
    for path in (
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "seguisym.ttf"),
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf"),
    ):
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except OSError:
                continue
    return _pil_font(size)


def _draw_greek(d, x, y, ch, font, color):
    if ch in ("\u03bd", "ν"):
        return _draw_greek_nu(d, x, y, int(getattr(font, "size", 12) or 12), color)
    g = _pil_greek_font(int(getattr(font, "size", 12) or 12))
    d.text((x, y), ch, fill=color, font=g)
    return x + _tw(d, ch, g)


def _draw_greek_nu(d, x, y, size, color):
    s = max(12, int(size) + 2)
    scale = 6
    for path in (
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "seguisym.ttf"),
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf"),
    ):
        if os.path.exists(path):
            break
    else:
        path = FONT_PATH
    if not path or not os.path.exists(path):
        return x + s
    big = ImageFont.truetype(path, s * scale)
    tmp = Image.new("RGBA", (s * scale * 2, s * scale * 2), (0, 0, 0, 0))
    td = ImageDraw.Draw(tmp)
    if isinstance(color, str) and color.startswith("#"):
        hx = color[1:]
        if len(hx) == 3:
            hx = "".join(c * 2 for c in hx)
        fill = tuple(int(hx[i:i + 2], 16) for i in (0, 2, 4)) + (255,)
    else:
        fill = (0, 0, 0, 255)
    td.text((scale, scale), "\U0001D708", fill=fill, font=big)
    bbox = tmp.getbbox()
    if not bbox:
        return x + s
    glyph = tmp.crop(bbox)
    tw = max(8, glyph.width // scale)
    th = max(8, glyph.height // scale)
    glyph = glyph.resize((tw, th), Image.Resampling.LANCZOS)
    host = getattr(d, "_image", None)
    if host is not None:
        host.paste(glyph, (int(x), int(y)), glyph)
    else:
        d.bitmap((x, y), glyph.convert("L"), fill=color)
    return x + tw + 1


def _to_unicode_sup(text: str) -> str:
    mp = str.maketrans("0123456789+-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻")
    return text.translate(mp)


def pretty_math(text):
    if not text:
        return text
    t = text
    t = re.sub(r"\blambda\b", "λ", t, flags=re.I)
    t = re.sub(r"\bpi\b", "π", t, flags=re.I)
    t = re.sub(r"\bDelta\s+n\b", "Δn", t)
    t = re.sub(r"\bDelta\b", "Δ", t)
    t = re.sub(r"\btheta(\d)\b", lambda m: "θ" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\btheta\b", "θ", t, flags=re.I)
    t = re.sub(r"см2\b", "см²", t)
    t = re.sub(r"мм2\b", "мм²", t)
    t = re.sub(r"/см2\b", "/см²", t)
    t = re.sub(r"/мм2\b", "/мм²", t)
    t = re.sub(r"(\d+(?:,\d+)?)·10\^(-?\d+)", lambda m: m.group(1) + "·10" + _to_unicode_sup(m.group(2)), t)
    t = re.sub(r"10\^(-?\d+)", lambda m: "10" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\^(-?\d+)", lambda m: _to_unicode_sup(m.group(1)), t)
    t = t.replace("NA^2", "NA²").replace("NA2", "NA²")
    t = re.sub(r"\bM2\b", "M²", t)
    return t


def _pil_safe(text):
    if not text:
        return text
    t = str(text)
    t = re.sub(r"\blambda\b", "λ", t, flags=re.I)
    t = re.sub(r"\bpi\b", "π", t, flags=re.I)
    t = re.sub(r"\btheta(\d)\b", lambda m: "θ" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\btheta\b", "θ", t, flags=re.I)
    t = re.sub(r"см2\b", "см²", t)
    t = re.sub(r"/см2\b", "/см²", t)
    t = re.sub(r"10\^(-?\d+)", lambda m: "10" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\^(-?\d+)", lambda m: _to_unicode_sup(m.group(1)), t)
    t = t.replace("NA^2", "NA²")
    return t


def pdf_sub(text):
    if not text:
        return text

    def math_ch(ch):
        return f'<font name="{PDF_MATH}"><i>{ch}</i></font>'

    t = text
    t = t.replace("CO₂", "CO<sub>2</sub>").replace("CO2", "CO<sub>2</sub>")
    t = re.sub(r"\bM2\b", "M²", t)
    t = t.replace("M²", "M<sup>2</sup>").replace("M^2", "M<sup>2</sup>")
    t = t.replace("−", "-")
    t = t.replace("·", "&middot;")
    t = re.sub(r"\blambda\b", "λ", t, flags=re.I)
    t = re.sub(r"\bpi\b", "π", t, flags=re.I)
    t = re.sub(r"\bDelta\s+n\b", "Δn", t)
    t = re.sub(r"\bDelta\b", "Δ", t)
    t = re.sub(r"\btheta(\d)\b", r"θ<sub>\1</sub>", t)
    t = re.sub(r"\btheta\b", "θ", t, flags=re.I)
    t = re.sub(r"см2\b", "см<sup>2</sup>", t)
    t = re.sub(r"мм2\b", "мм<sup>2</sup>", t)
    t = re.sub(r"/см2\b", "/см<sup>2</sup>", t)
    t = re.sub(r"/мм2\b", "/мм<sup>2</sup>", t)
    t = re.sub(r"NA\^2", "NA<sup>2</sup>", t)
    t = re.sub(r"NA2\b", "NA<sup>2</sup>", t)
    t = re.sub(r"\(([^()]+)\)\^(-?\d+)", r"(\1)<sup>\2</sup>", t)
    t = re.sub(r"(\d),(\d+)\^(-?\d+)", r"\1,\2<sup>\3</sup>", t)
    t = re.sub(r"10\^(-?\d+)", r"10<sup>\1</sup>", t)
    t = re.sub(r"10·10\^(-?\d+)", r"10&middot;10<sup>\1</sup>", t)
    t = re.sub(r"2·10\^(-?\d+)", r"2&middot;10<sup>\1</sup>", t)
    t = re.sub(r"(\d+(?:,\d+)?)·10\^(-?\d+)", r"\1&middot;10<sup>\2</sup>", t)
    t = re.sub(r"\^(-?\d+)", r"<sup>\1</sup>", t)
    uni_sup = str.maketrans({
        "⁰": "<sup>0</sup>", "¹": "<sup>1</sup>", "²": "<sup>2</sup>", "³": "<sup>3</sup>",
        "⁴": "<sup>4</sup>", "⁵": "<sup>5</sup>", "⁶": "<sup>6</sup>", "⁷": "<sup>7</sup>",
        "⁸": "<sup>8</sup>", "⁹": "<sup>9</sup>", "⁺": "<sup>+</sup>", "⁻": "<sup>-</sup>",
    })
    t = t.translate(uni_sup)
    for ch in ("ν", "ω", "π", "λ", "Λ", "ħ", "Δ", "θ"):
        t = t.replace(ch, math_ch(ch))
    return t


def save_pil_image(draw_fn, path):
    img = draw_fn()
    img.save(path, "PNG")
    return path


# ── Диаграммы ─────────────────────────────────────────────────────────────

def draw_intensity_comparison_pil():
    img = Image.new("RGB", (620, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(14, True), _pil_font(12), _pil_font(10)

    _draw_text(d, (12, 8), "Плотность мощности I = P / S (P = 200 Вт)", fill="#333", font=fh)

    cx1, cy1, r1 = 150, 210, 55
    d.ellipse([(cx1 - r1, cy1 - r1), (cx1 + r1, cy1 + r1)], outline="#1565c0", width=2)
    _draw_text(d, (70, 55), "Без фокусировки", fill="#1565c0", font=f)
    _draw_text(d, (70, 80), "d = 1 мм", fill="#333", font=fs)
    _draw_text(d, (70, 100), "S ~ 0,008 см²", fill="#666", font=fs)
    _draw_text(d, (70, 125), "I ~ 25 кВт/см²", fill="#e64a19", font=f)

    cx2, cy2, r2 = 430, 210, 18
    d.ellipse([(cx2 - r2, cy2 - r2), (cx2 + r2, cy2 + r2)], outline="#e64a19", width=3)
    _draw_text(d, (340, 55), "С фокусировкой", fill="#e64a19", font=f)
    _draw_text(d, (340, 80), "d = 50 мкм", fill="#333", font=fs)
    _draw_text(d, (340, 100), "S ~ 2·10⁻⁵ см²", fill="#666", font=fs)
    _draw_text(d, (340, 125), "I ~ 10 МВт/см²", fill="#c62828", font=f)

    d.line([(260, 210), (310, 210)], fill="#333", width=2)
    d.polygon([(310, 205), (325, 210), (310, 215)], fill="#333")
    _draw_text(d, (248, 175), "x 400", fill="#c62828", font=fh)

    _draw_text(d, (12, 340), "Та же мощность — другая площадь: оптика сжимает энергию в пространстве", fill="#555", font=fs)
    return img


def draw_rays_vs_waves_pil():
    """Геометрическая vs физическая оптика у фокуса."""
    img = Image.new("RGB", (640, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(13, True), _pil_font(11), _pil_font(9)

    _draw_text(d, (12, 8), "Два языка: лучи (геометрия) и волны (физика)", fill="#333", font=fh)

    # Left: rays
    _draw_text(d, (40, 45), "Геометрическая оптика", fill="#1565c0", font=f)
    d.line([(30, 200), (200, 200)], fill="#ccc", width=1)
    for y in (140, 200, 260):
        d.line([(40, y), (160, y)], fill="#1565c0", width=2)
        d.line([(160, y), (220, 200)], fill="#1565c0", width=2)
    d.ellipse([(216, 196), (224, 204)], fill="#c62828")
    _draw_text(d, (50, 300), "Лучи сходятся в «точку»", fill="#555", font=fs)
    _draw_text(d, (50, 320), "Хорошо далеко от фокуса", fill="#555", font=fs)

    # Divider
    d.line([(310, 50), (310, 360)], fill="#ddd", width=2)

    # Right: waves / diffraction
    _draw_text(d, (340, 45), "Физическая оптика", fill="#e64a19", font=f)
    # aperture
    d.rectangle([(360, 90), (520, 100)], fill="#333")
    d.rectangle([(420, 90), (460, 100)], fill="white")
    _draw_text(d, (430, 70), "апертура", fill="#666", font=fs)
    # wave fronts
    for i, r in enumerate((40, 70, 100, 130)):
        bbox = [440 - r, 100, 440 + r, 100 + 2 * r]
        d.arc(bbox, 20, 160, fill="#e64a19", width=2)
    # finite spot
    d.ellipse([(420, 280), (460, 310)], outline="#c62828", width=2, fill="#ffebee")
    _draw_text(d, (470, 285), "пятно d", fill="#c62828", font=f)
    _draw_text(d, (340, 330), "Волна «расплывается» — дифракция", fill="#555", font=fs)
    _draw_text(d, (340, 350), "Точки в фокусе нет", fill="#555", font=fs)
    return img


def draw_spot_size_pil():
    img = Image.new("RGB", (620, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(14, True), _pil_font(11), _pil_font(9)

    x = 12
    _draw_text(d, (x, 8), "Дифракционный предел: d ~ ", fill="#333", font=fh)
    x += _tw(d, "Дифракционный предел: d ~ ", fh)
    x = _draw_greek(d, x, 8, "\u03bb", fh, "#333")
    _draw_text(d, (x, 8), " / NA", fill="#333", font=fh)

    apex = (310, 280)
    d.polygon([(180, 80), (440, 80), apex[0], apex[1]], outline="#1565c0", fill="#e3f2fd")
    d.line([(180, 80), (440, 80)], fill="#1565c0", width=2)
    d.ellipse([(295, 268), (325, 288)], fill="#e64a19", outline="#c62828", width=2)
    _draw_text(d, (330, 265), "d", fill="#c62828", font=f)

    # angle mark
    d.line([(310, 280), (310, 100)], fill="#999", width=1)
    _draw_text(d, (318, 160), "θ", fill="#1565c0", font=f)
    _draw_text(d, (250, 95), "ширина пучка на линзе", fill="#666", font=fs)

    x = 12
    _draw_text(d, (x, 310), "d ~ ", fill="#1565c0", font=f)
    x += _tw(d, "d ~ ", f)
    x = _draw_greek(d, x, 310, "\u03bb", f, "#1565c0")
    _draw_text(d, (x, 310), " / NA    |    NA = n sin(θ)", fill="#1565c0", font=f)

    _draw_text(d, (12, 340), "NA = 0,1 -> d ~ 10,7 мкм  |  NA = 0,4 -> d ~ 2,7 мкм  (λ = 1,07 мкм)", fill="#333", font=fs)
    _draw_text(d, (12, 360), "Больше NA -> меньше пятно, но уже «коридор» фокуса (DOF)", fill="#666", font=fs)
    return img


def draw_optical_train_pil():
    """Схема оптического тракта СЛП (все блоки в пределах кадра)."""
    W, H = 640, 400
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(13, True), _pil_font(10), _pil_font(9)

    _draw_text(d, (12, 8), "Оптический тракт СЛП-станка (упрощённо)", fill="#333", font=fh)

    labels = [
        ("Волокно", "#00695c"),
        ("Коллиматор", "#1565c0"),
        ("Гальво-\nзеркала", "#6a1b9a"),
        ("F-θ\nлинза", "#e64a19"),
        ("Защитное\nстекло", "#455a64"),
        ("Порошок", "#c62828"),
    ]
    n = len(labels)
    box_w, box_h = 82, 70
    margin_l, margin_r = 14, 14
    y0 = 130
    span = W - margin_l - margin_r - box_w
    step = span / (n - 1) if n > 1 else 0
    xs = [int(round(margin_l + i * step)) for i in range(n)]

    for x, (label, col) in zip(xs, labels):
        d.rounded_rectangle([(x, y0), (x + box_w, y0 + box_h)], radius=8, outline=col, width=2, fill="#fafafa")
        lines = label.split("\n")
        for i, line in enumerate(lines):
            _draw_text(d, (x + 6, y0 + 18 + i * 16), line, fill=col, font=f)
    for i in range(n - 1):
        x1 = xs[i] + box_w
        x2 = xs[i + 1]
        yc = y0 + box_h // 2
        if x2 - x1 > 10:
            d.line([(x1 + 2, yc), (x2 - 6, yc)], fill="#333", width=2)
            d.polygon([(x2 - 10, yc - 5), (x2 - 2, yc), (x2 - 10, yc + 5)], fill="#333")

    _draw_text(d, (14, 230), "Волокно: доставка излучения, M² ≈ 1", fill="#555", font=fs)
    _draw_text(d, (14, 250), "Коллиматор: почти параллельный пучок", fill="#555", font=fs)
    _draw_text(d, (14, 270), "Гальво: отклонение по плоскости сканирования", fill="#555", font=fs)
    _draw_text(d, (14, 290), "F-θ: пятно по полю (напр. 250×250 мм)", fill="#555", font=fs)
    _draw_text(d, (14, 310), "Защитное стекло: загрязнение → нагрев → сдвиг фокуса", fill="#555", font=fs)
    _draw_text(d, (14, 345), "Итог: на порошок приходит пятно с заданными d и I", fill="#333", font=f)
    return img


def draw_dof_chart_pil():
    img = Image.new("RGB", (620, 420), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(13, True), _pil_font(10), _pil_font(9)

    _draw_text(d, (12, 6), "Компромисс NA: пятно d и глубина резкости DOF", fill="#333", font=fh)
    x = 12
    _draw_text(d, (x, 28), "DOF ~ ", fill="#666", font=fs)
    x += _tw(d, "DOF ~ ", fs)
    x = _draw_greek(d, x, 28, "\u03bb", fs, "#666")
    _draw_text(d, (x, 28), " / NA", fill="#666", font=fs)
    x += _tw(d, " / NA", fs)
    x = _draw_sup(d, x, 28, "", "2", fs, "#666")
    _draw_text(d, (x, 28), "  (падает быстрее, чем d)", fill="#666", font=fs)

    cases = [
        ("NA=0,1  пром.", 10.7, 107, "#009688"),
        ("NA=0,4  лаб.", 2.7, 6.7, "#e64a19"),
        ("NA=0,5  экстр.", 2.1, 4.3, "#c62828"),
    ]
    x0, y0, h = 70, 58, 230
    max_v = 120
    bar_w = 38
    gap = 95

    for i, (label, d_val, dof_val, col) in enumerate(cases):
        bx = x0 + i * gap
        _draw_text(d, (bx - 10, y0 - 20), label, fill=col, font=f)
        h_dof = int(dof_val / max_v * h)
        d.rectangle([(bx, y0 + h - h_dof), (bx + bar_w, y0 + h)], fill=col)
        _draw_text(d, (bx - 8, y0 + h + 6), f"DOF {dof_val} мкм", fill=col, font=fs)
        h_d = max(8, int(d_val / max_v * h))
        d.rectangle([(bx + bar_w + 10, y0 + h - h_d), (bx + bar_w + 24, y0 + h)], fill="#333")
        _draw_text(d, (bx + bar_w + 28, y0 + h - h_d - 2), f"d={d_val}", fill="#333", font=fs)

    _draw_text(d, (12, 395), "Малый NA: терпим к рельефу порошка", fill="#555", font=fs)
    _draw_text(d, (320, 395), "NA=0,5: DOF << неровности слоя", fill="#555", font=fs)
    return img


def draw_focus_practice_pil():
    """Положение фокуса относительно слоя порошка."""
    img = Image.new("RGB", (640, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(13, True), _pil_font(11), _pil_font(9)

    _draw_text(d, (12, 8), "Фокус на практике: где «горит» пятно относительно слоя", fill="#333", font=fh)

    # powder surface
    d.line([(40, 220), (600, 220)], fill="#8d6e63", width=3)
    _draw_text(d, (40, 230), "слой порошка", fill="#6d4c41", font=fs)
    # rough powder bumps
    for x in range(50, 580, 28):
        d.ellipse([(x, 208), (x + 14, 222)], fill="#bcaaa4", outline="#8d6e63")

    # three focus positions
    cases = [
        (120, "ниже слоя", "#1565c0", 260),
        (320, "в плоскости слоя", "#009688", 220),
        (520, "выше слоя", "#e64a19", 160),
    ]
    for cx, label, col, fy in cases:
        d.line([(cx - 50, 90), (cx, fy)], fill=col, width=2)
        d.line([(cx + 50, 90), (cx, fy)], fill=col, width=2)
        d.ellipse([(cx - 8, fy - 6), (cx + 8, fy + 6)], outline=col, width=2)
        _draw_text(d, (cx - 45, 55), label, fill=col, font=fs)

    # DOF band
    d.rectangle([(280, 190), (360, 250)], outline="#009688", width=1)
    _draw_text(d, (365, 205), "DOF", fill="#009688", font=f)

    _draw_text(d, (12, 300), "z-offset меняет d на поверхности → меняет I", fill="#333", font=f)
    _draw_text(d, (12, 325), "Загрязнённое защитное стекло → нагрев → сдвиг фокуса", fill="#555", font=fs)
    _draw_text(d, (12, 350), "Чистая оптика = стабильная геометрия пятна", fill="#c62828", font=f)
    return img


def draw_m2_comparison_pil():
    img = Image.new("RGB", (620, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(14, True), _pil_font(11), _pil_font(9)

    _draw_text(d, (12, 8), "Качество пучка M²: d_реал = M² · d_идеал", fill="#333", font=fh)

    cx1, cy1, r1 = 160, 200, 13
    d.ellipse([(cx1 - r1, cy1 - r1), (cx1 + r1, cy1 + r1)], fill="#009688", outline="#00695c", width=2)
    _draw_text(d, (80, 60), "Волоконный лазер", fill="#009688", font=f)
    _draw_text(d, (80, 85), "M² = 1,3; d_ид = 10 мкм", fill="#333", font=fs)
    _draw_text(d, (80, 110), "d_реал ~ 13 мкм", fill="#009688", font=f)

    cx2, cy2, r2 = 430, 200, 80
    d.ellipse([(cx2 - r2, cy2 - r2), (cx2 + r2, cy2 + r2)], outline="#e64a19", width=2, fill="#ffebee")
    d.ellipse([(cx2 - 13, cy2 - 13), (cx2 + 13, cy2 + 13)], fill="#009688", outline="#00695c", width=1)
    _draw_text(d, (350, 60), "Диодный лазер", fill="#e64a19", font=f)
    _draw_text(d, (350, 85), "M² = 20; d_ид = 10 мкм", fill="#333", font=fs)
    _draw_text(d, (350, 110), "d_реал ~ 200 мкм", fill="#e64a19", font=f)

    _draw_text(d, (12, 320), "I ~ 1/d²: плохой пучок → пятно ×20 → I в 400 раз меньше", fill="#c62828", font=fs)
    _draw_text(d, (12, 345), "В СЛП — волоконные лазеры с M² близким к 1 (связь с Лекцией 3)", fill="#555", font=fs)
    return img


NA_TABLE = {
    "headers": ["Параметр", "A: промышл.", "B: лаб.", "C: экстр."],
    "rows": [
        ["Числ. апертура (NA)", "0,1", "0,4", "0,5"],
        ["Пятно d, мкм", "~10,7", "~2,7", "~2,1"],
        ["DOF, мкм", "~107", "~6,7", "~4,3"],
        ["СЛП", "Да, стандарт", "Только лаб.", "Нереализуемо"],
        ["Почему?", "DOF > рельефа", "DOF < рельефа", "DOF << рельефа"],
    ],
}

# ── Раскадровка ───────────────────────────────────────────────────────────

READING_WPM = 130

SECTION_STAGE = {
    "1. ТИТУЛЬНЫЙ": (0.1, "Пауза на вход: слайд титула."),
    "2. ПЛАН ЛЕКЦИИ": (0.15, "Показать пункты плана на слайде."),
    "3. ЗАЧЕМ НАМ ОПТИКА? (С РАСЧЁТАМИ НА ДОСКЕ)": (
        3.5,
        "Доска: I = P/S; сценарий 1 (d = 1 мм) и сценарий 2 (d = 50 мкм); ×400.",
    ),
    "4. ДВА ЯЗЫКА ОПТИКИ: ЛУЧИ И ВОЛНЫ": (
        2.5,
        "Доска: слева лучи → «точка»; справа апертура и дифракционное пятно; слово «интерференция» — как сложение волн в фокусе.",
    ),
    "5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)": (
        3.5,
        "Доска: d ~ lambda/NA; NA = n sin theta; примеры 0,1 и 0,4; оговорка: оценка порядка величины.",
    ),
    "6. ОПТИЧЕСКИЙ ТРАКТ СЛП-СТАНКА": (
        2.5,
        "Доска/схема: волокно → коллиматор → гальво → F-θ («эф-тета») → защитное стекло → порошок.",
    ),
    "7. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2) И ТРИ СЛУЧАЯ": (
        6.0,
        "Доска: DOF ~ lambda/NA2; таблица A/B/C; почему NA = 0,5 не для СЛП.",
    ),
    "8. ФОКУС НА ПРАКТИКЕ: СМЕЩЕНИЕ И ЧИСТОТА ОПТИКИ": (
        2.5,
        "Доска: фокус ниже / в слое / выше; z-offset («зет-офсет»); защитное стекло.",
    ),
    "9. КАЧЕСТВО ПУЧКА M² И ПОДВЕДЕНИЕ ИТОГОВ": (
        3.0,
        "Доска: d_реал = M^2 · d_идеал; пять итоговых тезисов.",
    ),
}


def _count_words(text):
    return len(re.findall(r"[A-Za-zА-Яа-яЁё0-9]+(?:[-'][A-Za-zА-Яа-яЁё0-9]+)*", text))


def _fmt_clock(minutes):
    sec = max(0, int(round(minutes * 60)))
    return f"{sec // 60:02d}:{sec % 60:02d}"


def _fmt_duration(minutes):
    sec = max(0, int(round(minutes * 60)))
    if sec < 60:
        return f"{sec} сек"
    half = round(minutes * 2) / 2
    if abs(half - minutes) > 0.05:
        return f"{minutes:.1f}".replace(".", ",") + " мин"
    if half == int(half):
        return f"{int(half)} мин"
    return f"{str(half).replace('.', ',')} мин"


def _time_script(blocks):
    out = []
    cum = 0.0
    total_words = 0
    total_read = 0.0
    total_pause = 0.0
    for title, timing, paragraphs in blocks:
        words = sum(_count_words(p) for p in paragraphs)
        total_words += words
        pause_min, cue = SECTION_STAGE.get(title, (0.0, ""))
        if timing != "auto":
            m = re.search(r"(\d\d:\d\d)\s*[–-]\s*(\d\d:\d\d)", timing)
            wall_min = None
            if m:
                def _to_min(x):
                    mm, ss = x.split(":")
                    return int(mm) + int(ss) / 60.0
                wall_min = max(0.0, float(_to_min(m.group(2)) - _to_min(m.group(1))))
            if wall_min is None:
                wall_min = words / READING_WPM + float(pause_min)
            read_min = words / READING_WPM
            total_read += float(read_min)
            total_pause += max(0.0, float(wall_min) - float(read_min))
            out.append((title, timing, paragraphs, cue))
            cum += float(wall_min)
        else:
            read_min = words / READING_WPM
            wall_min = read_min + pause_min
            total_read += read_min
            total_pause += pause_min
            start, end = cum, cum + wall_min
            label = (
                f"{_fmt_clock(start)} – {_fmt_clock(end)} "
                f"(стенка {_fmt_duration(wall_min)} = чтение {_fmt_duration(read_min)}"
                f" + доска/паузы {_fmt_duration(pause_min)})"
            )
            out.append((title, label, paragraphs, cue))
            cum = end
    return out, total_words, cum, total_read, total_pause


EPILOGUE = (
    "Связь с курсом. В Лекции 2 мы выбрали язык «свет как энергия» и ввели "
    "I (интенсивность, плотность мощности) = P (мощность) / S (площадь пятна). "
    "В Лекции 3 — почему лазер даёт узкий направленный пучок. В Лекции 4 — почему важна длина волны lambda (лямбда). "
    "Сегодня мы закрыли цепочку: длина волны и качество пучка через оптику станка превращаются в размер пятна d, "
    "глубину резкости DOF (Depth of Focus, «глубина фокуса») и в итоге — в устойчивость плавления на слое порошка. "
    "Законы отражения, преломления и тонкой линзы мы уже знаем; здесь они работали как фон, "
    "а на передний план вышли дифракция, числовая апертура NA (Numerical Aperture) и инженерный компромисс СЛП "
    "(селективное лазерное плавление)."
)

SCRIPT = [
    ("1. ТИТУЛЬНЫЙ", "auto", [
        "Добрый день. Курс «Физика лазеров для аддитивных технологий» — лекция 5.",
        "Цепочка до сегодня: уровни и переходы; свет как энергия и плотность мощности "
        "I (интенсивность) = P (мощность) / S (площадь); устройство лазера и роль волоконного источника; "
        "длина волны и выбор лазера под материал.",
        "Тема: формирование лазерного пятна — геометрическая и физическая оптика. "
        "Базу геометрической оптики — отражение, преломление, построение в зеркале, тонкую линзу — мы уже разобрали ранее.",
        "Одной геометрии недостаточно: размер пятна и глубина резкости относятся уже к физической (волновой) оптике — "
        "к дифракции. Именно она объясняет, почему нельзя получить бесконечно малое пятно на практике.",
    ]),
    ("2. ПЛАН ЛЕКЦИИ", "auto", [
        "План на один академический час — ориентир 35–40 минут с работой у доски.",
        "Первое. Коротко вернёмся к I (интенсивности, плотности мощности) = P / S из Лекции 2 — два численных сценария для масштаба величин.",
        "Второе. Два языка описания: лучи и волны. Где работает геометрическая оптика и где проявляется дифракция.",
        "Третье. Формула №1: размер пятна d (диаметр) ≈ lambda (лямбда, длина волны) / NA (Numerical Aperture — числовая апертура).",
        "Четвёртое. Оптический тракт СЛП (селективного лазерного плавления): от волокна до порошка — "
        "F-θ (читается «эф-тета», плоскополевая фокусирующая линза), гальво-зеркала, защитное стекло.",
        "Пятое. Формула №2: глубина резкости DOF (Depth of Focus) ≈ lambda / NA² и три практических случая.",
        "Шестое. Практика: положение фокуса, смещение по оси z (z-offset — «зет-офсет», сдвиг фокуса по высоте), загрязнение оптики.",
        "Седьмое. Формула №3: качество пучка M² (читается «эм-квадрат») и итоги — связь с выбором волоконного лазера из Лекции 3.",
    ]),
    ("3. ЗАЧЕМ НАМ ОПТИКА? (С РАСЧЁТАМИ НА ДОСКЕ)", "auto", [
        "В СЛП (селективном лазерном плавлении) металл плавится не от мощности «вообще», а от плотности мощности — "
        "сколько ватт приходится на каждый квадратный сантиметр пятна. Это мы уже фиксировали в Лекции 2. "
        "Формула: I (интенсивность, плотность мощности) = P (мощность лазера) / S (площадь пятна).",
        "Размерность I (интенсивности): Вт/см2 или Вт/мм2. Оптика не создаёт энергию — она меняет S (площадь).",
        "Сценарий 1. Без фокусировки. d (диаметр пятна) = 1 мм = 0,1 см. "
        "Площадь круга S (площадь) = pi (d/2)^2 = pi (0,05)^2 ≈ 0,00785 см2 ≈ 0,008 см2.",
        "P (мощность) = 200 Вт. Тогда I (интенсивность) = 200 / 0,008 = 25 000 Вт/см2 = 25 кВт/см2. "
        "Металл нагреется, может покраснеть — стабильного расплава для печати обычно нет.",
        "Сценарий 2. Та же мощность P, но пятно d (диаметр) = 50 мкм = 0,005 см, радиус r = 0,0025 см. "
        "S (площадь) = pi r^2 ≈ 3,14 · 6,25·10^-6 ≈ 2·10^-5 см2.",
        "I (интенсивность) = 200 / (2·10^-5) = 10^7 Вт/см2 = 10 МВт/см2. Это примерно в 400 раз выше, чем в первом сценарии.",
        "Вывод для нас, как инженеров: показание мощности на пульте — ещё не всё. "
        "Главный усилитель станка — оптика, которая сжимает энергию в пространстве, уменьшая S (площадь пятна). "
        "Дальше вопрос физики: почему нельзя сжать пятно до нуля? Здесь геометрическая модель «лучи сходятся в точку» перестаёт быть достаточной.",
    ]),
    ("4. ДВА ЯЗЫКА ОПТИКИ: ЛУЧИ И ВОЛНЫ", "auto", [
        "В Лекции 2 мы говорили о корпускулярно-волновом дуализме: иногда удобнее думать о фотонах и энергии, иногда — о волне. "
        "В оптике лазерного станка нужна та же двойственность — уже в инженерной постановке.",
        "Геометрическая оптика описывает свет лучами. Луч распространяется прямолинейно, на границах сред отражается и преломляется. "
        "Этим описанием мы уже пользовались: зеркала, линза, фокус. Для коллиматора "
        "(устройства, делающего пучок почти параллельным), гальво-зеркал "
        "(гальванометрических зеркал сканера) и общей схемы оптического тракта этого достаточно.",
        "Вблизи фокуса геометрическая модель даёт неполный ответ. Она предсказывает математическую точку. "
        "Реальный свет — электромагнитная волна с длиной волны lambda (лямбда, λ). "
        "Волна проходит через конечную апертуру (световой диаметр) линзы не как идеальный луч: у краёв апертуры возникает дифракция — "
        "перераспределение энергии, и в фокусе образуется пятно конечного размера d (диаметра).",
        "Дифракция — не дефект оборудования, а фундаментальный предел. Чем меньшее пятно мы стремимся получить, "
        "тем заметнее проявляется волновая природа света.",
        "С понятием интерференции здесь можно обойтись без сложных формул. В фокальной плоскости волны, пришедшие от разных частей апертуры, "
        "складываются: где фазы согласованы — яркое ядро пятна, где фазы расходятся — спад интенсивности I к краям. "
        "Дифракционная картина в фокусе — результат такого сложения волн.",
        "Почему лазер здесь особенно удобен? В Лекции 3 мы видели: лазерное излучение узкополосное и обладает высокой степенью когерентности — "
        "фазы волн связаны между собой. Поэтому фокусировка предсказуема, а пятно воспроизводимо от слоя к слою. "
        "У теплового источника вроде лампы такой же степени сжатия излучения в пятно добиться нельзя.",
        "Итог раздела: геометрическая оптика описывает путь луча по станку; физическая (волновая) оптика — конечный размер пятна. "
        "Дальше — рабочие формулы.",
    ]),
    ("5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)", "auto", [
        "Минимальный диаметр пятна для идеального пучка оценивают формулой: "
        "d (диаметр пятна) ≈ lambda (лямбда, длина волны) / NA (Numerical Aperture — числовая апертура). "
        "Это оценка порядка величины дифракционного предела — удобная для инженерных оценок, но не замена полного расчёта дифракции.",
        "d — характерный диаметр пятна в фокусе. lambda (λ) — длина волны. Для типичного волоконного лазера СЛП "
        "(селективного лазерного плавления) берём lambda = 1,07 мкм (около 1070 нм — см. шкалу из Лекции 4).",
        "NA (числовая апертура) — безразмерная величина. NA = n · sin(theta), где n — показатель преломления среды между оптикой и деталью "
        "(в воздухе n ≈ 1), theta (тета) — половина угла схождения конуса лучей к фокусу.",
        "Смысл NA (числовой апертуры): насколько круто сходится световой конус. Больше угол theta — больше NA — меньше пятно d. "
        "Широкий пучок на линзе при том же фокусном расстоянии даёт больший theta (тета) и больший NA.",
        "Пример 1. NA = 0,1 — порядок, характерный для многих промышленных F-θ объективов "
        "(читается «эф-тета»; плоскополевая линза, которая фокусирует луч по плоскому полю сканирования). "
        "d ≈ 1,07 / 0,1 = 10,7 мкм. Это идеал при M² (параметре качества пучка, «эм-квадрат») = 1; реальное пятно будет больше.",
        "Пример 2. NA = 0,4. d ≈ 1,07 / 0,4 = 2,7 мкм. Казалось бы, можно взять оптику с большой NA и получить микронное разрешение. "
        "Но размер пятна — только часть задачи. Вторая часть — глубина резкости DOF (Depth of Focus), о которой речь дальше.",
        "Практическое замечание. На станках часто указывают рабочее пятно 40–80 мкм. Это не противоречие формуле d ≈ lambda / NA: "
        "рабочий размер задают конструкцией тракта, диаметром пучка, фокусным расстоянием и фактическим M² (качеством пучка). "
        "Дифракционный предел — нижняя физическая граница возможного, а не обязательная заводская настройка.",
    ]),
    ("6. ОПТИЧЕСКИЙ ТРАКТ СЛП-СТАНКА", "auto", [
        "Перенесём формулы на конструкцию станка. Упрощённый оптический тракт промышленного СЛП "
        "(селективного лазерного плавления) выглядит так.",
        "Волоконный лазер — источник излучения. Излучение идёт по оптическому волокну: это и доставка мощности, "
        "и сохранение хорошего качества пучка M² (малого «эм-квадрат»).",
        "Коллиматор — линза или блок, который превращает расходящийся пучок из волокна в почти параллельный. "
        "Здесь геометрическая оптика применяется напрямую: параллельность нужна, чтобы дальше корректно работали сканеры и "
        "F-θ линза (читается «эф-тета» — плоскополевая фокусирующая линза сканирующей системы).",
        "Гальванометрические зеркала — гальво (от англ. galvo). Два зеркала быстро отклоняют пучок по осям X и Y "
        "(горизонтальным направлениям в плоскости сканирования). Они задают траекторию сканирования, а не сам механизм плавления.",
        "F-θ линза («эф-тета»). Её задача — на большом поле, например 250×250 мм, удерживать пятно по возможности круглым "
        "и обеспечивать предсказуемую связь угла отклонения зеркал со скоростью перемещения пятна по плоскости. "
        "Без такой линзы на краю поля пятно сильно искажается, а I (интенсивность, плотность мощности) становится неоднородной по полю.",
        "Защитное стекло — последний прозрачный элемент перед рабочей камерой. Оно расходное: принимает брызги и пыль. "
        "Загрязнённое стекло поглощает часть мощности P, нагревается и начинает действовать как нежелательная линза: "
        "фокус смещается, пятно d растёт, стабильность процесса падает.",
        "И только затем — слой порошка. Вся цепочка нужна для одного: доставить на поверхность заданные "
        "d (диаметр пятна) и I (интенсивность), стабильно, слой за слоем.",
        "Практический вывод: если процесс стал нестабильным, проверяйте не только мощность и скорость сканирования, "
        "но и оптику — загрязнение, юстировку, положение фокуса.",
    ]),
    ("7. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2) И ТРИ СЛУЧАЯ", "auto", [
        "Глубина резкости DOF (Depth of Focus, буквально «глубина фокуса») — отрезок вдоль оси пучка, "
        "в пределах которого диаметр пятна d ещё близок к минимальному. "
        "Если выйти из этого интервала — пятно увеличивается, I (интенсивность, плотность мощности) падает, проплав становится нестабильным.",
        "Оценка: DOF ≈ lambda (длина волны) / NA^2 (числовая апертура в квадрате). Снова lambda = 1,07 мкм. "
        "Обратите внимание: NA (числовая апертура) стоит в квадрате. "
        "Уменьшили пятно d примерно вдвое за счёт роста NA — DOF (глубина резкости) упала примерно вчетверо. "
        "Это главный оптический компромисс СЛП (селективного лазерного плавления).",
        "Случай А. NA = 0,1. d ≈ 10,7 мкм; DOF ≈ 1,07 / 0,01 = 107 мкм ≈ 0,1 мм. "
        "При фракции порошка 20–40 мкм неровности слоя порядка десятков микрон всё ещё меньше или сравнимы с этим интервалом DOF. "
        "Процесс устойчив к рельефу слоя.",
        "Случай Б. NA = 0,4. d ≈ 2,7 мкм; DOF ≈ 1,07 / 0,16 ≈ 6,7 мкм. "
        "Сдвиг поверхности на несколько микрон уже заметно увеличивает пятно. "
        "Выравнивать порошковый слой с оптической точностью микронного уровня практически нельзя.",
        "Случай В. NA = 0,5. d ≈ 2,1 мкм; DOF ≈ 4,3 мкм. Для промышленного СЛП это нереализуемо: "
        "рельеф порошка 10–20 мкм больше DOF (глубины резкости), плюс поры между частицами и риск для короткофокусной оптики вблизи пыли.",
        "Промышленный выбор: NA ≈ 0,08–0,15. Не потому что нельзя сделать оптику «острее», "
        "а потому что слой порошка — не идеально плоская поверхность. Таблица трёх случаев — на слайде.",
    ]),
    ("8. ФОКУС НА ПРАКТИКЕ: СМЕЩЕНИЕ И ЧИСТОТА ОПТИКИ", "auto", [
        "Формулы для NA (числовой апертуры) и DOF (глубины резкости, Depth of Focus) прямо связаны с ежедневными настройками станка. "
        "Положение фокуса относительно слоя порошка задают смещением по оси z "
        "(ось z — вертикаль, перпендикуляр к плоскости слоя) или, в терминах ПО станка, z-offset "
        "(читается «зет-офсет» — сдвиг фокуса по высоте относительно поверхности).",
        "Фокус в плоскости слоя — базовый случай: минимальное пятно d на поверхности, максимальная "
        "I (интенсивность, плотность мощности) у порошка.",
        "Фокус чуть выше или ниже слоя — пятно на поверхности больше. Иногда это делают намеренно: шире ванна расплава, "
        "иные градиенты температуры, другая стратегия шва. Но это всегда обмен параметрами: меняются d (диаметр) и I (интенсивность). "
        "Без понимания DOF (глубины резкости) произвольное смещение фокуса приводит к плохо воспроизводимому результату.",
        "Если DOF ≈ 0,1 мм, то ошибка позиционирования платформы или тепловое смещение на десятки микрон уже расходует запас по фокусу. "
        "При DOF порядка 5 мкм станок требовал бы контроля высоты с оптической точностью на каждом слое.",
        "Второй важный фактор — загрязнение защитного стекла и оптики. Поглощение → нагрев → термолинза "
        "(нежелательное изменение фокусировки из-за нагрева) → уход фокуса и рост d (диаметра пятна). "
        "Типичные признаки: внезапно растёт пористость, ухудшается проплав, хотя мощность P на пульте не менялась.",
        "Вывод: чистота оптики — часть технологии процесса. Геометрия пятна — такой же параметр, как скорость сканирования и мощность.",
    ]),
    ("9. КАЧЕСТВО ПУЧКА M² И ПОДВЕДЕНИЕ ИТОГОВ", "auto", [
        "Даже идеальная линза не компенсирует плохой пучок. Параметр M² (читается «эм-квадрат», параметр качества пучка) "
        "показывает, насколько реальный пучок хуже идеального гауссова (наилучшего по дифракции). "
        "У идеала M² = 1. У хороших волоконных лазеров часто M² ≈ 1,1–1,5. У многомодовых диодных модулей M² может быть десятки.",
        "Связь с пятном: d_реал (реальный диаметр) ≈ M² · d_идеал, где d_идеал — оценка "
        "lambda (длина волны) / NA (числовая апертура).",
        "Числовой пример. d_идеал = 10 мкм. Волокно, M² = 1,3 → d_реал ≈ 13 мкм. Диод, M² = 20 → d_реал ≈ 200 мкм.",
        "Площадь пятна растёт как квадрат диаметра, поэтому I (интенсивность) падает резко: "
        "увеличение d в 20 раз даёт падение I примерно в 400 раз при той же мощности P. "
        "Отсюда вывод Лекции 3 в новой формулировке: в СЛП нужен пучок с малым M² — волоконный лазер.",
        "Итоговые тезисы.",
        "Первое. I (интенсивность) = P (мощность) / S (площадь) — ключевой параметр процесса; оптика управляет S.",
        "Второе. Геометрическая оптика описывает ход луча; дифракция запрещает нулевое пятно. "
        "d ≈ lambda / NA (диаметр ≈ длина волны / числовая апертура).",
        "Третье. DOF (глубина резкости) ≈ lambda / NA² падает быстрее, чем d. "
        "Малое пятно требует почти идеально ровного слоя — порошок этого не обеспечивает.",
        "Четвёртое. Оптический тракт СЛП — коллиматор, гальво (сканирующие зеркала), "
        "F-θ («эф-тета», плоскополевая линза), защитное стекло — определяет пятно не меньше, чем сама фокусирующая оптика.",
        "Пятое. d_реал ≈ M² · d_идеал. Большое M² («эм-квадрат») сводит на нет преимущества дорогой оптики. "
        "Промышленный компромисс — NA (числовая апертура) около 0,1 и чистая, юстированная оптическая система.",
    ]),
]



def _full_script():
    timed, *_rest = _time_script(SCRIPT)
    return timed


# ── PPTX ──────────────────────────────────────────────────────────────────

def _add_na_table(slide, ACCENT, LT):
    headers = NA_TABLE["headers"]
    rows = NA_TABLE["rows"]
    tbl = slide.shapes.add_table(
        1 + len(rows), len(headers),
        Inches(0.35), Inches(1.35), Inches(12.65), Inches(5.7),
    ).table
    col_w = [Inches(2.8), Inches(2.95), Inches(2.95), Inches(2.95)]
    for ci, w in enumerate(col_w[: len(headers)]):
        tbl.columns[ci].width = w
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = pretty_math(h)
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(13)
            p.font.color.rgb = LT
            p.font.name = "Calibri"
            p.alignment = PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb = ACCENT
    for ri, row in enumerate(rows, start=1):
        bg = RGBColor(0x24, 0x24, 0x3E) if ri % 2 else RGBColor(0x1E, 0x1E, 0x32)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = pretty_math(val)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12 if ci == 0 else 11)
                p.font.color.rgb = LT
                p.font.name = "Calibri"
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
                p.word_wrap = True
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg


def build_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    ACCENT = RGBColor(0x00, 0x96, 0x88)
    ACCENT2 = RGBColor(0xE6, 0x4A, 0x19)
    LT = RGBColor(0xEE, 0xEE, 0xEE)
    GT = RGBColor(0x75, 0x75, 0x75)

    def abg(s):
        f = s.background.fill
        f.solid()
        f.fore_color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

    def _font_for(tx):
        if any(ch in tx for ch in "νλτωθħΔπ"):
            return "Cambria Math"
        return "Calibri"

    def at(s, l, t, w, h, tx, sz=18, b=False, c=LT, a=PP_ALIGN.LEFT):
        tx = pretty_math(tx)
        tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = tx
        p.font.size = Pt(sz)
        p.font.bold = b
        p.font.color.rgb = c
        p.font.name = _font_for(tx)
        p.alignment = a

    def bullets(s, items, left=1.5, top=1.55, w=6.2, sz=17):
        tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(5.5))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            it = pretty_math(it)
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = it
            p.font.size = Pt(sz)
            p.font.color.rgb = LT
            p.font.name = _font_for(it)
            p.space_after = Pt(12 if w >= 10 else 6)

    def hdr(s, title, accent=ACCENT, sz=26):
        abg(s)
        at(s, 1.5, 0.30, 10.3, 0.9, pretty_math(title), sz, True)
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(1.22), Inches(10.3), Pt(3))
        ln.fill.solid()
        ln.fill.fore_color.rgb = accent
        ln.line.fill.background()

    tmpdir = os.path.join(BASE, "_diagrams_tmp_lecture5")
    os.makedirs(tmpdir, exist_ok=True)

    slide_defs = [
        (None, [
            ("hdr", (f"1. {LECTURE_TITLE_SHORT}", 24)),
            ("at", (1.5, 1.9, 10.3, 1.0, "Лекция 5: геометрическая и физическая оптика", 20)),
            ("at", (1.5, 2.9, 10.3, 1.2, "Пятно, дифракция, NA, DOF, тракт СЛП и M^2", 18, False, GT)),
            ("at", (1.5, 4.1, 10.3, 1.2, "Базу геометрической оптики уже знаем — сегодня предел пятна и практика СЛП", 16)),
        ]),
        (None, [
            ("hdr", ("2. ПЛАН ЛЕКЦИИ", 30)),
            ("bullets", ([
                "1. I = P/S",
                "2. Лучи и волны",
                "3. d ~ lambda/NA",
                "4. Тракт СЛП: F-θ, гальво, защита",
                "5. DOF ~ lambda/NA2",
                "6. Фокус: z-offset",
                "7. M^2 и итоги",
            ], 1.5, 1.45, 10.3, 16)),
        ]),
        (draw_intensity_comparison_pil, [
            ("hdr", ("3. ЗАЧЕМ НАМ ОПТИКА?", 26), ACCENT2),
            ("bullets", ([
                "I = P / S — плотность мощности",
                "d = 1 мм -> I ~ 25 кВт/см2",
                "d = 50 мкм -> I ~ 10 МВт/см2",
                "x400 при той же мощности 200 Вт",
                "Оптика сжимает энергию в пространстве",
            ], 1.5, 1.45, 5.4, 14)),
        ]),
        (draw_rays_vs_waves_pil, [
            ("hdr", ("4. ЛУЧИ И ВОЛНЫ", 26), ACCENT2),
            ("bullets", ([
                "Геометрия: лучи, зеркала, линзы",
                "У фокуса — дифракция",
                "Интерференция волн в апертуре",
                "Лазер: когерентность → предсказуемый фокус",
                "Точки в фокусе нет — есть пятно d",
            ], 1.5, 1.45, 5.2, 13)),
        ]),
        (draw_spot_size_pil, [
            ("hdr", ("5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)", 22), ACCENT2),
            ("bullets", ([
                "d ~ lambda / NA",
                "NA = n sin(theta)",
                "NA=0,1 -> d~10,7 мкм",
                "NA=0,4 -> d~2,7 мкм",
                "Рабочее пятно станка часто больше предела",
            ], 1.5, 1.45, 5.4, 13)),
        ]),
        (draw_optical_train_pil, [
            ("hdr", ("6. ОПТИЧЕСКИЙ ТРАКТ СЛП", 24), ACCENT2),
            ("bullets", ([
                "Волокно → коллиматор",
                "Гальво",
                "F-θ",
                "Защитное стекло",
                "Загрязнение → сдвиг фокуса",
            ], 1.5, 1.45, 5.0, 14)),
        ]),
        ("table", [
            ("hdr", ("7. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2)", 22), ACCENT2),
        ]),
        (draw_focus_practice_pil, [
            ("hdr", ("8. ФОКУС НА ПРАКТИКЕ", 26), ACCENT2),
            ("bullets", ([
                "Фокус: ниже / в слое / выше",
                "z-offset → d, I",
                "DOF",
                "Загрязнение → термолинза",
                "Чистота оптики = параметр процесса",
            ], 1.5, 1.45, 5.2, 13)),
        ]),
        (draw_m2_comparison_pil, [
            ("hdr", ("9. M^2 И ИТОГИ", 28), ACCENT2),
            ("bullets", ([
                "d_реал = M^2 * d_идеал",
                "Волокно M^2~1,3; диод M^2~20",
                "I ~ 1/d^2 — качество пучка критично",
                "NA~0,1 — компромисс для СЛП",
            ], 1.5, 1.45, 5.4, 13)),
        ]),
    ]

    for i, entry in enumerate(slide_defs):
        draw_fn = entry[0]
        actions = entry[1]
        pic = {"left": 6.8, "top": 1.45, "width": 5.9}
        s = prs.slides.add_slide(prs.slide_layouts[6])
        for act in actions:
            kind = act[0]
            if kind == "hdr":
                args = act[1]
                accent = act[2] if len(act) > 2 else ACCENT
                hdr(s, args[0], accent, args[1] if len(args) > 1 else 26)
            elif kind == "at":
                at(s, *act[1])
            elif kind == "bullets":
                a = act[1]
                bullets(s, a[0], a[1], a[2], a[3], a[4])
        if draw_fn == "table":
            _add_na_table(s, ACCENT, LT)
        elif draw_fn:
            png = save_pil_image(draw_fn, os.path.join(tmpdir, f"diag_{i}.png"))
            s.shapes.add_picture(
                png, Inches(pic["left"]), Inches(pic["top"]), width=Inches(pic["width"]),
            )

    prs.save(PPTX_PATH)
    print(f"OK: {PPTX_PATH}")


def build_pdf():
    doc = SimpleDocTemplate(
        PDF_PATH, pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=1.3 * cm, bottomMargin=1.3 * cm,
    )
    styles = getSampleStyleSheet()
    sh2 = ParagraphStyle("H2", parent=styles["Heading2"], fontName=PDF_FONT_BOLD, fontSize=13, leading=17)
    sh3 = ParagraphStyle("H3", parent=styles["Heading3"], fontName=PDF_FONT_BOLD, fontSize=11, leading=14)
    body = ParagraphStyle("B", parent=styles["Normal"], fontName=PDF_FONT, fontSize=10, leading=13, alignment=TA_JUSTIFY)
    hint = ParagraphStyle("H", parent=styles["Normal"], fontName=PDF_FONT, fontSize=9, leading=12, textColor=darkgray, alignment=TA_JUSTIFY)
    stage = ParagraphStyle(
        "Stage", parent=styles["Normal"], fontName=PDF_FONT, fontSize=9,
        leading=12, textColor=darkgray, alignment=TA_JUSTIFY, leftIndent=6,
    )

    pdf_tmp = os.path.join(BASE, "_diagrams_tmp_lecture5_pdf")
    os.makedirs(pdf_tmp, exist_ok=True)

    diagram_for_slide = {
        "3. ЗАЧЕМ НАМ ОПТИКА? (С РАСЧЁТАМИ НА ДОСКЕ)": draw_intensity_comparison_pil,
        "4. ДВА ЯЗЫКА ОПТИКИ: ЛУЧИ И ВОЛНЫ": draw_rays_vs_waves_pil,
        "5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)": draw_spot_size_pil,
        "6. ОПТИЧЕСКИЙ ТРАКТ СЛП-СТАНКА": draw_optical_train_pil,
        "7. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2) И ТРИ СЛУЧАЯ": draw_dof_chart_pil,
        "8. ФОКУС НА ПРАКТИКЕ: СМЕЩЕНИЕ И ЧИСТОТА ОПТИКИ": draw_focus_practice_pil,
        "9. КАЧЕСТВО ПУЧКА M² И ПОДВЕДЕНИЕ ИТОГОВ": draw_m2_comparison_pil,
    }

    script = _full_script()
    total_words = sum(sum(_count_words(p) for p in paras) for _, _, paras, _ in script)
    _, _, wall, read, pause = _time_script(SCRIPT)

    elements = []
    elements.append(Paragraph(pdf_sub(f"ЛЕКЦИЯ 5: {LECTURE_TITLE_FULL.upper()}"), sh2))
    elements.append(Paragraph(
        f"Раскадровка • (9 слайдов) • ~{total_words} слов • "
        f"стенка ~{_fmt_duration(wall)} (чтение ~{_fmt_duration(read)} + доска ~{_fmt_duration(pause)}) • "
        f"{datetime.now().strftime('%d.%m.%Y')}",
        hint,
    ))
    elements.append(Paragraph(
        "Тайминг «стенка» = произнесение текста (~130 слов/мин) + работа у доски/со схемой. "
        "Строки «Доска / пауза:» — ремарки лектору, не читаются вслух. Целевой хронометраж: 35–40 минут.",
        hint,
    ))
    elements.append(Spacer(1, 8))

    for si, (title, timing, paragraphs, cue) in enumerate(script):
        block = []
        block.append(Paragraph(pdf_sub(f"СЛАЙД: {title} | {timing}"), sh3))
        if cue:
            block.append(Paragraph(pdf_sub(f"Доска / пауза: {cue}"), stage))
        for para in paragraphs:
            block.append(Paragraph(pdf_sub(para), body))
        block.append(Spacer(1, 5))

        draw_fn = diagram_for_slide.get(title)
        if draw_fn:
            img = draw_fn()
            safe = re.sub(r"[^\w\-]+", "_", title)[:40]
            png = os.path.join(pdf_tmp, f"pdf_{safe}.png")
            img.save(png, "PNG")
            w_px, h_px = img.size
            img_w = 16.5 * cm if w_px >= 900 else 13 * cm
            block.append(RLImage(png, width=img_w, height=img_w * (h_px / w_px)))
            block.append(Spacer(1, 6))

        if si == len(script) - 1:
            block.append(Paragraph(pdf_sub(EPILOGUE), body))
            elements.append(KeepTogether(block))
        else:
            elements.extend(block)

    def _page_num(canvas, _doc):
        canvas.saveState()
        canvas.setFont(PDF_FONT, 10)
        canvas.setFillColor(gray)
        canvas.drawCentredString(A4[0] / 2, 0.85 * cm, str(canvas.getPageNumber()))
        canvas.restoreState()

    doc.build(elements, onFirstPage=_page_num, onLaterPages=_page_num)
    print(f"OK: {PDF_PATH} (~{total_words} слов, стенка ~{_fmt_duration(wall)})")


def build_docx():
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn

    script = _full_script()
    total_words = sum(sum(_count_words(p) for p in paras) for _, _, paras, _ in script)
    _, _, wall, read, pause = _time_script(SCRIPT)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(1.5)
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    style.font.size = Pt(11)

    def add(text, *, bold=False, size=11, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"
        return p

    add(f"ЛЕКЦИЯ 5: {LECTURE_TITLE_FULL.upper()}", bold=True, size=14, space_after=8)
    add(
        f"Раскадровка для лектора • ~{total_words} слов • "
        f"стенка ~{_fmt_duration(wall)} (чтение ~{_fmt_duration(read)} + доска ~{_fmt_duration(pause)}) • "
        f"{datetime.now().strftime('%d.%m.%Y')}",
        size=10, space_after=10,
    )
    add(
        "Тайминг «стенка» = произнесение текста (~130 слов/мин) + работа у доски/со схемой. "
        "Строки «Доска / пауза:» — ремарки лектору, не читаются вслух. Целевой хронометраж: 35–40 минут.",
        size=9, space_after=12,
    )

    for title, timing, paragraphs, cue in script:
        add(f"СЛАЙД: {pretty_math(title)}", bold=True, size=12, space_after=2)
        add(timing, size=10, space_after=4)
        if cue:
            add(f"Доска / пауза: {pretty_math(cue)}", size=10, space_after=6)
        for para in paragraphs:
            add(pretty_math(para), size=11, space_after=6)

    add(pretty_math(EPILOGUE), size=11, space_after=6)

    doc.save(DOCX_PATH)
    print(f"OK: {DOCX_PATH}")


if __name__ == "__main__":
    print("=" * 70)
    print("СБОРКА ЛЕКЦИИ 5")
    print("=" * 70)
    # preview timing
    timed, words, wall, read, pause = _time_script(SCRIPT)
    print(f"Слов: {words}; стенка: {_fmt_duration(wall)}; чтение: {_fmt_duration(read)}; доска: {_fmt_duration(pause)}")
    for title, timing, *_ in timed:
        print(f"  {title}: {timing}")
    print("=" * 70)
    build_pptx()
    build_pdf()
    build_docx()
    print("=" * 70)
    print("ГОТОВО")
    print("=" * 70)
