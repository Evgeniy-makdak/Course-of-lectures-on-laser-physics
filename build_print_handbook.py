# -*- coding: utf-8 -*-
"""
Сборка печатной методички по курсу.
Порядок глав: бывшая Лекция 3 → 1; Лекция 1 → 2; Лекция 2 → 3; Лекция 4 → 4;
Лекция 4-доп → 5; Лекция 5 → 6.
Текст = раскадровки без спикерских ремарок; со схемами, рисунками и таблицами.
"""
from __future__ import annotations

import ast
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, gray, black, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image as RLImage,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from _prepare_assets import ROOT, lecture4dop_dir, prepare_all

OUT_DIR = Path(__file__).resolve().parent
DOCX_PATH = OUT_DIR / "Методичка_Физика_лазеров_для_3D-технологий.docx"
PDF_PATH = OUT_DIR / "Методичка_Физика_лазеров_для_3D-технологий.pdf"
ASSETS = OUT_DIR / "assets"
# Распространяемая версия — папка «Методичка» рядом с исходниками лекций (Windows)
# или копия в корне репозитория (если исходников нет)
if (ROOT / "Lecture-1-main").exists():
    STUDENT_DIR = ROOT / "Методичка"
else:
    STUDENT_DIR = OUT_DIR
STUDENT_PDF_PATH = STUDENT_DIR / "Методичка_Физика_лазеров_для_3D-технологий.pdf"
STUDENT_DOCX_PATH = STUDENT_DIR / "Методичка_Физика_лазеров_для_3D-технологий.docx"
# ASCII-имя на Рабочем столе — удобно открывать и копировать в Яндекс Документы
DESKTOP_STUDENT_DOCX = Path.home() / "Desktop" / "Metodichka_lectures_1-6.docx"
# Копия в корне репозитория (для GitHub / Яндекс Документы)
REPO_STUDENT_DOCX = OUT_DIR / "Metodichka_lectures_1-6.docx"
AUTHOR = "Волков Е.В."

PREFACE = [
    (
        "Настоящая методичка собрана по материалам семинаров, которые автор читает для сотрудников АО "
        "«Лазерные Системы» в рамках курса по физике лазеров для практического применения в аддитивных "
        "(3D) технологиях в 2026 году."
    ),
    (
        "Текст глав соответствует содержанию лекций: сохранены определения, формулы, численные оценки, "
        "примеры и логика изложения."
    ),
    (
        "Методичка предназначена для самостоятельной работы студентов и инженеров аддитивных "
        "технологий как печатное сопровождение курса и будет дополняться после каждой новой прочитанной лекции."
    ),
]

# Дополнение к глоссарию из лекции 4 (термины, отсутствующие в раскадровке лекции 1)
GLOSSARY_CH4 = [
    "A(λ)\tСпектральная поглощательная способность — безразмерная доля падающей мощности, вошедшей в материал (0…1); A(λ) = P_погл / P_пад. Не путать с работой выхода металла",
    "СЛП\tСелективное лазерное плавление — послойное сплавление металлического порошка лазером (см. SLM)",
    "P/v\tЛинейная энергия — отношение мощности лазера к скорости сканирования; основной параметр режима СЛП",
    "M²\tПараметр качества пучка; M² ≈ 1 соответствует идеальной гауссовой моде TEM₀₀",
    "Фотоинициатор\tДобавка в фотополимер, поглощающая УФ и распадающаяся на радикалы, запускающие полимеризацию",
    "E_акт\tЭнергия активации — минимальный энергетический барьер химической реакции",
    "E (энергия фотона)\tE = h·ν = h·c/λ; в электрон-вольтах: E(эВ) ≈ 1240 / λ(нм)",
    "CO₂-лазер\tГазовый лазер с λ ≈ 10,6 мкм; эффективен для органики и ряда неметаллов",
    "Доза\tЭнергия излучения на единицу площади (экспозиция); ключевой параметр стереолитографии",
    "T_пл\tТемпература плавления материала",
    "Фотохимический механизм\tВзаимодействие, при котором один фотон достаточен для химического акта (E_ph ≥ порога)",
    "Тепловой механизм\tНакопление энергии многих фотонов в нагрев; типичен для СЛП металлов в ближнем ИК",
    "Стереолитография\tПослойная полимеризация жидкой смолы УФ-лазером (см. SLA)",
    "Режим глубокого проплавления\tПерегретая ванна расплава с удлинённым каналом; риск пор при избыточной плотности мощности",
    "Ближний ИК\tИК-диапазон около 1 мкм (волоконные лазеры ~1070 нм); основное окно СЛП металлов",
    "Дальний ИК\tИК-диапазон десятков микрометров (CO₂ ~10,6 мкм); эффективен для многих неметаллов",
]

# Дополнение к глоссарию из лекции 4-доп (глава 5)
GLOSSARY_CH5 = [
    "SHG\tГенерация второй гармоники (Second Harmonic Generation): 1070 нм (ω) → 535 нм (2ω) в нелинейном кристалле",
    "PPLN\tПериодически поляризованный ниобат лития (Periodically Poled Lithium Niobate); знак χ⁽²⁾ меняется каждые Λ",
    "«Зелёная щель»\tGreen gap: падение эффективности InGaN-диодов в зелёном диапазоне из-за дефектов решётки и QCSE",
    "InGaN\tНитрид индия-галлия — полупроводниковая система диодных лазеров видимого диапазона",
    "QCSE\tКвантово-размерный эффект Штарка: пьезополе в квантовой яме разводит электрон и дырку",
    "V-дефекты\tДефекты кристаллической решётки InGaN при высокой доле In; ловушки, переводящие энергию в тепло",
    "χ⁽¹⁾, χ⁽²⁾, χ⁽³⁾\tЛинейная, квадратичная и кубическая восприимчивости в разложении поляризации P(E)",
    "ε₀\tЭлектрическая постоянная ≈ 8,85·10⁻¹² Ф/м; Ф (фарад) = Кл/В",
    "ω, 2ω\tЦиклическая частота накачки и второй гармоники; ω = 2π·c/λ",
    "Λ\tПериод решётки переполяризации PPLN (прописная лямбда); для 1070→535 нм Λ ≈ 6,96 мкм",
    "Δk\tРассогласование волновых чисел: Δk = k₂ω − 2k_ω; условие синхронизма Δk = 2π/Λ",
    "sinc²(x)\tКвадрат sinc(x) = sin(x)/x; описывает падение эффективности SHG при фазовом рассогласовании",
    "η\tЭффективность преобразования SHG: η = P₂ω / P_ω",
    "d_eff\tЭффективный коэффициент нелинейности; для MgO:PPLN d₃₃ ≈ 25 пм/В",
    "Квазифазовый синхронизм\tКомпенсация набега фазы за счёт периодического переворота знака χ⁽²⁾ в PPLN",
    "Оптическое выпрямление\tПостоянная составляющая P⁽²⁾(t) (~ 1/2): поле не излучает, для SHG бесполезно",
    "dn/dT\tТермооптический коэффициент: изменение показателя преломления с температурой",
    "Плавленый кварц\tFused silica; dn/dT ≈ +9·10⁻⁶ К⁻¹, прозрачен на 1070 и 535 нм; материал клина термокомпенсации",
    "Двухпроходная схема\tSHG при прямом и обратном проходе через кристалл; удваивает длину взаимодействия",
]

# Дополнение к глоссарию из лекции 5 (геометрическая оптика) — глава 6
GLOSSARY_CH6 = [
    "I\tПлотность мощности (интенсивность): I = P/S, Вт/см²",
    "P\tМощность лазера, Вт",
    "S\tПлощадь лазерного пятна, см²",
    "d\tДиаметр пятна в фокусе, мкм",
    "d_идеал\tИдеальный (дифракционный) диаметр пятна: d ≈ λ/NA",
    "d_реал\tРеальный диаметр пятна с учётом качества пучка: d_реал = M²·d_идеал",
    "NA\tЧисловая апертура — мера ширины конуса сходящихся лучей; NA = n·sin θ",
    "DOF\tГлубина резкости (Depth of Focus) — диапазон вдоль оси луча, где d ≈ const",
    "n\tПоказатель преломления среды: n = c/v",
    "Δn\tРазность показателей преломления на границе двух сред",
    "θ\tПоловина угла схождения конуса лучей в фокусе",
    "F-theta\tТелецентрическая (эф-тета) линза: круглое пятно по всему полю сканирования",
]

# Повторяющиеся подсказки спикера в скобках — убираем в методичке (определение даётся один раз)
_HANDBOOK_GLOSS_RE = re.compile(
    r"\s*\("
    r"(?:"
    r"интенсивность(?:,\s*плотность мощности)?|интенсивности|"
    r"диаметр(?:\s+пятна)?|реальный диаметр|идеальный диаметр|"
    r"мощность|площадь|"
    r"глубина резкости|глубины резкости|"
    r"числов(?:ая|ой|ую)\s+аперт(?:ура|уре|урой|уру)(?:\s+в квадрате)?|"
    r"длина волны|"
    r"параметр качества пучка|параметр «эм-квадрат»|"
    r"показателем преломления|"
    r"показатель преломления первой среды|показатель преломления второй среды"
    r")\)",
    re.I,
)




def dedupe_handbook_terminology(sections: list[tuple[str, list[str]]]) -> list[tuple[str, list[str]]]:
    """Оставить 1–2 пояснения термина в главе; дальше — только символ."""
    na_long = 0
    out = []
    for title, paras in sections:
        new_paras = []
        for para in paras:
            p = para
            if re.search(r"NA\s*—\s*числовая апертура", p, re.I):
                na_long += 1
                if na_long > 2:
                    p = re.sub(r"NA\s*—\s*числовая апертура", "NA", p, flags=re.I)
            p = re.sub(
                r"обратно пропорционален\s+числов(?:ой|ую)\s+аперт(?:уре|уру)",
                "обратно пропорционален NA",
                p,
                flags=re.I,
            )
            p = re.sub(
                r"NA\s*\([^)]*числов[^)]*аперт[^)]*\)",
                "NA",
                p,
                flags=re.I,
            )
            p = re.sub(
                r"выбирает NA\s*\([^)]*числов[^)]*\)",
                "выбирает NA",
                p,
                flags=re.I,
            )
            new_paras.append(re.sub(r"\s+", " ", p).strip())
        out.append((title, new_paras))
    return out

def strip_handbook_glosses(text: str) -> str:
    """Убрать повторные расшифровки переменных в скобках (для печатной методички)."""
    t = _HANDBOOK_GLOSS_RE.sub("", text)
    t = re.sub(r"\s+", " ", t).strip()
    return t


# ── fonts / helpers ────────────────────────────────────────────────────────

def os_environ_windir() -> str:
    import os
    return os.environ.get("WINDIR", r"C:\Windows")


def _register_fonts():
    import os
    windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    candidates = [
        (
            windir / "Fonts" / "times.ttf",
            windir / "Fonts" / "timesbd.ttf",
            windir / "Fonts" / "timesi.ttf",
        ),
        (
            Path("/System/Library/Fonts/Supplemental/Times New Roman.ttf"),
            Path("/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf"),
            Path("/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf"),
        ),
        (
            windir / "Fonts" / "arial.ttf",
            windir / "Fonts" / "arialbd.ttf",
            windir / "Fonts" / "ariali.ttf",
        ),
        (
            Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
            Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
            Path("/System/Library/Fonts/Supplemental/Arial Italic.ttf"),
        ),
    ]
    regular = bold = italic = None
    for reg, bd, it in candidates:
        if reg.exists():
            regular, bold, italic = reg, bd, it
            break
    if regular is None:
        raise FileNotFoundError("Не найден Times New Roman / Arial для PDF")
    pdfmetrics.registerFont(TTFont("Book", str(regular)))
    pdfmetrics.registerFont(TTFont("BookBold", str(bold if bold.exists() else regular)))
    # ∝, → и др. нет в Times New Roman — отдельный шрифт
    sym = None
    for cand in (
        Path("/Library/Fonts/Arial Unicode.ttf"),
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
        Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
        Path("/System/Library/Fonts/Apple Symbols.ttf"),
        windir / "Fonts" / "seguisym.ttf",
        windir / "Fonts" / "arialuni.ttf",
        windir / "Fonts" / "arial.ttf",
    ):
        if cand.exists():
            try:
                pdfmetrics.registerFont(TTFont("BookSym", str(cand)))
                sym = "BookSym"
                break
            except Exception:
                continue
    if italic.exists():
        pdfmetrics.registerFont(TTFont("BookItalic", str(italic)))
        return "Book", "BookBold", "BookItalic", sym
    return "Book", "BookBold", "Book", sym


SPEAKER_LINE_RE = re.compile(
    r"^(?:ТАЙМИНГ\s*:|ДОСКА\s*/\s*ПАУЗА\s*:|□\s*Общее время|Общее время секции|"
    r"Длительность\s*:|Слайдов\s*:|Целевая аудитория\s*:|Файл презентации\s*:|"
    r"---\s*PAGE\s+\d+\s*---|---\s*КОНЕЦ|Страниц\s*:|Полная раскадровка|ПОЛНАЯ РАСКАДРОВКА)",
    re.I,
)
INLINE_TIMING_RE = re.compile(
    r"\s*[\|–-]?\s*\d{1,2}:\d{2}\s*[–-]\s*\d{1,2}:\d{2}(?:\s*\([^)]*\))?"
)
INLINE_HOLD_RE = re.compile(
    r"\s*\((?:держите на экране[^)]*|переключите[^)]*|≈\s*[^)]*минут[^)]*|"
    r"~\d+\s*мин[^)]*|быстро пройдите[^)]*)\)",
    re.I,
)
BOARD_CLOSING_RE = re.compile(
    r"^(?:Сейчас на доске|На доске)\b.*",
    re.I,
)

_SUP_MAP = str.maketrans("0123456789+-=()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾")
_SUB_MAP = str.maketrans("0123456789+-=()aehijklmnoprstuvx",
                         "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ")


_SUP_LETTERS = str.maketrans({
    "n": "ⁿ", "N": "ⁿ", "i": "ⁱ", "x": "ˣ", "y": "ʸ", "z": "ᶻ",
    "a": "ᵃ", "b": "ᵇ", "c": "ᶜ", "d": "ᵈ", "e": "ᵉ", "h": "ʰ",
    "j": "ʲ", "k": "ᵏ", "l": "ˡ", "m": "ᵐ", "o": "ᵒ", "p": "ᵖ",
    "r": "ʳ", "s": "ˢ", "t": "ᵗ", "u": "ᵘ", "v": "ᵛ", "w": "ʷ",
})


def _to_super(s: str) -> str:
    s = s.replace("−", "-").replace("–", "-")
    return s.translate(_SUP_MAP).translate(_SUP_LETTERS)


def _to_sub(s: str) -> str:
    s = s.replace("−", "-").replace("–", "-")
    return "".join(
        (ch.lower().translate(_SUB_MAP) if ch.lower() in "aehijklmnoprstuvx" else
         ch.translate(_SUB_MAP) if ch in "0123456789+-=()" else ch)
        for ch in s
    )


def _tex_markup_to_unicode(t: str) -> str:
    """Мини-разметка Лекции 4-доп: ^{...} степень, _{...} индекс."""
    def sup(m):
        return _to_super(m.group(1).replace("−", "-"))

    def sub(m):
        s = m.group(1).replace("−", "-")
        if re.fullmatch(r"[0-9+\-=()]+", s):
            return _to_sub(s)
        return "_" + s

    t = re.sub(r"\^\{([^{}]+)\}", sup, t)
    t = re.sub(r"_\{([^{}]+)\}", sub, t)
    return t


def typography_fix(text: str) -> str:
    """Normalize powers, indices, Greek nu, strip raw HTML sub/sup tags."""
    if not text:
        return text
    t = text.replace("\xa0", " ")

    # HTML tags -> unicode
    t = re.sub(r"<sup>\s*([^<]*?)\s*</sup>", lambda m: _to_super(m.group(1).replace("−", "-")), t, flags=re.I)
    t = re.sub(r"<sub>\s*([^<]*?)\s*</sub>", lambda m: _to_sub(m.group(1).replace("−", "-")), t, flags=re.I)
    # leftover angle-bracket artifacts
    t = t.replace("<sup>", "").replace("</sup>", "").replace("<sub>", "").replace("</sub>", "")
    t = _tex_markup_to_unicode(t)

    # Summation formula (after tag conversion or raw HTML leftovers)
    t = re.sub(
        r"Σ(?:ₗ₌₀|ₗ=₀)?(?:ⁿ⁻¹)?\s*\(2l\+1\)",
        "Σₗ₌₀ⁿ⁻¹ (2l+1)",
        t,
    )

    # 10^-15 / 10^−15 / 10^{ -15 } and 6,25·10^-6
    t = re.sub(r"(\d+(?:,\d+)?)·10\^\s*\{?\s*(−|-)?\s*(\d+)\s*\}?",
               lambda m: m.group(1) + "·10" + _to_super(("-" if m.group(2) else "") + m.group(3)), t)
    t = re.sub(r"10\^\s*\{?\s*(−|-)?\s*(\d+)\s*\}?",
               lambda m: "10" + _to_super(("-" if m.group(1) else "") + m.group(2)), t)
    # (d/2)^2, (0,05)^2, (0,0025)^2
    t = re.sub(r"\(([^()]+)\)\^\s*\{?\s*(−|-)?\s*(\d+)\s*\}?",
               lambda m: "(" + m.group(1) + ")" + _to_super(("-" if m.group(2) else "") + m.group(3)), t)
    # letter^number (n^2, r^2, NA^2, …)
    t = re.sub(r"([A-Za-z])\^\s*\{?\s*(−|-)?\s*(\d+)\s*\}?",
               lambda m: m.group(1) + _to_super(("-" if m.group(2) else "") + m.group(3)), t)
    # any remaining ^N (fallback)
    t = re.sub(r"\^\s*\{?\s*(−|-)?\s*(\d+)\s*\}?",
               lambda m: _to_super(("-" if m.group(1) else "") + m.group(2)), t)
    t = re.sub(r"\b2n2\b", "2n²", t)
    t = re.sub(r"\b2N2\b", "2N²", t)

    # Energy/population indices: E1..E3, N1/N2, and underscore forms
    t = re.sub(r"\bE_([0-3])\b", lambda m: "E" + _to_sub(m.group(1)), t)
    t = re.sub(r"\bN_([12])\b", lambda m: "N" + _to_sub(m.group(1)), t)
    t = re.sub(r"\bE([0-3])\b", lambda m: "E" + _to_sub(m.group(1)), t)
    t = re.sub(r"\bN([12])\b", lambda m: "N" + _to_sub(m.group(1)), t)
    t = t.replace("m_l", "mₗ").replace("m_s", "mₛ").replace("N_max", "Nₘₐₓ")
    # Intensity / Bouguer forms
    t = t.replace("I_x = I_0 · e^(−α · x)", "I(x) = I₀ · exp(−α · x)")
    t = t.replace("I_x = I_0 · e^(-α · x)", "I(x) = I₀ · exp(−α · x)")
    t = re.sub(r"\bI_([0x])\b", lambda m: "I" + _to_sub(m.group(1)), t)
    t = re.sub(r"\bI0\b", "I₀", t)
    t = re.sub(
        r"\be\^\s*\(([^)]*)\)",
        lambda m: "exp(" + m.group(1).strip() + ")",
        t,
    )
    t = re.sub(r"\{MOON_SPOT_KM:[^}]+\}", "3,8", t)
    t = re.sub(r"\{MOON_R_KM[^}]*\}", "384400", t)
    # Lecture 4-доп: смешанные индексы 2ω после _{2ω} → _2ω
    t = t.replace("P_2ω", "P₂ω").replace("k_2ω", "k₂ω").replace("n_2ω", "n₂ω")
    t = t.replace("λ_2ω", "λ₂ω").replace("d_33", "d₃₃")

    # Greek nu and multiplication
    t = t.replace("h*nu", "h·ν").replace("h·nu", "h·ν").replace("h ν", "h·ν")
    t = t.replace("h*ν", "h·ν")
    t = re.sub(r"(?<![A-Za-zА-Яа-я])nu(?![A-Za-zА-Яа-я])", "ν", t)

    # Lecture 5: pi, lambda, Delta n, M2, unit powers, angles
    t = re.sub(r"(?<![A-Za-zА-Яа-я])pi(?![A-Za-zА-Яа-я])", "π", t)
    t = re.sub(r"(?<![A-Za-zА-Яа-я])lambda(?![A-Za-zА-Яа-я])", "λ", t, flags=re.I)
    t = re.sub(r"\btheta(\d+)\b", lambda m: "θ" + _to_sub(m.group(1)), t, flags=re.I)
    t = re.sub(r"(?<![A-Za-z-])theta\b", "θ", t, flags=re.I)
    t = re.sub(r"\bn([12])\b", lambda m: "n" + _to_sub(m.group(1)), t)
    t = re.sub(r"\bDelta\s+n\b", "Δn", t)
    t = re.sub(r"\bM2\b", "M²", t)
    t = re.sub(r"\bNA2\b", "NA²", t)
    t = re.sub(r"см2\b", "см²", t)
    t = re.sub(r"мм2\b", "мм²", t)
    t = re.sub(r"/см2\b", "/см²", t)

    # пробелы после верхних индексов перед знаками и единицами
    t = re.sub(r"([²³⁴⁵⁶⁷⁸⁹⁰⁻⁺⁼⁽⁾ⁿ]+)([=≈])", r"\1 \2", t)
    t = re.sub(r"([²³⁴⁵⁶⁷⁸⁹⁰⁻⁺])([А-Яа-я])", r"\1 \2", t)

    # Delta E written as dE in formulas
    t = re.sub(r"(?<![A-Za-z])dE(?![A-Za-z])", "ΔE", t)
    t = t.replace("Delta E", "ΔE").replace("DeltaE", "ΔE")

    # normalize arrows / dots occasionally left as ASCII
    t = t.replace("<->", "↔").replace("->", "→")
    return t


def clean_heading(text: str) -> str:
    t = text.strip()
    t = re.sub(r"^СЛАЙД\s*:?\s*", "", t, flags=re.I)
    t = re.sub(r"^СЛАЙД\s+\d+\s*[-–—.:)]\s*", "", t, flags=re.I)
    t = INLINE_TIMING_RE.sub("", t)
    t = INLINE_HOLD_RE.sub("", t)
    # Ремарки раскадровки в заголовках (на доске, для лектора) — не для методички
    t = re.sub(r"\s*\([^)]*на\s+доске[^)]*\)", "", t, flags=re.I)
    # drop old numbering variants: "2.", "2 -", "2 —", "13 - ИТОГИ"
    t = re.sub(r"^\d+\s*[-–—.:)]\s*", "", t)
    t = re.sub(r"\s+", " ", t).strip(" |–-")
    return t


def clean_paragraph(text: str) -> str | None:
    t = text.strip()
    if not t or SPEAKER_LINE_RE.search(t) or re.fullmatch(r"\d{1,3}", t) or t.startswith("---"):
        return None
    if BOARD_CLOSING_RE.match(t):
        if t.lower().startswith("посмотрите на этот график"):
            t = re.sub(
                r"^Посмотрите на этот график\.\s*",
                "Рассмотрим график пороговой плотности энергии. ",
                t,
                flags=re.I,
            )
        else:
            return None
    t = INLINE_HOLD_RE.sub("", t)
    t = re.sub(r"\s+", " ", t).strip()
    reps = [
        (r"^Добрый день, коллеги\.\s*", ""),
        (r"^Добрый день\.\s*", ""),
        (r"^Поехали\.\s*", ""),
        (r"Проверьте сейчас:", "Проверим:"),
        (r"Давайте сделаем расчёт вместе", "Сделаем расчёт"),
        (r"Давайте вспомним", "Напомним"),
        (r"давайте вспомним", "напомним"),
        (r"Давайте посчитаем", "Посчитаем"),
        (r"давайте посчитаем", "посчитаем"),
        (r"Давайте посмотрим", "Рассмотрим"),
        (r"давайте посмотрим", "рассмотрим"),
        (r"Запишите гипотезу", "Сформулируйте гипотезу"),
        (r"Если есть вопросы по сегодняшнему материалу[^.]*\.\s*", ""),
        (r"Если вопросов нет,[^.]*\.\s*", ""),
        (r"короткий молчаливый повтор якорей на доске:[^.]*\.\s*", ""),
    ]
    for pat, rep in reps:
        t = re.sub(pat, rep, t)
    t = typography_fix(t.strip())
    return t or None


def load_script(py_path: Path):
    src = py_path.read_text(encoding="utf-8")
    tree = ast.parse(src)
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name) and t.id == "SCRIPT":
                    return ast.literal_eval(node.value)
    raise RuntimeError(f"SCRIPT not found in {py_path}")


def paragraphs_from_script_items(script) -> list[tuple[str, list[str]]]:
    sections = []
    for item in script:
        title = clean_heading(item[0])
        if "БОНУС" in title.upper():
            continue
        paras = [cp for p in item[2] if (cp := clean_paragraph(p))]
        if paras:
            sections.append((title, paras))
    return sections


def normalize_sections(sections: list[tuple[str, list[str]]]) -> list[tuple[str, list[str]]]:
    """Unify first/last titles and numbering: '1. ВВЕДЕНИЕ' ... 'N. ПОДВЕДЕНИЕ ИТОГОВ'."""
    if not sections:
        return sections
    out = []
    for title, paras in sections:
        key = title.upper()
        if "ГОЛОСА ФИЗИКИ" in key or "БОНУС" in key:
            continue
        out.append((title, paras))
    if not out:
        return out

    # First section -> Введение
    out[0] = ("ВВЕДЕНИЕ", out[0][1])

    # Collect all summary-like sections into one finale
    summary_paras: list[str] = []
    body = []
    for title, paras in out[1:]:
        key = title.upper()
        if any(x in key for x in ("ПОДВЕДЕНИЕ ИТОГ", "ИТОГ", "ФИНАЛ", "ЗАКЛЮЧ")):
            summary_paras.extend(paras)
        else:
            body.append((title, paras))
    out = [("ВВЕДЕНИЕ", out[0][1])] + body
    if summary_paras:
        # de-dup adjacent identical paragraphs
        merged = []
        for p in summary_paras:
            if not merged or merged[-1] != p:
                merged.append(p)
        out.append(("ПОДВЕДЕНИЕ ИТОГОВ", merged))
    elif any(x in out[-1][0].upper() for x in ("ИТОГ", "ФИНАЛ", "ЗАКЛЮЧ")):
        out[-1] = ("ПОДВЕДЕНИЕ ИТОГОВ", out[-1][1])

    numbered = []
    for i, (title, paras) in enumerate(out, 1):
        core = re.sub(r"^\d+\s*[-–—.:)]\s*", "", title).strip().upper()
        core = typography_fix(core)
        numbered.append((f"{i}. {core}", paras))
    return numbered


HANDBOOK_CHAPTER_KEYS = {
    1: "Строение атома",
    2: "Что такое свет",
    3: "Физические основы работы лазера",
    4: "Свет как поток фотонов",
    6: "Геометрическая оптика",
}


def _handbook_docx_candidates() -> list[Path]:
    return [
        DOCX_PATH,
        OUT_DIR / "Metodichka_lectures_1-6.docx",
        OUT_DIR / "Metodichka_lectures_1-5.docx",
    ]


def load_sections_from_existing_handbook(chapter_number: int) -> list[tuple[str, list[str]]]:
    """Запасной путь: текст главы из уже собранной методички (по заголовку, не по номеру)."""
    from docx import Document as Doc

    key = HANDBOOK_CHAPTER_KEYS[chapter_number]
    path = next((p for p in _handbook_docx_candidates() if p.exists()), None)
    if path is None:
        raise FileNotFoundError(f"Нет исходников лекции и нет DOCX методички для главы {chapter_number}")
    doc = Doc(path)
    paras_raw = [p.text.replace("\xa0", " ").strip() for p in doc.paragraphs]
    ch_re = re.compile(r"^Глава\s+\d+\.\s+(.+)$")
    hits = [i for i, t in enumerate(paras_raw) if ch_re.match(t) and key in ch_re.match(t).group(1)]
    if not hits:
        raise RuntimeError(f"В {path.name} не найдена глава «{key}»")
    start = hits[-1]
    end = len(paras_raw)
    for j in range(start + 1, len(paras_raw)):
        if paras_raw[j].startswith("Глава ") or paras_raw[j].startswith("Приложение."):
            end = j
            break
    body = paras_raw[start + 2:end]  # skip title + subtitle
    sec_re = re.compile(r"^\d+\.\s+\S")
    sections: list[tuple[str, list[str]]] = []
    title = None
    buf: list[str] = []

    def flush():
        nonlocal title, buf
        if title and buf:
            sections.append((title, buf[:]))
        title, buf = None, []

    for line in body:
        if not line:
            continue
        if line.startswith("Рис. ") or line.startswith("Таблица "):
            continue
        if sec_re.match(line):
            flush()
            title = line
            continue
        if title:
            buf.append(line)
    flush()
    return sections


def load_glossary_from_existing_handbook() -> list[str]:
    from docx import Document as Doc

    path = next((p for p in _handbook_docx_candidates() if p.exists()), None)
    if path is None:
        return []
    doc = Doc(path)
    lines = [p.text.replace("\xa0", " ").strip() for p in doc.paragraphs]
    try:
        i0 = next(i for i, t in enumerate(lines) if t.startswith("Приложение. Краткий глоссарий"))
    except StopIteration:
        return []
    # body appendix is the last occurrence
    idxs = [i for i, t in enumerate(lines) if t.startswith("Приложение. Краткий глоссарий")]
    i0 = idxs[-1] + 1
    out = []
    for t in lines[i0:]:
        if not t:
            continue
        if t.startswith("Глоссарий составлен"):
            continue
        out.append(t)
    return out


# ── loaders ────────────────────────────────────────────────────────────────

def load_lecture3_sections() -> list[tuple[str, list[str]]]:
    script = load_script(ROOT / "Lecture-3-main" / "build_lecture_3.py")
    sections = paragraphs_from_script_items(script)

    # Drop quotes section
    sections = [(t, p) for t, p in sections if "ГОЛОСА" not in t.upper()]

    # Logical order for chapter 1
    order_keys = [
        "ТИТУЛ",
        "ПЛАН",
        "УСТРОЙСТВО АТОМА",
        "ПАУЛИ",
        "НАТРИ",
        "ПУТАНИЦ",
        "ДВЕ СХЕМ",
        "СОБСТВЕННОЕ СОСТОЯНИЕ",
        "ТОЧКИ НА ЛИНИ",
        "ВСЕГДА ЛИ",
        "УРОВНИ В РАЗНЫХ",
        "ПОДВЕДЕНИЕ",
        "ИТОГ",
    ]

    def rank(title: str) -> int:
        u = title.upper()
        for i, key in enumerate(order_keys):
            if key in u:
                return i
        return 50

    # Merge duplicate-ish titles carefully; keep all unique by rank then original
    sections = sorted(sections, key=lambda x: rank(x[0]))

    if sections:
        title0, paras0 = sections[0]
        new_paras = []
        for p in paras0:
            if "внеплановая лекция" in p.lower() or "не повторяет прошлое" in p.lower():
                continue
            if p.startswith("Сегодня мы сосредоточимся"):
                p = (
                    "Этой лекцией мы начинаем курс семинаров по физике лазеров "
                    "для практического применения в аддитивных (3D) технологиях. "
                    + p.replace("Сегодня мы сосредоточимся", "Мы сосредоточимся", 1)
                )
            new_paras.append(p)
        if not any("начинаем курс семинаров" in p for p in new_paras):
            new_paras.insert(
                0,
                "Этой лекцией мы начинаем курс семинаров по физике лазеров "
                "для практического применения в аддитивных (3D) технологиях.",
            )
        sections[0] = (title0, new_paras)

    adapted = []
    for title, paras in sections:
        ap = []
        for p in paras:
            p = p.replace("из прошлой лекции", "из базовой схемы лазерной физики")
            p = p.replace(
                "Вернёмся к базовой схеме из прошлой лекции",
                "Вернёмся к базовой схеме лазерной физики",
            )
            ap.append(p)
        adapted.append((title, ap))
    return normalize_sections(adapted)


def load_lecture1_sections() -> tuple[list[tuple[str, list[str]]], list[str]]:
    from docx import Document as Doc

    path = next((ROOT / "Lecture-1-main").glob("*раскадровка*.docx"))
    doc = Doc(path)
    raw_lines = [p.text.replace("\xa0", " ") for p in doc.paragraphs]

    glossary: list[str] = []
    sections: list[tuple[str, list[str]]] = []
    current_title = None
    current_paras: list[str] = []
    mode = "preamble"

    def flush():
        nonlocal current_title, current_paras
        if current_title and current_paras:
            sections.append((current_title, current_paras[:]))
        current_title = None
        current_paras = []

    for line in raw_lines:
        s = line.strip()
        if not s:
            continue
        if s.upper().startswith("ГЛОССАРИЙ"):
            mode = "glossary"
            continue
        if s.startswith("РАЗДЕЛ") or re.match(r"^СЛАЙД\s+\d+", s, re.I):
            mode = "body"
            flush()
            current_title = clean_heading(s)
            current_paras = []
            continue
        if mode == "glossary":
            if not SPEAKER_LINE_RE.search(s):
                glossary.append(re.sub(r"\s+", " ", s))
            continue
        if mode == "preamble" or SPEAKER_LINE_RE.search(s):
            continue
        cp = clean_paragraph(s)
        if cp:
            current_paras.append(cp)
            if cp.startswith("Закон Бугера") and "затухает с глубиной" in cp:
                current_paras.append("I(x) = I₀ · exp(−α · x)")
    flush()

    adapted = []
    for title, paras in sections:
        # Skip bare "РАЗДЕЛ ..." wrappers if they have little content? keep them if useful
        if title.upper().startswith("РАЗДЕЛ"):
            # convert section wrappers into ordinary titled blocks only if they have text
            pass
        ap = []
        for p in paras:
            p = p.replace(
                "Сегодня мы начинаем курс «Оптика и лазерная физика». Это не тот курс, где вас завалят формулами на первой же минуте. Мы пойдём другим путём: от понимания физической сути - к инженерным решениям.",
                "После знакомства со строением атома перейдём к вопросу о природе света — тоже без перегрузки формулами на первых шагах: от физической сути к инженерным решениям.",
            )
            p = p.replace("Именно об этом наша первая лекция.", "Именно об этом эта глава.")
            p = p.replace("Подведём итог нашей первой лекции.", "Подведём итог этой главы.")
            p = p.replace("которые я хочу, чтобы вы унесли с собой", "которые важно зафиксировать")
            p = p.replace(
                "На следующей лекции мы заглянем внутрь лазерного луча: разберём, что такое гауссов пучок, почему распределение интенсивности в пятне не равномерное, а колоколообразное; что такое поляризация и как она влияет на поглощение; и что такое когерентность и зачем она нужна - а когда она мешает.",
                "В следующей главе разберём, как из взаимодействия фотона с атомом возникает лазерный луч: вынужденное излучение, инверсия населённостей, резонатор и типы лазеров, применяемых в аддитивных технологиях.",
            )
            p = p.replace("будем обсуждать в следующей лекции", "будет важно в дальнейших главах")
            p = p.replace("ВЫВОДЫ ДЛЯ SLM", "ВЫВОДЫ ДЛЯ СЕЛЕКТИВНОГО ЛАЗЕРНОГО ПЛАВЛЕНИЯ")
            p = p.replace("для SLM", "для селективного лазерного плавления")
            p = p.replace("для вашего станка", "для установки селективного лазерного плавления")
            ap.append(p)
        # Drop pure section banners like "РАЗДЕЛ 1. ..." with almost no body? keep
        if title.upper().startswith("РАЗДЕЛ") and len(ap) == 0:
            continue
        if title.upper().startswith("РАЗДЕЛ"):
            continue  # section banners without body — skip; slides carry content
        adapted.append((title, ap))
    return normalize_sections(adapted), glossary


def _strip_html(text: str) -> str:
    t = text.replace("<b>", "").replace("</b>", "").replace("<i>", "").replace("</i>", "")
    return re.sub(r"<[^>]+>", "", t).strip()


def _eval_py_string(raw: str) -> str:
    raw = raw.strip()
    # Known constants from Lecture-2 builder (for leftover f-string placeholders)
    moon_r = 384_400
    moon_spot = 4.0 * moon_r / 400_000  # ≈ 3.8 km
    ns = {
        "MOON_R_KM": moon_r,
        "MOON_SPOT_KM": moon_spot,
        "MOON_SPOT_M": moon_spot * 1000,
    }
    if raw.startswith("f"):
        lit = raw[1:].strip()
        try:
            template = ast.literal_eval(lit)
        except Exception:
            template = lit.strip("'\"")
        # Evaluate simple {NAME:.1f} / {NAME} placeholders
        def _repl(m):
            expr = m.group(1)
            try:
                return format(eval(expr, {"__builtins__": {}}, ns))  # noqa: S307
            except Exception:
                if "MOON_SPOT_KM" in expr:
                    return f"{moon_spot:.1f}".replace(".", ",")
                if "MOON_R_KM" in expr:
                    return str(moon_r)
                return "3,8"
        return re.sub(r"\{([^{}]+)\}", _repl, template)
    try:
        return ast.literal_eval(raw)
    except Exception:
        parts = re.findall(r"'((?:\\'|[^'])*)'|\"((?:\\\"|[^\"])*)\"", raw)
        return "".join(a or b for a, b in parts)


def load_lecture2_sections() -> list[tuple[str, list[str]]]:
    src = (ROOT / "Lecture-2-main" / "build_lecture_2.py").read_text(encoding="utf-8")
    m = re.search(r"def build_pdf\(.*?:\n(.*?)(?=\ndef build_docx|\nif __name__)", src, re.S)
    body = m.group(1) if m else src
    sections: list[tuple[str, list[str]]] = []
    title = None
    paras: list[str] = []

    def flush():
        nonlocal title, paras
        if title and paras:
            sections.append((title, paras[:]))
        paras = []

    for m in re.finditer(
        r"Paragraph\(\s*'(РАЗДЕЛ[^']+|СЛАЙД[^']+)'\s*,|"
        r"slide_h3\(\s*'(СЛАЙД[^']+)'\s*\)|"
        r"para\(\s*(f?'(?:\\'|[^'])*')\s*\)|"
        r"formula\(\s*(f?'(?:\\'|[^'])*')\s*\)|"
        r"hint\(\s*((?:.|\n)*?)\s*\)\s*\)",
        body,
    ):
        if m.group(1) or m.group(2):
            flush()
            title = clean_heading(m.group(1) or m.group(2))
            continue
        raw = m.group(3) or m.group(4) or m.group(5)
        cp = clean_paragraph(_strip_html(_eval_py_string(raw)))
        if cp:
            paras.append(cp)
    flush()

    adapted = []
    for sec_title, sec_paras in sections:
        if sec_title.upper().startswith("РАЗДЕЛ"):
            continue
        ap = []
        for p in sec_paras:
            p = p.replace(
                "Добрый день, коллеги. Мы продолжаем наш цикл лекций по общей теме «Оптика и физика лазеров». Сегодня вторая лекция — «Физические основы работы лазера».",
                "После разбора строения атома и природы света перейдём к теме «Физические основы работы лазера».",
            )
            p = p.replace(
                "На прошлой лекции мы узнали, что свет — это поток фотонов.",
                "Ранее мы установили, что свет — это поток фотонов.",
            )
            p = p.replace("изученные на прошлой лекции", "введённые ранее")
            p = p.replace("Об этом — во второй части лекции.", "Об этом — далее в этой главе.")
            p = p.replace("во второй части лекции", "далее в этой главе")
            ap.append(p)
        if ap:
            adapted.append((sec_title, ap))
    return normalize_sections(adapted)


def load_lecture4_sections() -> list[tuple[str, list[str]]]:
    """Глава 4: книжный стиль для читателя (не раскадровка лектора)."""
    script = load_script(ROOT / "Lecture-4-main" / "build_lecture_4.py")
    sections = paragraphs_from_script_items(script)
    adapted = []
    for title, paras in sections:
        ap = []
        for p in paras:
            p = p.replace(
                "Мы продолжаем курс. На прошлых занятиях разобрали двойственную природу света и устройство лазера: за счёт чего излучение становится когерентным, направленным и узкополосным.",
                "В предыдущих главах разобраны строение атома, двойственная природа света и устройство лазера: за счёт чего излучение становится когерентным, направленным и узкополосным.",
            )
            p = p.replace(
                "Сегодня — переход от физики источника к технологическому решению:",
                "Эта глава — переход от физики источника к технологическому решению:",
            )
            p = p.replace("Тема лекции:", "Тема главы:")
            p = p.replace("План — пять связанных шагов.", "План главы — пять связанных шагов.")
            p = p.replace("Это центр лекции:", "Это центр главы:")
            p = p.replace("Центральный раздел лекции.", "Центральный раздел главы.")
            p = p.replace(
                "селективного лазерного плавления.",
                "селективного лазерного плавления (далее — СЛП).",
                1,
            )
            p = p.replace("вопрос на следующее занятие", "переход к следующей теме курса")
            # Раскадровка → книга: обращение к слайду / доске / «вы»
            p = p.replace(
                "Центральный раздел лекции — два механизма",
                "Центральный раздел главы — два механизма",
            )
            p = p.replace(
                "Смотрим схему — две панели, две разные физики.",
                "На приведённом рисунке — две панели, две разные физики.",
            )
            p = p.replace(
                "на схеме справа кривая T(t) пересекает T_пл.",
                "на правой части рисунка кривая T(t) пересекает T_пл.",
            )
            p = p.replace("Чем управляют в фотохимии.", "Параметры управления в фотохимии.")
            p = p.replace(
                "Если квантовое условие уже выполнено (нужный фотон есть), дальше решают количество и геометрия процесса.",
                "Если квантовое условие уже выполнено (нужный фотон есть), далее определяют количество и геометрию процесса.",
            )
            p = p.replace("Не путайте словари двух процессов.", "Не следует смешивать словари двух процессов.")
            p = p.replace("Связь со слайдом механизмов.", "Связь с разделом о двух механизмах.")
            p = p.replace("мост к следующей лекции", "переход к следующей теме курса")
            p = p.replace(
                "На следующей лекции — геометрическая оптика:",
                "Далее по курсу логично перейти к геометрической оптике:",
            )
            p = p.replace(
                "На следующем занятии разберём её вместе с геометрией пучка.",
                "Эту цепочку полезно разобрать далее вместе с геометрией пучка.",
            )
            p = p.replace("Эти числа — опорные величины всей лекции.", "Эти числа — опорные величины всей главы.")
            p = p.replace("Выбирая λ (длину волны), вы задаёте", "Выбор λ (длины волны) задаёт")
            p = p.replace("Выбирая λ, вы задаёте", "Выбор λ задаёт")
            p = p.replace("Регулятором мощности меняется", "Регулятором мощности изменяют")
            p = p.replace("На доске: E = h·c/λ", "На рисунке и в формулах главы: E = h·c/λ")
            # Russianize residual English process jargon
            p = p.replace("lack of fusion", "несплавление")
            p = p.replace("keyhole-поры", "поры режима глубокого проплавления")
            p = p.replace("keyhole", "режим глубокого проплавления")
            p = p.replace("print-through", "пропечатывание насквозь")
            p = p.replace("шаг хэтча", "шаг штриховки")
            p = p.replace("green/blue", "зелёный/синий лазер")
            # Убрать остатки устной подачи
            p = p.replace("Формулировка для экзамена и практики: ", "")
            p = p.replace("Проверочный порядок вопросов", "Удобный порядок вопросов")
            if p.strip():
                ap.append(p)
        adapted.append((title, ap))
    return normalize_sections(adapted)


def load_lecture5_sections() -> list[tuple[str, list[str]]]:
    """Глава 6: геометрическая оптика — книжный стиль, без спикерских подсказок в скобках."""
    import importlib.util

    py_path = ROOT / "Lecture-5-main" / "build_lecture_5.py"
    script = load_script(py_path)
    spec = importlib.util.spec_from_file_location("bl5", py_path)
    bl5 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(bl5)
    epilogue = getattr(bl5, "EPILOGUE", "")

    sections = paragraphs_from_script_items(script)
    adapted = []
    for title, paras in sections:
        ap = []
        for p in paras:
            p = strip_handbook_glosses(p)
            p = p.replace(
                "Добрый день. Мы продолжаем курс. В Главе 2 мы выяснили главное:",
                "В предыдущих главах установлено главное:",
            )
            p = p.replace(
                "Сегодня мы разберём инструмент, которым мы управляем этой плотностью:",
                "В этой главе разбирается инструмент, которым управляют этой плотностью:",
            )
            p = p.replace("Мы ответим на три вопроса:", "Рассматриваются три вопроса:")
            p = p.replace("План будет таким.", "План главы:")
            p = p.replace("на примерах с расчётами на доске", "на примерах с расчётами")
            p = p.replace("вы не можете сделать его 1 мкм на вашем станке", "нельзя получить пятно 1 мкм на типичном СЛП-станке")
            p = p.replace("Давайте вспомним ключевое понятие", "Напомним ключевое понятие")
            p = p.replace("Я буду писать расчёты на доске.", "Ниже — расчёты двух сценариев.")
            p = p.replace("Теперь — два сценария.", "Рассмотрим два сценария.")
            p = p.replace("Теперь — как линза вообще работает.", "Рассмотрим, как линза работает.")
            p = p.replace(
                "Это базовая формула, которую полезно знать, даже если вы не будете её использовать для расчётов.",
                "Это базовая формула геометрической оптики.",
            )
            p = p.replace(
                "Delta n (разность показателей преломления)",
                "Δn (разность показателей преломления)",
            )
            p = p.replace("Теперь переходим к главному ограничению", "Перейдём к главному ограничению")
            p = p.replace("lambda (лямбda, λ)", "λ")
            p = p.replace("lambda (лямбда, λ)", "λ")
            p = p.replace("Обратите внимание:", "Важно:")
            p = p.replace("Закономерный вопрос:", "Логичный вопрос:")
            p = p.replace("Таблица трёх случаев — на слайде и в конспекте.", "Сводка трёх случаев приведена в таблице ниже.")
            p = p.replace("Итоговые тезисы.", "Итоговые тезисы главы:")
            p = p.replace(
                "Связь со следующей лекцией: мы говорили о линзах, как будто свет — идеальные прямые лучи.",
                "В этой главе линзы рассматривались в приближении геометрической оптики — свет как система прямых лучей.",
            )
            p = p.replace(
                "На следующем занятии («Волновая оптика»)",
                "В следующей главе («Волновая оптика»)",
            )
            if p.strip():
                ap.append(p)
        if ap:
            adapted.append((title, ap))

    if epilogue:
        ep = strip_handbook_glosses(clean_paragraph(epilogue) or epilogue)
        ep = ep.replace(
            "Связь со следующей лекцией: мы говорили о линзах, как будто свет — идеальные прямые лучи.",
            "В этой главе линзы рассматривались в приближении геометрической оптики — свет как система прямых лучей.",
        )
        ep = ep.replace(
            "На следующем занятии («Волновая оптика»)",
            "В следующей главе («Волновая оптика»)",
        )
        ep = typography_fix(ep)
        if adapted and adapted[-1][1]:
            adapted[-1][1].append(ep)

    return normalize_sections(dedupe_handbook_terminology(adapted))


def load_lecture4dop_sections() -> list[tuple[str, list[str]]]:
    """Глава 5: генерация зелёного излучения — книжный стиль по раскадровке Лекции 4-доп."""
    script = load_script(lecture4dop_dir() / "build_lecture_4dop.py")
    sections = paragraphs_from_script_items(script)
    adapted = []
    for title, paras in sections:
        title = title.replace("МОЁ ПРЕДЛОЖЕНИЕ:", "ПРЕДЛОЖЕНИЕ АВТОРА:")
        ap = []
        for p in paras:
            p = p.replace(
                "Добрый день. Мы продолжаем курс. В предыдущих лекциях мы разобрали, как оптика управляет плотностью мощности — тем, что определяет, плавится металл или нет.",
                "В предыдущих главах разобраны природа света, устройство лазера и спектральное поглощение — то, что определяет, плавится металл или нет.",
            )
            p = p.replace(
                "Мы продолжаем курс. В предыдущих лекциях мы разобрали, как оптика управляет плотностью мощности — тем, что определяет, плавится металл или нет.",
                "В предыдущих главах разобраны природа света, устройство лазера и спектральное поглощение — то, что определяет, плавится металл или нет.",
            )
            p = p.replace(
                "Сегодня мы разберём проблему, которая стоит перед всей отраслью аддитивных технологий:",
                "В этой главе разбирается проблема, которая стоит перед всей отраслью аддитивных технологий:",
            )
            p = p.replace(
                "И главное: как мы предлагаем эту проблему решить. Речь пойдёт о генерации зелёного излучения с длиной волны 535 нанометров — и о моём предложении, которое делает эту технологию экономически доступной.",
                "Главный вопрос главы: как эту проблему решить. Речь пойдёт о генерации зелёного излучения с длиной волны 535 нанометров — и о предложении, которое делает эту технологию экономически доступной.",
            )
            p = p.replace("План будет таким.", "План главы:")
            p = p.replace("Давайте начнём с физики.", "Начнём с физики.")
            p = p.replace("Вспомним ключевое понятие", "Напомним ключевое понятие")
            p = p.replace("Теперь посмотрим на зелёный диапазон.", "Рассмотрим зелёный диапазон.")
            p = p.replace("Объясню физику.", "Физика процесса такова.")
            p = p.replace("Давайте разберём его математику.", "Разберём его математику.")
            p = p.replace(
                "Мы пойдём другим путём. Мы возьмём мощный и дешёвый инфракрасный лазер",
                "Пойдём другим путём: возьмём мощный и дешёвый инфракрасный лазер",
            )
            p = p.replace("Расшифровка каждого символа — в таблице на слайде.", "Расшифровка каждого символа — в таблице ниже.")
            p = p.replace("Обратите внимание:", "Важно:")
            p = p.replace("Теперь — самое важное. Почему зелёные такие дорогие и что я предлагаю.", "Далее — ключевой практический вопрос: почему зелёные лазеры так дороги и что предлагается.")
            p = p.replace("Моё предложение — двухпроходная схема", "Предложение автора — двухпроходная схема")
            p = p.replace("Моё предложение — это применение", "Предложение автора — это применение")
            p = p.replace("Моё предложение:", "Предложение автора:")
            p = p.replace("Что именно является моим предложением?", "Что именно является предложением этой главы?")
            p = p.replace("Прежде чем описывать схему, я должен быть честен:", "Прежде чем описывать схему, важно оговорить:")
            p = p.replace("вместо пассивной термостабилизации кристалла в печи я предлагаю", "вместо пассивной термостабилизации кристалла в печи предлагается")
            p = p.replace("Вот моя логика как физика-теоретика:", "Логика выбора материала:")
            p = p.replace("Это означает, что я могу точно рассчитать", "Это означает, что можно точно рассчитать")
            p = p.replace("Моё предложение — использовать отдельный клин", "Предложение автора — использовать отдельный клин")
            p = p.replace("Расшифровка символов — в таблице на слайде.", "Расшифровка символов — в таблице ниже.")
            p = p.replace("Сфиксируем пять тезисов.", "Зафиксируем пять тезисов.")
            p = p.replace("Итак, мы разобрали физику", "Итак, разобрана физика")
            p = p.replace("Спасибо за внимание. Готов ответить на вопросы.", "")
            p = p.replace("в предыдущих лекциях", "в предыдущих главах")
            p = p.replace("на слайде", "на рисунке")
            if p.strip():
                ap.append(p)
        if ap:
            adapted.append((title, ap))
    return normalize_sections(adapted)


# ── attach media ───────────────────────────────────────────────────────────

def _match(title: str, *keys: str) -> bool:
    u = title.upper()
    return any(k.upper() in u for k in keys)


def attach_media(chapters, catalog):
    """Attach figures/tables to sections; return chapters with rich section dicts."""
    fig_counters = {ch["number"]: 0 for ch in chapters}
    tab_counters = {ch["number"]: 0 for ch in chapters}

    def add_fig(ch_num, key):
        if key not in catalog["figures"]:
            return None
        fig_counters[ch_num] += 1
        meta = catalog["figures"][key]
        return {
            "type": "figure",
            "path": Path(meta["path"]),
            "caption": typography_fix(
                f"Рис. {ch_num}.{fig_counters[ch_num]}. {meta['caption']}"
            ),
        }

    def add_tab(ch_num, key):
        if key not in catalog["tables"]:
            return None
        tab_counters[ch_num] += 1
        meta = catalog["tables"][key]
        return {
            "type": "table",
            "caption": typography_fix(
                f"Таблица {ch_num}.{tab_counters[ch_num]}. {meta['caption']}"
            ),
            "headers": [typography_fix(h) for h in meta["headers"]],
            "rows": [[typography_fix(c) for c in row] for row in meta["rows"]],
        }

    # Mapping rules per chapter number
    rules = {
        1: [
            (("УСТРОЙСТВО АТОМА",), ["ch1_atom"]),
            (("ПАУЛИ",), ["ch1_pauli", "tab:ch1_quantum"]),
            (("НАТРИ",), ["ch1_sodium"]),
            (("ПУТАНИЦ",), ["ch1_two_schemes"]),
            (("ДВЕ СХЕМ", "СОБСТВЕНН"), ["ch1_two_schemes"]),
            (("ТОЧКИ",), ["ch1_points"]),
            (("ВСЕГДА ЛИ",), ["ch1_not_always"]),
        ],
        2: [
            (("ПОРОГОВОЙ ПЛОТНОСТИ", "ГРАФИК ПОРОГОВ"), ["ch2_threshold"]),
        ],
        3: [
            (("ЭНЕРГЕТИЧЕСКИЕ УРОВНИ",), ["ch3_spontaneous"]),
            (("ФОТОННАЯ ЛАВИНА", "MASER"), ["ch3_avalanche"]),
            (("ИНВЕРСНАЯ",), ["ch3_inversion_prob"]),
            (("РУБИНОВ",), ["ch3_ruby"]),
            (("РЕЗОНАТОР", "ОБРАТНАЯ СВЯЗЬ"), ["ch3_pump", "ch3_resonator"]),
            (("СРАВНЕНИЕ",), ["tab:ch3_laser_compare"]),
        ],
        4: [
            (("ЧАСТОТА", "ЭНЕРГИЯ ФОТОНА", "ДЛИНА ВОЛНЫ И ЭНЕРГИЯ"), ["ch4_energy"]),
            (("ШКАЛА",), ["ch4_spectrum"]),
            (("ДВА МЕХАНИЗМА", "ФОТОХИМИЧ"), ["ch4_mechanisms"]),
            (("ПОГЛОЩАТЕЛЬН", "СПЕКТРАЛЬН"), ["ch4_absorption"]),
            (("КРИТЕРИИ", "ВЫБОР ЛАЗЕРА"), ["ch4_choice"]),
        ],
        5: [
            (("МЕДЬ", "ИК-ЛАЗЕР"), ["ch5_copper", "tab:ch5_copper"]),
            (("ПРЯМАЯ ГЕНЕРАЦИЯ", "ЗЕЛЁНОГО СВЕТА"), ["ch5_green_gap", "ch5_photon_energy", "tab:ch5_green_gap"]),
            (("НЕЛИНЕЙН", "ВТОРОЙ ГАРМОН"), ["ch5_asym", "ch5_shg", "ch5_cos2", "tab:ch5_chi"]),
            (("ДВУХПРОХОД", "ТЕРМОКОМПЕНС", "ПРЕДЛОЖЕН"), ["ch5_twopass", "ch5_sinc2", "ch5_eta", "tab:ch5_eta"]),
            (("ИТОГ",), ["ch5_economy"]),
        ],
        6: [
            (("ОПТИКА", "ПЛОТНОСТЬ", "РАСЧЁТ"), ["ch6_intensity"]),
            (("СНЕЛЛИУС",), ["ch6_snell"]),
            (("РАЗМЕР ПЯТНА", "ФОРМУЛА №1"), ["ch6_spot"]),
            (("ГЛУБИНА РЕЗКОСТИ", "ФОРМУЛА №2"), ["ch6_dof", "tab:ch6_na_table"]),
            (("M²", "M2", "КАЧЕСТВО ПУЧКА", "ИТОГ"), ["ch6_m2"]),
        ],
    }

    rich_chapters = []
    for ch in chapters:
        rich_sections = []
        used_figs = set()
        for title, paras in ch["sections"]:
            blocks = [{"type": "text", "text": p} for p in paras]
            pending = []
            for keys, media_keys in rules.get(ch["number"], []):
                if _match(title, *keys):
                    for mk in media_keys:
                        if mk.startswith("tab:"):
                            item = add_tab(ch["number"], mk[4:])
                        else:
                            if mk in used_figs:
                                continue
                            item = add_fig(ch["number"], mk)
                            if item:
                                used_figs.add(mk)
                        if item:
                            pending.append(item)
            insert_at = min(2, len(blocks))
            for i, item in enumerate(pending):
                blocks.insert(insert_at + i, item)
            rich_sections.append({"title": title, "blocks": blocks})
        # Ensure ch2 threshold appears if missed (attach to threshold section by fuzzy)
        if ch["number"] == 2 and "ch2_threshold" not in used_figs:
            for sec in rich_sections:
                if _match(sec["title"], "ПОРОГ", "ГРАФИК", "80 %"):
                    item = add_fig(2, "ch2_threshold")
                    if item:
                        sec["blocks"].insert(min(2, len(sec["blocks"])), item)
                        used_figs.add("ch2_threshold")
                    break
        rich_chapters.append({**ch, "sections": rich_sections})
    return rich_chapters


def _try_source_or_docx(loader, lecture_rel: str, handbook_ch: int):
    src = ROOT / lecture_rel
    if src.exists():
        return loader()
    return load_sections_from_existing_handbook(handbook_ch)


def build_book_model(catalog):
    ch1 = _try_source_or_docx(load_lecture3_sections, "Lecture-3-main/build_lecture_3.py", 1)
    if (ROOT / "Lecture-1-main").exists():
        ch2, glossary = load_lecture1_sections()
    else:
        ch2 = load_sections_from_existing_handbook(2)
        glossary = load_glossary_from_existing_handbook()
    ch3 = _try_source_or_docx(load_lecture2_sections, "Lecture-2-main/build_lecture_2.py", 3)
    ch4 = _try_source_or_docx(load_lecture4_sections, "Lecture-4-main/build_lecture_4.py", 4)
    ch5 = load_lecture4dop_sections()
    ch6 = _try_source_or_docx(load_lecture5_sections, "Lecture-5-main/build_lecture_5.py", 6)

    chapters = [
        {
            "number": 1,
            "title": "Строение атома и два языка лазерной физики",
            "subtitle": "Орбитали, квантовые числа и энергетические уровни E₁, E₂",
            "sections": ch1,
        },
        {
            "number": 2,
            "title": "Что такое свет, если забыть про формулы?",
            "subtitle": "Оптика и лазерная физика для инженеров аддитивных технологий",
            "sections": ch2,
        },
        {
            "number": 3,
            "title": "Физические основы работы лазера",
            "subtitle": "Вынужденное излучение, инверсия, резонатор и типы лазеров",
            "sections": ch3,
        },
        {
            "number": 4,
            "title": "Свет как поток фотонов: энергия, длина волны и выбор лазера",
            "subtitle": "Фотохимия и тепло, спектральное поглощение, критерии выбора источника",
            "sections": ch4,
        },
        {
            "number": 5,
            "title": "Генерация зелёного излучения для аддитивных технологий",
            "subtitle": "Медь, «зелёная щель», SHG в PPLN и двухпроходная схема с термокомпенсацией",
            "sections": ch5,
        },
        {
            "number": 6,
            "title": "Геометрическая оптика: как линзы собирают луч в пятно",
            "subtitle": "Плотность мощности I = P/S, числовая апертура, DOF и параметр M²",
            "sections": ch6,
        },
    ]
    rich = attach_media(chapters, catalog)
    glossary = _merge_glossary(list(glossary), GLOSSARY_CH4 + GLOSSARY_CH5 + GLOSSARY_CH6)
    return rich, list(PREFACE), glossary


def _merge_glossary(base: list[str], extras: list[str]) -> list[str]:
    keys = set()
    out = []
    for g in base + extras:
        key = re.split(r"\t| {2,}", g, maxsplit=1)[0].strip().lower()
        if len(key) < 1:
            continue
        if key in keys:
            continue
        keys.add(key)
        out.append(g)
    return out


# ── DOCX ───────────────────────────────────────────────────────────────────

def set_run_font(run, name="Times New Roman", size=12, bold=False, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


# Остаточные ASCII-индексы (ω не имеет юникод-подстрочной формы)
_DOCX_SUB_RE = re.compile(
    r"(₂ω|_2ω|_ω|_eff|_акт|_пл|_реал|_идеал|_max|_погл|_пад|_ph|_33)"
)
_DOCX_SUB_MAP = {
    "₂ω": "2ω",
    "_2ω": "2ω",
    "_ω": "ω",
    "_eff": "eff",
    "_акт": "акт",
    "_пл": "пл",
    "_реал": "реал",
    "_идеал": "идеал",
    "_max": "max",
    "_погл": "погл",
    "_пад": "пад",
    "_ph": "ph",
    "_33": "33",
}


def _add_script_runs(p, text, *, size=12, bold=False, italic=False):
    """Текст в run'ы; _ω / _eff / ₂ω → настоящий подстрочный индекс Word."""
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part:
            continue
        is_bold = part.startswith("**") and part.endswith("**") and len(part) >= 4
        chunk = part[2:-2] if is_bold else part
        pos = 0
        for m in _DOCX_SUB_RE.finditer(chunk):
            if m.start() > pos:
                run = p.add_run(chunk[pos:m.start()])
                set_run_font(run, size=size, bold=bold or is_bold, italic=italic)
            run = p.add_run(_DOCX_SUB_MAP[m.group()])
            set_run_font(run, size=size, bold=bold or is_bold, italic=italic)
            run.font.subscript = True
            pos = m.end()
        if pos < len(chunk):
            run = p.add_run(chunk[pos:])
            set_run_font(run, size=size, bold=bold or is_bold, italic=italic)


def add_para(doc, text, *, size=12, bold=False, italic=False,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=8, first_indent=True):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.15
    if first_indent and align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        p.paragraph_format.first_line_indent = Cm(1.25)
    _add_script_runs(p, text, size=size, bold=bold, italic=italic)
    return p


def add_figure_docx(doc, path: Path, caption: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run()
    try:
        from PIL import Image as PILImage
        with PILImage.open(path) as im:
            w_px, h_px = im.size
    except Exception:
        w_px, h_px = 1000, 700
    width_cm = 16.0 if w_px >= 900 else 14.0
    run.add_picture(str(path), width=Cm(width_cm))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(10)
    _add_script_runs(cap, caption, size=10, italic=True)


def _fill_cell(cell, text, *, size=9, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    _add_script_runs(p, text, size=size, bold=bold)


def add_table_docx(doc, headers, rows, caption: str):
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(8)
    cap.paragraph_format.space_after = Pt(4)
    _add_script_runs(cap, caption, size=10, italic=True)

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        _fill_cell(table.rows[0].cells[j], h, size=9, bold=True)
    for i, row in enumerate(rows, 1):
        for j, val in enumerate(row):
            _fill_cell(table.rows[i].cells[j], val, size=9)
    doc.add_paragraph()


def build_docx(chapters, preface, glossary, out_path: Path | None = None):
    out_path = out_path or DOCX_PATH
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(12)

    for _ in range(2):
        doc.add_paragraph()
    add_para(doc, "ФИЗИКА ЛАЗЕРОВ ДЛЯ АДДИТИВНЫХ (3D) ТЕХНОЛОГИЙ",
             size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, space_after=16)
    add_para(doc, "Методическое пособие по материалам семинарских лекций",
             size=14, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, space_after=10)
    add_para(doc, f"Автор: {AUTHOR}",
             size=13, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False, space_after=18)
    add_para(doc,
             "Для студентов и инженеров, осваивающих лазерные процессы в технологиях послойного синтеза",
             size=12, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=False)
    doc.add_page_break()

    add_para(doc, "Предисловие", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
             first_indent=False, space_after=12)
    for p in preface:
        add_para(doc, p)
    doc.add_page_break()

    add_para(doc, "Содержание", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
             first_indent=False, space_after=12)
    for ch in chapters:
        add_para(doc, f"Глава {ch['number']}. {ch['title']}",
                 align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, space_after=2)
        add_para(doc, ch["subtitle"], size=11, italic=True,
                 align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, space_after=4)
        for sec in ch["sections"]:
            add_para(doc, f"    {sec['title']}", size=11,
                     align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, space_after=1)
        add_para(doc, "", space_after=6, first_indent=False)
    add_para(doc, "Приложение. Краткий глоссарий обозначений",
             align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False)
    doc.add_page_break()

    for ch in chapters:
        add_para(doc, f"Глава {ch['number']}. {ch['title']}",
                 size=18, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, space_after=6)
        add_para(doc, ch["subtitle"], size=12, italic=True,
                 align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, space_after=14)
        for sec in ch["sections"]:
            add_para(doc, sec["title"], size=13, bold=True,
                     align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, space_after=8)
            for block in sec["blocks"]:
                if block["type"] == "text":
                    add_para(doc, block["text"])
                elif block["type"] == "figure":
                    if block["path"].exists():
                        add_figure_docx(doc, block["path"], block["caption"])
                elif block["type"] == "table":
                    add_table_docx(doc, block["headers"], block["rows"], block["caption"])
        doc.add_page_break()

    add_para(doc, "Приложение. Краткий глоссарий обозначений",
             size=16, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False, space_after=12)
    add_para(
        doc,
        "Глоссарий составлен по материалам главы о природе света и сохранён как справочник обозначений для всего курса.",
        italic=True, first_indent=False, space_after=10,
    )
    for g in glossary:
        if len(g) >= 3:
            add_para(doc, g, size=11, first_indent=False, space_after=4)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    print(f"OK DOCX: {out_path}")


# ── PDF ────────────────────────────────────────────────────────────────────

# Обратные таблицы: юникод-индексы → обычные символы для <super>/<sub> в PDF
_INV_SUP = str.maketrans(
    "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱˣʸᶻᵃᵇᶜᵈᵉʰʲᵏˡᵐᵒᵖʳˢᵗᵘᵛʷ",
    "0123456789+-=()nixyzabcdehjklmoprstuvw",
)
_INV_SUB = str.maketrans(
    "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ",
    "0123456789+-=()aehijklmnoprstuvx",
)
_SUPER_RUN_RE = re.compile(r"[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱˣʸᶻᵃᵇᶜᵈᵉʰʲᵏˡᵐᵒᵖʳˢᵗᵘᵛʷ]+")
_SUB_RUN_RE = re.compile(r"[₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ]+")
_PDF_MATH_RE = re.compile(r"[∝→←↔⇒×]")


def _pdf_scripts_to_rl(text: str) -> str:
    """Times New Roman в PDF не содержит ₀⁽⁾∝ — рисуем индексы тегами ReportLab."""
    t = _SUPER_RUN_RE.sub(lambda m: f"<super>{m.group().translate(_INV_SUP)}</super>", text)
    t = _SUB_RUN_RE.sub(lambda m: f"<sub>{m.group().translate(_INV_SUB)}</sub>", t)
    # _{ω} / _{eff} не имеют юникод-индекса — в PDF тоже подстрочно
    t = t.replace("_ω", "<sub>ω</sub>")
    t = t.replace("_eff", "<sub>eff</sub>")
    # λ<sub>2</sub>ω → λ<sub>2ω</sub>
    t = re.sub(r"</sub>ω", "ω</sub>", t)
    return t


def _pdf_wrap_math(text: str, sym_font: str | None) -> str:
    if not sym_font:
        return text
    return _PDF_MATH_RE.sub(lambda m: f'<font name="{sym_font}">{m.group()}</font>', text)


def pdf_escape(text: str, *, bold_font: str | None = None, sym_font: str | None = None) -> str:
    # typography first (may introduce unicode), then escape XML specials;
    # **...** → bold (via explicit bold font name when provided)
    t = typography_fix(text)
    parts = re.split(r"(\*\*.+?\*\*)", t)
    out = []
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            inner = part[2:-2].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if bold_font:
                out.append(f'<font name="{bold_font}">{inner}</font>')
            else:
                out.append(f"<b>{inner}</b>")
        else:
            out.append(part.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    s = "".join(out)
    s = _pdf_scripts_to_rl(s)
    s = _pdf_wrap_math(s, sym_font)
    return s


def build_pdf(chapters, preface, glossary, out_path: Path | None = None):
    out_path = out_path or PDF_PATH
    font, font_bold, font_italic, sym_font = _register_fonts()

    def esc(text: str) -> str:
        return pdf_escape(text, bold_font=font_bold, sym_font=sym_font)

    styles = getSampleStyleSheet()
    cover = ParagraphStyle("Cover", parent=styles["Normal"], fontName=font_bold, fontSize=18,
                           leading=24, alignment=TA_CENTER, spaceAfter=16, textColor=HexColor("#1a1a2e"))
    cover_sub = ParagraphStyle("CoverSub", parent=styles["Normal"], fontName=font_italic, fontSize=12,
                               leading=16, alignment=TA_CENTER, spaceAfter=10, textColor=HexColor("#333333"))
    h1 = ParagraphStyle("H1", parent=styles["Normal"], fontName=font_bold, fontSize=14,
                        leading=18, alignment=TA_LEFT, spaceBefore=6, spaceAfter=8, textColor=HexColor("#1a1a2e"))
    h2 = ParagraphStyle("H2", parent=styles["Normal"], fontName=font_bold, fontSize=11.5,
                        leading=15, alignment=TA_LEFT, spaceBefore=10, spaceAfter=6, textColor=HexColor("#222222"))
    body = ParagraphStyle("Body", parent=styles["Normal"], fontName=font, fontSize=10.5,
                          leading=14, alignment=TA_JUSTIFY, spaceAfter=6, firstLineIndent=14)
    body0 = ParagraphStyle("Body0", parent=body, firstLineIndent=0)
    note = ParagraphStyle("Note", parent=styles["Normal"], fontName=font_italic, fontSize=9.5,
                          leading=13, alignment=TA_CENTER, spaceAfter=10, textColor=HexColor("#444444"))
    toc = ParagraphStyle("TOC", parent=styles["Normal"], fontName=font, fontSize=11,
                         leading=15, alignment=TA_LEFT, spaceAfter=2)
    toc_sec = ParagraphStyle(
        "TOCSec", parent=toc, leftIndent=18, spaceAfter=1,
    )
    cell = ParagraphStyle("Cell", parent=styles["Normal"], fontName=font, fontSize=8,
                          leading=10, alignment=TA_LEFT)
    cell_h = ParagraphStyle("CellH", parent=cell, fontName=font_bold)

    story = []
    story.append(Spacer(1, 3 * cm))
    cover_author = ParagraphStyle(
        "CoverAuthor", parent=styles["Normal"], fontName=font, fontSize=12,
        leading=16, alignment=TA_CENTER, spaceAfter=14, textColor=HexColor("#222222"),
    )
    story.append(Paragraph(esc("ФИЗИКА ЛАЗЕРОВ ДЛЯ АДДИТИВНЫХ (3D) ТЕХНОЛОГИЙ"), cover))
    story.append(Paragraph(esc("Методическое пособие по материалам семинарских лекций"), cover_sub))
    story.append(Paragraph(esc(f"Автор: {AUTHOR}"), cover_author))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(
        esc("Для студентов и инженеров, осваивающих лазерные процессы в технологиях послойного синтеза"),
        cover_sub,
    ))
    story.append(PageBreak())

    story.append(Paragraph(esc("Предисловие"), h1))
    for p in preface:
        story.append(Paragraph(esc(p), body))
    story.append(PageBreak())

    story.append(Paragraph(esc("Содержание"), h1))
    for ch in chapters:
        story.append(Paragraph(esc(f"Глава {ch['number']}. {ch['title']}"), toc))
        story.append(Paragraph(esc(ch["subtitle"]), note))
        for sec in ch["sections"]:
            story.append(Paragraph(esc(sec["title"]), toc_sec))
    story.append(Paragraph(esc("Приложение. Краткий глоссарий обозначений"), toc))
    story.append(PageBreak())

    for ch in chapters:
        story.append(Paragraph(esc(f"Глава {ch['number']}. {ch['title']}"), h1))
        story.append(Paragraph(esc(ch["subtitle"]), note))
        for sec in ch["sections"]:
            story.append(Paragraph(esc(sec["title"]), h2))
            for block in sec["blocks"]:
                if block["type"] == "text":
                    story.append(Paragraph(esc(block["text"]), body))
                elif block["type"] == "figure" and block["path"].exists():
                    # Сохраняем пропорции; широкие схемы — на всю ширину полосы набора
                    try:
                        from PIL import Image as PILImage
                        with PILImage.open(block["path"]) as im:
                            w_px, h_px = im.size
                    except Exception:
                        w_px, h_px = 1000, 700
                    max_w = 16.5 * cm if w_px >= 900 else 14 * cm
                    aspect = h_px / max(1, w_px)
                    img_w = max_w
                    img_h = img_w * aspect
                    max_h = 21 * cm
                    if img_h > max_h:
                        img_h = max_h
                        img_w = img_h / aspect
                    img = RLImage(str(block["path"]), width=img_w, height=img_h)
                    story.append(Spacer(1, 6))
                    story.append(img)
                    story.append(Paragraph(esc(block["caption"]), note))
                elif block["type"] == "table":
                    story.append(Paragraph(esc(block["caption"]), note))
                    data = [[Paragraph(esc(h), cell_h) for h in block["headers"]]]
                    for row in block["rows"]:
                        data.append([Paragraph(esc(c), cell) for c in row])
                    n_cols = max(1, len(block["headers"]))
                    if n_cols == 6:
                        # Сценарии A(λ): узкие №/A, широкие «что/почему»
                        col_ws = [0.9 * cm, 2.9 * cm, 2.1 * cm, 2.2 * cm, 4.2 * cm, 4.2 * cm]
                    else:
                        col_ws = [16.5 * cm / n_cols] * n_cols
                    tbl = Table(data, colWidths=col_ws)
                    tbl.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#e8eef2")),
                        ("GRID", (0, 0), (-1, -1), 0.4, HexColor("#666666")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ]))
                    story.append(tbl)
                    story.append(Spacer(1, 8))
        story.append(PageBreak())

    story.append(Paragraph(esc("Приложение. Краткий глоссарий обозначений"), h1))
    story.append(Paragraph(
        esc("Глоссарий составлен по материалам главы о природе света и сохранён как справочник обозначений для всего курса."),
        note,
    ))
    for g in glossary:
        if len(g) >= 3:
            story.append(Paragraph(esc(g), body0))

    def _page(canvas, doc_):
        canvas.saveState()
        canvas.setFont(font, 9)
        canvas.setFillColor(gray)
        canvas.drawCentredString(A4[0] / 2, 1.2 * cm, str(canvas.getPageNumber()))
        canvas.restoreState()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(out_path), pagesize=A4,
        leftMargin=2.2 * cm, rightMargin=1.8 * cm,
        topMargin=1.6 * cm, bottomMargin=1.6 * cm,
        title="Физика лазеров для аддитивных (3D) технологий",
        author=AUTHOR,
    )
    doc.build(story, onFirstPage=_page, onLaterPages=_page)
    print(f"OK PDF: {out_path}")


def write_readme(chapters):
    lines = [
        "# Методичка: Физика лазеров для аддитивных (3D) технологий",
        "",
        "## Файлы",
        "",
        f"- `{DOCX_PATH.name}`",
        f"- `{PDF_PATH.name}`",
        f"- `{REPO_STUDENT_DOCX.name}` — Word для Яндекс Документов (главы 1–6, = PDF)",
        "- `assets/` — рисунки и схемы",
        "",
        "## Главы",
        "",
    ]
    for ch in chapters:
        lines.append(f"{ch['number']}. {ch['title']}")
        for sec in ch["sections"]:
            n_fig = sum(1 for b in sec["blocks"] if b["type"] == "figure")
            n_tab = sum(1 for b in sec["blocks"] if b["type"] == "table")
            extra = []
            if n_fig:
                extra.append(f"{n_fig} рис.")
            if n_tab:
                extra.append(f"{n_tab} табл.")
            suffix = f" ({', '.join(extra)})" if extra else ""
            lines.append(f"   - {sec['title']}{suffix}")
    lines += ["", "Пересборка: `python build_print_handbook.py`", ""]
    (OUT_DIR / "README.md").write_text("\n".join(lines), encoding="utf-8")
    print("OK README")


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    STUDENT_DIR.mkdir(parents=True, exist_ok=True)
    catalog = prepare_all()
    chapters, preface, glossary = build_book_model(catalog)
    n_fig = sum(1 for ch in chapters for sec in ch["sections"] for b in sec["blocks"] if b["type"] == "figure")
    n_tab = sum(1 for ch in chapters for sec in ch["sections"] for b in sec["blocks"] if b["type"] == "table")
    print(f"Chapters={len(chapters)}; figures={n_fig}; tables={n_tab}")
    for ch in chapters:
        print(f"  Ch{ch['number']}: {len(ch['sections'])} sections — first={ch['sections'][0]['title']}; last={ch['sections'][-1]['title']}")
    # Полная версия — исходники в Lectures_for_print / репозиторий
    build_docx(chapters, preface, glossary, DOCX_PATH)
    build_pdf(chapters, preface, glossary, PDF_PATH)
    if STUDENT_PDF_PATH.resolve() != PDF_PATH.resolve():
        build_pdf(chapters, preface, glossary, STUDENT_PDF_PATH)
    if STUDENT_DOCX_PATH.resolve() != DOCX_PATH.resolve():
        build_docx(chapters, preface, glossary, STUDENT_DOCX_PATH)
    # Копии с латинским именем — репозиторий и Рабочий стол (тот же текст, что PDF/DOCX)
    try:
        import shutil
        shutil.copy2(STUDENT_DOCX_PATH, REPO_STUDENT_DOCX)
        shutil.copy2(STUDENT_DOCX_PATH, DESKTOP_STUDENT_DOCX)
        print(f"OK REPO DOCX: {REPO_STUDENT_DOCX}")
        print(f"OK DESKTOP DOCX: {DESKTOP_STUDENT_DOCX}")
    except Exception as e:
        print(f"WARN student docx copies: {e}")
    write_readme(chapters)
    print("DONE")


if __name__ == "__main__":
    main()
