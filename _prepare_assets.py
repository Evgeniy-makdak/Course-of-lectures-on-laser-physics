# -*- coding: utf-8 -*-
"""Generate / extract figures and tables for the print handbook."""
from __future__ import annotations

import importlib.util
import io
import sys
from pathlib import Path

import os

from PIL import ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

_HERE = Path(__file__).resolve().parent
_WIN_ROOT = Path(r"C:\Users\Volkov\Desktop\Lecture-Unscheduled")


def _detect_root() -> Path:
    if _WIN_ROOT.exists() and (_WIN_ROOT / "Lecture-1-main").exists():
        return _WIN_ROOT
    parent = _HERE.parent
    if (parent / "Lecture-4_optional").exists() or (parent / "Lecture-1-main").exists():
        return parent
    return _HERE


ROOT = _detect_root()
ASSETS = _HERE / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)


def lecture4dop_dir() -> Path:
    for cand in (ROOT / "Lecture-4_optional", _HERE.parent / "Lecture-4_optional"):
        if (cand / "build_lecture_4dop.py").exists():
            return cand
    raise FileNotFoundError("Не найден build_lecture_4dop.py в Lecture-4_optional")


def _load(py_path: Path):
    name = py_path.stem + "_mod"
    spec = importlib.util.spec_from_file_location(name, py_path)
    mod = importlib.util.module_from_spec(spec)
    # Avoid running __main__
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _first_font(*paths: str) -> str | None:
    for p in paths:
        if p and Path(p).exists():
            return p
    return None


def _patch_lecture4dop_fonts(mod) -> None:
    """macOS (и запасные пути): греческие/индексы на картинках как в Лекции 4-доп."""
    regular = _first_font(
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf"),
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
    )
    bold = _first_font(
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arialbd.ttf"),
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        regular,
    )
    math = _first_font(
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "timesi.ttf"),
        "/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
    )
    symbol = _first_font(
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "seguisym.ttf"),
        "/Library/Fonts/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Apple Symbols.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    )
    if regular:
        mod.FONT_PATH = regular
        mod.FONT_BOLD_PATH = bold or regular
    if math:
        mod.MATH_FONT_PATH = math

    def _pil_font(size=14, bold=False):
        cands = []
        if bold:
            cands += [
                getattr(mod, "FONT_BOLD_PATH", None),
                "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            ]
        cands += [
            getattr(mod, "FONT_PATH", None),
            regular,
            "/System/Library/Fonts/Supplemental/Arial.ttf",
        ]
        for p in cands:
            if p and os.path.exists(p):
                try:
                    return ImageFont.truetype(p, size=size)
                except OSError:
                    pass
        return ImageFont.load_default()

    def _pil_math_font(size=14, bold=False):
        cands = [
            math,
            "/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
            "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
        ]
        if bold:
            cands.insert(0, "/System/Library/Fonts/Supplemental/Times New Roman Bold Italic.ttf")
        for p in cands:
            if p and os.path.exists(p):
                try:
                    return ImageFont.truetype(p, size=size)
                except OSError:
                    continue
        return _pil_font(size)

    def _pil_symbol_font(size=12):
        for p in (symbol, "/Library/Fonts/Arial Unicode.ttf",
                  "/System/Library/Fonts/Supplemental/Arial.ttf"):
            if p and os.path.exists(p):
                try:
                    return ImageFont.truetype(p, size=size)
                except OSError:
                    continue
        return _pil_font(size)

    mod._pil_font = _pil_font
    mod._pil_math_font = _pil_math_font
    mod._pil_symbol_font = _pil_symbol_font


def load_lecture4dop_module():
    mod = _load(lecture4dop_dir() / "build_lecture_4dop.py")
    _patch_lecture4dop_fonts(mod)
    return mod


def _add_fig(catalog: dict, key: str, path: Path, caption: str) -> None:
    if path and Path(path).exists():
        catalog["figures"][key] = {"path": Path(path), "caption": caption}


def _save(img, name: str) -> Path:
    path = ASSETS / name
    img.save(path, "PNG")
    print("fig", path.name)
    return path


def extract_pptx_pictures(pptx_path: Path, prefix: str) -> dict[int, Path]:
    """Return {1-based slide_index: png_path} for first picture on slide."""
    prs = Presentation(str(pptx_path))
    out = {}
    for i, slide in enumerate(prs.slides, 1):
        n = 0
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            n += 1
            blob = shape.image.blob
            ext = shape.image.ext or "png"
            path = ASSETS / f"{prefix}_s{i}_{n}.{ext}"
            path.write_bytes(blob)
            if i not in out:
                out[i] = path
            print("extract", path.name)
    return out


CH1_QUANTUM_TABLE = {
    "caption": "Квантовые числа электрона в атоме",
    "headers": ["Число", "Обозначение", "Смысл", "Возможные значения"],
    "rows": [
        ["Главное", "n", "Уровень / размер", "1, 2, 3, …"],
        ["Орбитальное", "l", "Форма орбитали", "0…n−1 (s, p, d, f, …)"],
        ["Магнитное", "m_l", "Ориентация", "−l … +l"],
        ["Спиновое", "m_s", "Спин электрона", "+1/2, −1/2"],
    ],
}

CH3_LASER_TABLE = {
    "caption": "Сравнение волоконного, диодного и CO₂-лазеров для аддитивных технологий",
    "headers": ["Параметр", "Волоконный", "Диодный", "CO₂"],
    "rows": [
        ["Длина волны λ", "≈ 1070 нм", "808–980 нм", "10,6 мкм"],
        ["КПД η", "30–40 %", "50–70 %", "10–20 %"],
        ["Качество пучка M²", "< 1,5", "> 10 (часто 20–30)", "1–1,5"],
        ["Поглощение Al", "≈ 40 % @ 1070 нм", "— (накачка)", "< 5 % (R > 95 %)"],
        ["Роль в СЛП", "Основной источник", "Накачка волоконных", "Редко; не для Al/Cu"],
    ],
}

CH6_NA_TABLE = {
    "caption": "Три случая числовой апертуры: промышленный, лабораторный и экстремальный",
    "headers": ["Параметр", "A: промышл.", "B: лаб.", "C: экстр."],
    "rows": [
        ["Числ. апертура (NA)", "0,1", "0,4", "0,5"],
        ["Пятно d, мкм", "~10,7", "~2,7", "~2,1"],
        ["DOF, мкм", "~107", "~6,7", "~4,3"],
        ["СЛП", "Да, стандарт", "Только лаб.", "Нереализуемо"],
        ["Почему?", "DOF > рельефа", "DOF < рельефа", "DOF << рельефа"],
    ],
}

EXISTING_FIGURES = [
    ("ch1_atom", "ch1_atom.png", "Эволюция моделей атома и квантовые числа орбиталей"),
    ("ch1_pauli", "ch1_pauli.png", "Принцип Паули и заполнение оболочек: N_max = 2n²"),
    ("ch1_sodium", "ch1_sodium.png", "Пример: переходы в атоме натрия на языке орбиталей и уровней"),
    ("ch1_two_schemes", "ch1_two_schemes.png", "Две схемы: орбитали (пространство) и уровни энергии"),
    ("ch1_points", "ch1_points.png", "Точки на линиях E₁, E₂ — частицы-излучатели, а не электроны"),
    ("ch1_not_always", "ch1_not_always.png", "Переход E₂→E₁ не всегда означает смену орбитали"),
    ("ch2_threshold", "ch2_s11_1.png", "График пороговой плотности энергии для стали, титана и алюминия"),
    ("ch3_spontaneous", "ch3_spontaneous.png", "Спонтанные переходы между уровнями E₂ и E₁"),
    ("ch3_avalanche", "ch3_avalanche.png", "Фотонная лавина: вынужденное излучение и умножение фотонов"),
    ("ch3_inversion_prob", "ch3_inversion_prob.png", "Равенство вероятностей вынужденных переходов Ω₁₂ = Ω₂₁"),
    ("ch3_ruby", "ch3_ruby.png", "Трёхуровневая схема рубинового лазера (ионы Cr³⁺)"),
    ("ch3_pump", "ch3_pump.png", "Эллиптический отражатель: лампа накачки и рубиновый стержень"),
    ("ch3_resonator", "ch3_resonator.png", "Оптический резонатор рубинового лазера"),
    ("ch4_energy", "ch4_energy.png", "Энергия фотона и длина волны: 1070 нм и 355 нм"),
    ("ch4_spectrum", "ch4_spectrum.png", "Шкала электромагнитных излучений и технологические маркеры лазеров"),
    ("ch4_mechanisms", "ch4_mechanisms.png",
     "Два механизма: слева — энергия E и ход реакции (E_ph ≥ E_акт); справа — температура T во времени до T_пл"),
    ("ch4_absorption", "ch4_absorption.png",
     "A(λ) (спектральная поглощательная способность) и четыре сценария: лазер → материал → результат"),
    ("ch4_choice", "ch4_choice.png", "Практические критерии выбора лазера и стратегии для алюминия"),
    # бывшая глава 5 → глава 6 (файлы ch5_*.png сохранены)
    ("ch6_intensity", "ch5_intensity.png", "Плотность мощности I = P/S: два сценария при P = 200 Вт"),
    ("ch6_snell", "ch5_snell.png", "Закон Снеллиуса и сборка параллельных лучей в фокус"),
    ("ch6_spot", "ch5_spot.png", "Дифракционный предел: d ≈ λ/NA"),
    ("ch6_dof", "ch5_dof.png", "Компромисс NA: диаметр пятна d и глубина резкости DOF"),
    ("ch6_m2", "ch5_m2.png", "Параметр M² и реальный диаметр пятна d_реал = M²·d_идеал"),
]


def _write_manifest(catalog: dict) -> None:
    lines = ["# Assets", ""]
    for k, v in catalog["figures"].items():
        lines.append(f"- FIG {k}: {v['path'].name} — {v['caption']}")
    for k, v in catalog["tables"].items():
        lines.append(f"- TAB {k}: {v['caption']}")
    (ASSETS / "manifest.txt").write_text("\n".join(lines), encoding="utf-8")


def prepare_all() -> dict:
    """Build asset catalog used by handbook builder."""
    catalog = {"figures": {}, "tables": {}}

    # Lecture 1 / Chapter 2 — graph from PPTX
    l1_dir = ROOT / "Lecture-1-main"
    if l1_dir.exists():
        l1_pptx = next(l1_dir.glob("Лекция_1_*.pptx"), None)
        if l1_pptx:
            pics1 = extract_pptx_pictures(l1_pptx, "ch2")
            if 11 in pics1:
                catalog["figures"]["ch2_threshold"] = {
                    "path": pics1[11],
                    "caption": "График пороговой плотности энергии для стали, титана и алюминия",
                }

    # Lecture 2 / Chapter 3
    l2_py = ROOT / "Lecture-2-main" / "build_lecture_2.py"
    if l2_py.exists():
        m2 = _load(l2_py)
        figs2 = [
            ("ch3_spontaneous", m2.draw_spontaneous_transitions_pil,
             "Спонтанные переходы между уровнями E₂ и E₁"),
            ("ch3_avalanche", m2.draw_photon_avalanche_pil,
             "Фотонная лавина: вынужденное излучение и умножение фотонов"),
            ("ch3_inversion_prob", m2.draw_inversion_probability_pil,
             "Равенство вероятностей вынужденных переходов Ω₁₂ = Ω₂₁"),
            ("ch3_ruby", m2.draw_ruby_three_level_pil,
             "Трёхуровневая схема рубинового лазера (ионы Cr³⁺)"),
            ("ch3_pump", m2.draw_elliptical_pump_cavity_pil,
             "Эллиптический отражатель: лампа накачки и рубиновый стержень"),
            ("ch3_resonator", m2.draw_ruby_resonator_pil,
             "Оптический резонатор рубинового лазера"),
        ]
        for key, fn, cap in figs2:
            catalog["figures"][key] = {"path": _save(fn(), f"{key}.png"), "caption": cap}

    catalog["tables"]["ch3_laser_compare"] = dict(CH3_LASER_TABLE)

    # Lecture 3 / Chapter 1
    l3_py = ROOT / "Lecture-3-main" / "build_lecture_3.py"
    if l3_py.exists():
        m3 = _load(l3_py)
        figs3 = [
            ("ch1_atom", m3.draw_atom_structure_pil,
             "Эволюция моделей атома и квантовые числа орбиталей"),
            ("ch1_pauli", m3.draw_pauli_shells_pil,
             "Принцип Паули и заполнение оболочек: N_max = 2n²"),
            ("ch1_sodium", m3.draw_sodium_example_pil,
             "Пример: переходы в атоме натрия на языке орбиталей и уровней"),
            ("ch1_two_schemes", m3.draw_two_diagrams_pil,
             "Две схемы: орбитали (пространство) и уровни энергии"),
            ("ch1_points", m3.draw_points_are_particles_pil,
             "Точки на линиях E₁, E₂ — частицы-излучатели, а не электроны"),
            ("ch1_not_always", m3.draw_not_always_orbital_pil,
             "Переход E₂→E₁ не всегда означает смену орбитали"),
        ]
        for key, fn, cap in figs3:
            catalog["figures"][key] = {"path": _save(fn(), f"{key}.png"), "caption": cap}

    catalog["tables"]["ch1_quantum"] = dict(CH1_QUANTUM_TABLE)

    # Lecture 4 / Chapter 4
    l4_py = ROOT / "Lecture-4-main" / "build_lecture_4.py"
    if l4_py.exists():
        m4 = _load(l4_py)
        figs4 = [
            ("ch4_energy", m4.draw_photon_energy_pil,
             "Энергия фотона и длина волны: 1070 нм и 355 нм"),
            ("ch4_spectrum", m4.draw_em_spectrum_pil,
             "Шкала электромагнитных излучений и технологические маркеры лазеров"),
            ("ch4_mechanisms", m4.draw_two_mechanisms_pil,
             "Два механизма: слева — энергия E и ход реакции (E_ph ≥ E_акт); справа — температура T во времени до T_пл"),
            ("ch4_absorption", m4.draw_absorption_scenarios_pil,
             "A(λ) (спектральная поглощательная способность) и четыре сценария: лазер → материал → результат"),
            ("ch4_choice", m4.draw_laser_choice_pil,
             "Практические критерии выбора лазера и стратегии для алюминия"),
        ]
        for key, fn, cap in figs4:
            catalog["figures"][key] = {"path": _save(fn(), f"{key}.png"), "caption": cap}

    # Lecture 5 / Chapter 6 (геометрическая оптика)
    l5_py = ROOT / "Lecture-5-main" / "build_lecture_5.py"
    if l5_py.exists():
        m5 = _load(l5_py)
        figs6 = [
            ("ch6_intensity", m5.draw_intensity_comparison_pil,
             "Плотность мощности I = P/S: два сценария при P = 200 Вт"),
            ("ch6_snell", m5.draw_snell_lens_pil,
             "Закон Снеллиуса и сборка параллельных лучей в фокус"),
            ("ch6_spot", m5.draw_spot_size_pil,
             "Дифракционный предел: d ≈ λ/NA"),
            ("ch6_dof", m5.draw_dof_chart_pil,
             "Компромисс NA: диаметр пятна d и глубина резкости DOF"),
            ("ch6_m2", m5.draw_m2_comparison_pil,
             "Параметр M² и реальный диаметр пятна d_реал = M²·d_идеал"),
        ]
        for key, fn, cap in figs6:
            catalog["figures"][key] = {"path": _save(fn(), f"{key}.png"), "caption": cap}
        catalog["tables"]["ch6_na_table"] = {
            "caption": CH6_NA_TABLE["caption"],
            "headers": list(m5.NA_TABLE["headers"]),
            "rows": [list(row) for row in m5.NA_TABLE["rows"]],
        }
    else:
        catalog["tables"]["ch6_na_table"] = dict(CH6_NA_TABLE)

    # Уже собранные PNG, если исходников лекций 1–5 нет
    for key, filename, caption in EXISTING_FIGURES:
        if key not in catalog["figures"]:
            _add_fig(catalog, key, ASSETS / filename, caption)

    # Lecture 4-доп / Chapter 5 — генерация зелёного излучения
    m4d = load_lecture4dop_module()
    figs5 = [
        ("ch5_copper", m4d.draw_copper_absorption_pil,
         "Спектральная поглощательная способность меди A(λ): 1070 нм и 535 нм"),
        ("ch5_green_gap", m4d.draw_green_gap_pil,
         "Прямая генерация 535 нм на InGaN: «проблема зелёного диапазона» и разрыв мощности"),
        ("ch5_photon_energy", m4d.formula_photon_energy,
         "Энергия фотона E = h·c/λ для λ = 535 нм"),
        ("ch5_asym", m4d.draw_asym_well_pil,
         "Почему в ниобате лития отклик нелинейный: симметричная и асимметричная ямы"),
        ("ch5_shg", m4d.draw_shg_ppln_pil,
         "Генерация второй гармоники (SHG) в кристалле PPLN: 1070 нм (ω) → 535 нм (2ω)"),
        ("ch5_cos2", m4d.formula_cos2,
         "Поле накачки E(t) и квадратичный отклик: cos²(ω·t) = (1 + cos(2ω·t))/2"),
        ("ch5_twopass", m4d.draw_twopass_pil,
         "Двухпроходная схема с активной клиновидной термокомпенсацией"),
        ("ch5_sinc2", m4d.draw_sinc2_pil,
         "Функция sinc²(Δk·L/2): максимум при синхронизме и гашение при рассогласовании фаз"),
        ("ch5_eta", m4d.formula_efficiency,
         "Эффективность преобразования η в приближении неистощённой накачки"),
        ("ch5_economy", m4d.draw_economy_pil,
         "Экономика решения: промышленный зелёный лазер и внешний SHG-модуль"),
    ]
    for key, fn, cap in figs5:
        catalog["figures"][key] = {"path": _save(fn(), f"{key}.png"), "caption": cap}

    def _tab_from(src: dict, caption: str) -> dict:
        return {
            "caption": caption,
            "headers": list(src["headers"]),
            "rows": [list(row) for row in src["rows"]],
        }

    catalog["tables"]["ch5_copper"] = _tab_from(
        m4d.COPPER_TABLE, "Поглощение меди на 1070 нм и 535 нм",
    )
    catalog["tables"]["ch5_green_gap"] = _tab_from(
        m4d.GREEN_GAP_TABLE, "Проблема зелёного диапазона InGaN: параметры",
    )
    catalog["tables"]["ch5_chi"] = _tab_from(
        m4d.CHI_TABLE, "Символы разложения поляризации P(E)",
    )
    catalog["tables"]["ch5_eta"] = _tab_from(
        m4d.ETA_TABLE, "Символы формулы эффективности преобразования η",
    )

    catalog["tables"].pop("ch4_absorption", None)

    _write_manifest(catalog)
    print("assets ready:", ASSETS)
    return catalog


if __name__ == "__main__":
    prepare_all()
