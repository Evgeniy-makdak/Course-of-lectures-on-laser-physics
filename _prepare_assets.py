# -*- coding: utf-8 -*-
"""Generate / extract figures and tables for the print handbook."""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

_ROOTS = [
    Path(r"C:\Users\Volkov\Desktop\Lecture-Unscheduled"),
    Path(__file__).resolve().parent.parent,
]


def _detect_root() -> Path:
    for root in _ROOTS:
        if (root / "Lecture-1-main").exists():
            return root
    return Path(__file__).resolve().parent


ROOT = _detect_root()
ASSETS = Path(__file__).resolve().parent / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)


def _load(py_path: Path):
    name = py_path.stem + "_mod"
    spec = importlib.util.spec_from_file_location(name, py_path)
    mod = importlib.util.module_from_spec(spec)
    # Avoid running __main__
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
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

CH5_NA_TABLE = {
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
    # глава 5 (формирование пятна)
    ("ch5_intensity", "ch5_intensity.png", "Плотность мощности I = P/S: два сценария при P = 200 Вт"),
    ("ch5_rays_waves", "ch5_rays_waves.png", "Два языка оптики: геометрические лучи и дифракционное пятно"),
    ("ch5_spot", "ch5_spot.png", "Дифракционный предел: d ≈ λ/NA"),
    ("ch5_train", "ch5_train.png", "Оптический тракт СЛП-станка: от волокна до порошка"),
    ("ch5_dof", "ch5_dof.png", "Компромисс NA: диаметр пятна d и глубина резкости DOF"),
    ("ch5_focus", "ch5_focus.png", "Положение фокуса относительно слоя порошка и z-offset"),
    ("ch5_m2", "ch5_m2.png", "Параметр M² и реальный диаметр пятна d_реал = M²·d_идеал"),
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

    # Lecture 5 / Chapter 5 (формирование пятна)
    l5_py = ROOT / "Lecture-5-main" / "build_lecture_5.py"
    if l5_py.exists():
        m5 = _load(l5_py)
        figs5 = [
            ("ch5_intensity", m5.draw_intensity_comparison_pil,
             "Плотность мощности I = P/S: два сценария при P = 200 Вт"),
            ("ch5_rays_waves", m5.draw_rays_vs_waves_pil,
             "Два языка оптики: геометрические лучи и дифракционное пятно"),
            ("ch5_spot", m5.draw_spot_size_pil,
             "Дифракционный предел: d ≈ λ/NA"),
            ("ch5_train", m5.draw_optical_train_pil,
             "Оптический тракт СЛП-станка: от волокна до порошка"),
            ("ch5_dof", m5.draw_dof_chart_pil,
             "Компромисс NA: диаметр пятна d и глубина резкости DOF"),
            ("ch5_focus", m5.draw_focus_practice_pil,
             "Положение фокуса относительно слоя порошка и z-offset"),
            ("ch5_m2", m5.draw_m2_comparison_pil,
             "Параметр M² и реальный диаметр пятна d_реал = M²·d_идеал"),
        ]
        for key, fn, cap in figs5:
            catalog["figures"][key] = {"path": _save(fn(), f"{key}.png"), "caption": cap}
        catalog["tables"]["ch5_na_table"] = {
            "caption": CH5_NA_TABLE["caption"],
            "headers": list(m5.NA_TABLE["headers"]),
            "rows": [list(row) for row in m5.NA_TABLE["rows"]],
        }
    else:
        catalog["tables"]["ch5_na_table"] = dict(CH5_NA_TABLE)

    # Уже собранные PNG, если исходников лекций 1–5 нет
    for key, filename, caption in EXISTING_FIGURES:
        if key not in catalog["figures"]:
            _add_fig(catalog, key, ASSETS / filename, caption)

    catalog["tables"].pop("ch4_absorption", None)

    _write_manifest(catalog)
    print("assets ready:", ASSETS)
    return catalog


if __name__ == "__main__":
    prepare_all()
