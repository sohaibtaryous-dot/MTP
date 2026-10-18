# -*- coding: utf-8 -*-
"""Conversion des poids vers le kilogramme (g, mg, t, lb, oz…) et lecture « nombre + unité »."""

from __future__ import annotations

import re

# facteur -> kg
WEIGHT_FACTORS = {
    "kg": 1.0, "kgs": 1.0, "kilo": 1.0, "kilos": 1.0, "kilogram": 1.0, "kilograms": 1.0, "kilogramme": 1.0,
    "kilogrammes": 1.0, "kilogramm": 1.0, "kgm": 1.0,
    "g": 0.001, "gr": 0.001, "gm": 0.001, "gms": 0.001, "gram": 0.001, "grams": 0.001, "gramme": 0.001,
    "grammes": 0.001, "gramm": 0.001, "grs": 0.001,
    "mg": 1e-6, "milligram": 1e-6, "milligrams": 1e-6, "milligramme": 1e-6, "milligrammes": 1e-6,
    "t": 1000.0, "to": 1000.0, "tn": 1000.0, "ton": 1000.0, "tons": 1000.0, "tonne": 1000.0, "tonnes": 1000.0,
    "tonnen": 1000.0, "tonelada": 1000.0, "toneladas": 1000.0, "mt": 1000.0,
    "lb": 0.45359237, "lbs": 0.45359237, "pound": 0.45359237, "pounds": 0.45359237, "livre": 0.45359237,
    "livres": 0.45359237,
    "oz": 0.028349523125, "ounce": 0.028349523125, "ounces": 0.028349523125, "once": 0.028349523125,
    "onces": 0.028349523125,
}
# Unités acceptées après un nombre dans un contexte de poids (le « t » et le « g » seuls ne sont acceptés que collés)
UNIT_ALT = "|".join(sorted((re.escape(u) for u in WEIGHT_FACTORS), key=len, reverse=True))
NUM_UNIT_RE = re.compile(rf"(?P<num>\d[\d.,]*(?:\s\d{{3}}(?![\d]))*)\s*(?P<unit>{UNIT_ALT})\b", re.I)
HEADER_UNIT_RE = re.compile(rf"[\(\[/]\s*(?P<unit>{UNIT_ALT})\s*[\)\]]?", re.I)


def unit_factor(unit: str | None) -> float | None:
    """Facteur de conversion vers le kg (None si l'unité n'est pas une unité de poids)."""
    if not unit:
        return None
    return WEIGHT_FACTORS.get(unit.strip().strip(".").lower())


def to_kg(value: float | None, unit: str | None) -> float | None:
    """Convertit une valeur dans son unité ; sans unité ou unité inconnue, la valeur est supposée en kg."""
    if value is None:
        return None
    f = unit_factor(unit)
    return round(value * (f if f is not None else 1.0), 6)


def unit_from_header(label: str) -> str:
    """Unité de poids annoncée dans un en-tête de colonne : « Net weight (g) », « Poids [lbs] », « Gewicht / t »."""
    m = HEADER_UNIT_RE.search(str(label or ""))
    if m and unit_factor(m.group("unit")) is not None:
        return m.group("unit").lower()
    return ""


def split_number_unit(text: str) -> tuple[str, str]:
    """« 12,5 lbs » -> ('12,5', 'lbs') ; sans unité -> (texte, '')."""
    s = str(text or "").strip()
    m = NUM_UNIT_RE.fullmatch(s) or NUM_UNIT_RE.match(s)
    if m and unit_factor(m.group("unit")) is not None:
        return m.group("num").strip(), m.group("unit").lower()
    return s, ""


def describe(unit: str) -> str:
    """Libellé court pour la colonne « Base / remarque poids » (ex. « converti de lbs »)."""
    f = unit_factor(unit)
    return "" if f is None or f == 1.0 else f"converti de {unit.lower()}"
