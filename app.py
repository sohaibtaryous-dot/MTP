# -*- coding: utf-8 -*-
"""
TIMAR Tanger - Ventilation Automatique Groupage Alstom Fez
==========================================================

Application Web Streamlit (aucun droit administrateur requis) qui :
  1. lit un dossier PDF de remorque de groupage (factures, packing lists,
     CMR, manifeste), y compris les pages scannées (OCR) ;
  2. extrait chaque ligne de facture de manière STRICTEMENT individuelle ;
  3. rapproche chaque référence avec le fichier Excel de référence (NGP) ;
  4. génère un classeur Excel à 3 feuilles : TOUS_CLIENTS, CONTROLE, RECAP_GLOBAL.

Lancement :  streamlit run app.py
"""

from __future__ import annotations

import difflib
import io
import math
import os
import re
import shutil
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st

import cleaning
import fileio
import ui
import units
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:  # PyPDF2 : extraction de secours si pdfplumber échoue sur une page
    from PyPDF2 import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None

try:  # lecture IA multilingue (facultative : paquet « anthropic »)
    import ai_engine
except Exception:  # pragma: no cover
    ai_engine = None


# =============================================================================
# CONSTANTES
# =============================================================================

APP_TITLE = "TIMAR Tanger - Ventilation Automatique Groupage Alstom Fez"
NON_TROUVE = "NON TROUVÉ"
REMARQUE_NOUVELLE_REF = "Nouveau Réf - NGP Facture appliqué"

PAGE_IN_PDF = "Page dans ce PDF"          # l'en-tête Excel devient « Page dans ce PDF (N p.) »
C_NET, C_NETU = "Poids net ligne (kg)", "Poids net unitaire imprimé (kg/pièce)"
C_NETF, C_GROSSF = "Poids net facture (kg)", "Poids brut facture (kg)"
C_COLIS, C_MANIF, C_BASE = "Nb colis / palettes (facture)", "Poids brut manifest – shipment (kg)", "Base / remarque poids"
COLUMNS = [
    "Client", "Job Ref manifest", "N° facture", "Date", "Item",
    "Réf article", "Désignation", "Origine",
    "Qté", "Prix unitaire", "Valeur HT",
    "NGP/code douanier (facture)", "NGP (REF_PAR_FRS)", "Onglet REF_PAR_FRS",
    "Contrôle NGP", "Remarque", "Page PDF", PAGE_IN_PDF,
    C_NET, C_NETU, C_NETF, C_GROSSF, C_COLIS, C_MANIF, C_BASE,
]
NUMERIC_COLUMNS = ["Qté", "Prix unitaire", "Valeur HT", C_NET, C_NETU, C_NETF, C_GROSSF, C_COLIS, C_MANIF]
INT_COLUMNS = ["Page PDF", PAGE_IN_PDF]
INTERNAL_COLUMNS = ["_réf_alstom", "_devise", "_source", "_inv", "_fee"]
# Colonnes saisies / corrigeables par l'utilisateur (le matching est recalculé)
EDITABLE_COLUMNS = [c for c in COLUMNS if c not in ("NGP (REF_PAR_FRS)", "Contrôle NGP", "Onglet REF_PAR_FRS")]
MANIFEST_COLUMNS = ["Page PDF", "Shipment", "Expéditeur", "Poids (kg)", "Colis", "Goods Value", "Devise"]

CURRENCIES = ["EUR", "USD", "GBP", "CHF", "CNY", "JPY", "SEK", "DKK", "NOK",
              "PLN", "CZK", "HUF", "TRY", "CAD", "MAD", "INR"]
INCOTERMS = ["EXW", "FCA", "FAS", "FOB", "CFR", "CIF", "CPT", "CIP", "DAP",
             "DPU", "DAT", "DDP", "DDU"]
ISO_COUNTRIES = set("""AD AE AF AL AM AO AR AT AU AZ BA BD BE BG BH BR BY CA CH CL CN CO CR CY CZ DE
DK DZ EC EE EG ES FI FR GB GE GR HK HR HU ID IE IL IN IQ IR IS IT JO JP KE KR KW KZ LB LK LT LU LV
LY MA MC MD ME MK MT MX MY NG NL NO NZ OM PE PH PK PL PT QA RO RS RU SA SE SG SI SK SN TH TN TR TW
UA US UY UZ VN ZA""".split())
COUNTRY_NAMES = {
    "GERMANY": "DE", "ALLEMAGNE": "DE", "DEUTSCHLAND": "DE", "FRANCE": "FR",
    "ITALY": "IT", "ITALIE": "IT", "ITALIA": "IT", "SPAIN": "ES", "ESPAGNE": "ES",
    "ESPANA": "ES", "CHINA": "CN", "CHINE": "CN", "UNITED STATES": "US", "USA": "US",
    "ETATS-UNIS": "US", "UNITED KINGDOM": "GB", "ROYAUME-UNI": "GB", "UK": "GB",
    "POLAND": "PL", "POLOGNE": "PL", "CZECH REPUBLIC": "CZ", "CZECHIA": "CZ",
    "TURKEY": "TR", "TURQUIE": "TR", "JAPAN": "JP", "JAPON": "JP", "INDIA": "IN",
    "INDE": "IN", "SWITZERLAND": "CH", "SUISSE": "CH", "BELGIUM": "BE",
    "BELGIQUE": "BE", "NETHERLANDS": "NL", "PAYS-BAS": "NL", "AUSTRIA": "AT",
    "AUTRICHE": "AT", "SWEDEN": "SE", "SUEDE": "SE", "PORTUGAL": "PT",
    "MOROCCO": "MA", "MAROC": "MA", "ROMANIA": "RO", "ROUMANIE": "RO",
    "HUNGARY": "HU", "HONGRIE": "HU", "SLOVAKIA": "SK", "SLOVAQUIE": "SK",
    "KOREA": "KR", "COREE": "KR", "TAIWAN": "TW", "MEXICO": "MX", "MEXIQUE": "MX",
    "EU": "EU", "EUROPEAN UNION": "EU", "UNION EUROPEENNE": "EU",
}
UNITS = {"PCS", "PC", "PCE", "PCES", "EA", "EACH", "ST", "STK", "STCK", "UN", "U",
         "UNIT", "UNITS", "PZ", "PZA", "KG", "KGS", "G", "M", "MM", "M2", "M3",
         "L", "LTR", "SET", "SETS", "KIT", "BOX", "ROL", "ROLL", "PR", "PAIR",
         "PAA", "LOT", "PK", "PAC", "PACK", "MTR", "NR", "NOS"}

FEE_WORDS = re.compile(
    r"\b(freight|fret|transport|shipping|carriage|packing\s*cost|packaging|emballage|"
    r"frais|surcharge|insurance|assurance|handling|verpackung|fracht|porto|"
    r"palette\s*charge|pallet\s*charge|discount|remise|rabatt)\b", re.I)
TOTAL_WORDS = re.compile(
    r"\b(sub\s*-?\s*total|total|sous[- ]?total|summe|gesamt\w*|zwischensumme|"
    r"net\s*amount|amount\s*due|montant\s*(?:net|total|ht|ttc)|tva|vat|mwst|"
    r"carried\s*forward|report|übertrag|a\s*reporter)\b", re.I)

# Mapping des en-têtes de colonnes (FR / EN / DE / ES / IT) -> champ interne.
# L'ordre est important : le premier motif qui correspond l'emporte.
LINE_COLMAP = [
    ("ngp", r"\b(hs|h\.s\.|ngp|tarif\w*|tariff|customs|douan\w*|zoll\w*|taric|commodity|nomencl\w*|statist\w*)"),
    ("po", r"(\bp\.?\s?o\.?\b(?!\s*box)|purchase\s*order|order\s*(no|n°|nr|number|ref)|commande|bestell\w*|\bpo\s*(no|n°|number))"),
    ("alstom", r"(alstom|customer|client|your\s*(ref|part|p/?n|art)|votre|kunden|cust\.?\s*p)"),
    ("net_unit", r"(unit\w*\s*(net\s*)?(weight|wt|mass)|poids\s*(net\s*)?unitaire|kg\s*/\s*(pc|pce|st|stk|pi[eè]ce)|"
                 r"gewicht\s*(pro|/|je)\s*st\w*|peso\s*unit\w*)"),
    ("net", r"(net\w*\s*(weight|wt|kg|mass)|poids\s*net|netto\w*|n\.?\s?w\.?\b|peso\s*neto|peso\s*netto)"),
    ("gross", r"(gross\s*(weight|wt|kg|mass)|poids\s*brut|brutto\w*|g\.?\s?w\.?\b|peso\s*bruto|peso\s*lordo)"),
    ("price", r"(unit\s*price|prix\s*unit\w*|p\.\s?u\.?|^pu$|price|prix|preis|precio|prezzo|einzel\w*|rate)"),
    ("amount", r"(amount|montant|total|valeur|value|betrag|importe|importo|gesamt|line\s*total|ext\w*\s*price)"),
    ("qty", r"(qty|q\.?t[eé]|quantit\w*|menge|cantidad|qt[aà]|anzahl|^pcs$|^nb$)"),
    ("unit", r"^(unit|unit[eé]|uom|u\.?m\.?|me|einheit|unidad|um|u\.?)$"),
    ("origin", r"(origin\w*|origine|coo\b|ursprung\w*|c\.?o\.?o\.?|made\s*in|pays|herkunft)"),
    ("brand", r"(brand|marque|manufacturer|fabricant|hersteller|make\b|marca)"),
    ("packages", r"(colis|packages?|pkgs?|cartons?|packst\w*|bultos|nb\s*col)"),
    ("desc", r"(descr\w*|d[eé]sign\w*|bezeichnung|libell\w*|denominaz\w*|item\s*name|goods|text)"),
    ("ref", r"(part|article|art\.?\s*(no|nr|n°)?|r[eé]f\w*|material|item\s*(no|code|number|#)|code|artikel|product|sku|p/?n|mod[eè]l\w*|modell\w*)"),
    ("pos", r"^(pos\w*|position|item|line|ligne|poste|no\.?|n°|#|nr\.?|lfd\.?|sr\.?\s*no\.?|s/n)$"),
]
MANIFEST_COLMAP = [
    ("shipment", r"(shipment|job\s*ref|bl\s*no|reference|ref\.?\s*exp|waybill|sxvf|\bawb\b|\bbno\b)"),
    ("invoice", r"(invoice|facture|inv\b|rechnung|factura)"),
    ("currency", r"(curr\w*|devise|cur\b|monnaie|w[aä]hrung)"),
    ("gross", r"(gross|brut|poids|weight|kg|gewicht)"),
    ("packages", r"(colis|pack\w*|pkgs?|pallets?|palettes?|cartons?|nb\b|qty|nombre)"),
    ("value", r"(value|valeur|montant|amount|goods|wert|prix|price)"),
    ("supplier", r"(shipper|supplier|fournisseur|exp[eé]diteur|exporter|vendor|lieferant|consignor|sender|soci[eé]t[eé]|company|name|nom)"),
]
REF_SHEET_COLMAP = [
    ("ngp", r"(ngp|\bhs\b|h\.s\.|nomencl\w*|tarif\w*|code\s*douan\w*|customs|\bsh\b|taric|commodity)"),
    ("alstom", r"(alstom|\bpn\b|p/n|part\s*number\s*alstom|client|customer)"),
    ("ref", r"(r[eé]f\w*|reference|article|part|material|code|artikel|item|sku)"),
]

HS_LABEL_RE = re.compile(
    r"(?:\bhs\b|h\.s\.|\bngp\b|tarif\w*|customs\s*(?:tariff|code)?|commodity\s*code|"
    r"code\s*douan\w*|zoll\w*|taric|nomenclature|statistical\s*n\w*)"
    r"\s*(?:code|no\.?|nr\.?|n°|number)?\s*[:#.]?\s*(\d[\d .]{5,14}\d)", re.I)
ORIGIN_LABEL_RE = re.compile(
    r"(?:country\s*of\s*origin|pays\s*d'?\s*origine|origin\w*|origine|ursprung\w*|\bcoo\b|made\s*in)"
    r"\s*[:#.]?\s*([A-Za-z][A-Za-z\- ]{1,25})", re.I)
ALSTOM_LABEL_RE = re.compile(
    r"(?:alstom\s*(?:p/?n|part\s*(?:no|number)?|r[eé]f\w*|code|art\w*)?|customer\s*(?:p/?n|part\s*(?:no\.?|number)?|r[eé]f\w*|item)|"
    r"your\s*(?:p/?n|part\s*(?:no\.?|number)?|r[eé]f\w*|item)|votre\s*r[eé]f\w*|r[eé]f\.?\s*client|kunden\w*)"
    r"\s*[:#.]?\s*([A-Z0-9][A-Z0-9\-_./]{2,})", re.I)
BRAND_LABEL_RE = re.compile(
    r"(?:brand|marque|manufacturer|fabricant|hersteller|marca)\s*[:#]\s*([A-Za-z0-9&\-. ]{2,30})", re.I)
WEIGHT_UNIT = r"(?:kgs?|kilo\\w*|grammes?|grams?|gr|gms?|g|mg|tonnes?|tons?|tn|t|lbs?|pounds?|oz)"
LINE_NET_RE = re.compile(r"(?:(?<!unit\s)net\s*weight|poids\s*net|nettogewicht|n\.w\.)\s*(?:\((?P<hu>" + WEIGHT_UNIT + r")\))?\s*[:=]?\s*"
                         r"(?P<num>[\d][\d.,]*)\s*(?P<unit>" + WEIGHT_UNIT + r")?\b", re.I)
LINE_GROSS_RE = re.compile(r"(?:gross\s*weight|poids\s*brut|bruttogewicht|g\.w\.)\s*(?:\((?P<hu>" + WEIGHT_UNIT + r")\))?\s*[:=]?\s*"
                           r"(?P<num>[\d][\d.,]*)\s*(?P<unit>" + WEIGHT_UNIT + r")?\b", re.I)
LINE_UNITW_RE = re.compile(r"(?:unit\s*(?:net\s*)?weight|poids\s*(?:net\s*)?unitaire|gewicht\s*(?:pro|je)\s*st[üu]ck)\s*"
                           r"(?:\((?P<hu>" + WEIGHT_UNIT + r")(?:\s*/\s*\w+)?\))?\s*[:=]?\s*(?P<num>[\d][\d.,]*)\s*"
                           r"(?P<unit>" + WEIGHT_UNIT + r")?\b", re.I)

INV_NO_RE = re.compile(
    r"(?:commercial\s+invoice|invoice|facture(?:\s+commerciale)?|rechnung|factura|fattura|inv\.)"
    r"\s*(?:no\.?|n[°ºo]\.?|nr\.?|number|num[ée]ro|#|nummer)?\s*[:.#]?\s*"
    r"([A-Z0-9][A-Z0-9\-/_.]*\d[A-Z0-9\-/_]*)", re.I)
INV_LABEL_ONLY_RE = re.compile(
    r"(?:invoice|facture|rechnung|factura|fattura)\s*(?:no\.?|n[°º]\.?|nr\.?|number|num[ée]ro|#|nummer)", re.I)
BL_NO_RE = re.compile(
    r"(?:delivery\s*note|packing\s*list|bon\s*de\s*livraison|liste\s*de\s*colisage|lieferschein|"
    r"albar[aá]n|\bb\.?l\.?\b|\bd\.?n\.?\b|\bp\.?l\.?\b|shipment|exp[ée]dition|despatch\s*note|dispatch\s*note)"
    r"\s*(?:no\.?|n[°º]\.?|nr\.?|number|num[ée]ro|#|nummer)\s*[:.#]?\s*"
    r"([A-Z0-9][A-Z0-9\-/_.]*\d[A-Z0-9\-/_]*)", re.I)
BL_NO_RE2 = re.compile(
    r"(?:delivery\s*note|packing\s*list|bon\s*de\s*livraison|lieferschein|despatch\s*note|dispatch\s*note)"
    r"\s*[:.#]?\s*([A-Z0-9][A-Z0-9\-/_]*\d[A-Z0-9\-/_]*)", re.I)
MONTHS = r"(?:jan|feb|f[eé]v|mar|m[aä]r|mrz|apr|avr|abr|may|mai|mag|jun|juin|giu|jul|juil|lug|aug|ao[uû]|ago|sep|set|oct|okt|ott|nov|d[eé]c|dec|dez|dic|ene|gen)"
DATE_RE = re.compile(
    r"\b(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}(?:st|nd|rd|th|er)?[\s./-]+" + MONTHS +
    r"[a-zéûäöüñ]*\.?[\s./,-]+\d{2,4}|" + MONTHS + r"[a-zéû]*\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})\b", re.I)
_UNIT_LABEL = r"(?:\((?:" + WEIGHT_UNIT + r")\)|" + WEIGHT_UNIT + r"|in\s*kg)?"
TOTAL_NET_RE = re.compile(
    r"(?:total\s*)?(?:net\s*weight|poids\s*net(?:\s*total)?|nettogewicht|peso\s*neto|net\s*wt\.?|n\.\s?w\.)"
    r"\s*(?:\((?P<hu>" + WEIGHT_UNIT + r")\)|" + WEIGHT_UNIT + r"|in\s*kg)?\s*[:=]?\s*(?P<num>\d[\d.,\s]*\d|\d)\s*(?P<unit>" + WEIGHT_UNIT + r")?\b", re.I)
TOTAL_GROSS_RE = re.compile(
    r"(?:total\s*)?(?:gross\s*weight|poids\s*brut(?:\s*total)?|bruttogewicht|peso\s*bruto|gross\s*wt\.?|g\.\s?w\.|"
    r"(?<!net\s)(?<!net)total\s*weight|poids\s*total|gesamtgewicht|peso\s*total)"
    r"\s*(?:\((?P<hu>" + WEIGHT_UNIT + r")\)|" + WEIGHT_UNIT + r"|in\s*kg)?\s*[:=]?\s*(?P<num>\d[\d.,\s]*\d|\d)\s*(?P<unit>" + WEIGHT_UNIT + r")?\b", re.I)
PACKAGES_RE = re.compile(
    r"(?:(?:total\s*)?(?:number\s*of\s*packages|no\.?\s*of\s*packages|packages|nombre\s*de\s*colis|nb\.?\s*colis|colis|"
    r"packst[üu]cke|bultos|pallets?|palettes?|cartons?)\s*[:=]?\s*(\d{1,4})\b)|"
    r"(?:\b(\d{1,4})\s*(?:colis|packages|pallets?|palettes?|cartons?|cases|caisses|crates)\b)", re.I)
SUPPLIER_LABEL_RE = re.compile(
    r"^\s*(?:seller|supplier|vendor|fournisseur|exp[ée]diteur|shipper|exporter|exportateur|lieferant|"
    r"verk[äa]ufer|absender|proveedor|exportador|fornitore|venditore|consignor|sender|"
    r"leverancier|leverancier|dostawca|dodavatel|fornecedor|ihracat[çc][ıi]|prod[áa]vaj[íi]c[íi]|"
    r"satici|tedarik[çc]i|izvo[žz]nik|prodavec|produc[aă]tor|izvoznik)"
    r"\s*(?:/\s*\w+\s*)?[:\-]?\s*(.*)$", re.I)
CLIENT_LABEL_RE = re.compile(
    r"^\s*(?:buyer|customer|consignee|bill\s*[:\-]?\s*to|ship\s*[:\-]?\s*to|sold\s*[:\-]?\s*to|"
    r"invoice\s*[:\-]?\s*to|deliver(?:y)?\s*[:\-]?\s*to|"
    r"client|acheteur|destinataire|livr[eé]\s*[àa]|factur[eé]\s*[àa]|import(?:ateur)?|"
    r"k[äa]ufer|empf[äa]nger|kunde|rechnungsempf[äa]nger|lieferadresse|warenempf[äa]nger|"
    r"comprador|consignatario|cliente|direcci[oó]n\s*de\s*(?:entrega|factura)|"
    r"acquirente|destinatario|importatore|"
    r"koper|ontvanger|afnemer|factuuradres|"
    r"nabywca|odbiorca|zamawiaj[aą]cy|"
    r"kupuj[íi]c[íi]|p[řr][íi]jemce|odb[eě]ratel|"
    r"comprador|destinat[áa]rio|importador|"
    r"al[ıi]c[ıi]|teslim\s*al[ıi]c[ıi]|ithalat[çc][ıi]|"
    r"покупатель|получатель|импортёр)"
    r"\s*(?:/\s*\w+\s*)?[:\-]?\s*(.*)$", re.I)
LEGAL_RE = re.compile(
    r"\b(GmbH|AG|S\.?A\.?S\.?U?|S\.?A\.?|S\.?p\.?A\.?|S\.?r\.?l\.?|SARL|EURL|Ltd\.?|Limited|LLC|Inc\.?|Corp\.?|"
    r"B\.?V\.?|N\.?V\.?|S\.?L\.?U?|KG|Co\.?\s*KG|Oy|AB|A/S|ApS|Sp\.?\s*z\s*o\.?\s*o\.?|s\.?r\.?o\.?|a\.?s\.?|"
    r"PLC|SE|S\.?C\.?A\.?|Kft\.?|Industries|Group)\b")
CONSIGNEE_WORDS = re.compile(
    r"(alstom|timar|ship\s*to|bill\s*to|invoice\s*to|sold\s*to|deliver\w*\s*to|consignee|destinataire|"
    r"client|customer|buyer|acheteur|empf[äa]nger|livr[ée]\s*[àa])", re.I)


# =============================================================================
# OUTILS GÉNÉRIQUES
# =============================================================================

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def clean_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    s = str(v).replace("\u00a0", " ").replace("\u202f", " ").strip()
    if s.lower() in ("nan", "none", "nat"):
        return ""
    return s


_DECIMAL_DEFAULT = ["."]


def decimal_votes(text: str) -> tuple[int, int]:
    """(votes virgule, votes point). « 0,216 » est forcément décimal (un séparateur de milliers ne suit jamais un 0)."""
    comma = len(re.findall(r"\d,\d{2}(?!\d)(?!,\d)", text)) \
        + 3 * len(re.findall(r"(?<![\d.,])0,\d{3}(?!\d)", text))
    dot = len(re.findall(r"\d\.\d{2}(?!\d)(?!\.\d)", text)) \
        + 3 * len(re.findall(r"(?<![\d.,])0\.\d{3}(?!\d)", text))
    return comma, dot


def set_decimal_default(texts) -> str:
    """Convention du dossier entier, utilisée quand un document n'a aucun indice (ex. BL : « 148,000 KG »)."""
    comma = dot = 0
    for t in texts:
        c, d = decimal_votes(t)
        comma, dot = comma + c, dot + d
    _DECIMAL_DEFAULT[0] = "," if comma > dot else "."
    return _DECIMAL_DEFAULT[0]


def detect_decimal(text: str) -> str:
    """Séparateur décimal d'un document (',' européen ou '.' anglo-saxon) ; à égalité : convention du dossier."""
    comma, dot = decimal_votes(text)
    if comma != dot:
        return "," if comma > dot else "."
    return _DECIMAL_DEFAULT[0]


def parse_number(value, decimal: str = ".") -> float | None:
    """Convertit '1.234,56' / '1,234.56' / '1 234,56' / 12.5 en float."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if (isinstance(value, float) and math.isnan(value)) else float(value)
    t = clean_cell(value)
    if not t:
        return None
    t = re.sub(r"(?i)(EUR|USD|GBP|CHF|MAD|€|\$|£|kgs?|pcs|%)", "", t).strip()
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1]
    if t.endswith("-"):
        neg, t = True, t[:-1]
    if t.startswith("-"):
        neg, t = True, t[1:]
    t = t.replace(" ", "").replace("'", "")
    if not t or not re.fullmatch(r"\d[\d.,]*", t) or t[-1] in ".,":
        return None
    has_dot, has_comma = "." in t, "," in t
    if has_dot and has_comma:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif has_comma or has_dot:
        sep = "," if has_comma else "."
        parts = t.split(sep)
        if len(parts) > 2:  # plusieurs séparateurs identiques = milliers
            if not all(len(p) == 3 for p in parts[1:]):
                return None
            t = t.replace(sep, "")
        elif sep == decimal:
            t = t.replace(",", ".")
        elif len(parts[1]) == 3 and len(parts[0]) <= 3 and parts[0] != "0":
            # milliers : ex. "1,204" ou "12,345" — seulement si la partie gauche ≤3 chiffres
            # "1204,023" ou "3500,000" → décimale (partie gauche > 3 chiffres)
            t = t.replace(sep, "")
        else:
            t = t.replace(",", ".")
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


def normalize_ngp(value) -> str:
    """Normalise un code NGP/HS : supprime espaces, points, tirets -> 8 ou 10 chiffres."""
    s = clean_cell(value)
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0+", s):  # nombre lu par Excel (84818099.0)
        s = s.split(".")[0]
    digits = re.sub(r"\D", "", s)
    if len(digits) < 6:
        return ""
    if len(digits) == 7:      # zéro de tête perdu par Excel (chapitres 01 à 09)
        digits = digits.zfill(8)
    elif len(digits) == 9:
        digits = digits.zfill(10)
    elif len(digits) > 10:
        digits = digits[:10]
    return digits


def norm_ref(value) -> str:
    s = clean_cell(value)
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return re.sub(r"\s+", "", s).upper()


def compact_ref(value) -> str:
    return re.sub(r"[^A-Z0-9]", "", norm_ref(value))


def ocr_ref(value) -> str:
    """Clé tolérante aux confusions OCR courantes (O/0, I/L/1, S/5, B/8)."""
    return compact_ref(value).translate(str.maketrans("OILSB", "01158"))


def norm_company(name: str) -> str:
    s = strip_accents(clean_cell(name)).upper()
    s = LEGAL_RE.sub(" ", s) if s else s
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = re.sub(r"\b(GMBH|AG|SA|SAS|SPA|SRL|SARL|LTD|LIMITED|LLC|INC|CORP|BV|NV|SL|KG|CO|THE|ET|AND|CIE)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def company_similarity(a: str, b: str) -> float:
    na, nb = norm_company(a), norm_company(b)
    if not na or not nb:
        return 0.0
    if na == nb or na in nb or nb in na:
        return 1.0
    ta, tb = set(na.split()), set(nb.split())
    jacc = len(ta & tb) / max(1, len(ta | tb))
    return max(difflib.SequenceMatcher(None, na, nb).ratio(), jacc)


MONTH_NUM = {"jan": 1, "feb": 2, "fev": 2, "mar": 3, "mrz": 3, "apr": 4, "avr": 4, "abr": 4, "may": 5, "mai": 5, "mag": 5,
             "jun": 6, "giu": 6, "jul": 7, "lug": 7, "aug": 8, "aou": 8, "ago": 8, "sep": 9, "set": 9, "oct": 10,
             "okt": 10, "ott": 10, "nov": 11, "dec": 12, "dez": 12, "dic": 12, "ene": 1, "gen": 1}


def _month_from_word(word: str) -> int | None:
    w = strip_accents(word).lower().strip(".")
    if w.startswith("juil"):
        return 7
    if w.startswith("juin"):
        return 6
    return MONTH_NUM.get(w[:3])


def _make_date(d, m, y) -> str | None:
    try:
        y = int(y)
        y = y + (1900 if y >= 70 else 2000) if y < 100 else y
        return datetime(y, int(m), int(d)).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return None


def normalize_date(raw: str) -> str:
    """Date -> JJ/MM/AAAA (formats numériques, ISO avec heure, mois en toutes lettres FR/EN/DE/ES/IT)."""
    raw = clean_cell(raw)
    if not raw:
        return ""
    t = re.sub(r"[T\s]+\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?$", "", raw.strip())
    m = re.fullmatch(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", t)
    if m:
        return _make_date(m.group(3), m.group(2), m.group(1)) or raw
    m = re.fullmatch(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})", t)
    if m:
        d, mo, y = m.groups()
        if int(mo) > 12 >= int(d):      # 10/13/2026 : format américain mois/jour
            d, mo = mo, d
        return _make_date(d, mo, y) or raw
    m = re.fullmatch(r"(\d{1,2})(?:st|nd|rd|th|er)?[\s./-]*([A-Za-zÀ-ÿ]{3,9})\.?[\s./,-]*(\d{2,4})", t)
    if m and _month_from_word(m.group(2)):
        return _make_date(m.group(1), _month_from_word(m.group(2)), m.group(3)) or raw
    m = re.fullmatch(r"([A-Za-zÀ-ÿ]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", t)
    if m and _month_from_word(m.group(1)):
        return _make_date(m.group(2), _month_from_word(m.group(1)), m.group(3)) or raw
    return raw


ORIGIN_FILLER = {"DE", "LA", "LE", "LES", "DU", "DES", "L", "OF", "THE", "GOODS", "MARCHANDISE", "MARCHANDISES",
                 "PRODUIT", "PRODUITS", "ORIGINE", "ORIGIN", "WARE", "WAREN", "ET", "AND"}


def normalize_origin(raw: str) -> str:
    s = strip_accents(clean_cell(raw)).upper().strip(" .:-")
    if not s or s.split()[0] in ORIGIN_FILLER:   # « de la marchandise » : mots de liaison, pas un pays
        return ""
    if s in ISO_COUNTRIES or s == "EU":
        return s
    for name, code in COUNTRY_NAMES.items():
        if s.startswith(name):
            return code
    first = s.split()[0]
    if first in ISO_COUNTRIES:
        return first
    return s[:20]


ORIGIN_LINE_RE = re.compile(
    r"(?:country\s*of\s*origin|pays\s*d'?\s*origine|origin\w*\s*of\s*goods|origine\s*des\s*marchandises|ursprungsland|"
    r"herkunftsland|paese\s*d'?\s*origine|pa[ií]s\s*de\s*origen)"
    r"(?:\s*(?:de\s*la\s*marchandise|des\s*marchandises|of\s*(?:the\s*)?goods|du\s*produit))?", re.I)


def _first_country(text: str) -> str:
    toks = re.findall(r"[A-Za-zÀ-ÿ'\-]+", text)[:8]
    for i, tok in enumerate(toks):
        t = strip_accents(tok).upper()
        if t in INCOTERMS:
            continue
        two = f"{t} {strip_accents(toks[i + 1]).upper()}" if i + 1 < len(toks) else ""
        if two in COUNTRY_NAMES:
            return COUNTRY_NAMES[two]
        if t in COUNTRY_NAMES:
            return COUNTRY_NAMES[t]
        if len(tok) == 2 and tok.isupper() and t in ISO_COUNTRIES and t != "DE" or (t == "DE" and tok == "DE" and len(toks) == 1):
            return t
    return ""


def find_origin(text: str) -> str:
    """Pays d'origine de la marchandise : « Country of origin: China », ou en-tête de colonne puis valeur à la ligne suivante."""
    lines = text_lines(text)
    for i, ln in enumerate(lines):
        m = ORIGIN_LINE_RE.search(ln)
        if not m:
            continue
        cand = _first_country(ln[m.end():].strip(" :#.-"))
        if not cand and i + 1 < len(lines):
            cand = _first_country(lines[i + 1])
        if cand:
            return cand
    return ""


def text_lines(text: str) -> list[str]:
    return [re.sub(r"[ \t]+", " ", ln).strip() for ln in (text or "").splitlines() if ln.strip()]


# =============================================================================
# LECTURE PDF + OCR
# =============================================================================

def tesseract_ready(custom_path: str = "") -> bool:
    try:
        import pytesseract
    except Exception:
        return False
    if custom_path:
        pytesseract.pytesseract.tesseract_cmd = custom_path
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def easyocr_ready() -> bool:
    try:
        import easyocr  # noqa: F401
        return True
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def get_easyocr_reader(langs: tuple):
    import easyocr
    return easyocr.Reader(list(langs), gpu=False)


def remove_table_lines(image, min_frac: float = 0.04):
    """Prétraitement OCR : niveaux de gris + effacement des filets de tableau (traits longs
    horizontaux / verticaux) qui perturbent la reconnaissance des caractères."""
    import numpy as np
    from PIL import Image
    gray = np.array(image.convert("L"))
    dark = gray < 160
    out = gray.copy()
    for axis in (1, 0):
        k = max(25, int(dark.shape[axis] * min_frac))
        arr = dark if axis == 1 else dark.T
        cs = np.cumsum(np.pad(arr.astype(np.int32), ((0, 0), (1, 0))), axis=1)
        full = (cs[:, k:] - cs[:, :-k]) == k          # fenêtre entièrement noire
        mark = np.zeros(arr.shape[1] + 1, dtype=np.int32)[None, :].repeat(arr.shape[0], 0)
        idx = np.nonzero(full)
        np.add.at(mark, (idx[0], idx[1]), 1)
        np.add.at(mark, (idx[0], idx[1] + k), -1)
        line_px = np.cumsum(mark, axis=1)[:, :-1] > 0
        if axis == 0:
            line_px = line_px.T
        out[line_px] = 255
    return Image.fromarray(out)


class OCREngine:
    """Encapsule pytesseract (si le binaire est présent) ou easyocr (100 % pip)."""

    def __init__(self, mode: str = "Automatique", tess_path: str = "", dpi: int = 300):
        self.dpi = dpi
        self.engine = None
        self.tess_path = tess_path
        if mode in ("Automatique", "Tesseract") and tesseract_ready(tess_path):
            self.engine = "tesseract"
        elif mode in ("Automatique", "EasyOCR") and easyocr_ready():
            self.engine = "easyocr"

    @property
    def available(self) -> bool:
        return self.engine is not None

    def image_to_words(self, image) -> list[dict]:
        """OCR d'une image -> liste de mots positionnés {x0, x1, top, bottom, text}."""
        words = []
        try:
            image = remove_table_lines(image)
        except Exception:
            pass
        if self.engine == "tesseract":
            import pytesseract
            data = None
            for lang in ("fra+eng+deu", "fra+eng", "eng"):
                try:
                    data = pytesseract.image_to_data(image, lang=lang, config="--psm 6",
                                                     output_type=pytesseract.Output.DICT)
                    break
                except Exception:
                    continue
            if data:
                for i, txt in enumerate(data["text"]):
                    txt = (txt or "").strip()
                    if txt and float(data["conf"][i]) > 0:
                        x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
                        words.append({"x0": x, "x1": x + w, "top": y, "bottom": y + h, "text": txt})
        elif self.engine == "easyocr":
            import numpy as np
            reader = get_easyocr_reader(("fr", "en", "de"))
            for box, txt, _conf in reader.readtext(np.array(image.convert("RGB")), detail=1, paragraph=False):
                xs, ys = [p[0] for p in box], [p[1] for p in box]
                for k, part in enumerate(str(txt).split()):
                    # découpage approximatif des boîtes multi-mots
                    n = len(str(txt).split())
                    w = (max(xs) - min(xs)) / max(n, 1)
                    words.append({"x0": min(xs) + k * w, "x1": min(xs) + (k + 1) * w,
                                  "top": min(ys), "bottom": max(ys), "text": part})
        out = []
        for w in words:
            if re.fullmatch(r"[|\[\]_—–=:;!]+", w["text"]):
                continue
            if len(re.findall(r"\d", w["text"])) >= 3:  # « DTRO00573478 » -> « DTR000573478 »
                w["text"] = re.sub(r"(?<=\d)[Oo]|[Oo](?=\d)", "0", w["text"])
                w["text"] = re.sub(r"(?<=\d)[Il|](?=\d)", "1", w["text"])
            out.append(w)
        return out


# =============================================================================
# RECONSTRUCTION DES LIGNES ET TABLEAUX À PARTIR DES POSITIONS DES MOTS
# =============================================================================

def group_word_lines(words: list[dict]) -> list[list[dict]]:
    if not words:
        return []
    heights = sorted(w["bottom"] - w["top"] for w in words)
    med_h = max(heights[len(heights) // 2], 1)
    lines, current, cur_y = [], [], None
    for w in sorted(words, key=lambda w: ((w["top"] + w["bottom"]) / 2, w["x0"])):
        y = (w["top"] + w["bottom"]) / 2
        if cur_y is None or abs(y - cur_y) <= 0.55 * med_h:
            current.append(w)
            cur_y = y if cur_y is None else (cur_y * (len(current) - 1) + y) / len(current)
        else:
            lines.append(sorted(current, key=lambda w: w["x0"]))
            current, cur_y = [w], y
    if current:
        lines.append(sorted(current, key=lambda w: w["x0"]))
    return lines


def words_to_text(words: list[dict]) -> str:
    return "\n".join(" ".join(w["text"] for w in line) for line in group_word_lines(words))


def _phrases(line: list[dict], gap: float) -> list[dict]:
    out = []
    for w in line:
        if out and w["x0"] - out[-1]["x1"] <= gap:
            out[-1]["x1"] = w["x1"]
            out[-1]["text"] += " " + w["text"]
        else:
            out.append(dict(w))
    return out


def words_to_tables(words: list[dict]) -> list[list[list[str]]]:
    """Reconstitue des tableaux (sans filets) : une ligne d'en-tête reconnue définit les colonnes,
    chaque mot des lignes suivantes est affecté à la colonne la plus proche."""
    lines = group_word_lines(words)
    if not lines:
        return []
    heights = sorted(w["bottom"] - w["top"] for w in words)
    gap = 0.8 * max(heights[len(heights) // 2], 1)
    header_idx = []
    for i, line in enumerate(lines):
        cells = [p["text"] for p in _phrases(line, gap)]
        ml, mm = map_header(cells, LINE_COLMAP), map_header(cells, MANIFEST_COLMAP)
        if (len(ml) >= 3 and ("qty" in ml or "amount" in ml or "net" in ml or "gross" in ml)) or \
                (len(mm) >= 3 and "supplier" in mm):
            header_idx.append(i)
    tables = []
    for n, h in enumerate(header_idx):
        header = _phrases(lines[h], gap)
        bounds = [(header[k]["x1"] + header[k + 1]["x0"]) / 2 for k in range(len(header) - 1)]
        hmap = map_header([p["text"] for p in header], LINE_COLMAP)
        numeric_cols = sorted(v for k, v in hmap.items() if k in ("qty", "price", "amount", "net", "gross", "packages"))
        stop = header_idx[n + 1] if n + 1 < len(header_idx) else len(lines)
        table = [[p["text"] for p in header]]
        for line in lines[h + 1:stop]:
            row = [""] * len(header)
            for w in line:
                cx = (w["x0"] + w["x1"]) / 2
                col = sum(1 for b in bounds if cx > b)
                row[col] = (row[col] + " " + w["text"]).strip()
            # mots de désignation qui débordent dans une colonne numérique : on les rend à la colonne de gauche
            for col in numeric_cols:
                toks = row[col].split()
                lead = []
                while len(toks) > 1 and not re.search(r"\d", toks[0]):
                    lead.append(toks.pop(0))
                if lead and col > 0:
                    row[col - 1] = (row[col - 1] + " " + " ".join(lead)).strip()
                    row[col] = " ".join(toks)
            table.append(row)
        tables.append(table)
    return tables


@dataclass
class PageData:
    file: str
    number: int
    text: str
    tables: list               # tableaux à filets (pdfplumber)
    ptables: list = field(default_factory=list)  # tableaux reconstitués par position des mots
    ocr: bool = False
    kind: str = "AUTRE"
    decimal: str = ""          # '.' pour de vrais nombres Excel ; '' = à détecter
    src: str = "pdf"           # pdf | excel | csv | word | text | image
    seq: int = 0               # n° de page dans l'ensemble du dossier (tous fichiers confondus)

    def table_sets(self) -> list:
        return [t for t in (self.tables, self.ptables) if t]


def read_pdf_pages(name: str, data: bytes, ocr: OCREngine, log: list, progress_cb=None,
                   errors: list | None = None) -> list[PageData]:
    pages: list[PageData] = []
    pypdf_reader = None
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text, tables, ptables, used_ocr = "", [], [], False
                try:
                    text = page.extract_text(x_tolerance=1.5, y_tolerance=3) or ""
                except Exception:
                    text = ""
                if not text.strip() and PdfReader is not None:
                    try:
                        pypdf_reader = pypdf_reader or PdfReader(io.BytesIO(data))
                        text = pypdf_reader.pages[i - 1].extract_text() or ""
                    except Exception:
                        pass
                try:
                    tables = page.extract_tables() or []
                except Exception:
                    tables = []
                try:
                    ptables = words_to_tables(page.extract_words(x_tolerance=1.5, y_tolerance=3))
                except Exception:
                    ptables = []
                if len(re.sub(r"\s", "", text)) < 40:  # page scannée -> OCR
                    if ocr.available:
                        try:
                            img = page.to_image(resolution=ocr.dpi).original
                            words = ocr.image_to_words(img)
                            text, tables, ptables, used_ocr = words_to_text(words), [], words_to_tables(words), True
                        except Exception as exc:
                            log.append(f"⚠️ {name} p.{i} : échec OCR ({exc})")
                    else:
                        log.append(f"⚠️ {name} p.{i} : page scannée mais aucun moteur OCR disponible")
                pages.append(PageData(name, i, text, tables, ptables, used_ocr))
                if progress_cb:
                    progress_cb(i, total)
    except Exception as exc:
        log.append(f"❌ Impossible de lire {name} : {exc}")
        if errors is not None:
            errors.append(f"{type(exc).__name__}: {exc}")
    return pages


# =============================================================================
# CLASSIFICATION ET SEGMENTATION DES DOCUMENTS
# =============================================================================

KIND_PATTERNS = [
    ("MANIFESTE", re.compile(r"(manifest|manifeste|liste\s*de\s*chargement|loading\s*list|groupage\s*list|"
                             r"r[ée]capitulatif\s*groupage|borderau|bordereau\s*de\s*groupage|ladeliste)", re.I)),
    ("CMR", re.compile(r"(\bcmr\b|lettre\s*de\s*voiture|consignment\s*note|frachtbrief|carta\s*de\s*porte)", re.I)),
    ("PACKING", re.compile(r"(packing\s*list|liste\s*de\s*colisage|colisage|packliste|delivery\s*note|"
                           r"bon\s*de\s*livraison|lieferschein|albar[aá]n|despatch\s*note|dispatch\s*note|"
                           r"lista\s*de\s*empaque|packing\s*slip)", re.I)),
    ("FACTURE", re.compile(r"(invoice|facture|rechnung|factura|fattura|proforma)", re.I)),
]


def classify_page(page: PageData) -> str:
    lines = text_lines(page.text)
    top = "\n".join(lines[:12])
    # 1) titre du document dans le haut de page
    for kind, rx in KIND_PATTERNS:
        if kind == "FACTURE":
            # « Invoice No » sur un BL ne fait pas une facture : on cherche un titre
            if re.search(r"(^|\n)\s*(commercial\s+|proforma\s+|tax\s+)?(invoice|facture|rechnung|factura|fattura)\b"
                         r"(?!\s*(no|n°|nr|number|num))", top, re.I):
                return "FACTURE"
            continue
        if rx.search(top):
            # un BL qui comporte des prix et un titre « facture » reste une facture
            if kind == "PACKING" and re.search(r"(^|\n)\s*(commercial\s+)?(invoice|facture|rechnung)\b(?!\s*(no|n°|nr|number|num|ref)\b)", top, re.I) \
                    and count_priced_lines(lines, page.decimal) >= 1:
                return "FACTURE"
            return kind
    # 2) contenu : présence de lignes valorisées + mot facture
    priced = count_priced_lines(lines, page.decimal)
    if priced and KIND_PATTERNS[3][1].search(page.text):
        return "FACTURE"
    for kind, rx in KIND_PATTERNS:
        if rx.search(page.text):
            return kind
    return "AUTRE"


def count_priced_lines(lines: list[str], decimal: str = "") -> int:
    dec = decimal or detect_decimal("\n".join(lines))
    return sum(1 for ln in lines if parse_item_line(ln, dec) is not None)


def find_invoice_number(text: str) -> str:
    lines = text_lines(text)
    for i, ln in enumerate(lines):
        for m in INV_NO_RE.finditer(ln):
            val = m.group(1).strip(".-/")
            if DATE_RE.fullmatch(val) or len(val) < 3:
                continue
            return val
        # libellé seul ; valeur sur la ligne suivante (mise en page tabulaire)
        if INV_LABEL_ONLY_RE.search(ln) and i + 1 < len(lines):
            for tok in lines[i + 1].split():
                if re.search(r"\d", tok) and not DATE_RE.fullmatch(tok) and len(tok) >= 3:
                    return tok.strip(".,;")
    return ""


def find_bl_number(text: str) -> str:
    for rx in (BL_NO_RE, BL_NO_RE2):
        for m in rx.finditer(text):
            val = m.group(1).strip(".-/")
            if not DATE_RE.fullmatch(val) and len(val) >= 3:
                return val
    return ""


@dataclass
class Document:
    kind: str
    file: str
    pages: list = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.pages)

    @property
    def decimal(self) -> str:
        """Séparateur décimal du document : vrais nombres Excel, sinon détection sur le texte."""
        hints = {p.decimal for p in self.pages if p.decimal}
        return hints.pop() if len(hints) == 1 and all(p.decimal for p in self.pages) else detect_decimal(self.text)

    @property
    def page_label(self) -> str:
        nums = [p.number for p in self.pages]
        return f"{nums[0]}" if len(nums) == 1 else f"{nums[0]}-{nums[-1]}"


def segment_documents(pages: list[PageData]) -> list[Document]:
    docs: list[Document] = []
    for page in pages:
        page.kind = classify_page(page)
        prev = docs[-1] if docs else None
        same_file = prev is not None and prev.file == page.file
        if page.kind == "AUTRE" and same_file:
            prev.pages.append(page)  # page de suite (page 2/3, conditions générales...)
            continue
        if same_file and prev.kind == page.kind:
            if page.kind == "FACTURE":
                n_prev, n_new = find_invoice_number(prev.text), find_invoice_number(page.text)
                if not n_new or n_new == n_prev or re.search(r"page\s*[2-9]|seite\s*[2-9]|\b[2-9]\s*/\s*\d\b",
                                                              page.text[:400], re.I) and not n_new:
                    prev.pages.append(page)
                    continue
            elif page.kind in ("PACKING", "CMR"):
                b_prev, b_new = find_bl_number(prev.text), find_bl_number(page.text)
                if not b_new or b_new == b_prev:
                    prev.pages.append(page)
                    continue
            else:
                prev.pages.append(page)
                continue
        docs.append(Document(page.kind, page.file, [page]))
    return docs


# =============================================================================
# EXTRACTION DES EN-TÊTES
# =============================================================================

NOT_A_NAME = re.compile(r"page|date|tel|fax|www|@|:|poids|weight|gewicht|\bkgs?\b|information|adress|address|livraison|delivery|"
                        r"printed|packing|facture|invoice|bon de|d\u00e9tails|details|contact|total", re.I)


def find_supplier(text: str) -> str:
    lines = text_lines(text)
    # 1) libellé explicite (Seller:, Fournisseur:, Shipper: ...)
    for i, ln in enumerate(lines[:40]):
        m = SUPPLIER_LABEL_RE.match(ln)
        if m:
            cand = m.group(1).strip(" :-")
            if not re.search(r"[A-Za-z]{3}", cand) and i + 1 < len(lines):
                cand = lines[i + 1]
            cand = re.split(r"\s{2,}|\s+(?:invoice|facture|date|n°|no\.)\b", cand, flags=re.I)[0]
            if cand and not CONSIGNEE_WORDS.search(cand):
                return cand.strip(" ,;:-")[:80]
    # 2) raison sociale avec forme juridique dans le haut du document
    for ln in lines[:25]:
        if CONSIGNEE_WORDS.search(ln) or KIND_PATTERNS[3][1].fullmatch(ln.strip()):
            continue
        if re.search(r"poids|weight|gewicht|\d[.,]\d{3}\b|\d{5,}", ln, re.I):  # « Poids brut total : 148,000 KG » : KG n'est pas une société
            continue
        m = LEGAL_RE.search(ln)
        if m and re.search(r"[A-Za-z]{3}", ln[:m.start()] or ""):
            return ln[:m.end() + (1 if ln[m.end():m.end() + 1] == "." else 0)].strip(" ,;:-")[:80]
    # 3) bloc « Société / Company » (pied de page de la facture), même si le nom contient le client habituel
    for i, ln in enumerate(lines):
        if re.match(r"(?i)^(soci[ée]t[ée]|company|firma|empresa|societ[àa])\b\s*(?:contact)?\s*$", ln) and i + 1 < len(lines):
            nxt = lines[i + 1]
            lm = LEGAL_RE.search(nxt)
            cand = nxt[:lm.end() + (1 if nxt[lm.end():lm.end() + 1] == "." else 0)] if lm else re.split(r"\s{2,}", nxt)[0]
            if re.search(r"[A-Za-z]{3}", cand):
                return cand.strip(" ,;:-")[:80]
    # 4) première ligne « texte » qui n'est pas un titre ni une adresse
    for ln in lines[:8]:
        if re.search(r"[A-Za-z]{3}", ln) and not any(rx.search(ln) for _, rx in KIND_PATTERNS) \
                and not CONSIGNEE_WORDS.search(ln) and not NOT_A_NAME.search(ln) and not re.search(r"\d{5,}", ln):
            return ln[:80]
    return ""


def _extract_client_from_two_col(line: str) -> str:
    """Sur une ligne 'FOURNISSEUR_SA  CLIENT_SA', extrait la partie client (après le 1er nom)."""
    lm = list(LEGAL_RE.finditer(line))
    if len(lm) >= 2:
        return line[lm[0].end():].strip(" ,;:-")
    parts = re.split(r'\s{2,}', line)
    if len(parts) >= 2:
        return parts[-1].strip(" ,;:-")
    if lm:
        return line[lm[0].end():].strip(" ,;:-")
    return ""


def find_client(text: str) -> str:
    """Détecte le nom du client/acheteur/destinataire dans une facture (toute langue)."""
    lines = text_lines(text)
    # 0) Format Maroc : "EXPORTATEUR: DESTINATAIRE:" sur une seule ligne (deux colonnes)
    for i, ln in enumerate(lines[:30]):
        if re.search(r"\bexport\w+\s*:.*\bdestinataire\s*:", ln, re.I):
            nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
            if nxt:
                cand = _extract_client_from_two_col(nxt)
                if cand and re.search(r"[A-Za-z]{3}", cand):
                    return cand[:80]
    # 1) libellé explicite (Bill to:, Client:, Käufer:, Nabywca:, ...)
    for i, ln in enumerate(lines[:80]):
        m = CLIENT_LABEL_RE.match(ln)
        if m:
            cand = m.group(1).strip(" :-")
            if not re.search(r"[A-Za-z]{3}", cand) and i + 1 < len(lines):
                nxt = lines[i + 1]
                cand = _extract_client_from_two_col(nxt) or nxt
            cand = re.split(r"\s{2,}|\s+(?:invoice|facture|date|n°|no\.)\b", cand, flags=re.I)[0]
            cand = cand.strip(" ,;:-")[:80]
            if cand and re.search(r"[A-Za-z]{3}", cand):
                return cand
    # 2) mot-clé consignee/destinataire dans la ligne → extraire le nom après
    for i, ln in enumerate(lines[:80]):
        if CONSIGNEE_WORDS.search(ln):
            clean = re.sub(
                r"^.*?(?:bill\s*to|ship\s*to|sold\s*to|invoice\s*to|deliver\w*\s*to|"
                r"consignee|destinataire|empf[äa]nger|acheteur|odbiorca|nabywca|"
                r"ontvanger|destinatario|koper)\s*[:\-]?\s*", "",
                ln, flags=re.I).strip()
            if not clean or len(clean) < 4:
                nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
                clean = _extract_client_from_two_col(nxt) or nxt
            if clean and re.search(r"[A-Za-z]{3}", clean) and not re.search(r"\b(timar)\b", clean, re.I):
                return clean[:80]
    return ""


def find_date(text: str) -> str:
    for ln in text_lines(text)[:60]:
        if re.search(r"(date|datum|fecha|data)\b", ln, re.I) and not re.search(r"(due|échéance|delivery|livraison|order)", ln, re.I):
            m = DATE_RE.search(ln)
            if m:
                return normalize_date(m.group(1))
    m = DATE_RE.search(text)
    return normalize_date(m.group(1)) if m else ""


def find_currency(text: str) -> str:
    counts = Counter()
    for cur in CURRENCIES:
        counts[cur] += len(re.findall(rf"\b{cur}\b", text))
    counts["EUR"] += text.count("€")
    counts["USD"] += len(re.findall(r"US\$|\$\s?\d", text))
    counts["GBP"] += text.count("£")
    cur, n = counts.most_common(1)[0]
    return cur if n else ""


def find_incoterm(text: str) -> str:
    m = re.search(r"\b(" + "|".join(INCOTERMS) + r")\b", text)
    return m.group(1) if m else ""


def weight_cell(value, decimal: str, header_unit: str = "") -> float | None:
    """Poids d'une cellule (« 250 g », « 1,2 t », « 12.5 lbs », 0.216) converti en kg ; l'unité de l'en-tête sert de défaut."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return None if (isinstance(value, float) and math.isnan(value)) else units.to_kg(float(value), header_unit)
    text = clean_cell(value)
    if not text:
        return None
    num, unit = units.split_number_unit(text)
    v = parse_number(num, decimal)
    if v is None and num.split():
        v = parse_number(num.split()[0], decimal)
    return units.to_kg(v, unit or header_unit)


def find_weight(text: str, rx: re.Pattern, decimal: str) -> float | None:
    """Poids total d'un document en kg (unités g / t / lb… converties automatiquement)."""
    best = None
    for m in rx.finditer(text):
        raw = m.group("num").strip()
        unit = (m.group("unit") or m.groupdict().get("hu") or "").lower()
        v = parse_number(raw, decimal)
        if v is None:
            v = parse_number(raw.split()[0], decimal) if raw.split() else None
        if v is None or v <= 0:
            continue
        factor = units.unit_factor(unit)
        if factor in (None, 1.0) and v > 40_000:   # > 40 t : « 148,000 » lu comme 148 000 -> autre convention décimale
            alt = parse_number(raw.split()[0], "," if decimal == "." else ".")
            if alt is not None and 0 < alt <= 40_000:
                v = alt
        v = units.to_kg(v, unit)
        is_total = "total" in m.group(0).lower()
        if best is None or is_total:
            best = v
    return best


HDR_WEIGHT_LABELS = [
    ("gross", r"(?:poids\s*brut|gross\s*weight|bruttogewicht|peso\s*bruto)"),
    ("net", r"(?:poids\s*net|net\s*weight|nettogewicht|peso\s*neto)"),
]


def find_header_weights(text: str, decimal: str) -> tuple[float | None, float | None]:
    """Mise en page « en-tête de colonnes puis valeurs » : « Poids brut (kg) Poids net (kg) RPP » puis « 280,00 220,00 … ».
    Renvoie (poids net, poids brut) en kg."""
    lines = text_lines(text)
    for i, ln in enumerate(lines[:-1]):
        if re.search(r"\d", re.sub(r"\((?:" + WEIGHT_UNIT + r")\)", "", ln)):
            continue
        found = []
        for key, pat in HDR_WEIGHT_LABELS:
            m = re.search(pat + r"\s*(?:\((" + WEIGHT_UNIT + r")\))?", ln, re.I)
            if m:
                found.append((m.start(), key, (m.group(1) or "").lower()))
        if not found:
            continue
        found.sort()
        if not re.match(r"\s*(?:" + "|".join(p for _, p in HDR_WEIGHT_LABELS) + ")", ln, re.I):
            continue
        vals = re.findall(r"\d[\d.,]*", lines[i + 1])
        if len(vals) < len(found) or not re.match(r"\s*\d", lines[i + 1]):
            continue
        out = {}
        for (_, key, unit), raw in zip(found, vals):
            v = parse_number(raw, decimal)
            if v is not None and v > 0:
                out[key] = units.to_kg(v, unit)
        if out:
            return out.get("net"), out.get("gross")
    return None, None


def find_packages(text: str) -> int | None:
    best = None
    for m in PACKAGES_RE.finditer(text):
        val = m.group(1) or m.group(2)
        if val and 0 < int(val) < 10000:
            if best is None or "total" in m.group(0).lower():
                best = int(val)
    return best


# =============================================================================
# EXTRACTION DES LIGNES DE FACTURE
# =============================================================================

def map_header(cells: list, colmap: list) -> dict:
    """Associe les indices de colonnes d'un tableau aux champs internes."""
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(cells):
        label = strip_accents(clean_cell(cell)).lower().replace("\n", " ").strip()
        if not label or len(label) > 60:
            continue
        for key, pattern in colmap:
            if key in mapping:
                continue
            if re.search(pattern, label, re.I):
                mapping[key] = idx
                break
    return mapping


def extract_detail_fields(text: str, decimal: str) -> dict:
    """Champs complémentaires présents dans une désignation ou une ligne de détail."""
    out = {}
    m = HS_LABEL_RE.search(text)
    if m:
        ngp = normalize_ngp(m.group(1))
        if ngp:
            out["ngp"] = ngp
    m = ORIGIN_LABEL_RE.search(text)
    if m:
        out["origin"] = normalize_origin(m.group(1))
    m = ALSTOM_LABEL_RE.search(text)
    if m:
        out["alstom"] = m.group(1).strip(".,;")
    m = BRAND_LABEL_RE.search(text)
    if m:
        out["brand"] = m.group(1).strip()
    for key, rx in (("net", LINE_NET_RE), ("gross", LINE_GROSS_RE), ("net_unit", LINE_UNITW_RE)):
        m = rx.search(text)
        if m:
            v = parse_number(m.group("num"), decimal)
            if v is not None:
                out[key] = units.to_kg(v, (m.group("unit") or m.group("hu") or "").lower())
                u = (m.group("unit") or m.group("hu") or "").lower()
                if units.describe(u):
                    out["w_note"] = units.describe(u)
    return out


NGP_TOKEN_RE = re.compile(r"^(\d{4}[.\s]?\d{2}[.\s]?\d{2}(?:[.\s]?\d{2})?)$")


def _numeric_tokens(tokens: list[str], decimal: str):
    nums = []
    for i, tok in enumerate(tokens):
        raw = tok.strip(";:")
        if DATE_RE.fullmatch(raw) or re.fullmatch(r"\d{8,}", raw) or re.fullmatch(r"\d{4}\.\d{2}\.\d{2}(\.\d{2})?", raw):
            continue
        v = parse_number(raw, decimal)
        if v is not None and v >= 0:
            nums.append((i, v))
    return nums


def _find_triplet(nums):
    """Cherche (qté, prix, montant) tels que qté × prix ≈ montant."""
    candidates_k = nums[-4:][::-1]
    for k_idx, (k, a) in enumerate(candidates_k):
        if a <= 0:
            continue
        before_k = [n for n in nums if n[0] < k]
        for j, p in before_k[::-1]:
            if p <= 0:
                continue
            for i, q in [n for n in nums if n[0] < j][::-1]:
                if q <= 0:
                    continue
                for per in (1, 100, 1000):
                    calc = q * p / per
                    if abs(calc - a) <= max(0.011, 0.005 * a):
                        return i, j, k, q, p, a, per
    return None


def parse_item_line(line: str, decimal: str):
    """Analyse une ligne de texte « libre » ; retourne un dict si c'est une ligne article."""
    if TOTAL_WORDS.search(line) and not FEE_WORDS.search(line):
        return None
    tokens = line.split()
    if len(tokens) < 3:
        return None
    nums = _numeric_tokens(tokens, decimal)
    if len(nums) < 3:
        # séparateurs de milliers par espace : « 1 234,56 »
        merged = re.sub(r"(?<=\d) (?=\d{3}(?:[.,]\d+)?(?:\s|$))", "", line)
        if merged != line:
            return parse_item_line(merged, decimal)
        return None
    trip = _find_triplet(nums)
    if trip is None:
        merged = re.sub(r"(?<=\d) (?=\d{3}(?:[.,]\d+)(?:\s|$))", "", line)
        if merged != line:
            return parse_item_line(merged, decimal)
        return None
    i, j, k, q, p, a, per = trip
    before = tokens[:i]
    if not before:
        return None
    pos = ""
    if re.fullmatch(r"\d{1,5}[.)]?", before[0]) and len(before) > 1:
        pos = before[0].rstrip(".)")
        before = before[1:]
    ngp = ""
    rest = []
    for tok in before:
        if not ngp and NGP_TOKEN_RE.match(tok) and ("." in tok or len(tok) in (8, 10)) and rest:
            ngp = normalize_ngp(tok)
            continue
        rest.append(tok)
    ref = ""
    for idx, tok in enumerate(rest[:3]):
        if re.search(r"\d", tok) and len(tok) >= 3 and re.fullmatch(r"[A-Za-z0-9\-_./#]+", tok):
            ref = tok
            rest = rest[:idx] + rest[idx + 1:]
            break
    if not ref and rest and re.fullmatch(r"[A-Z0-9\-_./]{4,}", rest[0]):
        ref = rest.pop(0)
    middle = tokens[i + 1:j]
    # Format « Poste Qté Réf PU Montant » : le 1er nombre est un n° de poste, la vraie réf est entre qté et prix
    mid_ref = next((t for t in middle if re.search(r"[A-Za-z]", t) and re.search(r"\d", t) and len(t) >= 6
                    and re.fullmatch(r"[A-Za-z0-9\-_./#]+", t)), "")
    if mid_ref and (not ref or re.fullmatch(r"\d{3,6}", ref)):
        if ref and not pos:
            pos = ref.lstrip("0") or ref
        ref = mid_ref
        middle = [t for t in middle if t != mid_ref]
    desc = " ".join(rest).strip(" -|")
    if not desc and not ref:
        return None
    unit = ""
    for tok in middle:
        if tok.upper().strip(".") in UNITS:
            unit = tok.upper().strip(".")
    if not unit and before and before[-1].upper().strip(".") in UNITS:
        unit = before[-1].upper().strip(".")
        desc = desc.rsplit(" ", 1)[0] if desc.endswith(before[-1]) else desc
    after = tokens[k + 1:]
    origin = ""
    for tok in after + middle:
        t = tok.upper().strip(".,;")
        if not ngp and NGP_TOKEN_RE.match(tok):
            ngp = normalize_ngp(tok)
        elif t in ISO_COUNTRIES and not origin and t not in UNITS:
            origin = t
    return {
        "pos": pos, "ref": ref, "desc": desc, "qty": q, "unit": unit,
        "price": p, "amount": a, "ngp": ngp, "origin": origin,
        "price_per": per,
    }


FOOTER_RE = re.compile(r"(CNSS|\bICE\s*:|Patente|\bRIB\b|IBAN|SWIFT|Soci[ée]t[ée]\s*Anonyme|Capital\s*de|R\.?C\.?\s*Fes)", re.I)
MAX_QTY, MAX_AMOUNT = 1_000_000, 50_000_000


def plausible_item(it: dict) -> bool:
    """Écarte les « lignes » qui sont en réalité un pied de page, un n° ICE/CNSS ou un RIB lus comme des nombres."""
    for key, lim in (("qty", MAX_QTY), ("price", MAX_AMOUNT), ("amount", MAX_AMOUNT)):
        v = it.get(key)
        if v is not None and abs(v) > lim:
            return False
    ref = clean_cell(it.get("ref"))
    if re.match(r"(?i)(kgs?|net|gross|serial|poids|weight|total|montant)\b", ref):  # poids / total lus comme un article
        return False
    return not FOOTER_RE.search(f"{ref} {it.get('desc', '')} {it.get('desc_full', '')}")


def rows_from_tables(tables: list, decimal: str, carry_map: dict | None):
    """Extrait les lignes article des tableaux pdfplumber. Retourne (lignes, mapping)."""
    items = []
    mapping = carry_map
    for table in tables:
        if not table:
            continue
        start = 0
        local_map = None
        for r_idx, row in enumerate(table[:6]):
            m = map_header(row, LINE_COLMAP)
            if len(m) >= 3 and ("qty" in m or "amount" in m) and ("ref" in m or "desc" in m):
                m["_hints"] = {k: units.unit_from_header(row[m[k]]) for k in ("net", "gross", "net_unit")
                               if k in m and m[k] < len(row)}
                local_map, start = m, r_idx + 1
                break
        if local_map is None:
            if mapping is None or len(table[0]) <= max(v for v in mapping.values() if isinstance(v, int)):
                continue
            local_map = mapping
        mapping = local_map
        hints = local_map.get("_hints", {})
        for row in table[start:]:
            def cell(key):
                idx = local_map.get(key)
                return clean_cell(row[idx]) if isinstance(idx, int) and idx < len(row) else ""
            joined = " ".join(clean_cell(c) for c in row)
            if not joined.strip() or FOOTER_RE.search(joined):
                continue
            qty = parse_number(cell("qty").split("\n")[0].split(" ")[0], decimal) if cell("qty") else None
            amount = parse_number(cell("amount"), decimal)
            price = parse_number(cell("price"), decimal)
            # colonnes fusionnées par l'OCR / la mise en page : « 45,50 455,00 »
            for src in ("price", "amount"):
                parts = [parse_number(t, decimal) for t in cell(src).split()]
                parts = [v for v in parts if v is not None]
                if len(parts) >= 2 and (price is None or amount is None):
                    price, amount = (price if price is not None else parts[0]), parts[-1]
            ref = cell("ref").split("\n")[0]
            desc_full = cell("desc")
            desc = desc_full.split("\n")[0] if desc_full else ""
            if qty is None and amount is None:
                # ligne de suite de la désignation (même tableau) : rattachée à l'article précédent
                if items and desc_full and not ref and not TOTAL_WORDS.search(joined) \
                        and not any(cell(k) for k in ("qty", "price", "amount", "pos")):
                    items[-1]["desc_full"] = f"{items[-1].get('desc_full') or ''}\n{desc_full}".strip()
                continue
            if not (ref or desc):
                continue
            if TOTAL_WORDS.search(f"{ref} {desc}") and not FEE_WORDS.search(f"{ref} {desc}"):
                continue
            if amount is None and qty is not None and price is not None:
                amount = round(qty * price, 2)
            if price is None and qty and amount is not None:
                price = round(amount / qty, 4)
            details = extract_detail_fields(joined, decimal)
            unit = cell("unit")
            if not unit:
                um = re.search(r"\b(" + "|".join(sorted(UNITS, key=len, reverse=True)) + r")\b", cell("qty").upper())
                unit = um.group(1) if um else ""
            items.append({
                "pos": cell("pos"), "ref": ref, "desc": desc or desc_full,
                "desc_full": desc_full, "po": cell("po"), "unit_raw": cell("unit"),
                "alstom": cell("alstom") or details.get("alstom", ""),
                "brand": cell("brand") or details.get("brand", ""),
                "qty": qty, "unit": unit.upper(), "price": price, "amount": amount,
                "net": weight_cell(cell("net"), decimal, hints.get("net", "")) if cell("net") else details.get("net"),
                "gross": weight_cell(cell("gross"), decimal, hints.get("gross", "")) if cell("gross") else details.get("gross"),
                "net_unit": (weight_cell(cell("net_unit"), decimal, hints.get("net_unit", "")) if cell("net_unit")
                             else details.get("net_unit")),
                "w_note": details.get("w_note") or units.describe(hints.get("net", "")),
                "packages": parse_number(cell("packages"), decimal) if cell("packages") else None,
                "origin": normalize_origin(cell("origin")) if cell("origin") else details.get("origin", ""),
                "ngp": normalize_ngp(cell("ngp")) if cell("ngp") else details.get("ngp", ""),
            })
    return items, mapping


def rows_from_text(lines: list[str], decimal: str):
    """Analyse ligne à ligne (PDF sans tableau ou OCR). Retourne (lignes, NGP d'en-tête libres)."""
    items = []
    free_ngps = []
    last = None
    for ln in lines:
        item = parse_item_line(ln, decimal)
        if item:
            details = extract_detail_fields(ln, decimal)
            item.setdefault("alstom", details.get("alstom", ""))
            item.setdefault("brand", details.get("brand", ""))
            item["net"] = details.get("net")
            item["gross"] = details.get("gross")
            item["packages"] = None
            if not item["ngp"] and details.get("ngp"):
                item["ngp"] = details["ngp"]
            if not item["origin"] and details.get("origin"):
                item["origin"] = details["origin"]
            items.append(item)
            last = item
            continue
        details = extract_detail_fields(ln, decimal)
        if (last is not None and not details and not TOTAL_WORDS.search(ln) and len(ln) < 90
                and len(last.setdefault("desc_extra", [])) < 3
                and sum(c.isalpha() for c in ln) >= 0.6 * len(ln.replace(" ", ""))):
            last["desc_extra"].append(ln)  # suite de désignation sur la ligne suivante
            continue
        if last is not None and details and not TOTAL_WORDS.search(ln):
            # ligne de détail rattachée à l'article précédent
            for key, target in (("ngp", "ngp"), ("origin", "origin"), ("alstom", "alstom"),
                                ("brand", "brand"), ("net", "net"), ("gross", "gross")):
                if details.get(key) and not last.get(target):
                    last[target] = details[key]
            continue
        if details.get("ngp"):
            free_ngps.append(details["ngp"])
        if TOTAL_WORDS.search(ln):
            last = None
    return items, free_ngps


def find_header_ngp(text: str, items: list) -> str:
    line_ngps = {it.get("ngp") for it in items if it.get("ngp")}
    found = []
    for m in HS_LABEL_RE.finditer(text):
        n = normalize_ngp(m.group(1))
        if n:
            found.append(n)
    free = [n for n in found if n not in line_ngps]
    if free:
        return free[0]
    if found:
        return Counter(found).most_common(1)[0][0]
    return ""


@dataclass
class Invoice:
    doc: Document
    supplier: str = ""
    client: str = ""
    number: str = ""
    date: str = ""
    currency: str = ""
    incoterm: str = ""
    bl: str = ""
    header_ngp: str = ""
    origin: str = ""
    net_total: float | None = None
    gross_total: float | None = None
    packages: int | None = None
    items: list = field(default_factory=list)
    packing_lists: list = field(default_factory=list)


def parse_invoice(doc: Document) -> Invoice:
    text = doc.text
    dec = doc.decimal
    inv = Invoice(doc=doc)
    inv.supplier = find_supplier(text)
    inv.client = find_client(text)
    inv.number = find_invoice_number(text)
    inv.date = find_date(text)
    inv.currency = find_currency(text)
    inv.incoterm = find_incoterm(text)
    inv.bl = find_bl_number(text)
    inv.net_total = find_weight(text, TOTAL_NET_RE, dec)
    inv.gross_total = find_weight(text, TOTAL_GROSS_RE, dec)
    if inv.net_total is None or inv.gross_total is None:
        hn, hg = find_header_weights(text, dec)
        inv.net_total = inv.net_total if inv.net_total is not None else hn
        inv.gross_total = inv.gross_total if inv.gross_total is not None else hg
    inv.packages = find_packages(text)
    inv.origin = find_origin(text)

    items: list = []
    carry = None
    for page in doc.pages:
        # Candidats : tableaux à filets, tableaux reconstitués, analyse ligne à ligne.
        # On retient celui qui fournit le plus de lignes valorisées (montant présent).
        candidates = []
        for tset in page.table_sets():
            is_ruled = tset is page.tables
            found, new_carry = rows_from_tables(tset, dec, carry if is_ruled else None)
            if is_ruled:
                carry = new_carry
            candidates.append(found)
        candidates.append(rows_from_text(text_lines(page.text), dec)[0])

        def score(found):
            valued = sum(1 for it in found if it.get("amount") is not None)
            return valued, -sum(1 for it in found if it.get("amount") is None)
        page_items = max(candidates, key=score)  # à égalité : le premier (tableau) l'emporte
        for it in page_items:
            it.setdefault("page", page.number)
        items.extend(page_items)
    inv.items = [it for it in items if plausible_item(it)]
    inv.header_ngp = find_header_ngp(text, inv.items)
    return inv


# =============================================================================
# PACKING LISTS, CMR ET MANIFESTE
# =============================================================================

def row_cell(row: list, header: dict, key: str) -> str:
    idx = header.get(key)
    return clean_cell(row[idx]) if idx is not None and idx < len(row) else ""


def iter_table_rows(tables: list, colmap: list, accept):
    """Parcourt les lignes de données des tableaux dont l'en-tête satisfait `accept(mapping)`."""
    for table in tables:
        for r_idx, row in enumerate(table[:6]):
            m = map_header(row, colmap)
            if accept(m):
                m["_hints"] = {k: units.unit_from_header(row[m[k]]) for k in ("net", "gross", "net_unit")
                               if k in m and m[k] < len(row)}
                for data_row in table[r_idx + 1:]:
                    if any(clean_cell(c) for c in data_row):
                        yield m, data_row
                break


HU_RE = re.compile(r"^(\d{6,8})\s+[A-Za-zÀ-ÿ]", re.M)


def count_handling_units(text: str) -> int | None:
    """Colis = unités de manutention du BL (« 1524298 CARTON STANDARD … »), comptées une fois chacune."""
    n = len(set(HU_RE.findall(text)))
    return n or None


@dataclass
class PackingList:
    doc: Document
    number: str = ""
    invoice_ref: str = ""
    supplier: str = ""
    net_total: float | None = None
    gross_total: float | None = None
    packages: int | None = None
    lines: list = field(default_factory=list)  # {ref, qty, net, gross, packages}
    link: str = ""                              # méthode de rattachement à la facture


def parse_packing(doc: Document) -> PackingList:
    text = doc.text
    dec = doc.decimal
    pl = PackingList(doc=doc)
    pl.number = find_bl_number(text)
    pl.invoice_ref = find_invoice_number(text)
    pl.supplier = find_supplier(text)
    pl.net_total = find_weight(text, TOTAL_NET_RE, dec)
    pl.gross_total = find_weight(text, TOTAL_GROSS_RE, dec)
    if pl.net_total is None or pl.gross_total is None:
        hn, hg = find_header_weights(text, dec)
        pl.net_total = pl.net_total if pl.net_total is not None else hn
        pl.gross_total = pl.gross_total if pl.gross_total is not None else hg
    pl.packages = find_packages(text) or count_handling_units(text)
    for page in doc.pages:
        for tset in page.table_sets():  # tableaux à filets d'abord, puis reconstitués
            found = []
            for header, row in iter_table_rows(tset, LINE_COLMAP,
                                               lambda m: ("ref" in m or "alstom" in m) and ("net" in m or "gross" in m)):
                ref = row_cell(row, header, "ref").split("\n")[0] or row_cell(row, header, "alstom")
                if not ref or TOTAL_WORDS.search(ref):
                    continue
                hints = header.get("_hints", {})
                found.append({
                    "ref": ref,
                    "qty": parse_number(row_cell(row, header, "qty"), dec),
                    "net": weight_cell(row_cell(row, header, "net"), dec, hints.get("net", "")),
                    "gross": weight_cell(row_cell(row, header, "gross"), dec, hints.get("gross", "")),
                    "packages": parse_number(row_cell(row, header, "packages"), dec),
                })
            if found:
                pl.lines.extend(found)
                break
    return pl


def parse_manifest(doc: Document) -> list[dict]:
    dec = doc.decimal
    rows = []
    for page in doc.pages:
        for tset in page.table_sets():  # tableaux à filets d'abord, puis reconstitués
            found = []
            for header, row in iter_table_rows(tset, MANIFEST_COLMAP,
                                               lambda m: "supplier" in m and ("value" in m or "gross" in m)):
                def cell(key, row=row, header=header):
                    return row_cell(row, header, key)
                supplier = cell("supplier").replace("\n", " ")
                if not supplier or TOTAL_WORDS.search(supplier):
                    continue
                found.append({
                    "Page PDF": page.number,
                    "Shipment": cell("shipment").replace("\n", " "),
                    "Expéditeur": supplier,
                    "Poids (kg)": weight_cell(cell("gross"), dec, header.get("_hints", {}).get("gross", "")),
                    "Colis": parse_number(cell("packages"), dec),
                    "Goods Value": parse_number(cell("value"), dec),
                    "Devise": find_currency(cell("currency") + " " + cell("value")),
                })
            if found:
                rows.extend(found)
                break
    return rows


# =============================================================================
# CONSTRUCTION DES LIGNES + VENTILATION DES POIDS
# =============================================================================

MAX_TRUCK_KG = 40_000
BASIS_LABEL = {"Valeur HT": "valeur HT", "Quantité": "quantité", "Lignes": "nombre de lignes (égal)"}


def weight_tol(a: float, b: float) -> float:
    """Tolérance de comparaison de deux poids (kg) : 50 g ou 0,5 %."""
    return max(0.05, 0.005 * max(abs(a), abs(b)))


def weights_differ(a, b) -> bool:
    return a is not None and b is not None and abs(a - b) > weight_tol(a, b)


def diff_severity(a: float, b: float) -> str:
    """Écart de poids : critique au-delà de 10 %, sinon avertissement."""
    return "Critique" if abs(a - b) > 0.10 * max(abs(a), abs(b), 1e-9) else "Avertissement"


def link_packing_lists(invoices: list[Invoice], packings: list[PackingList], docs: list[Document]):
    """Rattache chaque packing list / BL à sa facture (n° cités, fournisseur, proximité) et mémorise la méthode."""
    order = {id(d): i for i, d in enumerate(docs)}
    for pl in packings:
        target, how = None, ""
        if pl.invoice_ref:
            target = next((inv for inv in invoices if inv.number and
                           compact_ref(inv.number) == compact_ref(pl.invoice_ref)), None)
            how = "n° de facture cité dans le BL / PL"
        if target is None and pl.number and len(pl.number) >= 5:   # n° de BL cité dans la facture
            cited = [inv for inv in invoices if pl.number in inv.doc.text]
            if len(cited) == 1:
                target, how = cited[0], "n° de BL cité dans la facture"
        if target is None and pl.supplier:
            same = [inv for inv in invoices if company_similarity(inv.supplier, pl.supplier) >= 0.8]
            if len(same) == 1:
                target, how = same[0], "même fournisseur"
        if target is None and invoices:
            pos = order[id(pl.doc)]
            same_file = [inv for inv in invoices if inv.doc.file == pl.doc.file]
            pool = same_file or invoices
            before = [inv for inv in pool if order[id(inv.doc)] < pos]
            target = before[-1] if before else min(pool, key=lambda inv: abs(order[id(inv.doc)] - pos))
            how = "proximité dans le dossier (à vérifier)"
        pl.link = how
        if target is not None:
            target.packing_lists.append(pl)
            if not target.bl and pl.number:
                target.bl = pl.number


def allocate(total: float, weights: list[float], decimals: int = 3) -> list[float]:
    s = sum(weights)
    if not s:
        weights = [1.0] * len(weights)
        s = float(len(weights))
    parts = [round(total * w / s, decimals) for w in weights]
    if parts:
        parts[-1] = round(total - sum(parts[:-1]), decimals)
    return parts


def _combine_totals(vals: list) -> tuple[float | None, str]:
    """Total de plusieurs documents liés : une seule valeur si identiques (doublon PL / BL), sinon la somme."""
    vals = [v for v in vals if v is not None]
    if not vals:
        return None, ""
    if len(vals) == 1 or all(not weights_differ(vals[0], v) for v in vals[1:]):
        return vals[0], ""
    return sum(vals), f"somme de {len(vals)} documents"


def build_invoice_rows(inv: Invoice, basis: str, alerts: list | None = None) -> list[dict]:
    """Lignes TOUS_CLIENTS d'une facture, poids compris :
    1) poids de ligne imprimé sur la facture ; 2) à défaut, poids de la même référence dans la packing list / le BL ;
    3) à défaut, poids net total (facture, sinon PL / BL) ventilé au prorata. Chaque valeur garde sa « base »."""
    A = alerts if alerts is not None else []

    def alert(sev, typ, msg, item="", ref=""):
        A.append({"Sévérité": sev, "Type": typ, "Fournisseur": inv.supplier, "N° facture": inv.number,
                  "Item": item, "Réf article": ref, "Message": msg})
    page_seq = {p.number: p.seq for p in inv.doc.pages}
    first_page = inv.doc.pages[0] if inv.doc.pages else None
    rows = []
    for n, it in enumerate(inv.items, start=1):
        desc = re.sub(r"\s+", " ", clean_cell(it.get("desc")))
        is_fee = bool(it.get("_fee")) or bool(FEE_WORDS.search(f"{desc} {it.get('ref', '')}"))
        page = it.get("page") or (first_page.number if first_page else 0)
        net, base = it.get("net"), []
        if net is not None:
            base.append("Facture (ligne imprimée)")
            if it.get("w_note"):
                base.append(it["w_note"])
        rows.append({
            "Client": inv.supplier, "Job Ref manifest": "", "N° facture": inv.number, "Date": inv.date,
            "Item": clean_cell(it.get("pos")) or str(n), "Réf article": clean_cell(it.get("ref")), "Désignation": desc,
            "Origine": clean_cell(it.get("origin")) or inv.origin,
            "Qté": it.get("qty"), "Prix unitaire": it.get("price"), "Valeur HT": it.get("amount"),
            "NGP/code douanier (facture)": it.get("ngp") or inv.header_ngp,
            "NGP (REF_PAR_FRS)": "", "Onglet REF_PAR_FRS": "", "Contrôle NGP": "", "Remarque": "",
            "Page PDF": page, PAGE_IN_PDF: page_seq.get(page, first_page.seq if first_page else 0),
            C_NET: net, C_NETU: it.get("net_unit"), C_NETF: None, C_GROSSF: None, C_COLIS: None, C_MANIF: None,
            C_BASE: "",
            "_réf_alstom": clean_cell(it.get("alstom")), "_devise": inv.currency, "_source": inv.doc.file,
            "_inv": f"{inv.supplier}§{inv.number}§{id(inv)}", "_fee": is_fee,
            "_is_fee": is_fee, "_line_ngp": bool(it.get("ngp")), "_base": base, "_pkg": it.get("packages"),
        })
    if not rows:
        return rows
    goods = [r for r in rows if not r["_is_fee"]] or rows
    pls = inv.packing_lists
    pl_label = lambda pl: pl.number or "PL"  # noqa: E731

    # --- 1) Poids de la packing list / du BL pour la même référence (consommés ligne à ligne)
    pl_lines = [dict(l, _pl=pl) for pl in pls for l in pl.lines]

    def take(r):
        for key_col in ("Réf article", "_réf_alstom"):
            key = compact_ref(r[key_col])
            if not key:
                continue
            cands = [l for l in pl_lines if not l.get("_used") and compact_ref(l["ref"]) == key]
            if not cands:
                continue
            q = r["Qté"]
            exact = next((l for l in cands if q is not None and l.get("qty") is not None and abs(l["qty"] - q) < 1e-9), None)
            if exact:
                group = [exact]
            elif len(cands) > 1 and q is not None and abs(sum(l.get("qty") or 0 for l in cands) - q) < 1e-9:
                group = cands
            else:
                group = [cands[0]]
            for l in group:
                l["_used"] = True
            return group
        return None
    for r in goods:
        group = take(r)
        if not group:
            continue
        nets = [l.get("net") for l in group if l.get("net") is not None]
        pl_net = sum(nets) if nets else None
        names = ", ".join(dict.fromkeys(pl_label(l["_pl"]) for l in group))
        pq = sum(l.get("qty") or 0 for l in group) if any(l.get("qty") is not None for l in group) else None
        if pq is not None and r["Qté"] is not None and abs(pq - r["Qté"]) > 1e-9:
            alert("Info", "Quantité", f"Qté facture {r['Qté']:g} ≠ qté PL/BL {names} {pq:g}", r["Item"], r["Réf article"])
        if pl_net is not None:
            if r[C_NET] is None:
                r[C_NET] = pl_net
                r["_base"].append(f"Packing list {names} (même réf)" + (f" — qté PL {pq:g} ≠ facture {r['Qté']:g}"
                                                                       if pq is not None and r["Qté"] is not None and abs(pq - r["Qté"]) > 1e-9 else ""))
            elif weights_differ(r[C_NET], pl_net):
                alert(diff_severity(r[C_NET], pl_net), "Écart de poids",
                      f"Poids net ligne facture {r[C_NET]:g} kg ≠ packing list/BL {names} {pl_net:g} kg", r["Item"], r["Réf article"])
        if any(l.get("packages") is not None for l in group) and r["_pkg"] is None:
            r["_pkg"] = sum(l.get("packages") or 0 for l in group)
    unused = [l for l in pl_lines if not l.get("_used") and l.get("ref")]
    if unused and pls:
        alert("Info", "Liaison documents", f"{len(unused)} ligne(s) de packing list / BL sans ligne de facture : "
                                           + ", ".join(dict.fromkeys(str(l['ref']) for l in unused))[:140])

    # --- 2) Totaux de la facture, à défaut de la packing list / du BL
    net_total, net_note = (inv.net_total, "") if inv.net_total is not None else _combine_totals([pl.net_total for pl in pls])
    gross_total, gross_note = (inv.gross_total, "") if inv.gross_total is not None else _combine_totals([pl.gross_total for pl in pls])
    pkg_total = inv.packages if inv.packages is not None else _combine_totals([pl.packages for pl in pls])[0]
    net_src = ("facture" if inv.net_total is not None else
               (f"PL/BL {', '.join(dict.fromkeys(pl_label(pl) for pl in pls if pl.net_total is not None))}"
                + (f" ({net_note})" if net_note else "") if net_total is not None else ""))
    for pl in pls:   # alertes : facture ≠ packing list / BL
        for lab, a, b in (("net", inv.net_total, pl.net_total), ("brut", inv.gross_total, pl.gross_total)):
            if weights_differ(a, b):
                alert(diff_severity(a, b), "Écart de poids",
                      f"Poids {lab} facture {a:g} kg ≠ packing list/BL {pl_label(pl)} {b:g} kg (écart {b - a:+.3f} kg)")
    if net_total is not None and gross_total is not None and net_total > gross_total + 1e-6:
        alert("Critique", "Poids", f"Poids net {net_total:g} kg > poids brut {gross_total:g} kg")
    for lab, v in (("net", net_total), ("brut", gross_total)):
        if v is not None and v > MAX_TRUCK_KG:
            alert("Critique", "Poids", f"Poids {lab} {v:,.1f} kg invraisemblable (> 40 t) — unité ou virgule mal lue ?")

    # --- 3) Prorata du poids net pour les lignes sans poids
    known = sum(r[C_NET] for r in goods if r[C_NET] is not None)
    missing = [r for r in goods if r[C_NET] is None]
    if missing:
        if net_total is not None:
            remaining = net_total - known
            if remaining > 1e-9:
                if basis == "Quantité":
                    w = [abs(r["Qté"] or 0) for r in missing]
                elif basis == "Lignes":
                    w = [1.0] * len(missing)
                else:
                    w = [abs(r["Valeur HT"] or 0) for r in missing]
                for r, v in zip(missing, allocate(remaining, w)):
                    r[C_NET] = v
                    r["_base"].append(f"Prorata {BASIS_LABEL.get(basis, 'valeur HT')} du poids net {net_src} ({net_total:g} kg)")
            else:
                for r in missing:
                    r["_base"].append("Poids net non ventilable (total déjà atteint)")
                alert("Avertissement", "Poids", f"Σ poids nets imprimés ({known:g} kg) ≥ poids net total {net_src} "
                                                  f"({net_total:g} kg) : {len(missing)} ligne(s) sans poids")
        else:
            for r in missing:
                r["_base"].append("Poids net non trouvé")
            alert("Avertissement", "Poids", f"{len(missing)} ligne(s) sans poids net : aucun poids net total "
                                              "dans la facture ni dans une packing list / un BL lié")
    elif net_total is not None and weights_differ(known, net_total):
        alert("Avertissement", "Écart de poids",
              f"Σ poids nets des lignes {known:.3f} kg ≠ poids net total {net_src} {net_total:.3f} kg")
    if gross_total is None:
        alert("Info", "Poids", "Poids brut de la facture non trouvé (facture et packing list / BL)")

    # --- 4) Poids net unitaire (imprimé, sinon calculé)
    for r in goods:
        if r[C_NETU] is None and r[C_NET] is not None and r["Qté"] and r["Qté"] > 0:
            r[C_NETU] = round(r[C_NET] / r["Qté"], 6)
            r["_base"].append("unitaire calculé (net ÷ qté)")

    # --- 5) Colonnes « facture » : portées par la 1re ligne seulement (jamais dupliquées)
    if pkg_total is None and any(r["_pkg"] is not None for r in rows):
        pkg_total = sum(r["_pkg"] or 0 for r in rows)
    rows[0][C_NETF], rows[0][C_GROSSF], rows[0][C_COLIS] = net_total, gross_total, pkg_total
    for r in rows:
        if not r["_line_ngp"] and r["NGP/code douanier (facture)"]:
            r["Remarque"] = "NGP en-tête facture"
        r[C_BASE] = " | ".join(dict.fromkeys(r["_base"]))
        for k in ("_is_fee", "_line_ngp", "_base", "_pkg"):
            r.pop(k, None)
    return rows


# =============================================================================
# FICHIER EXCEL DE RÉFÉRENCE ET MATCHING NGP
# =============================================================================

def load_reference(data: bytes, filename: str, log: list) -> dict:
    """Retourne {'exact': {ref: ngp}, 'compact': {ref: ngp}, 'ocr': {ref: ngp}, 'sheet_names': {ref: sheet}, 'count': n}."""
    exact, compact, ocrk, sheet_names = {}, {}, {}, {}
    try:
        if filename.lower().endswith(".csv"):
            sheets = {"CSV": pd.read_csv(io.BytesIO(data), header=None, dtype=str, sep=None, engine="python")}
        else:
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str)
    except Exception as exc:
        log.append(f"❌ Lecture du fichier de référence impossible : {exc}")
        return {"exact": exact, "compact": compact, "ocr": ocrk, "count": 0}
    for sheet, df in sheets.items():
        if df.empty:
            continue
        header_idx, mapping, ref_cols = None, {}, []
        for r in range(min(25, len(df))):
            cells = list(df.iloc[r].values)
            m = map_header(cells, REF_SHEET_COLMAP)
            if "ngp" in m and ("ref" in m or "alstom" in m):
                header_idx, mapping = r, m
                # toutes les colonnes « référence » de la ligne d'en-tête
                for c_idx, c in enumerate(cells):
                    lab = strip_accents(clean_cell(c)).lower()
                    if c_idx != m["ngp"] and lab and re.search(REF_SHEET_COLMAP[1][1] + "|" + REF_SHEET_COLMAP[2][1], lab) \
                            and not re.search(r"(desig|descr|libell|ngp|\bhs\b)", lab):
                        ref_cols.append(c_idx)
                break
        if header_idx is None:
            # pas d'en-tête reconnu : colonne NGP = la plus riche en codes à 8-10 chiffres
            scores = {c: df[c].map(lambda v: len(normalize_ngp(v)) in (8, 10)).sum() for c in df.columns}
            ngp_col = max(scores, key=scores.get)
            if scores[ngp_col] == 0:
                log.append(f"⚠️ Feuille « {sheet} » ignorée (aucune colonne NGP détectée)")
                continue
            header_idx, mapping = -1, {"ngp": ngp_col}
            ref_cols = [c for c in df.columns if c != ngp_col][:1]
        body = df.iloc[header_idx + 1:]
        added = 0
        for _, row in body.iterrows():
            ngp = normalize_ngp(row.iloc[mapping["ngp"]])
            if not ngp:
                continue
            for c in ref_cols:
                key = norm_ref(row.iloc[c])
                if not key:
                    continue
                exact.setdefault(key, ngp)
                sheet_names.setdefault(key, sheet)
                ck = compact_ref(key)
                if ck:
                    compact.setdefault(ck, ngp)
                    ocrk.setdefault(ocr_ref(key), ngp)
                added += 1
        log.append(f"✅ Référentiel « {sheet} » : {added} références chargées")
    return {"exact": exact, "compact": compact, "ocr": ocrk, "sheet_names": sheet_names, "count": len(exact)}


def lookup_ngp(refdb: dict, *refs) -> tuple[str, str]:
    for label, ref in refs:
        key = norm_ref(ref)
        if key and key in refdb["exact"]:
            return refdb["exact"][key], label
    for label, ref in refs:
        key = compact_ref(ref)
        if key and key in refdb["compact"]:
            return refdb["compact"][key], label
    for label, ref in refs:
        key = ocr_ref(ref)
        if len(key) >= 5 and key in refdb.get("ocr", {}):
            return refdb["ocr"][key], f"{label} (tolérance OCR)"
    return "", ""


def control_ngp(ngp_ref: str, code_fact: str) -> str:
    """Réplique Python de la formule Excel « Contrôle NGP »."""
    if not ngp_ref or ngp_ref == NON_TROUVE or not code_fact:
        return "À vérifier"
    return "OK" if str(ngp_ref)[:8] == str(code_fact)[:8] else "ÉCART"


def apply_matching(df: pd.DataFrame, refdb: dict) -> pd.DataFrame:
    df = df.copy()
    ngp_ref_list, onglet_list, ctrl, remarks = [], [], [], []
    sheet_names = refdb.get("sheet_names", {})
    for _, r in df.iterrows():
        code_fact = normalize_ngp(r.get("NGP/code douanier (facture)", ""))
        rem = clean_cell(r.get("Remarque", ""))
        rem = " | ".join(p for p in rem.split(" | ") if p and p != REMARQUE_NOUVELLE_REF
                         and not p.startswith("Trouvé via"))
        ref_art = r.get("Réf article", "")
        ref_alstom = r.get("_réf_alstom", "")
        ngp, via = lookup_ngp(refdb, ("Réf article", ref_art), ("_réf_alstom", ref_alstom))
        if ngp:
            if via != "Réf article":
                rem = f"{rem} | Trouvé via {via}".strip(" |")
            # get sheet name for this reference
            found_key = norm_ref(ref_art) if via == "Réf article" else norm_ref(ref_alstom)
            onglet = sheet_names.get(found_key, "")
        else:
            ngp = NON_TROUVE
            onglet = ""
            rem = f"{rem} | {REMARQUE_NOUVELLE_REF}".strip(" |")
        ngp_ref_list.append(ngp)
        onglet_list.append(onglet)
        ctrl.append(control_ngp(ngp, code_fact))
        remarks.append(rem)
    df["NGP/code douanier (facture)"] = df["NGP/code douanier (facture)"].map(
        lambda v: normalize_ngp(v) if pd.notna(v) else "")
    df["NGP (REF_PAR_FRS)"] = ngp_ref_list
    df["Onglet REF_PAR_FRS"] = onglet_list
    df["Contrôle NGP"] = ctrl
    df["Remarque"] = remarks
    # ensure all COLUMNS exist (colonnes internes conservées : devise, fichier source…)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df[COLUMNS + [c for c in INTERNAL_COLUMNS if c in df.columns]]


# =============================================================================
# PIPELINE COMPLET
# =============================================================================

# =============================================================================
# LECTURE IA (Claude) — adaptateurs vers les structures du moteur
# =============================================================================

def _ai_num(v):
    return ai_engine.num(v) if ai_engine else parse_number(v)


def _ai_int(v):
    x = _ai_num(v)
    return int(round(x)) if x is not None else None


def invoice_from_ai(d: dict, doc: Document) -> Invoice:
    inv = Invoice(doc=doc)
    inv.supplier = ai_engine.party(d, "supplier")
    inv.client = ai_engine.party(d, "client")
    inv.number = clean_cell(d.get("number"))
    inv.date = normalize_date(ai_engine.iso_date_to_fr(d.get("date")))
    inv.currency = clean_cell(d.get("currency")).upper()
    inv.incoterm = clean_cell(d.get("incoterm")).upper()
    inv.bl = next(iter(d.get("related_delivery_numbers") or []), "")
    inv.origin = normalize_origin(d.get("origin_country") or "")
    inv.net_total = _ai_num(d.get("net_weight_kg"))
    inv.gross_total = _ai_num(d.get("gross_weight_kg"))
    inv.packages = _ai_int(d.get("packages"))
    inv.header_ngp = normalize_ngp(d.get("header_hs_code"))
    for l in d.get("lines") or []:
        amount = _ai_num(l.get("amount"))
        if l.get("line_type") == "remise" and amount is not None:
            amount = -abs(amount)
        inv.items.append({
            "pos": clean_cell(l.get("position")),
            "ref": clean_cell(l.get("reference")) or clean_cell(l.get("customer_reference")),
            "alstom": clean_cell(l.get("customer_reference")),
            "desc": clean_cell(l.get("description")),
            "qty": _ai_num(l.get("quantity")), "price": _ai_num(l.get("unit_price")), "amount": amount,
            "unit": clean_cell(l.get("unit")), "unit_raw": clean_cell(l.get("unit")),
            "ngp": normalize_ngp(l.get("hs_code")), "origin": normalize_origin(l.get("origin_country") or ""),
            "net": _ai_num(l.get("net_weight_kg")), "gross": _ai_num(l.get("gross_weight_kg")),
            "net_unit": _ai_num(l.get("unit_net_weight_kg")),
            "packages": _ai_num(l.get("packages")), "po": clean_cell(l.get("po_number")),
            "page": l.get("page"), "_fee": l.get("line_type") in ("frais", "remise"),
        })
    return inv


def packing_from_ai(d: dict, doc: Document) -> PackingList:
    return PackingList(
        doc=doc, number=clean_cell(d.get("number")) or next(iter(d.get("related_delivery_numbers") or []), ""),
        invoice_ref=next(iter(d.get("related_invoice_numbers") or []), ""),
        supplier=ai_engine.party(d, "supplier"),
        net_total=_ai_num(d.get("net_weight_kg")), gross_total=_ai_num(d.get("gross_weight_kg")),
        packages=_ai_int(d.get("packages")),
        lines=[{"ref": clean_cell(l.get("reference")) or clean_cell(l.get("customer_reference")),
                "qty": _ai_num(l.get("quantity")), "net": _ai_num(l.get("net_weight_kg")),
                "gross": _ai_num(l.get("gross_weight_kg")), "packages": _ai_num(l.get("packages"))}
               for l in d.get("lines") or [] if l.get("line_type") == "article"])


def manifest_from_ai(d: dict) -> list[dict]:
    return [{"Page PDF": r.get("page") or (d["pages"][0] if d.get("pages") else 0),
             "Shipment": clean_cell(r.get("shipment")), "Expéditeur": clean_cell(r.get("shipper")),
             "Poids (kg)": _ai_num(r.get("gross_weight_kg")), "Colis": _ai_num(r.get("packages")),
             "Goods Value": _ai_num(r.get("goods_value")), "Devise": clean_cell(r.get("currency")).upper()}
            for r in d.get("manifest_rows") or [] if clean_cell(r.get("shipper"))]


def add_issue(issues: list | None, file: str, severity: str, message: str, hint: str = ""):
    if issues is not None:
        issues.append({"Fichier": file, "Gravité": severity, "Problème": message, "Que faire": hint})


def read_file_pages(name: str, data: bytes, ocr: OCREngine, log: list, issues: list | None,
                    progress_cb=None) -> list[PageData]:
    """Lit un fichier téléversé (PDF numérique/scanné, Excel, CSV, Word, texte, image) en pages texte + tableaux.
    Ne lève jamais d'exception : un fichier illisible est signalé dans `issues` avec un message clair."""
    try:
        kind = fileio.validate_upload(name, data)
    except fileio.FileReadError as exc:
        add_issue(issues, name, "Erreur", str(exc), "Ce fichier est ignoré : corrigez-le ou retirez-le de la liste.")
        log.append(f"❌ {name} ignoré : {exc}")
        return []
    try:
        if kind == "pdf":
            errs: list[str] = []
            pages = read_pdf_pages(name, data, ocr, log, progress_cb, errs)
            for p in pages:
                p.src = "pdf"
            if not pages:
                reason = errs[0] if errs else "aucune page lisible"
                low = reason.lower()
                hint = ("Le PDF est protégé par mot de passe : retirez la protection puis réimportez-le."
                        if any(w in low for w in ("password", "encrypt", "decrypt")) else
                        "Le PDF semble corrompu : réexportez-le depuis son application d'origine.")
                add_issue(issues, name, "Erreur", f"PDF illisible ({reason[:120]})", hint)
            else:
                blank = [p.number for p in pages if len(re.sub(r"\s", "", p.text)) < 40]
                if blank:
                    add_issue(issues, name, "Avertissement",
                              f"{len(blank)} page(s) sans texte lisible (p. {', '.join(map(str, blank[:8]))}"
                              f"{'…' if len(blank) > 8 else ''}) — probablement des scans",
                              "Activez un moteur OCR (barre latérale) ou la lecture IA pour les lire.")
            return pages
        if kind == "image":
            if not ocr.available:
                add_issue(issues, name, "Erreur", "image : un moteur OCR est nécessaire pour la lire",
                          "Activez Tesseract/EasyOCR (barre latérale) ou utilisez un PDF texte / la lecture IA.")
                return []
            from PIL import Image, ImageSequence
            img = Image.open(io.BytesIO(data))
            pages = []
            for i, frame in enumerate(ImageSequence.Iterator(img), start=1):
                words = ocr.image_to_words(frame.convert("RGB"))
                pages.append(PageData(name, i, words_to_text(words), [], words_to_tables(words), True, "AUTRE", "", "image"))
            if not any(p.text.strip() for p in pages):
                add_issue(issues, name, "Avertissement", "aucun texte reconnu sur l'image",
                          "Vérifiez la netteté / la résolution (≥ 200 dpi) ou augmentez la résolution OCR.")
            return pages
        sheets = fileio.read_any(name, kind, data)
        return [PageData(name, i, sp.text, sp.tables, [], False, "AUTRE", sp.decimal, kind)
                for i, sp in enumerate(sheets, start=1)]
    except fileio.FileReadError as exc:
        add_issue(issues, name, "Erreur", str(exc), "Ce fichier est ignoré : corrigez-le ou retirez-le de la liste.")
        log.append(f"❌ {name} ignoré : {exc}")
    except Exception as exc:  # garde-fou : aucun fichier ne doit faire planter l'application
        add_issue(issues, name, "Erreur", f"lecture impossible ({type(exc).__name__}: {str(exc)[:120]})",
                  "Réexportez le fichier ou convertissez-le en PDF / Excel (.xlsx).")
        log.append(f"❌ {name} : {type(exc).__name__}: {exc}")
    return []


def page_text_for_ai(p: PageData) -> str:
    """Texte d'une page transmis à l'IA : les feuilles Excel / CSV / Word gardent leur structure en colonnes."""
    if p.src in ("excel", "csv", "word") and p.tables:
        return "\n".join("\t".join(r) for t in p.tables for r in t)
    return p.text


def analyse_documents(files: list, ocr: OCREngine, ai_cfg, log: list, progress, p0: float, p1: float,
                      issues: list | None = None, meta: dict | None = None):
    """Lit tous les fichiers (PDF, scans, images, Excel, CSV, Word, texte). Avec la lecture IA : Claude découpe et
    extrait chaque document, contrôles croisés sur le texte ; sans clé ou en cas d'échec : moteur de règles.
    `issues` reçoit les problèmes de fichiers (illisible, vide, protégé, scan sans OCR…).
    Renvoie (documents, {id(document): données IA}, lignes d'audit)."""
    pages_by_file: dict[str, list[PageData]] = {}
    data_by_file: dict[str, bytes] = {}
    span = (p1 - p0) * (0.35 if ai_cfg else 1.0)
    for f_idx, f in enumerate(files):
        try:
            data_by_file[f.name] = f.getvalue()
        except Exception as exc:
            add_issue(issues, f.name, "Erreur", f"fichier impossible à lire ({type(exc).__name__})", "Réimportez-le.")
            pages_by_file[f.name] = []
            continue

        def cb(i, total, f_idx=f_idx, name=f.name):
            progress.progress(min(1.0, p0 + span * (f_idx + i / max(total, 1)) / len(files)),
                              text=f"Lecture des fichiers… ({name} p.{i}/{total})")
        pages_by_file[f.name] = read_file_pages(f.name, data_by_file[f.name], ocr, log, issues, cb)
    all_pages = [p for f in files for p in pages_by_file.get(f.name, [])]
    for i, p in enumerate(all_pages, start=1):
        p.seq = i
    if meta is not None:
        meta["n_pages"] = len(all_pages)
        meta["n_files"] = sum(1 for f in files if pages_by_file.get(f.name))
    dec = set_decimal_default(p.text for p in all_pages)
    kinds = Counter(p.src for p in all_pages)
    log.append(f"📄 {len(all_pages)} pages lues ({sum(p.ocr for p in all_pages)} via OCR) — "
               + ", ".join(f"{fileio.KIND_LABEL.get(k, k)} {n}" for k, n in kinds.items()) + f" — décimale « {dec} »")

    ai_docs: dict[str, list] = {}
    todo = [f for f in files if pages_by_file.get(f.name)]
    if ai_cfg and ai_engine and todo:
        import concurrent.futures as cf
        done = 0
        p_ai = p0 + span
        progress.progress(min(1.0, p_ai), text=f"Lecture IA multilingue et contrôles croisés… (0/{len(todo)} fichier(s))")

        def run_ai(f):
            pg = pages_by_file[f.name]
            texts = {p.number: page_text_for_ai(p) for p in pg}
            if pg[0].src == "pdf":
                return ai_engine.extract_file(f.name, data_by_file[f.name], {p.number: p.text for p in pg}, ai_cfg)
            return ai_engine.extract_text_file(f.name, texts, ai_cfg)

        with cf.ThreadPoolExecutor(max_workers=min(3, len(todo))) as ex:
            futs = {ex.submit(run_ai, f): f.name for f in todo}
            for fut in cf.as_completed(futs):
                name = futs[fut]
                try:
                    ai_docs[name] = fut.result()
                    n_doc = len(ai_docs[name])
                    n_rel = sum(1 for d in ai_docs[name] if d.get("_rounds", 1) > 1)
                    log.append(f"🧠 {name} : {n_doc} document(s) lus par l'IA"
                               + (f", {n_rel} relu(s) en profondeur après contrôle" if n_rel else ""))
                except Exception as exc:
                    msg = ai_engine.explain_error(exc)
                    log.append(f"⚠️ {name} : lecture IA impossible ({msg}) — moteur de règles utilisé pour ce fichier")
                    add_issue(issues, name, "Avertissement", f"lecture IA impossible ({msg})",
                              "Le moteur de règles local a été utilisé pour ce fichier : vérifiez les valeurs.")
                done += 1
                progress.progress(min(1.0, p_ai + (p1 - p_ai) * done / len(todo)),
                                  text=f"Lecture IA multilingue et contrôles croisés… ({done}/{len(todo)} fichier(s))")

    docs: list[Document] = []
    ai_map: dict[int, dict] = {}
    audit: list[dict] = []
    for f in files:
        fpages = pages_by_file.get(f.name, [])
        if not fpages:
            continue
        if f.name in ai_docs and ai_docs[f.name]:
            by_num = {p.number: p for p in fpages}
            for d in ai_docs[f.name]:
                kind = d.get("doc_type") if d.get("doc_type") in ("FACTURE", "PACKING", "MANIFESTE", "CMR") else "AUTRE"
                pg = [by_num[n] for n in d["pages"] if n in by_num] or \
                     [PageData(f.name, n, "", []) for n in d["pages"]]
                doc = Document(kind, f.name, pg)
                docs.append(doc)
                ai_map[id(doc)] = d
                audit.append(ai_engine.audit_row(f.name, d))
        else:
            docs.extend(segment_documents(fpages))
    return docs, ai_map, audit


def docs_row(d: Document, ai: dict | None) -> dict:
    """Ligne de l'inventaire des documents (feuille INVENTAIRE_PAGES et onglet Documents)."""
    if ai:
        return {"Fichier": d.file, "Pages": d.page_label, "Type": d.kind,
                "N° détecté": clean_cell(ai.get("number")) or ", ".join(ai.get("related_delivery_numbers") or []),
                "Fournisseur (expéditeur)": ai_engine.party(ai, "supplier"),
                "Client (destinataire)": ai_engine.party(ai, "client"),
                "OCR": "Oui" if any(p.ocr for p in d.pages) else "Non",
                "Lecture": f"IA ({(ai.get('language') or '?').upper()})",
                "Contrôles": "OK" if not ai.get("_issues") else f"{len(ai['_issues'])} à vérifier"}
    return {"Fichier": d.file, "Pages": d.page_label, "Type": d.kind,
            "N° détecté": find_invoice_number(d.text) if d.kind == "FACTURE" else find_bl_number(d.text),
            "Fournisseur (expéditeur)": find_supplier(d.text) if d.kind in ("FACTURE", "PACKING") else "",
            "Client (destinataire)": find_client(d.text) if d.kind in ("FACTURE", "PACKING") else "",
            "OCR": "Oui" if any(p.ocr for p in d.pages) else "Non", "Lecture": "Règles", "Contrôles": ""}


def run_pipeline(pdf_files: list, ref_file, ocr: OCREngine, basis: str, progress, ai_cfg=None) -> dict:
    import shipment
    log: list[str] = []
    issues: list[dict] = []
    meta: dict = {}

    # ---------- Étape 1 : lecture (IA multilingue + contrôles, ou moteur de règles) et extraction
    progress.progress(0.02, text="Étape 1/3 : Lecture et extraction des fichiers...")
    docs, ai_map, audit = analyse_documents(pdf_files, ocr, ai_cfg, log, progress, 0.02, 0.6, issues, meta)
    invoices = [invoice_from_ai(ai_map[id(d)], d) if id(d) in ai_map else parse_invoice(d)
                for d in docs if d.kind == "FACTURE"]
    packings = [packing_from_ai(ai_map[id(d)], d) if id(d) in ai_map else parse_packing(d)
                for d in docs if d.kind == "PACKING"]
    manifest_rows = [r for d in docs if d.kind == "MANIFESTE"
                     for r in (manifest_from_ai(ai_map[id(d)]) if id(d) in ai_map else parse_manifest(d))]
    link_packing_lists(invoices, packings, docs)

    alerts: list[dict] = []
    # Noms de tiers : nettoyage (libellés parasites, OCR corrompu, adresses) puis repli sur le client principal
    for pl in packings:
        pl.supplier = cleaning.clean_supplier_name(pl.supplier, fallback="")
    for inv in invoices:
        inv.supplier = cleaning.clean_supplier_name(inv.supplier, fallback="")
        if not inv.supplier:
            inv.supplier = (next((pl.supplier for pl in inv.packing_lists if pl.supplier), "")
                            or cleaning.MAIN_CLIENT)
        if not inv.items:
            log.append(f"⚠️ Facture {inv.number or '?'} ({inv.supplier}) : aucune ligne article détectée "
                       f"({inv.doc.file} p.{inv.doc.page_label}) — saisie manuelle possible")
            alerts.append(shipment.make_alert("Critique", "Extraction", "Aucune ligne article détectée — saisie manuelle "
                                              f"possible ({inv.doc.file} p.{inv.doc.page_label})", inv.supplier, inv.number))
    seen_nums: dict[str, list] = {}
    for inv in invoices:
        if inv.number:
            seen_nums.setdefault(compact_ref(inv.number), []).append(inv)
    for key, group in seen_nums.items():
        if len(group) > 1:
            alerts.append(shipment.make_alert(
                "Critique", "Doublon", f"Facture {group[0].number} lue {len(group)} fois ({', '.join(dict.fromkeys(i.doc.file + ' p.' + i.doc.page_label for i in group))}) "
                "— risque de double comptage", group[0].supplier, group[0].number))
    if not invoices and packings:
        alerts.append(shipment.make_alert("Avertissement", "Liaison documents",
                                          "Packing lists / BL reçus sans aucune facture : aucun poids ne peut être ventilé"))

    docs_df = pd.DataFrame([docs_row(d, ai_map.get(id(d))) for d in docs])
    log.extend(f"⚠️ Contrôle lecture IA — {i}" for r in audit for i in r["_issues"])
    groups_df = shipment.build_groups(invoices, packings, manifest_rows, docs)

    # ---------- Étape 2 : réconciliation + matching NGP
    progress.progress(0.62, text="Étape 2/3 : Réconciliation, ventilation des poids et matching NGP...")
    refdb = load_reference(ref_file.getvalue(), ref_file.name, log) if ref_file else {"exact": {}, "compact": {}, "ocr": {}, "sheet_names": {}, "count": 0}
    rows = [r for inv in invoices for r in build_invoice_rows(inv, basis, alerts)]
    init_cols = [c for c in COLUMNS if c not in ("NGP (REF_PAR_FRS)", "Contrôle NGP", "Onglet REF_PAR_FRS")]
    all_cols = init_cols + INTERNAL_COLUMNS
    lines_df = pd.DataFrame(rows, columns=all_cols) if rows else pd.DataFrame(columns=all_cols)
    for c in NUMERIC_COLUMNS:
        if c in lines_df.columns:
            lines_df[c] = pd.to_numeric(lines_df[c], errors="coerce")
    for c in INT_COLUMNS:
        if c in lines_df.columns:
            lines_df[c] = pd.to_numeric(lines_df[c], errors="coerce").fillna(0).astype(int)
    for c in lines_df.columns:
        if c not in NUMERIC_COLUMNS and c not in INT_COLUMNS:
            lines_df[c] = lines_df[c].map(clean_cell)

    manifest_df = pd.DataFrame(manifest_rows, columns=MANIFEST_COLUMNS) if manifest_rows else pd.DataFrame(columns=MANIFEST_COLUMNS)
    for c in ("Poids (kg)", "Colis", "Goods Value"):
        if c in manifest_df.columns:
            manifest_df[c] = pd.to_numeric(manifest_df[c], errors="coerce")
    if manifest_rows and not lines_df.empty:
        # Job Ref manifest (références d'expédition) + poids brut manifest porté par la 1re ligne du fournisseur
        by_sup, _orph = shipment.match_manifest(manifest_rows, list(dict.fromkeys(lines_df["Client"])))
        for sup, mrows in by_sup.items():
            idx = lines_df.index[lines_df["Client"] == sup]
            if not len(idx) or not mrows:
                continue
            lines_df.loc[idx, "Job Ref manifest"] = ", ".join(dict.fromkeys(m["Shipment"] for m in mrows if m.get("Shipment")))
            wts = [m["Poids (kg)"] for m in mrows if m.get("Poids (kg)") is not None]
            if wts:
                w_man = float(sum(wts))
                lines_df.loc[idx[0], C_MANIF] = w_man
                g_inv = lines_df.loc[idx, C_GROSSF].sum(min_count=1)
                if pd.notna(g_inv) and weights_differ(float(g_inv), w_man):
                    alerts.append(shipment.make_alert(
                        diff_severity(float(g_inv), w_man), "Écart de poids",
                        f"Σ poids brut des factures {g_inv:,.3f} kg ≠ poids brut manifest {w_man:,.3f} kg "
                        f"(écart {w_man - g_inv:+,.3f} kg)", sup))
    # Poids : distribution obligatoire au prorata (aucun poids net ligne / facture laissé vide si un poids global existe)
    if not lines_df.empty:
        man_total = pd.to_numeric(manifest_df["Poids (kg)"], errors="coerce").sum(min_count=1) if not manifest_df.empty else None
        lines_df, wlog = cleaning.force_prorata_weight_distribution(
            lines_df, declared_gross_kg=float(man_total) if pd.notna(man_total) else None, basis=basis,
            col_net=C_NET, col_netf=C_NETF, col_grossf=C_GROSSF, col_manif=C_MANIF, col_base=C_BASE)
        for w in wlog:
            if w["Niveau"] == 0:
                alerts.append(shipment.make_alert("Avertissement", "Poids", f"{w['Lignes']} ligne(s) sans poids net : "
                                                  "aucun poids global disponible", w["Fournisseur"], w["N° facture"]))
            else:
                alerts.append(shipment.make_alert("Info", "Poids", f"Poids {w['Poids (kg)']:,.3f} kg ({w['Source']}) réparti "
                                                  f"au prorata sur {w['Lignes']} ligne(s)", w["Fournisseur"], w["N° facture"]))
    if manifest_df.empty:
        log.append("ℹ️ Aucun tableau de manifeste exploitable : saisissez les Goods Value dans l'onglet « Manifeste »")
    alerts.extend(shipment.group_alerts(groups_df))
    progress.progress(0.85, text="Étape 3/3 : Génération du classeur Excel...")
    return {"lines": lines_df, "manifest": manifest_df, "docs": docs_df, "refdb": refdb, "log": log,
            "audit": audit, "ai": bool(ai_map), "alerts": alerts, "groups": groups_df, "issues": issues,
            "n_pages": meta.get("n_pages", 0), "n_files": meta.get("n_files", 0), "basis": basis}


# =============================================================================
# GÉNÉRATION DU CLASSEUR EXCEL
# =============================================================================

HEADER_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name="Arial", size=10)
BOLD_FONT = Font(name="Arial", size=10, bold=True)
RED_FONT = Font(name="Arial", size=10, bold=True, color="C00000")
INPUT_FONT = Font(name="Arial", size=10, color="0000FF")
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
INPUT_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_OK = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
FILL_KO = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
FILL_WARN = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")


def style_header(ws, row: int, ncols: int, height: float = 30):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def assign_suppliers_to_manifest(manifest: pd.DataFrame, lines: pd.DataFrame) -> pd.DataFrame:
    """Rattache chaque ligne du manifeste au fournisseur des factures (nom similaire)."""
    suppliers = [s for s in lines["Client"].dropna().unique() if s]
    out = manifest.copy()
    mapped = []
    for _, r in out.iterrows():
        sup = ""
        exp = clean_cell(r.get("Expéditeur", ""))
        if sup == "" and suppliers and exp:
            best = max(suppliers, key=lambda s: company_similarity(s, exp))
            if company_similarity(best, exp) >= 0.6:
                sup = best
        mapped.append(sup or exp)
    out["_supplier"] = mapped
    return out


def build_workbook(lines: pd.DataFrame, manifest: pd.DataFrame, docs: pd.DataFrame = None, n_pages: int = 0,
                   alerts: pd.DataFrame | None = None, groups: pd.DataFrame | None = None) -> bytes:
    import shipment
    wb = Workbook()
    YELLOW_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")

    # ------------------------------------------------------------------ TOUS_CLIENTS
    ws = wb.active
    ws.title = "TOUS_CLIENTS"
    headers = [f"{PAGE_IN_PDF} ({n_pages} p.)" if (c == PAGE_IN_PDF and n_pages) else c for c in COLUMNS]
    ws.append(headers)
    # header styling: 1F3864 fill, white bold Arial 10, vertically centered, wrap_text, height 35
    for c_idx in range(1, len(COLUMNS) + 1):
        hcell = ws.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 35

    col_idx = {name: i + 1 for i, name in enumerate(COLUMNS)}
    col_ltr = {name: get_column_letter(i + 1) for i, name in enumerate(COLUMNS)}
    num_formats_tc = {
        "Qté": "#,##0.###", "Prix unitaire": "#,##0.00", "Valeur HT": "#,##0.00",
        C_NET: "#,##0.000", C_NETU: "#,##0.000", C_NETF: "#,##0.000", C_GROSSF: "#,##0.000", C_MANIF: "#,##0.000",
    }
    text_format_cols = {"N° facture", "Réf article", "NGP/code douanier (facture)"}
    sev_map = shipment.severity_map(alerts) if alerts is not None else {}
    msgs: dict = {}
    if alerts is not None and not alerts.empty:
        for _, al in alerts.iterrows():
            msgs.setdefault((clean_cell(al["N° facture"]), clean_cell(al["Item"])), []).append(f"{al['Sévérité']} : {al['Message']}")
    SEV_FILL = {0: PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE"),
                1: PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")}

    n = len(lines)
    for i, (_, r) in enumerate(lines.iterrows(), start=2):
        key = (clean_cell(r.get("N° facture")), clean_cell(r.get("Item")))
        sev = sev_map.get(key, 9)
        for j, name in enumerate(COLUMNS, start=1):
            val = r.get(name)
            if name == "Contrôle NGP":
                m_col = col_ltr["NGP (REF_PAR_FRS)"]
                l_col = col_ltr["NGP/code douanier (facture)"]
                val = (f'=IF(OR({m_col}{i}="{NON_TROUVE}",{m_col}{i}="",{l_col}{i}=""),"À vérifier",'
                       f'IF(LEFT({m_col}{i},8)=LEFT({l_col}{i},8),"OK","ÉCART"))')
            elif name in NUMERIC_COLUMNS:
                if val is None or (isinstance(val, float) and math.isnan(val)):
                    val = None
                else:
                    try:
                        val = float(val)
                        if name == C_COLIS and val.is_integer():
                            val = int(val)
                    except (ValueError, TypeError):
                        val = None
            elif name in INT_COLUMNS:
                try:
                    val = int(val) if val is not None and not (isinstance(val, float) and math.isnan(val)) else None
                except (ValueError, TypeError):
                    val = None
            else:
                val = clean_cell(val) or None
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = BODY_FONT
            if name in num_formats_tc:
                cell.number_format = num_formats_tc[name]
            if name in text_format_cols:
                cell.number_format = "@"
            if name == "Réf article" and val == NON_TROUVE:
                cell.font = RED_FONT
            if name in ("NGP (REF_PAR_FRS)", "Onglet REF_PAR_FRS"):
                cell.fill = YELLOW_FILL
            if sev in SEV_FILL and name in ("Remarque", C_BASE, "Item"):
                cell.fill = SEV_FILL[sev]
            if name == "Item" and key in msgs:
                from openpyxl.comments import Comment
                cell.comment = Comment("\n".join(msgs[key][:6]), "TIMAR")

    last = max(n + 1, 2)
    tc_widths = {
        "Client": 30, "Job Ref manifest": 24, "N° facture": 14, "Date": 12,
        "Item": 8, "Réf article": 20, "Désignation": 42, "Origine": 14,
        "Qté": 10, "Prix unitaire": 13, "Valeur HT": 13,
        "NGP/code douanier (facture)": 18, "NGP (REF_PAR_FRS)": 14,
        "Onglet REF_PAR_FRS": 16, "Contrôle NGP": 14, "Remarque": 50, "Page PDF": 9, PAGE_IN_PDF: 11,
        C_NET: 13, C_NETU: 15, C_NETF: 13, C_GROSSF: 13, C_COLIS: 11, C_MANIF: 15, C_BASE: 60,
    }
    for name, w in tc_widths.items():
        if name in col_ltr:
            ws.column_dimensions[col_ltr[name]].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    # Ranges used in CONTROLE / CONTROLE_FACTURES formulas (safe up to row 5000)
    MAX_ROW = 5000
    def tc_range(col_name):
        c = col_ltr[col_name]
        return f"TOUS_CLIENTS!${c}$2:${c}${MAX_ROW}"
    R_CLIENT = tc_range("Client")
    R_FACTURE = tc_range("N° facture")
    R_HT = tc_range("Valeur HT")
    R_CTRL = tc_range("Contrôle NGP")

    # ------------------------------------------------------------------ CONTROLE (8 cols)
    wc = wb.create_sheet("CONTROLE")
    ctrl_headers = ["Client", "Σ Valeur HT extraite", "Goods Value manifest",
                    "Écart", "Nb lignes extraites", "Nb À vérifier", "Nb ÉCART", "Statut"]
    wc.append(ctrl_headers)
    for c_idx in range(1, len(ctrl_headers) + 1):
        hcell = wc.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wc.row_dimensions[1].height = 24

    # collect unique suppliers using matched names (so manifest long names map to invoice short names)
    man_matched = assign_suppliers_to_manifest(manifest, lines) if not manifest.empty else pd.DataFrame(columns=["_supplier", "Goods Value"])
    man_sups = list(dict.fromkeys(man_matched["_supplier"].dropna().tolist())) if not man_matched.empty else []
    line_sups = list(lines["Client"].dropna().unique()) if not lines.empty else []
    all_sups = list(dict.fromkeys(man_sups + [s for s in line_sups if s not in man_sups]))

    for i, sup in enumerate(all_sups, start=2):
        msub = man_matched[man_matched["_supplier"] == sup] if not man_matched.empty else pd.DataFrame()
        gv = msub["Goods Value"].dropna().sum() if not msub.empty and "Goods Value" in msub.columns and msub["Goods Value"].notna().any() else None
        # determine statut
        nb_inv_found = int(lines[lines["Client"] == sup]["N° facture"].replace("", pd.NA).dropna().nunique()) if not lines.empty else 0
        if nb_inv_found > 0:
            statut_txt = "Toutes les factures extraites."
        else:
            statut_txt = "⚠ AUCUNE facture extraite pour l'instant"
        values = [
            sup,
            f"=SUMIFS({R_HT},{R_CLIENT},A{i})",
            gv,
            f"=ROUND(B{i}-C{i},2)",
            f"=COUNTIFS({R_CLIENT},A{i})",
            f'=COUNTIFS({R_CLIENT},A{i},{R_CTRL},"À vérifier")',
            f'=COUNTIFS({R_CLIENT},A{i},{R_CTRL},"ÉCART")',
            statut_txt,
        ]
        for j, v in enumerate(values, start=1):
            cell = wc.cell(row=i, column=j, value=v)
            cell.font = BODY_FONT
            if j == 3:  # Goods Value manifest — always yellow (user-editable)
                cell.fill = YELLOW_FILL
        for j, fmt in [(2, "#,##0.00"), (3, "#,##0.00"), (4, "#,##0.00"),
                       (5, "#,##0"), (6, "#,##0"), (7, "#,##0")]:
            wc.cell(row=i, column=j).number_format = fmt

    # TOTAL row
    tot_row = len(all_sups) + 2
    wc.cell(row=tot_row, column=1, value="TOTAL")
    for j in range(2, 8):
        L = get_column_letter(j)
        if j == 4:  # Écart = SUM
            wc.cell(row=tot_row, column=j, value=f"=SUM({L}2:{L}{tot_row - 1})")
        elif j == 3:  # Goods Value = SUM
            wc.cell(row=tot_row, column=j, value=f"=SUM({L}2:{L}{tot_row - 1})")
        else:
            wc.cell(row=tot_row, column=j, value=f"=SUM({L}2:{L}{tot_row - 1})")
    for j in range(1, len(ctrl_headers) + 1):
        c = wc.cell(row=tot_row, column=j)
        c.font = BOLD_FONT
    for j, fmt in [(2, "#,##0.00"), (3, "#,##0.00"), (4, "#,##0.00"),
                   (5, "#,##0"), (6, "#,##0"), (7, "#,##0")]:
        wc.cell(row=tot_row, column=j).number_format = fmt

    ctrl_widths = [38, 18, 14, 14, 18, 12, 10, 60]
    for j, w in enumerate(ctrl_widths, start=1):
        wc.column_dimensions[get_column_letter(j)].width = w
    wc.freeze_panes = "A2"
    wc.auto_filter.ref = f"A1:{get_column_letter(len(ctrl_headers))}1"

    # ------------------------------------------------------------------ CONTROLE_FACTURES (5 cols)
    wcf = wb.create_sheet("CONTROLE_FACTURES")
    cf_headers = ["Client", "N° facture", "Σ lignes extraites", "Total imprimé facture", "Écart"]
    wcf.append(cf_headers)
    for c_idx in range(1, len(cf_headers) + 1):
        hcell = wcf.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wcf.row_dimensions[1].height = 24

    # one row per unique invoice
    if not lines.empty and "N° facture" in lines.columns and "Client" in lines.columns:
        inv_pairs = lines[["Client", "N° facture"]].drop_duplicates().dropna(subset=["N° facture"])
        inv_pairs = inv_pairs[inv_pairs["N° facture"] != ""]
    else:
        inv_pairs = pd.DataFrame(columns=["Client", "N° facture"])

    for i, (_, row_inv) in enumerate(inv_pairs.iterrows(), start=2):
        sup_val = clean_cell(row_inv["Client"])
        inv_val = clean_cell(row_inv["N° facture"])
        values = [
            sup_val,
            inv_val,
            f"=SUMIFS({R_HT},{R_FACTURE},B{i})",
            None,  # printed total - empty by default
            f"=ROUND(C{i}-D{i},2)",
        ]
        for j, v in enumerate(values, start=1):
            cell = wcf.cell(row=i, column=j, value=v)
            cell.font = BODY_FONT
            if j == 2:
                cell.number_format = "@"
        for j, fmt in [(3, "#,##0.00"), (4, "#,##0.00"), (5, "#,##0.00")]:
            wcf.cell(row=i, column=j).number_format = fmt

    cf_widths = [34, 16, 18, 20, 12]
    for j, w in enumerate(cf_widths, start=1):
        wcf.column_dimensions[get_column_letter(j)].width = w
    wcf.freeze_panes = "A2"
    wcf.auto_filter.ref = f"A1:{get_column_letter(len(cf_headers))}1"

    # ------------------------------------------------------------------ MANIFEST (7 cols)
    wm = wb.create_sheet("MANIFEST")
    wm.append(MANIFEST_COLUMNS)
    for c_idx in range(1, len(MANIFEST_COLUMNS) + 1):
        hcell = wm.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wm.row_dimensions[1].height = 24

    man_num_cols = {"Poids (kg)", "Colis", "Goods Value"}
    if not manifest.empty:
        for _, mr in manifest.iterrows():
            row_vals = []
            for mc in MANIFEST_COLUMNS:
                v = mr.get(mc)
                if mc in man_num_cols:
                    try:
                        v = float(v) if v is not None and not (isinstance(v, float) and math.isnan(v)) else None
                    except (ValueError, TypeError):
                        v = None
                elif mc == "Page PDF":
                    try:
                        v = int(v) if v is not None and not (isinstance(v, float) and math.isnan(v)) else None
                    except (ValueError, TypeError):
                        v = None
                else:
                    v = clean_cell(v) or None
                row_vals.append(v)
            wm.append(row_vals)
    man_last = max(len(manifest) + 1, 2) if not manifest.empty else 2
    # format Goods Value column F
    gv_col = get_column_letter(MANIFEST_COLUMNS.index("Goods Value") + 1)
    for row_m in wm.iter_rows(min_row=2, max_row=man_last):
        for cell_m in row_m:
            cell_m.font = BODY_FONT
        row_m[MANIFEST_COLUMNS.index("Goods Value")].number_format = "#,##0.00"
    man_widths = [9, 20, 38, 12, 8, 14, 9]
    for j, w in enumerate(man_widths, start=1):
        wm.column_dimensions[get_column_letter(j)].width = w
    wm.freeze_panes = "A2"
    wm.auto_filter.ref = f"A1:{get_column_letter(len(MANIFEST_COLUMNS))}1"

    # ------------------------------------------------------------------ INVENTAIRE_PAGES (4 cols)
    wi = wb.create_sheet("INVENTAIRE_PAGES")
    ip_headers = ["Expéditeur / type", "Pages", "Type", "Statut"]
    wi.append(ip_headers)
    for c_idx in range(1, len(ip_headers) + 1):
        hcell = wi.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wi.row_dimensions[1].height = 24

    if docs is not None and not docs.empty:
        # get set of suppliers that have extracted lines
        extracted_sups = set(lines["Client"].dropna().unique()) if not lines.empty else set()
        for _, doc_row in docs.iterrows():
            doc_type = str(doc_row.get("Type", ""))
            if doc_type not in ("FACTURE", "PACKING"):
                continue
            sup_name = str(doc_row.get("Fournisseur (expéditeur)", "") or doc_row.get("Fournisseur", "") or "")
            pages = str(doc_row.get("Pages", "") or "")
            ocr_val = str(doc_row.get("OCR", "") or "")
            type_label = "Scanné" if ocr_val == "Oui" else "Texte"
            # check if extracted
            found_sup = any(company_similarity(sup_name, s) >= 0.6 for s in extracted_sups) if sup_name else False
            statut = "Extrait" if found_sup else "À lire (probable facture manquante...)"
            wi.append([sup_name or doc_type, pages, type_label, statut])
        ip_last = max(len(docs) + 1, 2)
    else:
        ip_last = 2
    for row_i in wi.iter_rows(min_row=2, max_row=ip_last):
        for cell_i in row_i:
            cell_i.font = BODY_FONT
    ip_widths = [34, 26, 16, 70]
    for j, w in enumerate(ip_widths, start=1):
        wi.column_dimensions[get_column_letter(j)].width = w
    wi.freeze_panes = "A2"

    # ------------------------------------------------------------------ ALERTES
    wa = wb.create_sheet("ALERTES")
    wa.append(shipment.ALERT_COLUMNS)
    style_header(wa, 1, len(shipment.ALERT_COLUMNS), 24)
    SEV_ROW = {"Critique": FILL_KO, "Avertissement": FILL_WARN}
    if alerts is not None and not alerts.empty:
        for _, al in alerts.iterrows():
            wa.append([clean_cell(al[c]) or None for c in shipment.ALERT_COLUMNS])
            for cell_a in wa[wa.max_row]:
                cell_a.font = BODY_FONT
                cell_a.alignment = Alignment(vertical="top", wrap_text=cell_a.column == 7)
            if al["Sévérité"] in SEV_ROW:
                wa.cell(row=wa.max_row, column=1).fill = SEV_ROW[al["Sévérité"]]
    for j, w in enumerate([14, 18, 30, 16, 8, 20, 110], start=1):
        wa.column_dimensions[get_column_letter(j)].width = w
    wa.freeze_panes = "A2"
    wa.auto_filter.ref = f"A1:G{max(wa.max_row, 2)}"

    # ------------------------------------------------------------------ GROUPES
    wg = wb.create_sheet("GROUPES")
    gcols = list(groups.columns) if groups is not None and len(groups.columns) else [
        "Fournisseur", "N° facture", "Facture", "Packing list / BL liés", "Manifeste", "Poids net facture",
        "Poids net PL / BL", "Poids brut facture", "Poids brut PL / BL", "Statut"]
    wg.append(gcols)
    style_header(wg, 1, len(gcols), 30)
    if groups is not None and not groups.empty:
        for _, g in groups.iterrows():
            vals = []
            for c in gcols:
                v = g[c]
                if isinstance(v, float) and math.isnan(v):
                    v = None
                vals.append(v if isinstance(v, (int, float)) or v is None else clean_cell(v) or None)
            wg.append(vals)
            for cell_g in wg[wg.max_row]:
                cell_g.font = BODY_FONT
                cell_g.alignment = Alignment(vertical="top", wrap_text=True)
                if isinstance(cell_g.value, float):
                    cell_g.number_format = "#,##0.000"
    for j, w in enumerate([30, 14, 28, 60, 34, 14, 14, 14, 14, 44], start=1):
        wg.column_dimensions[get_column_letter(j)].width = w
    wg.freeze_panes = "A2"

    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =============================================================================
# INTERFACE STREAMLIT
# =============================================================================

def compute_summary(lines: pd.DataFrame) -> dict:
    # group by internal _devise if present, otherwise no currency grouping
    if "_devise" in lines.columns and not lines.empty:
        by_cur = lines.groupby("_devise")["Valeur HT"].sum().to_dict()
    else:
        by_cur = {}
    return {
        "lignes": len(lines),
        "factures": len(lines[["Client", "N° facture"]].drop_duplicates()) if len(lines) else 0,
        "fournisseurs": lines["Client"].replace("", pd.NA).dropna().nunique() if len(lines) else 0,
        "non_trouves": int((lines["NGP (REF_PAR_FRS)"] == NON_TROUVE).sum()) if len(lines) else 0,
        "ecarts": int((lines["Contrôle NGP"] == "ÉCART").sum()) if len(lines) else 0,
        "a_verifier": int((lines["Contrôle NGP"] == "À vérifier").sum()) if len(lines) else 0,
        "ok": int((lines["Contrôle NGP"] == "OK").sum()) if len(lines) else 0,
        "par_devise": by_cur,
        "poids_net": float(pd.to_numeric(lines[C_NET], errors="coerce").dropna().sum()) if len(lines) else 0.0,
        "poids_brut": float(pd.to_numeric(lines[C_GROSSF], errors="coerce").dropna().sum()) if len(lines) else 0.0,
    }


def fmt(v: float, d: int = 2) -> str:
    return f"{v:,.{d}f}".replace(",", " ").replace(".", ",")


def highlight_rows(row):
    styles = [""] * len(row)
    cols = list(row.index)
    if "NGP (REF_PAR_FRS)" in cols and row.get("NGP (REF_PAR_FRS)") == NON_TROUVE:
        styles[cols.index("NGP (REF_PAR_FRS)")] = "color:#C00000;font-weight:bold"
    if "Réf article" in cols and row.get("Réf article") == NON_TROUVE:
        styles[cols.index("Réf article")] = "color:#C00000;font-weight:bold"
    ctrl = row.get("Contrôle NGP")
    color = {"OK": "#C6EFCE", "ÉCART": "#FFC7CE", "À vérifier": "#FFEB9C"}.get(ctrl)
    if color and "Contrôle NGP" in cols:
        styles[cols.index("Contrôle NGP")] = f"background-color:{color}"
    return styles


def render_audit(audit: list[dict]):
    """Onglet « Analyse IA & contrôles » : synthèse par document + anomalies détaillées."""
    if not audit:
        ui.status_banner("warn", "", "Lecture IA non utilisée pour ce dossier",
                         "Les documents ont été lus par le moteur de règles. Activez la lecture IA dans la barre "
                         "latérale (clé API Anthropic) pour la lecture multilingue et les contrôles croisés.")
        return
    n_ok = sum(1 for r in audit if not r["_issues"])
    langs = sorted({r["Langue"] for r in audit if r["Langue"]})
    ui.kpi_cards([
        {"icon": "file", "label": "Documents lus", "value": len(audit), "sub": "factures, BL, manifestes…", "tone": "blue"},
        {"icon": "check", "label": "Vérifiés sans anomalie", "value": n_ok, "sub": f"sur {len(audit)}", "tone": "good"},
        {"icon": "alert", "label": "À vérifier", "value": len(audit) - n_ok, "sub": "voir le détail ci-dessous",
         "tone": "warn" if len(audit) - n_ok else "good"},
        {"icon": "sparkles", "label": "Relectures approfondies", "value": sum(1 for r in audit if r["Lectures"] > 1),
         "sub": "après contrôle automatique", "tone": "cyan"},
        {"icon": "globe", "label": "Langues détectées", "value": " · ".join(langs) or "—", "sub": "", "tone": "gold"},
    ])
    ui.section("Synthèse par document", "layers")
    ui.audit_table(audit)
    st.caption("Contrôles automatiques : quantité × prix = montant, somme des lignes = total HT, poids net ≤ brut, "
               "somme des poids des lignes = poids total, et présence de chaque numéro, référence et montant dans le "
               "texte du PDF (protection contre toute valeur inventée).")
    if any(r["_issues"] or r["_notes"] for r in audit):
        ui.section("Points à vérifier et remarques du lecteur", "alert")
        ui.audit_details(audit)


def sidebar_ai():
    """Réglages de la lecture IA ; renvoie un ai_engine.AIConfig ou None (moteur de règles seul)."""
    ui.sidebar_cap("Lecture IA multilingue", "sparkles")
    if ai_engine is None:
        ui.sidebar_pill(ui.WARNING, "Module IA absent — installez : <code>pip install --user anthropic</code>")
        return None
    if not st.toggle("Lire les documents avec Claude", value=True, key="ai_on",
                     help="Toutes langues, scans et mises en page complexes ; chaque valeur est ensuite contrôlée "
                          "(calculs, totaux, poids, présence dans le PDF) et relue en profondeur si besoin."):
        ui.sidebar_pill(ui.INK_3, "Moteur de règles seul (hors ligne)")
        return None
    preset = os.environ.get("ANTHROPIC_API_KEY", "")
    if not preset:
        try:
            preset = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:
            preset = ""
    typed = st.text_input("Clé API Anthropic", type="password", key="ai_key", placeholder="sk-ant-…",
                          help="Ou définissez ANTHROPIC_API_KEY (variable d'environnement ou .streamlit/secrets.toml).")
    key = typed.strip() or preset
    label = st.selectbox("Modèle", list(ai_engine.MODELS), key="ai_model")
    deep = st.toggle("Vérification approfondie", value=True, key="ai_deep",
                     help="Si un contrôle échoue, le document est relu entièrement avec la liste des anomalies.")
    if not key:
        ui.sidebar_pill(ui.WARNING, "Clé API requise — sans clé, le moteur de règles est utilisé")
        return None
    cfg = ai_engine.AIConfig(api_key=key, model=ai_engine.MODELS[label], deep=deep)
    if st.button("Tester la connexion", key="ai_test", icon=":material/wifi_tethering:", use_container_width=True):
        st.session_state["ai_test_result"] = ai_engine.check_key(cfg)
    ok, msg = st.session_state.get("ai_test_result", (True, ""))
    if not ok:
        ui.sidebar_pill(ui.CRITICAL, msg)
    else:
        ui.sidebar_pill(ui.GOOD, f"Lecture IA active · <b>&nbsp;{label.split(' — ')[0]}</b>" + (" · connexion OK" if msg else ""))
    st.markdown('<div class="side-note">Les PDF sont transmis chiffrés à l\'API Anthropic pour lecture ; '
                'le coût est facturé sur votre compte Anthropic.</div>', unsafe_allow_html=True)
    return cfg


def main():
    st.set_page_config(page_title="TIMAR · Ventilation Groupage Alstom Fez", page_icon="🚛", layout="wide",
                       initial_sidebar_state="expanded")
    if "theme" not in st.session_state:
        st.session_state["theme"] = "Bleu Port"
    ui.inject_css()

    import database
    import views  # import local : évite l'import circulaire avec app.py
    database.init_db()

    # ---------------- Authentification : rien n'est accessible sans connexion
    user = views.current_user()
    if not user:
        with st.sidebar:
            ui.sidebar_brand()
        views.render_auth()
        ui.footer()
        return
    if user["role"] != "admin":                                   # utilisateur validé : consultation + téléchargement seulement
        with st.sidebar:
            ui.sidebar_brand()
            views.sidebar_account(user)
        views.render_user_space(user)
        ui.footer()
        return

    # ---------------- Paramètres (barre latérale) — Admin : application complète
    with st.sidebar:
        ui.sidebar_brand()
        views.sidebar_account(user)
        st.selectbox("🎨 Thème", list(ui.THEMES), key="theme", help="Change les couleurs de l'application.")
        ui.sidebar_cap("Module", "layers")
        module = st.radio("Module", [":material/move_to_inbox: IMPORT — Ventilation groupage",
                                     ":material/outbox: EXPORT — BL & Factures",
                                     ":material/admin_panel_settings: ADMIN — Gestion des Accès & Dispatch"], key="module")
        is_admin_panel = "ADMIN" in module
        is_export = "EXPORT" in module
        ai_cfg, ocr, basis = None, None, "Valeur HT"
        if not is_admin_panel:
            ai_cfg = sidebar_ai()
            if not is_export:
                ui.sidebar_cap("Ventilation des poids", "scale")
                basis = st.radio("Poids globaux répartis au prorata de :", ["Valeur HT", "Quantité", "Lignes"],
                                 format_func=lambda b: {"Valeur HT": "Valeur HT", "Quantité": "Quantité",
                                                        "Lignes": "Lignes (parts égales)"}[b],
                                 help="Utilisé quand une ligne n'a pas de poids propre : le poids net total de la facture "
                                      "(ou de la packing list / du BL) est réparti sur ces lignes.")
            ui.sidebar_cap("OCR — pages scannées (moteur de règles)", "search")
            ocr_mode = st.selectbox("Moteur OCR", ["Automatique", "Tesseract", "EasyOCR", "Désactivé"])
            tess_path = st.text_input("Chemin tesseract.exe (portable, optionnel)",
                                      value=shutil.which("tesseract") or "",
                                      help="Ex. C:\\Users\\moi\\Tesseract-OCR\\tesseract.exe — aucune installation admin requise.")
            dpi = st.slider("Résolution OCR (DPI)", 150, 400, 300, 50)
            ocr = OCREngine(ocr_mode if ocr_mode != "Désactivé" else "none", tess_path, dpi)
            if ocr.available:
                ui.sidebar_pill(ui.GOOD, f"OCR actif : <b>&nbsp;{ocr.engine}</b>")
            elif ocr_mode == "Désactivé":
                ui.sidebar_pill(ui.INK_3, "OCR désactivé")
            else:
                ui.sidebar_pill(ui.WARNING, "Aucun moteur OCR : PDF texte uniquement")
                st.caption("Pour les scans : `pip install --user easyocr`, ou indiquez un Tesseract portable ci-dessus.")
            st.write("")
            if st.button("Nouveau dossier", use_container_width=True, icon=":material/restart_alt:"):
                protected = ("module", "auth_user", "theme")
                for k in [k for k in st.session_state if (k.startswith("exp") if is_export else not k.startswith("exp"))]:
                    if k not in protected and not k.startswith("ai_"):
                        del st.session_state[k]
                st.rerun()

    if is_admin_panel:
        views.render_admin_panel()
        ui.footer()
        return

    ai_label = f"Lecture IA · {next(k for k, v in ai_engine.MODELS.items() if v == ai_cfg.model).split(' — ')[0]}" \
        if ai_cfg else ""
    if is_export:
        import export_engine  # import local : évite l'import circulaire avec app.py
        export_engine.render_export(ocr, ai_cfg, ai_label)
        ui.footer(ai_active=bool(ai_cfg))
        return

    views.render_import(ocr, ai_cfg, ai_label, basis)
    ui.footer(ai_active=bool(ai_cfg))


if __name__ == "__main__":
    main()
