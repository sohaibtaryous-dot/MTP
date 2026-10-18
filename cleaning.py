# -*- coding: utf-8 -*-
"""Nettoyage des tiers (fournisseur / client) et distribution obligatoire des poids au prorata.

  clean_supplier_name(name)                : nom propre d'un tiers, ou repli (« CLIENT_NON_IDENTIFIE » / client principal).
  force_prorata_weight_distribution(df, …) : aucun poids net de ligne ni de facture laissé vide dès qu'un poids global
                                             existe (facture, packing list / BL, manifeste ou total de la déclaration).
Ce module ne dépend d'aucun autre module du projet (pandas + re uniquement).
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

FALLBACK_SUPPLIER = "CLIENT_NON_IDENTIFIE"
MAIN_CLIENT = "TIMAR / ALSTOM FES"

# --------------------------------------------------------------------------- libellés / listes noires
_STRONG_LABELS = (
    r"adresse\s+de\s+(?:livraison|facturation|exp[ée]dition)|delivery\s+address|invoice\s+address|billing\s+address|"
    r"shipping\s+address|ship(?:ping)?\s*[- ]?to|bill(?:ed)?\s*[- ]?to|sold\s*[- ]?to|deliver(?:y)?\s*to|invoice\s*to|"
    r"livr[ée]\s*[àa]|factur[ée]\s*[àa]|importer|exporter|consignee|consignor|lieferadresse|rechnungsadresse|"
    r"warenempf[äa]nger|rechnungsempf[äa]nger"
)
_WEAK_LABELS = (r"seller|supplier|vendor|buyer|customer|client|fournisseur|exp[ée]diteur|destinataire|acheteur|"
                r"empf[äa]nger|absender|verk[äa]ufer|from|to|von|soci[ée]t[ée]|company|firma")
LABEL_PREFIX_RE = re.compile(
    rf"^\s*(?:\[\d+\]\s*|\(\d+\)\s*)?(?:(?:{_STRONG_LABELS})\s*(?:[:\-–]\s*|\s+(?=\S))|(?:{_WEAK_LABELS})\s*[:\-–]\s*)",
    re.I)

GENERIC_ONLY = {
    "invoice", "facture", "importer", "exporter", "consignee", "consignor", "seller", "supplier", "vendor", "buyer",
    "customer", "client", "fournisseur", "destinataire", "expediteur", "acheteur", "adresse", "address", "delivery",
    "livraison", "date", "page", "total", "description", "details", "information", "contact", "packing list",
    "bon de livraison", "commercial invoice", "proforma", "importateur", "exportateur", "pour monsieur", "pour madame",
    "societe", "company", "non identifie", "inconnu", "unknown", "n a", "na", "none",
}
BLACKLIST_RE = re.compile(
    r"(?:adresse\s+de\s+(?:livraison|facturation)|delivery\s+address|billing\s+address|pour\s+(?:monsieur|madame|mme)|"
    r"[àa]\s+l['’]attention|\battn\b|deliver\s+to|ship\s+to|bill\s+to|invoice\s+(?:no|n°|number|date)|facture\s+n[°o]|"
    r"bon\s+de\s+livraison|packing\s+list|page\s+\d|bureau\s+d['’]\s*(?:exp[ée]dition|exportation)|en\s+fin\s+de\s+vie|"
    r"\biban\b|\bswift\b|\brib\b|\bn\s*°\s*tva|vat\s*(?:no|number)|date\s+d['’]?\s*(?:[ée]ch[ée]ance|facture)|"
    r"conditions\s+(?:de\s+)?(?:paiement|livraison)|total\s+(?:ht|ttc|invoice|amount)|montant|incoterm)", re.I)
BANK_RE = re.compile(r"\b(?:bank|banque|kantonalbank|sparkasse|financiers?|finances?|services\s+financiers|accounting|"
                     r"comptabilit[ée]|buchhaltung|tr[ée]sorerie|treasury|payment\s+services)\b", re.I)

# --------------------------------------------------------------------------- formes juridiques / adresses
LEGAL_FORM_RE = re.compile(
    r"(?:\b(?:GmbH|SARL|SASU?|SRL|SpA|S\.p\.A\.?|Ltd|Limited|LLC|Inc|Corp|Corporation|PLC|Kft|s\.?r\.?o|ApS|EURL|"
    r"S\.?A\.?S\.?|S\.?C\.?A\.?|Sp\.?\s*z\s*o\.?\s*o\.?|Co\.?\s*KG|& Co\b|Holding|Group|Groupe|Industries|Industrie)\b)|"
    r"(?-i:\b(?:AG|SA|S\.A\.|SE|NV|N\.V\.|BV|B\.V\.|SL|S\.L\.|KG|AB|Oy|A/S)\b)", re.I)
STREET_WORD = (r"rue|avenue|av\.|boulevard|bd\.?|blvd|chemin|route|impasse|all[ée]e|quai|place|parc|zone\s+indus\w*|"
               r"zone\s+industrielle|z\.?i\.?|z\.?a\.?c\.?|z\.?a\.?|lot(?:issement)?|bp|b\.p\.|cedex|po\s*box|p\.o\.\s*box|"
               r"street|st\.|road|rd\.|lane|way|drive|calle|c/|via|viale|piazza|strasse|straße|str\.|weg|laan|platz|"
               r"straat|quartier|km|n°|s/n|rue|carretera|polígono|poligono|industrial\s+estate|industriepark")
ADDRESS_RE = re.compile(rf"(?<![A-Za-zÀ-ÿ])(?:{STREET_WORD})(?![A-Za-zÀ-ÿ])", re.I)
STREET_PREFIX_RE = re.compile(
    rf"^.*?(?:straat|strasse|straße|str\.|weg|laan|platz|(?<![A-Za-zÀ-ÿ])(?:{STREET_WORD}))\s*\.?\s*"
    r"(?:\d+[A-Za-z]?(?:[-/]\d+)?)?[\s,;]*", re.I)
PERSON_PREFIX_RE = re.compile(
    r"^\s*(?:pour|à\s+l['’]attention\s+de|attn\.?:?|to\s+the\s+attention\s+of|zu\s+händen\s+von|for)\s+"
    r"(?:monsieur|madame|mme|m\.|mr\.?|mrs\.?|ms\.?|herr|frau)\b", re.I)
FEE_RE = re.compile(r"\b(freight|fret|transport|shipping|carriage|packing\s*cost|packaging|emballage|frais|surcharge|"
                    r"insurance|assurance|handling|verpackung|fracht|porto|palette\s*charge|pallet\s*charge|discount|"
                    r"remise|rabatt)\b", re.I)


def _plain(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    return re.sub(r"[^a-z ]", " ", "".join(c for c in s if not unicodedata.combining(c)).lower()).strip()


def has_legal_form(s: str) -> bool:
    return bool(LEGAL_FORM_RE.search(s or ""))


def _is_garbage(s: str) -> bool:
    """Chaînes corrompues par l'OCR : lettres répétées (« IIIINNNNVVVV »), symboles, mots de 1-2 lettres."""
    letters = re.sub(r"[^A-Za-zÀ-ÿ]", "", s)
    if not letters:
        return True
    if re.search(r"([A-Za-zÀ-ÿ])\1{2,}", s):                       # triple lettre : jamais dans un nom réel
        return True
    collapsed = re.sub(r"(.)\1+", r"\1", letters)
    if len(letters) >= 6 and (len(letters) - len(collapsed)) / len(letters) >= 0.30:   # IIIINNNNOOOO
        return True
    if sum(ch.isalpha() for ch in s) / max(len(s), 1) < 0.60:        # trop de symboles / chiffres
        return True
    if len(re.findall(r"[~|_^•¦¬\\{}<>=]", s)) >= 2:
        return True
    toks = re.findall(r"[A-Za-zÀ-ÿ0-9]+", s)
    if len(toks) >= 3 and sum(len(t) <= 2 for t in toks) / len(toks) > 0.6 and not has_legal_form(s):
        return True
    return False


def _strip_labels(s: str) -> str:
    for _ in range(3):
        m = LABEL_PREFIX_RE.match(s)
        if not m:
            break
        s = s[m.end():].strip(" :-–,;")
    return s


def clean_supplier_name(name: str, fallback: str = FALLBACK_SUPPLIER) -> str:
    """Nom propre d'un tiers (fournisseur / client) ou `fallback` si le texte n'est pas un nom de société.

    - retire les libellés (« Adresse de livraison : », « Deliver to », « Importer »…) en gardant le nom qui suit ;
    - « Pour Monsieur X, Directeur Général de Logitrade SAS » -> « Logitrade SAS » ;
    - rejette : chaînes OCR corrompues, adresses postales (sauf si une société avec forme juridique suit),
      services bancaires / financiers, libellés génériques (« Invoice », « Importer »…), numéros isolés ;
    - sinon : `fallback` (par défaut « CLIENT_NON_IDENTIFIE », ou le client principal de la déclaration)."""
    s = re.sub(r"[\u00a0\t]+", " ", str(name or "")).strip()
    s = re.sub(r"\s+", " ", s).strip(" :-–,;|•*~")
    if not s or s.lower() in ("nan", "none", "null", "à compléter", "a completer"):
        return fallback

    m = PERSON_PREFIX_RE.match(s)                                    # « Pour Monsieur X … de Logitrade SAS »
    if m:
        tail = re.search(r".*\b(?:de|d['’]|chez|of|at|bei)\s+([A-ZÀ-Ý][^,;]{2,70})$", s)
        s = tail.group(1).strip() if tail else ""
    s = _strip_labels(s)
    s = re.sub(r"^[\[(]\d+[\])]\s*", "", s)                          # « [2] KLN FREIGHT SPAIN »
    dup = re.fullmatch(r"(.{6,}?)\s+\1", s, flags=re.I)              # « Lot 106 ZI … Lot 106 ZI … »
    if dup:
        s = dup.group(1)
    s = s.strip(" :-–,;|•*~")
    if not s:
        return fallback
    if _plain(s) in GENERIC_ONLY or len(_plain(s)) < 3:
        return fallback

    legal = has_legal_form(s)
    if ADDRESS_RE.search(s) or re.match(r"^\d{1,5}\b", s):           # adresse postale : on cherche la société qui suit
        rest = STREET_PREFIX_RE.sub("", s, count=1).strip(" ,;:-")
        if rest and rest != s and has_legal_form(rest) and not ADDRESS_RE.search(rest) and not _is_garbage(rest):
            s, legal = rest, True
        elif not legal or ADDRESS_RE.search(re.sub(LEGAL_FORM_RE, "", s)):
            return fallback
    if _is_garbage(s):
        return fallback
    if BLACKLIST_RE.search(s) and not legal:
        return fallback
    if BANK_RE.search(s):
        return fallback
    if not legal:
        if re.match(r"^\d{4,}\b", s):                                # « 26619 en fin de vie »
            return fallback
        words = [w for w in re.findall(r"[A-Za-zÀ-ÿ&'’.-]+", s) if len(w) >= 3]
        if not words:
            return fallback
        if len(words) == 1 and (len(words[0]) < 4 or _plain(words[0]) in GENERIC_ONLY):
            return fallback
    return s[:80]


def is_valid_supplier_name(name: str) -> bool:
    return clean_supplier_name(name, fallback="") != ""


# =============================================================================
# POIDS : DISTRIBUTION OBLIGATOIRE AU PRORATA
# =============================================================================

def _allocate(total: float, weights: list[float], decimals: int = 3) -> list[float]:
    """Répartit `total` au prorata de `weights` ; le dernier absorbe l'arrondi (somme exacte)."""
    s = float(sum(weights))
    if s <= 0:
        weights, s = [1.0] * len(weights), float(len(weights))
    parts = [round(total * w / s, decimals) for w in weights]
    if parts:
        parts[-1] = round(total - sum(parts[:-1]), decimals)
    return parts


def _weights(sub: pd.DataFrame, col_value: str, col_qty: str, basis: str) -> list[float]:
    """Poids de répartition : valeur HT (défaut), sinon quantité, sinon parts égales."""
    val = pd.to_numeric(sub[col_value], errors="coerce").abs().fillna(0).tolist() if col_value in sub else [0] * len(sub)
    qty = pd.to_numeric(sub[col_qty], errors="coerce").abs().fillna(0).tolist() if col_qty in sub else [0] * len(sub)
    if basis == "Lignes":
        return [1.0] * len(sub)
    first, second = (qty, val) if basis == "Quantité" else (val, qty)
    for cand in (first, second):
        if sum(cand) > 0:
            return cand
    return [1.0] * len(sub)


def force_prorata_weight_distribution(
        df: pd.DataFrame, declared_net_kg: float | None = None, declared_gross_kg: float | None = None,
        basis: str = "Valeur HT", *, col_net="Poids net ligne (kg)", col_netf="Poids net facture (kg)",
        col_grossf="Poids brut facture (kg)", col_manif="Poids brut manifest – shipment (kg)",
        col_base="Base / remarque poids", col_value="Valeur HT", col_qty="Qté", col_supplier="Client",
        col_invoice="N° facture", col_desc="Désignation", col_ref="Réf article", group_col="_inv",
) -> tuple[pd.DataFrame, list[dict]]:
    """Garantit un poids net sur chaque ligne marchandise et un poids net facture sur chaque facture.

    Une ligne est « sans poids » si son poids net est NaN ou 0,0 (frais annexes exclus : poids 0, signalés). Cascade :
      1. poids net de la facture, sinon son poids brut        -> réparti sur les lignes sans poids de la facture ;
      2. poids brut du manifeste du fournisseur               -> réparti sur ses factures sans poids ;
      3. poids total de la déclaration (net, sinon brut)      -> réparti sur toutes les lignes encore sans poids.
    Formule : Poids net ligne = Poids à répartir × (Valeur HT ligne ÷ Σ Valeur HT des lignes concernées)
    (repli quantité puis parts égales si les valeurs sont absentes). Les poids déjà connus sont conservés : on ventile
    le reliquat, donc le total reste exact. Si aucune source de poids n'existe (niveau 1, 2 ni 3), la ligne reste vide
    et est signalée : un poids n'est jamais inventé.
    Retourne (tableau complété, journal des opérations [{Niveau, Fournisseur, N° facture, Poids (kg), Lignes, Source}])."""
    df = df.copy()
    log: list[dict] = []
    if df.empty:
        return df, log
    for c in (col_net, col_netf, col_grossf, col_manif, col_value, col_qty):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            df[c] = float("nan")
    if col_base not in df.columns:
        df[col_base] = ""
    df[col_base] = df[col_base].fillna("").astype(str)

    desc = df[col_desc].astype(str) if col_desc in df.columns else pd.Series("", index=df.index)
    refc = df[col_ref].astype(str) if col_ref in df.columns else pd.Series("", index=df.index)
    if "_fee" in df.columns:
        fee = df["_fee"].map(lambda v: str(v).lower() in ("true", "1", "1.0", "oui"))
    else:
        fee = (desc + " " + refc).map(lambda t: bool(FEE_RE.search(t)) and not re.search(r"\d", t.split()[-1] if t.split() else ""))
    goods = ~fee

    def tag(idx, text):
        base = df.loc[idx, col_base].map(lambda b: re.sub(r"(?:\s*\|\s*)?Poids net non trouvé", "", b).strip(" |"))
        df.loc[idx, col_base] = base.map(lambda b: f"{b} | {text}" if b else text)

    df.loc[goods & (df[col_net] <= 0), col_net] = float("nan")          # 0,0 = manquant
    fee_idx = df.index[fee & df[col_net].isna()]
    if len(fee_idx):
        df.loc[fee_idx, col_net] = 0.0
        tag(fee_idx, "Frais annexes (hors marchandise) : poids 0")

    inv_key = (df[group_col].astype(str) if group_col in df.columns
               else df[col_supplier].astype(str) + "§" + df[col_invoice].astype(str))
    inv_groups = {k: list(ix) for k, ix in df.groupby(inv_key, sort=False).groups.items()}

    def missing(ix):
        return [i for i in ix if goods[i] and pd.isna(df.at[i, col_net])]

    def distribute(ix, total, label, level, sup, fac, src):
        ix = list(ix)
        if not ix or total is None or total <= 0:
            return False
        sub = df.loc[ix]
        parts = _allocate(float(total), _weights(sub, col_value, col_qty, basis))
        for i, v in zip(ix, parts):
            df.at[i, col_net] = v
        tag(ix, f"Prorata {'valeur HT' if basis == 'Valeur HT' else basis.lower()} du {label} ({total:,.3f} kg)".replace(",", " "))
        log.append({"Niveau": level, "Fournisseur": sup, "N° facture": fac, "Poids (kg)": round(float(total), 3),
                    "Lignes": len(ix), "Source": src})
        return True

    # ---- 1) poids de la facture (net, sinon brut)
    for k, ix in inv_groups.items():
        miss = missing(ix)
        if not miss:
            continue
        sup, fac = df.at[ix[0], col_supplier], df.at[ix[0], col_invoice]
        netf = df.loc[ix, col_netf].dropna()
        grossf = df.loc[ix, col_grossf].dropna()
        if len(netf) and netf.iloc[0] > 0:
            total, label, src = float(netf.iloc[0]), "poids net facture", "poids net facture"
        elif len(grossf) and grossf.iloc[0] > 0:
            total, label, src = float(grossf.iloc[0]), "poids brut facture (net absent : brut utilisé)", "poids brut facture"
        else:
            continue
        known = float(df.loc[[i for i in ix if goods[i]], col_net].sum(skipna=True))
        remaining = max(total - known, 0.0)
        distribute(miss, remaining, label, 1, sup, fac, src)

    # ---- 2) poids brut du manifeste (par fournisseur), réparti sur ses factures encore sans poids
    sup_groups: dict = {}
    for k, ix in inv_groups.items():
        sup_groups.setdefault(df.at[ix[0], col_supplier], []).append(ix)
    for sup, groups in sup_groups.items():
        miss_groups = [g for g in groups if missing(g)]
        if not miss_groups:
            continue
        man = df.loc[[i for g in groups for i in g], col_manif].dropna()
        if not len(man) or man.sum() <= 0:
            continue
        known = float(df.loc[[i for g in groups for i in g if goods[i]], col_net].sum(skipna=True))
        remaining = max(float(man.sum()) - known, 0.0)
        ix_all = [i for g in miss_groups for i in missing(g)]
        distribute(ix_all, remaining, "poids brut manifest du fournisseur", 2, sup, "(toutes factures)", "poids brut manifest")

    # ---- 3) poids total de la déclaration, réparti sur tout ce qui reste sans poids
    still = [i for i in df.index if goods[i] and pd.isna(df.at[i, col_net])]
    glob = None
    if declared_net_kg and declared_net_kg > 0:
        glob, glabel = float(declared_net_kg), "poids net total de la déclaration"
    elif declared_gross_kg and declared_gross_kg > 0:
        glob, glabel = float(declared_gross_kg), "poids brut total de la déclaration (net absent : brut utilisé)"
    if still and glob:
        known = float(df.loc[[i for i in df.index if goods[i]], col_net].sum(skipna=True))
        distribute(still, max(glob - known, 0.0), glabel, 3, "(toute la déclaration)", "(toutes factures)", "déclaration")

    # ---- 4) poids net facture : rempli (Σ des lignes) sur la 1re ligne de chaque facture qui n'en a pas
    for k, ix in inv_groups.items():
        if df.loc[ix, col_netf].notna().any():
            continue
        s = df.loc[[i for i in ix if goods[i]], col_net].sum(skipna=True)
        if pd.notna(s) and s > 0 and not missing(ix):
            df.at[ix[0], col_netf] = round(float(s), 3)
            tag([ix[0]], "poids net facture = Σ poids nets des lignes")

    # ---- 5) ce qui reste sans aucune source : signalé, jamais inventé
    for k, ix in inv_groups.items():
        miss = missing(ix)
        if miss:
            sup, fac = df.at[ix[0], col_supplier], df.at[ix[0], col_invoice]
            tag(miss, "Poids net non trouvé (aucun poids global disponible)")
            log.append({"Niveau": 0, "Fournisseur": sup, "N° facture": fac, "Poids (kg)": None, "Lignes": len(miss),
                        "Source": "AUCUNE — renseignez le poids total de la déclaration"})
    return df, log
