# -*- coding: utf-8 -*-
"""Dossier d'expédition : groupes de documents (facture ↔ packing list ↔ BL ↔ manifeste), alertes de cohérence
et synthèse par fournisseur. Les alertes statiques sont produites à l'extraction (comparaison facture / packing
list / manifeste) ; les alertes « live » sont recalculées sur le tableau tel que l'utilisateur l'a corrigé."""

from __future__ import annotations

import math
import re
from datetime import datetime, timedelta

import pandas as pd

import app as core

SEV_ORDER = {"Critique": 0, "Avertissement": 1, "Info": 2}
SEV_ICON = {"Critique": "🔴", "Avertissement": "🟠", "Info": "🔵"}
ALERT_COLUMNS = ["Sévérité", "Type", "Fournisseur", "N° facture", "Item", "Réf article", "Message"]


def make_alert(sev: str, typ: str, msg: str, supplier: str = "", invoice: str = "", item: str = "", ref: str = "") -> dict:
    return {"Sévérité": sev, "Type": typ, "Fournisseur": supplier, "N° facture": invoice, "Item": item,
            "Réf article": ref, "Message": msg}


weight_tol = core.weight_tol
weights_differ = core.weights_differ
severity_for_diff = core.diff_severity


# =============================================================================
# MANIFESTE
# =============================================================================

def match_manifest(manifest_rows: list[dict], suppliers: list[str], threshold: float = 0.6):
    """Rattache chaque ligne de manifeste à un fournisseur de facture (noms proches). -> ({fournisseur: [lignes]}, orphelines)"""
    by_sup: dict[str, list[dict]] = {s: [] for s in suppliers}
    orphans = []
    for mr in manifest_rows:
        exp = core.clean_cell(mr.get("Expéditeur", ""))
        best, score = None, 0.0
        for s in suppliers:
            sc = core.company_similarity(s, exp) if exp and s else 0.0
            if sc > score:
                best, score = s, sc
        if best is not None and score >= threshold:
            by_sup[best].append(mr)
        else:
            orphans.append(mr)
    return by_sup, orphans


# =============================================================================
# GROUPES DE DOCUMENTS
# =============================================================================

def _doc_ref(doc) -> str:
    return f"{doc.file} p.{doc.page_label}"


def build_groups(invoices: list, packings: list, manifest_rows: list[dict], docs: list) -> pd.DataFrame:
    """Une ligne par facture avec ses documents liés ; puis les documents orphelins."""
    suppliers = [i.supplier for i in invoices if i.supplier]
    by_sup, orphan_man = match_manifest(manifest_rows, list(dict.fromkeys(suppliers)))
    rows = []
    linked_pl = set()
    for inv in invoices:
        pls = inv.packing_lists
        for pl in pls:
            linked_pl.add(id(pl))
        man = by_sup.get(inv.supplier, [])
        pl_txt = " · ".join(f"{pl.number or 'PL'} ({_doc_ref(pl.doc)}) — {pl.link or 'lié'}" for pl in pls) or "—"
        man_txt = " · ".join(f"{m.get('Shipment') or '?'} ({core.fmt(m['Poids (kg)'], 3) + ' kg' if m.get('Poids (kg)') is not None else 'poids ?'})"
                             for m in man) or "—"
        pl_net = next((pl.net_total for pl in pls if pl.net_total is not None), None)
        pl_gross = next((pl.gross_total for pl in pls if pl.gross_total is not None), None)
        diffs = []
        if weights_differ(inv.net_total, pl_net):
            diffs.append(f"net {inv.net_total:g} ≠ {pl_net:g}")
        if weights_differ(inv.gross_total, pl_gross):
            diffs.append(f"brut {inv.gross_total:g} ≠ {pl_gross:g}")
        if not pls:
            status = "⚠ Sans packing list / BL"
        elif diffs:
            status = "⚠ Écart de poids : " + " ; ".join(diffs)
        elif manifest_rows and not man:
            status = "⚠ Absent du manifeste"
        else:
            status = "✅ Complet"
        rows.append({"Fournisseur": inv.supplier, "N° facture": inv.number, "Facture": _doc_ref(inv.doc),
                     "Packing list / BL liés": pl_txt, "Manifeste": man_txt,
                     "Poids net facture": inv.net_total, "Poids net PL / BL": pl_net,
                     "Poids brut facture": inv.gross_total, "Poids brut PL / BL": pl_gross, "Statut": status})
    for pl in packings:
        if id(pl) not in linked_pl:
            rows.append({"Fournisseur": pl.supplier, "N° facture": pl.invoice_ref or "", "Facture": "—",
                         "Packing list / BL liés": f"{pl.number or 'PL'} ({_doc_ref(pl.doc)})", "Manifeste": "—",
                         "Poids net facture": None, "Poids net PL / BL": pl.net_total,
                         "Poids brut facture": None, "Poids brut PL / BL": pl.gross_total,
                         "Statut": "⚠ Packing list / BL sans facture"})
    for m in orphan_man:
        rows.append({"Fournisseur": m.get("Expéditeur", ""), "N° facture": "", "Facture": "—",
                     "Packing list / BL liés": "—", "Manifeste": f"{m.get('Shipment') or '?'}",
                     "Poids net facture": None, "Poids net PL / BL": None,
                     "Poids brut facture": None, "Poids brut PL / BL": m.get("Poids (kg)"),
                     "Statut": "⚠ Manifeste sans facture"})
    cols = ["Fournisseur", "N° facture", "Facture", "Packing list / BL liés", "Manifeste", "Poids net facture",
            "Poids net PL / BL", "Poids brut facture", "Poids brut PL / BL", "Statut"]
    return pd.DataFrame(rows, columns=cols)


def group_alerts(groups: pd.DataFrame) -> list[dict]:
    out = []
    any_pl = bool((groups["Packing list / BL liés"] != "—").any()) if len(groups) else False
    for _, g in groups.iterrows():
        st = str(g["Statut"])
        if st.startswith("✅") or st.startswith("⚠ Écart"):   # écarts de poids déjà signalés ligne à ligne / par facture
            continue
        no_pl = "sans packing" in st.lower()
        sev = "Info" if (no_pl and not any_pl) else ("Avertissement" if ("sans facture" in st.lower() or no_pl) else "Info")
        typ = "Liaison documents" if "Écart" not in st else "Écart de poids"
        out.append(make_alert(sev, typ, st.replace("⚠ ", ""), supplier=g["Fournisseur"], invoice=g["N° facture"]))
    return out


# =============================================================================
# ALERTES « LIVE » (recalculées sur le tableau édité)
# =============================================================================

def _num(v):
    try:
        f = float(v)
        return None if math.isnan(f) else f
    except (TypeError, ValueError):
        return None


def live_alerts(df: pd.DataFrame) -> list[dict]:
    out: list[dict] = []
    if df is None or df.empty:
        return out
    C_NET, C_NETF, C_GROSSF = core.C_NET, core.C_NETF, core.C_GROSSF
    today = datetime.now()
    seen_lines: dict = {}
    inv_cur: dict = {}
    inv_sup: dict = {}
    for _, r in df.iterrows():
        sup, fac, item = core.clean_cell(r.get("Client")), core.clean_cell(r.get("N° facture")), core.clean_cell(r.get("Item"))
        ref = core.clean_cell(r.get("Réf article"))
        q, pu, ht = _num(r.get("Qté")), _num(r.get("Prix unitaire")), _num(r.get("Valeur HT"))
        net = _num(r.get(C_NET))
        base = core.clean_cell(r.get(core.C_BASE))
        is_fee = bool(core.FEE_WORDS.search(f"{r.get('Désignation', '')} {ref}"))

        def add(sev, typ, msg):
            out.append(make_alert(sev, typ, msg, sup, fac, item, ref))
        if q is not None and pu is not None and ht is not None and ht != 0:
            calc = q * pu
            if all(abs(calc / k - ht) > max(0.02, abs(ht) * 0.005) for k in (1, 100, 1000)):
                add("Avertissement", "Calcul", f"Qté × PU = {calc:,.2f} ≠ Valeur HT {ht:,.2f} (écart {ht - calc:+,.2f})")
        if ht is None and not is_fee:
            add("Avertissement", "Donnée manquante", "Valeur HT absente")
        if not ref and not is_fee:
            add("Avertissement", "Donnée manquante", "Réf article absente")
        if q is not None and q <= 0 and not is_fee:
            add("Info", "Valeur suspecte", f"Quantité ≤ 0 ({q:g})")
        if r.get("Contrôle NGP") == "ÉCART":
            add("Avertissement", "NGP", f"NGP référentiel {r.get('NGP (REF_PAR_FRS)')} ≠ NGP facture "
                                         f"{r.get('NGP/code douanier (facture)')}")
        if net is None and not is_fee:
            add("Avertissement", "Poids", "Poids net de ligne absent")
        elif net is not None and net < 0:
            add("Critique", "Poids", f"Poids net négatif ({net:g} kg)")
        nf, gf = _num(r.get(C_NETF)), _num(r.get(C_GROSSF))
        if nf is not None and gf is not None and nf > gf + 1e-6:
            add("Critique", "Poids", f"Poids net facture {nf:g} kg > poids brut facture {gf:g} kg")
        if gf is not None and gf > core.MAX_TRUCK_KG:
            add("Critique", "Poids", f"Poids brut facture {gf:,.1f} kg invraisemblable (> 40 t) — unité ou virgule mal lue ?")
        # date
        d = core.clean_cell(r.get("Date"))
        if d and (nf is not None or gf is not None or True):
            try:
                dt = datetime.strptime(d, "%d/%m/%Y")
                if dt > today + timedelta(days=3):
                    add("Avertissement", "Date", f"Date de facture dans le futur ({d})")
                elif dt < today - timedelta(days=3 * 365):
                    add("Info", "Date", f"Facture ancienne ({d})")
            except ValueError:
                add("Avertissement", "Date", f"Date « {d} » non reconnue (format JJ/MM/AAAA attendu)")
        elif not d and item in ("1", ""):
            add("Avertissement", "Donnée manquante", "Date de facture absente")
        if not fac and item in ("1", ""):
            add("Critique", "Donnée manquante", "N° de facture absent")
        cur = core.clean_cell(r.get("_devise"))
        inv_cur.setdefault((sup, fac), set()).add(cur)
        inv_sup.setdefault(fac, set()).add(sup)
        key = (sup, fac, ref, q, ht)
        if ref and key in seen_lines and ht is not None:
            add("Info", "Doublon", "Ligne identique déjà présente sur cette facture (même réf, quantité, montant)")
        seen_lines[key] = True
    for (sup, fac), curs in inv_cur.items():
        real = {c for c in curs if c}
        if len(real) > 1:
            out.append(make_alert("Avertissement", "Devise", f"Plusieurs devises sur la même facture : {', '.join(sorted(real))}", sup, fac))
        elif not real:
            out.append(make_alert("Avertissement", "Devise", "Devise non détectée", sup, fac))
    for fac, sups in inv_sup.items():
        if fac and len({s for s in sups if s}) > 1:
            out.append(make_alert("Avertissement", "Doublon", f"Le N° de facture {fac} apparaît chez {len(sups)} fournisseurs", "", fac))
    return out


def collect_alerts(static: list[dict], live: list[dict], issues: list[dict] | None = None) -> pd.DataFrame:
    rows = list(static or []) + list(live or [])
    for it in issues or []:
        sev = {"Erreur": "Critique", "Avertissement": "Avertissement"}.get(it.get("Gravité"), "Info")
        rows.append(make_alert(sev, "Fichier", f"{it.get('Fichier')} : {it.get('Problème')}"))
    df = pd.DataFrame(rows, columns=ALERT_COLUMNS) if rows else pd.DataFrame(columns=ALERT_COLUMNS)
    if df.empty:
        return df
    df = df.drop_duplicates(subset=["Type", "N° facture", "Item", "Message"]).copy()
    df["_o"] = df["Sévérité"].map(SEV_ORDER).fillna(3)
    return df.sort_values(["_o", "Fournisseur", "N° facture"], kind="stable").drop(columns="_o").reset_index(drop=True)


def severity_map(alerts: pd.DataFrame) -> dict:
    """(N° facture, Item) -> sévérité la plus haute (0 critique, 1 avertissement, 2 info). Item '' = niveau facture."""
    out: dict = {}
    if alerts is None or alerts.empty:
        return out
    for _, a in alerts.iterrows():
        k = (core.clean_cell(a["N° facture"]), core.clean_cell(a["Item"]))
        out[k] = min(out.get(k, 9), SEV_ORDER.get(a["Sévérité"], 3))
    return out


def row_severity(row: pd.Series, sev: dict) -> int:
    fac, item = core.clean_cell(row.get("N° facture")), core.clean_cell(row.get("Item"))
    return min(sev.get((fac, item), 9), 9)


# =============================================================================
# SYNTHÈSE PAR FOURNISSEUR
# =============================================================================

def supplier_summary(df: pd.DataFrame, alerts: pd.DataFrame | None = None) -> list[dict]:
    """Un dict par fournisseur : nb de lignes, poids net / brut (kg), valeur HT par devise, alertes."""
    out = []
    if df is None or df.empty:
        return out
    for sup, g in df.groupby("Client", sort=False):
        name = core.clean_cell(sup) or "À compléter"
        nums = lambda col: pd.to_numeric(g[col], errors="coerce")  # noqa: E731
        by_cur: dict = {}
        if "_devise" in g.columns:
            for cur, gg in g.groupby(g["_devise"].fillna("")):
                by_cur[core.clean_cell(cur) or "?"] = float(pd.to_numeric(gg["Valeur HT"], errors="coerce").sum())
        else:
            by_cur["?"] = float(nums("Valeur HT").sum())
        n_crit = n_warn = 0
        if alerts is not None and not alerts.empty:
            sub = alerts[alerts["Fournisseur"].map(core.clean_cell) == core.clean_cell(sup)]
            n_crit = int((sub["Sévérité"] == "Critique").sum())
            n_warn = int((sub["Sévérité"] == "Avertissement").sum())
        out.append({"Fournisseur": name, "Factures": int(g["N° facture"].replace("", pd.NA).dropna().nunique()),
                    "Lignes": int(len(g)), "Poids net (kg)": float(nums(core.C_NET).sum()),
                    "Poids brut (kg)": float(nums(core.C_GROSSF).sum()), "Valeur HT": by_cur,
                    "Critiques": n_crit, "Avertissements": n_warn})
    return out
