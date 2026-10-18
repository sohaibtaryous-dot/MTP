# -*- coding: utf-8 -*-
"""
Lecture IA multilingue des documents de transport / douane (Claude, API Anthropic).

Principe « deep research » :
  1. Claude lit le PDF complet (texte + image : scans, tampons, tableaux, toutes langues)
     et renvoie un JSON structuré (schéma strict).
  2. Contrôles automatiques indépendants : qté × PU = montant, Σ lignes = total, net ≤ brut,
     Σ poids lignes = poids total, et présence réelle des références / numéros / totaux
     dans la couche texte du PDF (anti-invention).
  3. S'il reste des anomalies : relecture approfondie ciblée avec la liste des anomalies,
     puis nouveau contrôle. Le résultat le plus cohérent est retenu.
"""

from __future__ import annotations

import base64
import concurrent.futures as cf
import hashlib
import io
import json
import re
from dataclasses import dataclass

import anthropic

try:
    from PyPDF2 import PdfReader, PdfWriter
except Exception:  # pragma: no cover
    PdfReader = PdfWriter = None

MODELS = {
    "Claude Opus 5.5 — précision maximale": "claude-opus-5-5",
    "Claude Sonnet 5.5 — rapide": "claude-sonnet-5-5",
    "Claude Haiku 5.5 — économique": "claude-haiku-5-5",
}
DEFAULT_MODEL = "claude-opus-5-5"
# Repli serveur en cas de refus du filtre de sécurité (non disponible sur Haiku)
FALLBACK_MODELS = {"claude-opus-5-5", "claude-sonnet-5-5"}
FALLBACK_BETA = "server-side-fallback-2026-07-01"

CHUNK_PAGES = 40            # au-delà : le PDF est découpé en lots de pages
MAX_CHUNK_BYTES = 24_000_000
MAX_WORKERS = 4

DOC_TYPES = ["FACTURE", "PACKING", "MANIFESTE", "CMR", "AUTRE"]
LINE_TYPES = ["article", "frais", "remise"]


# =============================================================================
# SCHÉMA DE SORTIE (tout en chaînes : pas de type union, nombres normalisés par le modèle)
# =============================================================================

def _obj(props: dict) -> dict:
    return {"type": "object", "additionalProperties": False, "properties": props, "required": list(props)}


S = {"type": "string"}
PARTY = _obj({"name": S, "address": S, "country": S, "tax_id": S})
LINE = _obj({
    "page": {"type": "integer"},
    "line_type": {"type": "string", "enum": LINE_TYPES},
    "position": S, "reference": S, "customer_reference": S, "description": S,
    "quantity": S, "unit": S, "unit_price": S, "amount": S,
    "hs_code": S, "origin_country": S,
    "net_weight_kg": S, "unit_net_weight_kg": S, "gross_weight_kg": S, "packages": S, "po_number": S,
})
MANIFEST_ROW = _obj({
    "page": {"type": "integer"}, "shipment": S, "shipper": S, "consignee": S, "invoice_number": S,
    "gross_weight_kg": S, "packages": S, "goods_value": S, "currency": S,
})
DOCUMENT = _obj({
    "doc_type": {"type": "string", "enum": DOC_TYPES},
    "title_as_printed": S,
    "language": S,
    "pages": {"type": "array", "items": {"type": "integer"}},
    "number": S,
    "date": S,
    "currency": S,
    "incoterm": S,
    "supplier": PARTY,
    "client": PARTY,
    "related_invoice_numbers": {"type": "array", "items": S},
    "related_delivery_numbers": {"type": "array", "items": S},
    "po_numbers": {"type": "array", "items": S},
    "origin_country": S,
    "header_hs_code": S,
    "total_quantity": S,
    "total_ht": S,
    "vat_amount": S,
    "total_ttc": S,
    "net_weight_kg": S,
    "gross_weight_kg": S,
    "packages": S,
    "lines": {"type": "array", "items": LINE},
    "manifest_rows": {"type": "array", "items": MANIFEST_ROW},
    "confidence": {"type": "string", "enum": ["haute", "moyenne", "faible"]},
    "notes": {"type": "array", "items": S},
})
SCHEMA = _obj({"documents": {"type": "array", "items": DOCUMENT}})

SYSTEM_PROMPT = """Tu es l'analyste documentaire senior d'un commissionnaire en douane marocain (TIMAR, Tanger).
Tu lis des dossiers PDF de transport international : factures commerciales, proformas, packing lists,
bons de livraison, CMR, manifestes de groupage. Ils peuvent être dans n'importe quelle langue
(français, anglais, allemand, espagnol, italien, néerlandais, polonais, tchèque, portugais, turc, chinois,
arabe…), être scannés, tamponnés, en plusieurs colonnes ou sur plusieurs pages.

Ces données servent directement à une déclaration en douane : une erreur coûte cher. Exactitude avant tout.

Règles d'extraction :
1. Lis TOUTES les pages. Un PDF peut contenir plusieurs documents : crée une entrée par document distinct
   (une facture de 3 pages = un seul document). `pages` = numéros de pages (1 = première page du PDF fourni).
2. Recopie les identifiants exactement comme imprimés (numéros de facture/BL, références articles, codes SH,
   numéros de commande). N'invente jamais : si une information est absente ou illisible, mets "" et
   explique dans `notes`.
3. Nombres : écris-les normalisés, point décimal, sans séparateur de milliers, sans devise ni unité.
   Interprète la convention du document : "1.234,56" → "1234.56" ; "1,234.56" → "1234.56" ;
   "1 204,023 KG" → "1204.023". Garde toutes les décimales imprimées.
4. Poids toujours en kilogrammes : convertis grammes (÷1000), tonnes (×1000), livres (×0,45359237), onces
   (×0,0283495) et signale chaque conversion dans `notes`. `unit_net_weight_kg` = poids net UNITAIRE (par pièce)
   uniquement s'il est imprimé sur la ligne, sinon "".
5. Chaque ligne d'article imprimée = une entrée dans `lines`, dans l'ordre, sur toutes les pages.
   Ne fusionne jamais deux lignes, même si la référence se répète. N'inclus pas les lignes de sous-total,
   total, report ou TVA. Les frais (fret, emballage, assurance…) ont line_type "frais", les remises "remise".
   Les informations écrites sous une ligne (code SH, origine, poids, n° de commande) appartiennent à cette ligne.
6. `supplier` = émetteur / vendeur / exportateur / expéditeur. `client` = acheteur / facturé à / destinataire /
   importateur. Attention aux mises en page en deux colonnes (ex. « EXPORTATEUR: DESTINATAIRE: » puis les deux
   noms sur la ligne suivante) : ne mélange pas les deux sociétés.
7. `reference` = référence article du fournisseur (ou modèle / code article) ; `customer_reference` = référence
   du client (ex. « Alstom P/N », « your ref », « Kunden-Nr. »).
8. `total_ht` = total hors taxes de la facture ; s'il n'y a qu'un total et que la TVA est nulle ou exonérée,
   c'est ce total. `date` au format AAAA-MM-JJ. `currency` en code ISO 4217 (EUR, USD…).
   `origin_country` et les pays en code ISO 3166 alpha-2 (MA, FR, DE, CN…). `language` = code ISO 639-1.
9. Manifeste de groupage : une entrée par ligne du tableau dans `manifest_rows` (laisse `lines` vide).
   Packing list / BL : les lignes vont dans `lines` (quantités, poids, colis par ligne si imprimés).
10. Avant de répondre, vérifie toi-même : quantité × prix unitaire ≈ montant, somme des montants ≈ total,
    poids net ≤ poids brut, somme des poids des lignes ≈ poids total. En cas d'écart, relis la page.
    Si l'écart existe réellement dans le document, garde les valeurs imprimées et signale-le dans `notes`.
11. Fichiers Excel / Word / CSV : ils te sont remis en texte, pages séparées par « === PAGE n === » (n = numéro de
    feuille ou de page), cellules d'une ligne séparées par une tabulation. Applique les mêmes règles.
12. `confidence` : "haute" si tout est lisible et cohérent, "moyenne" ou "faible" sinon."""

FIRST_PASS = ("Analyse ce PDF et extrais toutes les données de chaque document selon le schéma. "
              "Fichier : {name}{chunk}.")
DEEP_PASS = """Deuxième lecture approfondie du même PDF (fichier : {name}{chunk}).

Une première extraction a donné le JSON ci-dessous. Des contrôles automatiques indépendants ont relevé
ces anomalies :
{issues}

Relis le document intégralement, page par page, en vérifiant chaque anomalie, chaque ligne, chaque nombre
et chaque identifiant. Corrige ce qui est faux, complète ce qui manque, retire ce qui n'est pas dans le
document. Si une anomalie correspond à la réalité du document (écart réellement imprimé), conserve les
valeurs imprimées et explique-le dans `notes`. Renvoie le JSON complet corrigé.

Première extraction :
{previous}"""


# =============================================================================
# CONFIGURATION / ERREURS
# =============================================================================

@dataclass
class AIConfig:
    api_key: str = ""
    model: str = DEFAULT_MODEL
    deep: bool = True
    max_rounds: int = 2


class AIError(Exception):
    pass


_CACHE: dict = {}


def make_client(cfg: AIConfig) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=cfg.api_key or None, max_retries=3, timeout=900.0)


def check_key(cfg: AIConfig) -> tuple[bool, str]:
    """Vérifie la clé avec un appel minimal (aucun document envoyé)."""
    try:
        client = make_client(cfg)
        client.models.retrieve(cfg.model)
        return True, "Clé valide"
    except anthropic.AuthenticationError:
        return False, "Clé API invalide"
    except anthropic.PermissionDeniedError:
        return False, "Clé sans accès à ce modèle"
    except anthropic.NotFoundError:
        return False, "Modèle indisponible pour ce compte"
    except anthropic.APIConnectionError:
        return False, "Pas de connexion Internet vers api.anthropic.com"
    except anthropic.APIStatusError as e:
        return False, f"Erreur API ({e.status_code})"
    except Exception as e:  # clé absente, etc.
        return False, str(e)[:120]


# =============================================================================
# NOMBRES / TEXTE
# =============================================================================

def num(v) -> float | None:
    """Convertit une chaîne normalisée par le modèle (avec tolérance aux formats locaux)."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v).replace(" ", "").replace(" ", ""))
    if not s or not re.search(r"\d", s):
        return None
    neg = s.startswith("-")
    s = s.lstrip("-")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = s.replace(",", "") if len(parts) > 2 else s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        x = float(s)
    except ValueError:
        return None
    return -x if neg else x


def compact(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def _digit_tokens(text: str) -> set[str]:
    """Nombres présents dans le texte du PDF, réduits à leurs chiffres (« 1.234,56 » → « 123456 »)."""
    out = set()
    for m in re.finditer(r"\d[\d.,'  ]*\d|\d", text or ""):
        tok = m.group(0)
        out.add(re.sub(r"\D", "", tok))
        for part in re.split(r"[  ]", tok):  # « 841 5,06 » : deux nombres séparés
            if part:
                out.add(re.sub(r"\D", "", part))
    return out


def _number_in_text(value: float, tokens: set[str]) -> bool:
    for d in (0, 1, 2, 3, 4):
        digits = re.sub(r"\D", "", f"{abs(value):.{d}f}")
        if digits in tokens or digits.lstrip("0") in tokens:
            return True
    return False


# =============================================================================
# CONTRÔLES (indépendants du modèle)
# =============================================================================

def verify_document(d: dict, page_texts: dict[int, str]) -> list[str]:
    issues: list[str] = []
    kind = d.get("doc_type", "AUTRE")
    label = f"{kind} {d.get('number') or '(sans n°)'}"
    lines = d.get("lines") or []
    articles = [l for l in lines if l.get("line_type") == "article"]
    text = "\n".join(page_texts.get(p, "") for p in (d.get("pages") or [])) or "\n".join(page_texts.values())
    has_text = len(re.sub(r"\s", "", text)) > 200  # PDF avec couche texte (pas un pur scan)
    tokens = _digit_tokens(text) if has_text else set()
    ctext = compact(text) if has_text else ""

    if kind == "FACTURE":
        for fld, lab in (("number", "numéro de facture"), ("date", "date"), ("currency", "devise"),
                         ("total_ht", "total HT")):
            if not str(d.get(fld) or "").strip():
                issues.append(f"{label} : {lab} absent")
        if not (d.get("supplier") or {}).get("name"):
            issues.append(f"{label} : fournisseur (émetteur) absent")
        if not (d.get("client") or {}).get("name"):
            issues.append(f"{label} : client (destinataire) absent")
        if not articles:
            issues.append(f"{label} : aucune ligne d'article extraite")
        if d.get("date") and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(d["date"])):
            issues.append(f"{label} : date « {d['date']} » non normalisée AAAA-MM-JJ")

    # Lignes : qté × PU ≈ montant
    bad_math = []
    for i, l in enumerate(lines, 1):
        q, p, a = num(l.get("quantity")), num(l.get("unit_price")), num(l.get("amount"))
        if q is not None and p is not None and a is not None and a:
            calc = q * p
            # tolère les prix « pour 100 / 1000 unités »
            if all(abs(calc / k - a) > max(0.02, abs(a) * 0.005) for k in (1, 100, 1000)):
                bad_math.append(f"ligne {i} ({l.get('reference') or '?'}) : {q:g} × {p:g} = {calc:.2f} ≠ {a:.2f}")
    if bad_math:
        issues.append(f"{label} : quantité × prix ≠ montant — " + " ; ".join(bad_math[:8]))

    # Σ montants ≈ total HT
    total = num(d.get("total_ht"))
    amounts = [num(l.get("amount")) for l in lines]
    if total is not None and lines and all(a is not None for a in amounts):
        s = sum(a if l.get("line_type") != "remise" else -abs(a) for a, l in zip(amounts, lines))
        s_art = sum(num(l.get("amount")) or 0 for l in articles)
        if abs(s - total) > max(0.05, abs(total) * 0.001) and abs(s_art - total) > max(0.05, abs(total) * 0.001):
            issues.append(f"{label} : somme des lignes {s:.2f} ≠ total HT {total:.2f} (écart {total - s:+.2f}) "
                          "— ligne manquante, dupliquée ou mal lue ?")

    # Poids
    net, gross = num(d.get("net_weight_kg")), num(d.get("gross_weight_kg"))
    if net is not None and gross is not None and net > gross + 1e-6:
        issues.append(f"{label} : poids net {net:g} kg > poids brut {gross:g} kg (colonnes inversées ?)")
    for key, lab, tot in (("net_weight_kg", "net", net), ("gross_weight_kg", "brut", gross)):
        vals = [num(l.get(key)) for l in articles]
        if tot is not None and vals and all(v is not None for v in vals):
            s = sum(vals)
            if abs(s - tot) > max(0.05, tot * 0.01):
                issues.append(f"{label} : somme des poids {lab} des lignes {s:.3f} ≠ total {tot:.3f} kg")
    for i, l in enumerate(articles, 1):
        n_, g_ = num(l.get("net_weight_kg")), num(l.get("gross_weight_kg"))
        if n_ is not None and g_ is not None and n_ > g_ + 1e-6:
            issues.append(f"{label} : ligne {i} poids net {n_:g} > brut {g_:g}")
            break

    # Anti-invention : tout doit exister dans la couche texte du PDF
    if has_text:
        if d.get("number") and compact(d["number"]) not in ctext:
            issues.append(f"{label} : le numéro « {d['number']} » est introuvable dans le texte du PDF")
        missing = [l.get("reference") for l in lines
                   if l.get("reference") and len(compact(l["reference"])) >= 3 and compact(l["reference"]) not in ctext]
        if missing:
            issues.append(f"{label} : référence(s) introuvable(s) dans le PDF : {', '.join(missing[:10])}")
        for fld, lab in (("total_ht", "total HT"), ("gross_weight_kg", "poids brut"), ("net_weight_kg", "poids net")):
            v = num(d.get(fld))
            if v and not _number_in_text(v, tokens):
                issues.append(f"{label} : {lab} {v:g} introuvable tel quel dans le PDF (calculé ou mal lu ?)")
        bad_amt = [f"{num(l.get('amount')):g}" for l in lines
                   if num(l.get("amount")) and not _number_in_text(num(l.get("amount")), tokens)]
        if bad_amt:
            issues.append(f"{label} : montant(s) de ligne introuvable(s) dans le PDF : {', '.join(bad_amt[:10])}")
    return issues


# =============================================================================
# APPELS API
# =============================================================================

def _doc_source(pdf_b64: str | None = None, text: str | None = None) -> dict:
    if text is not None:
        return {"type": "text", "media_type": "text/plain", "data": text}
    return {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}


def _call(client: anthropic.Anthropic, cfg: AIConfig, pdf_b64: str | None, instruction: str, effort: str,
          source: dict | None = None) -> dict:
    kwargs = dict(
        model=cfg.model,
        max_tokens=64000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            {"type": "document", "source": source or _doc_source(pdf_b64),
             "cache_control": {"type": "ephemeral"}},  # la relecture réutilise le document en cache
            {"type": "text", "text": instruction},
        ]}],
        output_config={"effort": effort, "format": {"type": "json_schema", "schema": SCHEMA}},
    )
    if cfg.model in FALLBACK_MODELS:
        kwargs.update(betas=[FALLBACK_BETA], fallbacks="default")
    try:
        msg = _stream(client, kwargs)
    except anthropic.BadRequestError as e:
        # Option non acceptée par ce compte / cette version : on réessaie sans elle plutôt que d'échouer
        text = str(e.message).lower()
        kwargs.pop("betas", None)
        kwargs.pop("fallbacks", None)
        if any(w in text for w in ("schema", "format", "output_config")):
            kwargs["output_config"] = {"effort": effort}
            kwargs["messages"][0]["content"][1]["text"] = (
                instruction + "\n\nRéponds uniquement par un objet JSON conforme à ce schéma, sans aucun autre texte :\n"
                + json.dumps(SCHEMA, ensure_ascii=False))
        msg = _stream(client, kwargs)
    if msg.stop_reason == "refusal":
        raise AIError("lecture refusée par le filtre de sécurité du modèle")
    if msg.stop_reason == "max_tokens":
        raise AIError("document trop long pour une seule réponse (sortie tronquée)")
    text = next((b.text for b in msg.content if b.type == "text"), "")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, re.S)  # mode sans schéma : JSON entouré de texte
        try:
            return json.loads(m.group(0)) if m else {}
        except json.JSONDecodeError as e:
            raise AIError(f"réponse JSON invalide ({e})") from e


def _stream(client: anthropic.Anthropic, kwargs: dict):
    if "betas" in kwargs or "fallbacks" in kwargs:
        with client.beta.messages.stream(**kwargs) as stream:
            return stream.get_final_message()
    with client.messages.stream(**kwargs) as stream:
        return stream.get_final_message()


def _split_pdf(data: bytes, n_known: int) -> list[tuple[int, bytes, int]]:
    """[(décalage de page, octets, nb pages)] — découpe les gros PDF en lots de pages."""
    if PdfReader is None:
        return [(0, data, max(n_known, 1))]
    try:
        reader = PdfReader(io.BytesIO(data))
        n = len(reader.pages)
    except Exception:
        return [(0, data, max(n_known, 1))]
    if n <= CHUNK_PAGES and len(data) <= MAX_CHUNK_BYTES:
        return [(0, data, n)]
    size = CHUNK_PAGES
    if len(data) > MAX_CHUNK_BYTES:  # scans lourds : lots plus petits
        size = max(5, min(CHUNK_PAGES, int(n * MAX_CHUNK_BYTES / len(data))))
    out = []
    for start in range(0, n, size):
        w = PdfWriter()
        end = min(n, start + size)
        for i in range(start, end):
            w.add_page(reader.pages[i])
        buf = io.BytesIO()
        w.write(buf)
        out.append((start, buf.getvalue(), end - start))
    return out


def _normalize(docs: list[dict], offset: int, n_pages: int) -> list[dict]:
    for d in docs:
        pages = sorted({int(p) + offset for p in (d.get("pages") or []) if 1 <= int(p) <= n_pages})
        d["pages"] = pages or [offset + 1]
        for coll in ("lines", "manifest_rows"):
            for r in d.get(coll) or []:
                if r.get("page"):
                    r["page"] = int(r["page"]) + offset
    return docs


def _read_chunk(client, cfg: AIConfig, name: str, offset: int, source: dict, n_chunk_pages: int,
                page_texts: dict[int, str], multi: bool) -> list[dict]:
    chunk = f" — pages {offset + 1} à {offset + n_chunk_pages}" if multi else ""
    docs = _normalize(_call(client, cfg, None, FIRST_PASS.format(name=name, chunk=chunk), "high", source)
                      .get("documents", []), offset, n_chunk_pages)
    issues = [i for d in docs for i in verify_document(d, page_texts)]
    rounds = 0
    while cfg.deep and issues and rounds < cfg.max_rounds:
        rounds += 1
        rel = [dict(d, pages=[p - offset for p in d["pages"]]) for d in docs]
        prompt = DEEP_PASS.format(name=name, chunk=chunk, issues="\n".join(f"- {i}" for i in issues),
                                  previous=json.dumps({"documents": rel}, ensure_ascii=False))
        try:
            again = _normalize(_call(client, cfg, None, prompt, "xhigh", source).get("documents", []), offset, n_chunk_pages)
        except AIError:
            break
        new_issues = [i for d in again for i in verify_document(d, page_texts)]
        if len(new_issues) <= len(issues):
            docs, issues = again, new_issues
        else:
            break
    for d in docs:
        d["_issues"] = verify_document(d, page_texts)
        d["_rounds"] = 1 + rounds
    return docs


def _merge_chunks(docs: list[dict]) -> list[dict]:
    """Recolle un document coupé entre deux lots de pages."""
    out: list[dict] = []
    for d in sorted(docs, key=lambda x: x["pages"][0]):
        prev = out[-1] if out else None
        if prev and prev["doc_type"] == d["doc_type"] and d["pages"][0] == prev["pages"][-1] + 1 \
                and d.get("_chunk_start") and (not d.get("number") or compact(d["number"]) == compact(prev.get("number"))):
            prev["pages"] = sorted(set(prev["pages"] + d["pages"]))
            prev["lines"] = (prev.get("lines") or []) + (d.get("lines") or [])
            prev["manifest_rows"] = (prev.get("manifest_rows") or []) + (d.get("manifest_rows") or [])
            for k in ("total_ht", "vat_amount", "total_ttc", "net_weight_kg", "gross_weight_kg", "packages",
                      "total_quantity"):
                if str(d.get(k) or "").strip():
                    prev[k] = d[k]   # les totaux sont en fin de document
            prev["notes"] = (prev.get("notes") or []) + (d.get("notes") or [])
            prev["_issues"] = (prev.get("_issues") or []) + (d.get("_issues") or [])
            continue
        out.append(d)
    return out


def extract_file(name: str, data: bytes, page_texts: dict[int, str], cfg: AIConfig) -> list[dict]:
    """Lit un PDF complet ; renvoie la liste des documents (dict au format SCHEMA + _issues, _rounds)."""
    key = (hashlib.sha256(data).hexdigest(), cfg.model, cfg.deep, cfg.max_rounds)
    if key in _CACHE:
        return json.loads(json.dumps(_CACHE[key]))
    client = make_client(cfg)
    chunks = _split_pdf(data, len(page_texts))
    multi = len(chunks) > 1
    results: list[dict] = []
    with cf.ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(chunks))) as ex:
        futs = {ex.submit(_read_chunk, client, cfg, name, off,
                          _doc_source(base64.standard_b64encode(blob).decode()), n, page_texts, multi): off
                for off, blob, n in chunks}
        for fut in cf.as_completed(futs):
            docs = fut.result()
            if multi and docs:
                docs[0]["_chunk_start"] = futs[fut] > 0
            results.extend(docs)
    results = _merge_chunks(results) if multi else results
    for d in results:
        d.pop("_chunk_start", None)
    _CACHE[key] = json.loads(json.dumps(results))
    return results


MAX_TEXT_CHARS = 450_000


def extract_text_file(name: str, page_blobs: dict[int, str], cfg: AIConfig) -> list[dict]:
    """Lit un fichier Excel / CSV / Word / texte déjà converti en texte (pages = feuilles) avec les mêmes contrôles."""
    nums = sorted(page_blobs)
    key = (hashlib.sha256("\x00".join(f"{n}:{page_blobs[n]}" for n in nums).encode()).hexdigest(),
           cfg.model, cfg.deep, cfg.max_rounds)
    if key in _CACHE:
        return json.loads(json.dumps(_CACHE[key]))
    chunks, cur, size = [], [], 0
    for n in nums:
        if cur and size + len(page_blobs[n]) > MAX_TEXT_CHARS:
            chunks.append(cur)
            cur, size = [], 0
        cur.append(n)
        size += len(page_blobs[n])
    if cur:
        chunks.append(cur)
    client = make_client(cfg)
    multi = len(chunks) > 1
    results: list[dict] = []
    offset = 0
    jobs = []
    for group in chunks:
        text = "\n".join(f"=== PAGE {i} ===\n{page_blobs[n]}" for i, n in enumerate(group, 1))
        jobs.append((offset, _doc_source(text=text), len(group), {i: page_blobs[n] for i, n in enumerate(group, 1)}))
        offset += len(group)
    with cf.ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(jobs))) as ex:
        futs = {ex.submit(_read_chunk, client, cfg, name, off, src, n, texts, multi): off for off, src, n, texts in jobs}
        for fut in cf.as_completed(futs):
            docs = fut.result()
            if multi and docs:
                docs[0]["_chunk_start"] = futs[fut] > 0
            results.extend(docs)
    results = _merge_chunks(results) if multi else results
    for d in results:
        d.pop("_chunk_start", None)
    _CACHE[key] = json.loads(json.dumps(results))
    return results


def explain_error(exc: Exception) -> str:
    if isinstance(exc, anthropic.AuthenticationError):
        return "clé API invalide"
    if isinstance(exc, anthropic.PermissionDeniedError):
        return "clé API sans accès au modèle choisi"
    if isinstance(exc, anthropic.RateLimitError):
        return "limite de débit de l'API atteinte — réessayez dans une minute"
    if isinstance(exc, anthropic.BadRequestError):
        return f"PDF refusé par l'API ({exc.message[:160]})"
    if isinstance(exc, anthropic.APIConnectionError):
        return "pas de connexion à api.anthropic.com"
    if isinstance(exc, anthropic.APIStatusError):
        return f"erreur API {exc.status_code}"
    return str(exc)[:200]


def party(d: dict, key: str) -> str:
    return ((d.get(key) or {}).get("name") or "").strip()


def iso_date_to_fr(s: str) -> str:
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", str(s or "").strip())
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else str(s or "").strip()


def audit_row(file: str, d: dict) -> dict:
    """Ligne du tableau « Analyse IA & contrôles »."""
    issues = d.get("_issues") or []
    return {
        "Fichier": file,
        "Pages": (f"{d['pages'][0]}" if len(d["pages"]) == 1 else f"{d['pages'][0]}-{d['pages'][-1]}"),
        "Type": d.get("doc_type", ""),
        "Langue": (d.get("language") or "").upper(),
        "N°": d.get("number", ""),
        "Fournisseur": party(d, "supplier"),
        "Client": party(d, "client"),
        "Date": iso_date_to_fr(d.get("date", "")),
        "Devise": d.get("currency", ""),
        "Total HT": num(d.get("total_ht")),
        "Poids net (kg)": num(d.get("net_weight_kg")),
        "Poids brut (kg)": num(d.get("gross_weight_kg")),
        "Colis": num(d.get("packages")),
        "Lignes": len(d.get("lines") or []) + len(d.get("manifest_rows") or []),
        "Lectures": d.get("_rounds", 1),
        "Confiance": d.get("confidence", ""),
        "Statut": "Vérifié" if not issues else f"{len(issues)} point(s) à vérifier",
        "_issues": issues,
        "_notes": d.get("notes") or [],
    }
