# -*- coding: utf-8 -*-
"""Base de données SQLite : utilisateurs, validation par l'Admin, répartition (dispatch) des fichiers Excel.

Tables
  users              : id, email, password_hash, role ('admin'|'user'), status ('pending'|'active'),
                       assigned_files (JSON : liste des noms de dossiers attribués), created_at
  dossiers_dispatch  : id, dossier_name, assigned_to_email, file_paths (JSON : {"FACTURES.xlsx": chemin, ...}), created_at
  login_attempts     : protection contre les essais de mot de passe répétés (5 échecs / 15 min)

Sécurité : mots de passe hachés (PBKDF2-HMAC-SHA256, sel aléatoire, 260 000 itérations) — jamais stockés en clair.
Fichiers : conservés sur disque dans  data/dispatch/<id>/  ; un utilisateur ne peut lire que les siens.
Emplacement : variable d'environnement TIMAR_DATA_DIR (défaut : ./data).
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import time
import uuid
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from pathlib import Path

DATA_DIR = Path(os.environ.get("TIMAR_DATA_DIR", Path(__file__).resolve().parent / "data"))
DB_PATH = DATA_DIR / "timar.db"
DISPATCH_DIR = DATA_DIR / "dispatch"

ROLES = ("admin", "user")
STATUSES = ("pending", "active")
PBKDF2_ITERATIONS = 260_000
MIN_PASSWORD = 8
MAX_FAILS, FAIL_WINDOW_S = 5, 15 * 60
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL UNIQUE COLLATE NOCASE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'user' CHECK (role IN ('admin','user')),
    status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending','active')),
    assigned_files TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    email_verified INTEGER NOT NULL DEFAULT 1,
    verify_code_hash TEXT,
    verify_expires REAL,
    verify_sent_at REAL,
    verify_attempts INTEGER NOT NULL DEFAULT 0,
    subscription_end TEXT,
    plan TEXT NOT NULL DEFAULT 'free'
);
CREATE TABLE IF NOT EXISTS dossiers_dispatch (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dossier_name TEXT NOT NULL,
    assigned_to_email TEXT NOT NULL COLLATE NOCASE,
    file_paths TEXT NOT NULL,
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_dispatch_email ON dossiers_dispatch(assigned_to_email);
CREATE TABLE IF NOT EXISTS login_attempts (
    email TEXT NOT NULL COLLATE NOCASE,
    ts REAL NOT NULL
);
"""


# ----------------------------------------------------------------------------- connexion
@contextmanager
def connect():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=15)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    try:
        yield con
        con.commit()
    finally:
        con.close()


def init_db():
    """Crée les tables si besoin (idempotent) et, si TIMAR_ADMIN_EMAIL / TIMAR_ADMIN_PASSWORD sont définies,
    le compte Admin initial."""
    with connect() as con:
        con.executescript(SCHEMA)
        cols = {r["name"] for r in con.execute("PRAGMA table_info(users)").fetchall()}      # migration des bases existantes
        for name, ddl in (("email_verified", "INTEGER NOT NULL DEFAULT 1"), ("verify_code_hash", "TEXT"),
                          ("verify_expires", "REAL"), ("verify_sent_at", "REAL"),
                          ("verify_attempts", "INTEGER NOT NULL DEFAULT 0"), ("subscription_end", "TEXT"),
                          ("plan", "TEXT NOT NULL DEFAULT 'free'")):
            if name not in cols:
                con.execute(f"ALTER TABLE users ADD COLUMN {name} {ddl}")
        try:
            con.execute("PRAGMA journal_mode=WAL")
        except sqlite3.DatabaseError:
            pass
    email, pwd = os.environ.get("TIMAR_ADMIN_EMAIL", ""), os.environ.get("TIMAR_ADMIN_PASSWORD", "")
    if email and pwd and not admin_exists():
        create_user(email, pwd, role="admin", status="active")


# ----------------------------------------------------------------------------- mots de passe
def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return "pbkdf2_sha256${}${}${}".format(PBKDF2_ITERATIONS, base64.b64encode(salt).decode(), base64.b64encode(dk).decode())


def verify_password(password: str, stored: str) -> bool:
    try:
        algo, iters, salt_b64, hash_b64 = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), base64.b64decode(salt_b64), int(iters))
        return hmac.compare_digest(dk, base64.b64decode(hash_b64))
    except (ValueError, TypeError):
        return False


def normalize_email(email: str) -> str:
    return (email or "").strip().lower()


def validate_credentials(email: str, password: str) -> str:
    """Retourne un message d'erreur, ou '' si valide."""
    if not EMAIL_RE.match(normalize_email(email)):
        return "Adresse e-mail invalide."
    if len(password or "") < MIN_PASSWORD:
        return f"Le mot de passe doit contenir au moins {MIN_PASSWORD} caractères."
    return ""


# ----------------------------------------------------------------------------- utilisateurs
def _user(row) -> dict | None:
    if row is None:
        return None
    d = dict(row)
    try:
        d["assigned_files"] = json.loads(d.get("assigned_files") or "[]")
    except json.JSONDecodeError:
        d["assigned_files"] = []
    return d


def admin_exists() -> bool:
    with connect() as con:
        return con.execute("SELECT 1 FROM users WHERE role='admin' LIMIT 1").fetchone() is not None


def create_user(email: str, password: str, role: str = "user", status: str = "pending",
                verified: bool = True) -> tuple[bool, str]:
    """Crée un compte. Inscription publique : role='user', status='pending'. Retourne (ok, message)."""
    email = normalize_email(email)
    err = validate_credentials(email, password)
    if err:
        return False, err
    if role not in ROLES or status not in STATUSES:
        return False, "Rôle ou statut invalide."
    try:
        with connect() as con:
            con.execute("INSERT INTO users (email, password_hash, role, status, assigned_files, created_at, email_verified) VALUES (?,?,?,?,?,?,?)",
                        (email, hash_password(password), role, status, "[]", datetime.now().isoformat(timespec="seconds"), int(verified)))
    except sqlite3.IntegrityError:
        return False, "Un compte existe déjà avec cette adresse e-mail."
    return True, "Compte créé avec succès."


def get_user(email: str) -> dict | None:
    with connect() as con:
        return _user(con.execute("SELECT * FROM users WHERE email=?", (normalize_email(email),)).fetchone())


def list_users(status: str | None = None, role: str | None = None) -> list[dict]:
    q, args = "SELECT * FROM users WHERE email_verified=1", []
    if status:
        q += " AND status=?"
        args.append(status)
    if role:
        q += " AND role=?"
        args.append(role)
    with connect() as con:
        return [_user(r) for r in con.execute(q + " ORDER BY created_at DESC", args).fetchall()]


def set_password(email: str, new_password: str) -> tuple[bool, str]:
    """Nouveau mot de passe (réinitialisation par l'Admin) ; débloque aussi les essais échoués."""
    email = normalize_email(email)
    err = validate_credentials(email, new_password)
    if err:
        return False, err
    with connect() as con:
        n = con.execute("UPDATE users SET password_hash=? WHERE email=?", (hash_password(new_password), email)).rowcount
        con.execute("DELETE FROM login_attempts WHERE email=?", (email,))
    return (True, "Mot de passe modifié.") if n else (False, "Compte introuvable.")


def change_password(email: str, old_password: str, new_password: str) -> tuple[bool, str]:
    """Changement de mot de passe par l'utilisateur lui-même (ancien mot de passe exigé)."""
    u = get_user(email)
    if not u or not verify_password(old_password or "", u["password_hash"]):
        return False, "Mot de passe actuel incorrect."
    return set_password(email, new_password)


def approve_user(email: str, months: int | None = None, plan: str = "free") -> bool:
    """Active le compte. plan='free' : consultation / téléchargement. plan='pro' : outils complets pendant `months` mois
    (None = sans limite) à partir d'aujourd'hui."""
    with connect() as con:
        ok = con.execute("UPDATE users SET status='active' WHERE email=? AND role='user'", (normalize_email(email),)).rowcount > 0
    if ok:
        set_plan(email, plan, months)
    return ok


# ----------------------------------------------------------------------------- formules : Gratuit / Pro
def _add_months(d: date, months: int) -> date:
    y, m = divmod(d.month - 1 + months, 12)
    y, m = d.year + y, m + 1
    last = (date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1)).day
    return date(y, m, min(d.day, last))


def set_plan(email: str, plan: str, months: int | None = None) -> str | None:
    """Fixe la formule. 'free' : consultation seule. 'pro' : outils complets ; `months` prolonge la date de fin
    (depuis la fin actuelle si encore valide, sinon depuis aujourd'hui) ; months=None : Pro sans limite."""
    email = normalize_email(email)
    if plan not in ("free", "pro", "collab"):
        raise ValueError("plan invalide")
    end = None
    if plan == "pro" and months:
        u = get_user(email)
        cur = None
        if u and u.get("plan") == "pro" and u.get("subscription_end"):
            try:
                cur = date.fromisoformat(u["subscription_end"])
            except ValueError:
                cur = None
        start = cur if cur and cur >= date.today() else date.today()
        end = _add_months(start, int(months)).isoformat()
    with connect() as con:
        con.execute("UPDATE users SET plan=?, subscription_end=? WHERE email=? AND role='user'", (plan, end, email))
    return end


def set_subscription_end(email: str, end: str):
    """Fixe directement la date de fin de l'abonnement Pro (AAAA-MM-JJ)."""
    with connect() as con:
        con.execute("UPDATE users SET plan='pro', subscription_end=? WHERE email=? AND role='user'", (end, normalize_email(email)))


def subscription_info(user: dict) -> dict:
    """{'plan': 'admin'|'free'|'pro', 'effective': 'admin'|'free'|'pro', 'state': ..., 'end', 'days_left'}
    state : 'admin' | 'gratuit' | 'pro_illimite' | 'actif' | 'bientot' | 'expire'  (expire = Pro terminé → retombe en Gratuit)"""
    if user.get("role") == "admin":
        return {"plan": "admin", "effective": "admin", "state": "admin", "end": None, "days_left": None}
    if user.get("plan") == "collab":      # Pro offert par l'Admin : mêmes outils que Pro, jamais d'accès Admin
        return {"plan": "collab", "effective": "pro", "state": "collab", "end": None, "days_left": None}
    if (user.get("plan") or "free") != "pro":
        return {"plan": "free", "effective": "free", "state": "gratuit", "end": None, "days_left": None}
    end = user.get("subscription_end")
    if not end:
        return {"plan": "pro", "effective": "pro", "state": "pro_illimite", "end": None, "days_left": None}
    try:
        left = (date.fromisoformat(end) - date.today()).days
    except ValueError:
        return {"plan": "pro", "effective": "pro", "state": "pro_illimite", "end": None, "days_left": None}
    state = "expire" if left < 0 else ("bientot" if left <= 7 else "actif")
    return {"plan": "pro", "effective": "free" if left < 0 else "pro", "state": state, "end": end, "days_left": left}


def effective_plan(user: dict) -> str:
    return subscription_info(user)["effective"]


def suspend_user(email: str) -> bool:
    with connect() as con:
        return con.execute("UPDATE users SET status='pending' WHERE email=? AND role='user'", (normalize_email(email),)).rowcount > 0


def delete_user(email: str) -> bool:
    """Supprime un compte 'user' (jamais un admin) et ses attributions de dossiers."""
    email = normalize_email(email)
    with connect() as con:
        row = con.execute("SELECT role FROM users WHERE email=?", (email,)).fetchone()
        if not row or row["role"] == "admin":
            return False
        for r in con.execute("SELECT id FROM dossiers_dispatch WHERE assigned_to_email=?", (email,)).fetchall():
            _delete_dispatch_row(con, r["id"])
        con.execute("DELETE FROM users WHERE email=?", (email,))
    return True


# ----------------------------------------------------------------------------- connexion / anti-force-brute
def _recent_fails(con, email: str) -> int:
    con.execute("DELETE FROM login_attempts WHERE ts < ?", (time.time() - FAIL_WINDOW_S,))
    return con.execute("SELECT COUNT(*) FROM login_attempts WHERE email=?", (email,)).fetchone()[0]


def authenticate(email: str, password: str) -> tuple[dict | None, str]:
    """Retourne (utilisateur, code) — code : 'ok' | 'invalid' | 'unverified' | 'pending' | 'locked'."""
    email = normalize_email(email)
    with connect() as con:
        if _recent_fails(con, email) >= MAX_FAILS:
            return None, "locked"
        row = con.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        ok = bool(row) and verify_password(password or "", row["password_hash"])
        if not row:
            verify_password(password or "", hash_password("dummy"))        # temps de réponse constant
        if not ok:
            con.execute("INSERT INTO login_attempts (email, ts) VALUES (?, ?)", (email, time.time()))
            return None, "invalid"
        con.execute("DELETE FROM login_attempts WHERE email=?", (email,))
        user = _user(row)
    if user["role"] != "admin" and not user.get("email_verified", 1):
        return user, "unverified"
    if user["role"] != "admin" and user["status"] != "active":
        return None, "pending"
    return user, "ok"


# ----------------------------------------------------------------------------- dispatch
def _safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._\- ]", "_", Path(name).name) or "fichier"


def dispatch_dossier(dossier_name: str, emails: list[str], files: dict[str, bytes]) -> tuple[bool, str]:
    """Enregistre les fichiers sur disque (une seule copie) et crée une ligne par utilisateur destinataire.
    Seuls les comptes 'user' actifs peuvent recevoir un dossier."""
    dossier_name = (dossier_name or "").strip()
    if not dossier_name:
        return False, "Donnez un nom au dossier."
    if not files:
        return False, "Aucun fichier à envoyer."
    emails = [normalize_email(e) for e in dict.fromkeys(emails) if e]
    if not emails:
        return False, "Choisissez au moins un utilisateur."
    with connect() as con:
        for e in emails:
            row = con.execute("SELECT status, role FROM users WHERE email=?", (e,)).fetchone()
            if not row or row["role"] != "user" or row["status"] != "active":
                return False, f"« {e} » n'est pas un utilisateur actif."
        folder = DISPATCH_DIR / uuid.uuid4().hex
        folder.mkdir(parents=True, exist_ok=True)
        paths = {}
        for name, data in files.items():
            p = folder / _safe_name(name)
            p.write_bytes(data)
            paths[_safe_name(name)] = str(p)
        now = datetime.now().isoformat(timespec="seconds")
        for e in emails:
            con.execute("INSERT INTO dossiers_dispatch (dossier_name, assigned_to_email, file_paths, created_at) VALUES (?,?,?,?)",
                        (dossier_name, e, json.dumps(paths), now))
            row = con.execute("SELECT assigned_files FROM users WHERE email=?", (e,)).fetchone()
            lst = json.loads(row["assigned_files"] or "[]")
            if dossier_name not in lst:
                lst.append(dossier_name)
            con.execute("UPDATE users SET assigned_files=? WHERE email=?", (json.dumps(lst, ensure_ascii=False), e))
    return True, f"Dossier « {dossier_name} » attribué à {len(emails)} utilisateur(s)."


def _row_dispatch(r) -> dict:
    d = dict(r)
    try:
        d["file_paths"] = json.loads(d["file_paths"])
    except json.JSONDecodeError:
        d["file_paths"] = {}
    return d


def list_dispatch(email: str | None = None) -> list[dict]:
    with connect() as con:
        if email:
            rows = con.execute("SELECT * FROM dossiers_dispatch WHERE assigned_to_email=? ORDER BY created_at DESC, id DESC",
                               (normalize_email(email),)).fetchall()
        else:
            rows = con.execute("SELECT * FROM dossiers_dispatch ORDER BY created_at DESC, id DESC").fetchall()
    return [_row_dispatch(r) for r in rows]


def read_dispatch_file(dispatch_id: int, filename: str, email: str) -> bytes | None:
    """Contenu d'un fichier, uniquement si le dossier est attribué à `email` (et chemin à l'intérieur de data/dispatch)."""
    with connect() as con:
        r = con.execute("SELECT * FROM dossiers_dispatch WHERE id=? AND assigned_to_email=?",
                        (dispatch_id, normalize_email(email))).fetchone()
    if not r:
        return None
    path = _row_dispatch(r)["file_paths"].get(filename)
    if not path:
        return None
    p = Path(path).resolve()
    if DISPATCH_DIR.resolve() not in p.parents or not p.is_file():
        return None
    return p.read_bytes()


def _delete_dispatch_row(con, dispatch_id: int):
    r = con.execute("SELECT * FROM dossiers_dispatch WHERE id=?", (dispatch_id,)).fetchone()
    if not r:
        return
    d = _row_dispatch(r)
    con.execute("DELETE FROM dossiers_dispatch WHERE id=?", (dispatch_id,))
    row = con.execute("SELECT assigned_files FROM users WHERE email=?", (d["assigned_to_email"],)).fetchone()
    if row:
        remaining = {x["dossier_name"] for x in con.execute(
            "SELECT dossier_name FROM dossiers_dispatch WHERE assigned_to_email=?", (d["assigned_to_email"],)).fetchall()}
        lst = [n for n in json.loads(row["assigned_files"] or "[]") if n in remaining]
        con.execute("UPDATE users SET assigned_files=? WHERE email=?", (json.dumps(lst, ensure_ascii=False), d["assigned_to_email"]))
    for path in d["file_paths"].values():                           # supprime les fichiers si plus personne n'y fait référence
        if not con.execute("SELECT 1 FROM dossiers_dispatch WHERE file_paths LIKE ? LIMIT 1",
                           ("%" + json.dumps(path)[1:-1] + "%",)).fetchone():
            try:
                p = Path(path)
                p.unlink(missing_ok=True)
                if p.parent != DISPATCH_DIR and not any(p.parent.iterdir()):
                    p.parent.rmdir()
            except OSError:
                pass


def delete_dispatch(dispatch_id: int):
    with connect() as con:
        _delete_dispatch_row(con, dispatch_id)


# ----------------------------------------------------------------------------- vérification de l'e-mail
VERIFY_TTL_S, VERIFY_RESEND_S, VERIFY_MAX_TRIES = 15 * 60, 60, 5


def _code_hash(email: str, code: str) -> str:
    return hmac.new(email.encode(), code.encode(), hashlib.sha256).hexdigest()


def issue_verification_code(email: str) -> tuple[str | None, str]:
    """Génère un code à 6 chiffres (stocké haché, valable 15 min). Retourne (code, message) ; code=None si refusé."""
    email = normalize_email(email)
    with connect() as con:
        row = con.execute("SELECT email_verified, verify_sent_at FROM users WHERE email=?", (email,)).fetchone()
        if not row or row["email_verified"]:
            return None, "Compte introuvable ou déjà vérifié."
        wait = VERIFY_RESEND_S - (time.time() - (row["verify_sent_at"] or 0))
        if wait > 0:
            return None, f"Patientez {int(wait) + 1} s avant de redemander un code."
        code = f"{secrets.randbelow(1_000_000):06d}"
        con.execute("UPDATE users SET verify_code_hash=?, verify_expires=?, verify_sent_at=?, verify_attempts=0 WHERE email=?",
                    (_code_hash(email, code), time.time() + VERIFY_TTL_S, time.time(), email))
    return code, "Code généré."


def check_verification_code(email: str, code: str) -> tuple[bool, str]:
    email = normalize_email(email)
    code = re.sub(r"\D", "", code or "")
    with connect() as con:
        row = con.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if not row:
            return False, "Compte introuvable."
        if row["email_verified"]:
            return True, "E-mail déjà vérifié."
        if not row["verify_code_hash"] or time.time() > (row["verify_expires"] or 0):
            return False, "Code expiré : demandez un nouveau code."
        if row["verify_attempts"] >= VERIFY_MAX_TRIES:
            return False, "Trop d'essais : demandez un nouveau code."
        if not hmac.compare_digest(_code_hash(email, code), row["verify_code_hash"]):
            con.execute("UPDATE users SET verify_attempts=verify_attempts+1 WHERE email=?", (email,))
            return False, "Code incorrect."
        con.execute("UPDATE users SET email_verified=1, verify_code_hash=NULL, verify_expires=NULL, verify_attempts=0 WHERE email=?", (email,))
    return True, "E-mail vérifié."
