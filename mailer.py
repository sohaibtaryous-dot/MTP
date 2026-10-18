# -*- coding: utf-8 -*-
"""Envoi du code de vérification par e-mail (SMTP).

Configuration (variables d'environnement, ou Secrets Streamlit / .streamlit/secrets.toml) :
  SMTP_USER      adresse expéditrice, ex. timar.app@gmail.com
  SMTP_PASSWORD  mot de passe d'application Gmail (16 caractères) — pas le mot de passe du compte
  SMTP_HOST      défaut smtp.gmail.com
  SMTP_PORT      défaut 465 (SSL) ; 587 = STARTTLS
  SMTP_FROM      défaut = SMTP_USER
"""

from __future__ import annotations

import os
import smtplib
import ssl
from email.message import EmailMessage


def _setting(name: str, default: str = "") -> str:
    val = os.environ.get(name)
    if val:
        return val
    try:
        import streamlit as st
        return str(st.secrets.get(name, default) or default)
    except Exception:  # noqa: BLE001  (pas de secrets.toml)
        return default


def is_configured() -> bool:
    return bool(_setting("SMTP_USER") and _setting("SMTP_PASSWORD"))


def send_email(to: str, subject: str, body: str) -> tuple[bool, str]:
    user, password = _setting("SMTP_USER"), _setting("SMTP_PASSWORD")
    if not (user and password):
        return False, "L'envoi d'e-mails n'est pas configuré (SMTP_USER / SMTP_PASSWORD)."
    host, port = _setting("SMTP_HOST", "smtp.gmail.com"), int(_setting("SMTP_PORT", "465") or 465)
    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = _setting("SMTP_FROM", user) or user, to, subject
    msg.set_content(body)
    try:
        if port == 465:
            with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context(), timeout=20) as s:
                s.login(user, password)
                s.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=20) as s:
                s.starttls(context=ssl.create_default_context())
                s.login(user, password)
                s.send_message(msg)
    except smtplib.SMTPAuthenticationError:
        return False, "Connexion SMTP refusée (identifiants ou mot de passe d'application incorrects)."
    except (smtplib.SMTPException, OSError) as exc:
        return False, f"Envoi impossible : {exc}"
    return True, "E-mail envoyé."


def send_verification_code(to: str, code: str) -> tuple[bool, str]:
    body = (f"Bonjour,\n\nVotre code de vérification TIMAR Tanger est : {code}\n\n"
            "Ce code est valable 15 minutes. Si vous n'êtes pas à l'origine de cette demande, ignorez ce message.\n\n"
            "Après vérification, votre compte devra encore être validé par l'administrateur.\n")
    return send_email(to, f"TIMAR Tanger — code de vérification : {code}", body)
