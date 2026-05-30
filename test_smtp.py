#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Test SMTP settings from settings.txt without fiscalizing a real receipt.

Usage:
    python test_smtp.py --to client@example.ru
    python test_smtp.py --settings settings.txt --to client@example.ru

SMTP password is stored as SMTP_PASSWORD_ENC and encrypted with a key derived
from DREAMKAS_TOKEN. This is portable together with settings.txt, but if someone
gets the whole settings.txt, they get both the token and the decryption key.
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import getpass
import hashlib
import hmac
import os
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
from typing import Dict

SETTINGS_ORDER = [
    "DREAMKAS_TOKEN",
    "SHOP_ID", "SHOP_NAME", "DEVICE_ID", "DEVICE_NAME",
    "API_BASE_URL", "OPERATION_TYPE", "TIMEOUT_MINUTES", "POLL_INTERVAL_SECONDS", "REQUEST_TIMEOUT_SECONDS",
    "DUPLICATE_WINDOW_MINUTES", "DUPLICATE_LOOKBACK_DAYS", "CLEAR_EXCEL_AFTER_SUCCESS", "PDF_ENABLED",
    "EMAIL_ENABLED_AFTER_SUCCESS",
    "SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASSWORD_ENC", "SMTP_FROM",
    "SMTP_SSL", "SMTP_STARTTLS", "SMTP_VERIFY_CERT",
]


def load_settings(path: Path) -> Dict[str, str]:
    settings: Dict[str, str] = {}
    if not path.exists():
        raise FileNotFoundError(f"Settings file not found: {path}")

    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or line.startswith(";"):
            continue
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        settings[key.strip().upper()] = value.strip().strip('"').strip("'")
    return settings


def save_settings(path: Path, settings: Dict[str, str]) -> None:
    lines = [
        "# Settings for Dreamkas Receipt Tool",
        "# SMTP_PASSWORD is not stored in plain text.",
        "# SMTP_PASSWORD_ENC is encrypted with a key derived from DREAMKAS_TOKEN.",
        "",
    ]
    for key in SETTINGS_ORDER:
        if key == "SMTP_PASSWORD" or key not in settings and key == "SMTP_PASSWORD_ENC":
            pass
        lines.append(f"{key} = {settings.get(key, '')}")
    extra = sorted(k for k in settings if k not in SETTINGS_ORDER and k != "SMTP_PASSWORD")
    if extra:
        lines.append("")
        lines.append("# Extra settings")
        for key in extra:
            lines.append(f"{key} = {settings[key]}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def bool_setting(settings: Dict[str, str], key: str, default: bool = False) -> bool:
    value = settings.get(key, "").strip().lower()
    if value == "":
        return default
    return value in {"1", "true", "yes", "да", "y", "on", "вкл"}


def require(settings: Dict[str, str], key: str) -> str:
    value = settings.get(key, "").strip()
    if not value:
        raise ValueError(f"Missing required setting: {key}")
    return value


def smtp_password_kdf(token: str, salt: bytes) -> bytes:
    if not token:
        raise ValueError("DREAMKAS_TOKEN is required to decrypt SMTP_PASSWORD_ENC")
    return hashlib.pbkdf2_hmac("sha256", token.encode("utf-8"), salt, 200_000, dklen=32)


def xor_bytes(data: bytes, stream: bytes) -> bytes:
    return bytes(a ^ b for a, b in zip(data, stream))


def hmac_stream(key: bytes, nonce: bytes, length: int) -> bytes:
    out = bytearray()
    counter = 0
    while len(out) < length:
        out.extend(hmac.new(key, nonce + counter.to_bytes(4, "big"), hashlib.sha256).digest())
        counter += 1
    return bytes(out[:length])


def encrypt_smtp_password(plain_password: str, dreamkas_token: str) -> str:
    salt = os.urandom(16)
    nonce = os.urandom(16)
    key = smtp_password_kdf(dreamkas_token, salt)
    plain = plain_password.encode("utf-8")
    cipher = xor_bytes(plain, hmac_stream(key, nonce, len(plain)))
    payload_without_tag = b"SMTPPW1" + salt + nonce + cipher
    tag = hmac.new(key, payload_without_tag, hashlib.sha256).digest()
    payload = salt + nonce + tag + cipher
    return "enc-v1:" + base64.urlsafe_b64encode(payload).decode("ascii")


def decrypt_smtp_password(encrypted_value: str, dreamkas_token: str) -> str:
    value = (encrypted_value or "").strip()
    if not value:
        return ""
    if not value.startswith("enc-v1:"):
        return value
    raw = base64.urlsafe_b64decode(value.split(":", 1)[1].encode("ascii"))
    if len(raw) < 16 + 16 + 32:
        raise ValueError("Invalid SMTP_PASSWORD_ENC")
    salt = raw[:16]
    nonce = raw[16:32]
    tag = raw[32:64]
    cipher = raw[64:]
    key = smtp_password_kdf(dreamkas_token, salt)
    payload_without_tag = b"SMTPPW1" + salt + nonce + cipher
    expected = hmac.new(key, payload_without_tag, hashlib.sha256).digest()
    if not hmac.compare_digest(tag, expected):
        raise ValueError("Cannot decrypt SMTP_PASSWORD_ENC: token changed or value corrupted")
    return xor_bytes(cipher, hmac_stream(key, nonce, len(cipher))).decode("utf-8")


def get_smtp_password(settings_path: Path, settings: Dict[str, str]) -> str:
    token = require(settings, "DREAMKAS_TOKEN")
    encrypted = settings.get("SMTP_PASSWORD_ENC", "").strip()
    if encrypted:
        return decrypt_smtp_password(encrypted, token)

    plain_legacy = settings.get("SMTP_PASSWORD", "").strip()
    if plain_legacy:
        settings["SMTP_PASSWORD_ENC"] = encrypt_smtp_password(plain_legacy, token)
        settings.pop("SMTP_PASSWORD", None)
        save_settings(settings_path, settings)
        print("SMTP_PASSWORD migrated to encrypted SMTP_PASSWORD_ENC.")
        return plain_legacy

    password = getpass.getpass("SMTP_PASSWORD: ").strip()
    if not password:
        raise ValueError("SMTP password not entered")
    settings["SMTP_PASSWORD_ENC"] = encrypt_smtp_password(password, token)
    settings.pop("SMTP_PASSWORD", None)
    save_settings(settings_path, settings)
    print("SMTP password saved encrypted to settings.txt as SMTP_PASSWORD_ENC.")
    return password


def smtp_ssl_context(settings: Dict[str, str]) -> ssl.SSLContext:
    if bool_setting(settings, "SMTP_VERIFY_CERT", True):
        return ssl.create_default_context()
    print("WARNING: SMTP_VERIFY_CERT = 0, SMTP certificate verification is disabled.")
    return ssl._create_unverified_context()


def send_test_email(settings_path: Path, settings: Dict[str, str], recipient: str) -> None:
    host = require(settings, "SMTP_HOST")
    port = int(require(settings, "SMTP_PORT"))
    user = require(settings, "SMTP_USER")
    password = get_smtp_password(settings_path, settings)
    sender = settings.get("SMTP_FROM", "").strip() or user
    use_ssl = bool_setting(settings, "SMTP_SSL", False)
    use_starttls = bool_setting(settings, "SMTP_STARTTLS", False)

    if use_ssl and use_starttls:
        raise ValueError("SMTP_SSL and SMTP_STARTTLS cannot both be enabled")

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = "Dreamkas SMTP test"
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg.set_content(
        "SMTP test from Dreamkas Receipt Tool.\n\n"
        f"Time: {now}\n"
        f"SMTP_HOST: {host}\n"
        f"SMTP_PORT: {port}\n"
        f"SMTP_SSL: {int(use_ssl)}\n"
        f"SMTP_STARTTLS: {int(use_starttls)}\n\n"
        "If you received this email, SMTP settings are working.\n"
    )

    print("Connecting to SMTP server...")
    print(f"Host: {host}")
    print(f"Port: {port}")
    print(f"SSL: {int(use_ssl)}")
    print(f"STARTTLS: {int(use_starttls)}")
    print(f"VERIFY_CERT: {int(bool_setting(settings, 'SMTP_VERIFY_CERT', True))}")
    print(f"User: {user}")
    print(f"From: {sender}")
    print(f"To: {recipient}")
    print()

    context = smtp_ssl_context(settings)

    if use_ssl:
        with smtplib.SMTP_SSL(host, port, timeout=30, context=context) as server:
            server.login(user, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.ehlo()
            if use_starttls:
                server.starttls(context=context)
                server.ehlo()
            server.login(user, password)
            server.send_message(msg)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--settings", default="settings.txt", help="Path to settings.txt")
    parser.add_argument("--to", required=True, help="Recipient email address for test")
    args = parser.parse_args()

    try:
        settings_path = Path(args.settings)
        settings = load_settings(settings_path)
        send_test_email(settings_path, settings, args.to)
    except Exception as exc:
        print()
        print("SMTP TEST FAILED")
        print("Error:")
        print(exc)
        print()
        print("Check SMTP_SETUP.txt for examples and troubleshooting.")
        input("Press Enter to exit...")
        return 1

    print()
    print("SMTP TEST OK")
    print("Test email sent successfully.")
    input("Press Enter to exit...")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
