#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MailerLite: group, subscriber, HTML campaign (gosafe_vc_outreach_email.html), instant schedule.

Env: MAILES_API_KEY or MAILERLITE_API_KEY; optional CALENDLY_URL, PUBLIC_BASE_URL (logo in email),
     MAILERLITE_GROUP_NAME,
     MAILERLITE_FROM_EMAIL, MAILERLITE_FROM_NAME.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

BASE = "https://connect.mailerlite.com/api"
ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "gosafe_vc_outreach_email.html"
DEFAULT_SUBJECT = "GoSafe AI - healthtech pre-anestesico (Startco / 15 min si hay fit)"


def load_dotenv(path: Path) -> None:
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        k, v = k.strip(), v.strip().strip('"').strip("'")
        if k and k not in os.environ:
            os.environ[k] = v


def api_request(
    method: str,
    path: str,
    token: str,
    body: dict | None = None,
) -> tuple[int, dict]:
    url = f"{BASE}{path}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(
        url,
        data=data,
        method=method,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            raw = resp.read().decode("utf-8")
            return resp.status, json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        try:
            parsed = json.loads(err) if err else {}
        except json.JSONDecodeError:
            parsed = {"raw": err}
        return e.code, parsed


def get_token() -> str:
    t = os.environ.get("MAILERLITE_API_KEY") or os.environ.get("MAILES_API_KEY")
    if not t:
        sys.exit("Missing MAILES_API_KEY or MAILERLITE_API_KEY.")
    return t.strip()


def get_account_sender(token: str) -> tuple[str, str]:
    code, data = api_request("GET", "/account", token)
    if code != 200:
        sys.exit(f"GET /account failed: {code} {data}")
    d = data.get("data", {})
    return str(d.get("sender_email") or ""), str(d.get("sender_name") or "Go Safe AI")


def find_group_id(token: str, name: str) -> str | None:
    code, data = api_request("GET", "/groups", token)
    if code != 200:
        sys.exit(f"GET /groups failed: {code} {data}")
    for g in data.get("data", []):
        if g.get("name") == name:
            return str(g["id"])
    return None


def create_group(token: str, name: str) -> str:
    code, data = api_request("POST", "/groups", token, {"name": name})
    if code not in (200, 201):
        sys.exit(f"POST /groups failed: {code} {data}")
    return str(data["data"]["id"])


def ensure_group(token: str, name: str) -> str:
    gid = find_group_id(token, name)
    return gid or create_group(token, name)


def upsert_subscriber(token: str, email: str, name: str, group_id: str) -> None:
    payload = {
        "email": email,
        "fields": {"name": name},
        "groups": [group_id],
        "resubscribe": True,
    }
    code, data = api_request("POST", "/subscribers", token, payload)
    if code not in (200, 201):
        sys.exit(f"POST /subscribers failed: {code} {data}")


def build_html(calendly_url: str) -> str:
    import base64

    raw = TEMPLATE.read_text(encoding="utf-8")
    raw = raw.replace("__CALENDLY_URL__", calendly_url.rstrip("/") + "/")
    pub = os.environ.get("PUBLIC_BASE_URL", "").strip().rstrip("/")
    if pub:
        logo_src = f"{pub}/go-safe-logo.png"
    else:
        logo_path = ROOT / "go-safe-logo.png"
        if not logo_path.is_file():
            logo_path = ROOT / "Go Safe Logo.png"
        if logo_path.is_file():
            logo_src = "data:image/png;base64," + base64.b64encode(logo_path.read_bytes()).decode(
                "ascii"
            )
        else:
            logo_src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
    return raw.replace("__LOGO_SRC__", logo_src)


def create_campaign(
    token: str,
    group_id: str,
    html: str,
    from_email: str,
    from_name: str,
    subject: str,
) -> str:
    body = {
        "name": f"GoSafe VC outreach - {subject[:40]}",
        "type": "regular",
        "language_id": 9,
        "groups": [group_id],
        "emails": [
            {
                "subject": subject,
                "from_name": from_name,
                "from": from_email,
                "reply_to": from_email,
                "content": html,
            }
        ],
    }
    code, data = api_request("POST", "/campaigns", token, body)
    if code not in (200, 201):
        sys.exit(f"POST /campaigns failed: {code} {data}")
    return str(data["data"]["id"])


def schedule_instant(token: str, campaign_id: str) -> None:
    code, data = api_request(
        "POST", f"/campaigns/{campaign_id}/schedule", token, {"delivery": "instant"}
    )
    if code != 200:
        sys.exit(f"POST /schedule failed: {code} {data}")


def main() -> None:
    parser = argparse.ArgumentParser(description="GoSafe VC outreach via MailerLite API")
    parser.add_argument("--email", required=True, help="Recipient email")
    parser.add_argument("--name", default="Contacto", help="Subscriber Name field in MailerLite")
    parser.add_argument("--subject", default=DEFAULT_SUBJECT)
    parser.add_argument("--dry-run", action="store_true", help="Print HTML only")
    args = parser.parse_args()

    load_dotenv(ROOT / ".env")
    token = get_token()
    calendly = os.environ.get("CALENDLY_URL", "https://calendly.com/").strip()
    group_name = os.environ.get(
        "MAILERLITE_GROUP_NAME", "GoSafe VC Outreach (Startco 2026)"
    ).strip()

    html = build_html(calendly)
    if args.dry_run:
        print(html.replace("{$name}", args.name))
        return

    from_email = os.environ.get("MAILERLITE_FROM_EMAIL", "").strip()
    from_name = os.environ.get("MAILERLITE_FROM_NAME", "").strip()
    if not from_email:
        from_email, from_name = get_account_sender(token)
    if not from_name:
        _, from_name = get_account_sender(token)

    gid = ensure_group(token, group_name)
    upsert_subscriber(token, args.email.strip(), args.name.strip(), gid)
    cid = create_campaign(token, gid, html, from_email, from_name, args.subject)
    schedule_instant(token, cid)

    print("Done.")
    print(f"  Group: {group_name} (id {gid})")
    print(f"  Subscriber: {args.email}")
    print(f"  Campaign id: {cid}")
    print("  If status stays 'ready', open MailerLite > Campaigns and confirm send / domain auth.")
    print("  https://dashboard.mailerlite.com/")


if __name__ == "__main__":
    main()
