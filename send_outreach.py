#!/usr/bin/env python3
"""Send personalized outreach emails from your own Microsoft 365 mailbox (for example a UW account) through the
Microsoft Graph API, a few at a time, and record each send in the recipient CSV.

Inputs
  email_template.md   the subject ("**Subject:** ..." line) and the body (between the first two "---" lines), with
                      {{Column}} placeholders filled from the CSV columns of the same name, or from --var
  recipients.csv      one row per person: an "Email" column, a "Status" column and whatever columns the template uses

Sign-in (pick one)
  Pasted token (works without registering an app): copy the access token from Graph Explorer, then
      read -rs OUTREACH_TOKEN && export OUTREACH_TOKEN
  Your own app registration (only if your organization lets you register apps): --client-id and --tenant, or set
  OUTREACH_CLIENT_ID / OUTREACH_TENANT. This needs `pip install msal`.

Typical order
  python send_outreach.py --dry-run --limit 3      # preview the next 3 emails (no sign-in, nothing sent)
  python send_outreach.py --check                  # check token, mailbox, template and CSV (nothing sent)
  python send_outreach.py --test-self              # send the next email to yourself, marked [TEST]
  python send_outreach.py --limit 10               # send to the next 10 people (asks to confirm first)

A row is sent only when its Status is empty or "Not contacted", its address looks valid, it is not a duplicate and
every placeholder has a value. After each email its Status becomes "Sent YYYY-MM-DD HH:MM" and the CSV is saved, so
an interrupted run resumes where it stopped. Rows that fail get "Error: ..." and are not retried until you clear it.
See README.md for the full setup.
"""
from __future__ import annotations

import argparse
import base64
import csv
import json
import os
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit("This script needs the 'requests' package:  pip install -r requirements.txt")

HERE = Path(__file__).resolve().parent
GRAPH = "https://graph.microsoft.com/v1.0"
PLACEHOLDER = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
EMAIL_RE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[A-Za-z]{2,}$")
SENDABLE = {"", "not contacted"}


class AuthError(Exception):
    """The token or mailbox can't send: stop the run (retrying won't help)."""


class RecipientError(Exception):
    """This one email was rejected (for example an invalid address): mark the row and go on."""


def die(msg: str) -> None:
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(1)


# --- template --------------------------------------------------------------------------------------------------------
def load_template(path: Path) -> tuple[str, str]:
    if not path.exists():
        die(f"template not found: {path}")
    text = path.read_text(encoding="utf-8")
    m = re.search(r"^\*\*Subject:\*\*\s*(.+)$", text, re.M)
    parts = re.split(r"^---\s*$", text, flags=re.M)
    if not m or len(parts) < 3 or not parts[1].strip():
        die(f"{path.name} needs a '**Subject:** ...' line and a body between two '---' lines (see the example template)")
    return m.group(1).strip(), parts[1].strip("\n")


def render(template: str, values: dict) -> tuple[str, list[str]]:
    """Fill {{Name}} placeholders; return the text and the names that had no value."""
    missing = []

    def sub(m):
        val = str(values.get(m.group(1), "") or "").strip()
        if not val:
            missing.append(m.group(1))
            return m.group(0)
        return val
    return PLACEHOLDER.sub(sub, template), missing


# --- recipient list --------------------------------------------------------------------------------------------------
def load_csv(path: Path) -> tuple[list[str], list[dict]]:
    if not path.exists():
        die(f"recipient list not found: {path}")
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            cols, rows = list(reader.fieldnames or []), list(reader)
    except UnicodeDecodeError:
        die(f"{path.name} is not UTF-8. In Excel use 'Save As → CSV UTF-8', or in Google Sheets 'Download → .csv'.")
    except csv.Error as e:
        die(f"{path.name} could not be read as CSV: {e}")
    cols = [c.strip() if c else c for c in cols]
    rows = [{(k.strip() if k else k): v for k, v in r.items()} for r in rows]
    if "Email" not in cols:
        die(f"{path.name} needs a column named exactly 'Email' (found: {', '.join(c for c in cols if c)})")
    if "Status" not in cols:
        cols.append("Status")
        for r in rows:
            r["Status"] = ""
    return cols, rows


def save_csv(path: Path, cols: list[str], rows: list[dict]) -> None:
    """Write the CSV back; if it is locked (e.g. open in Excel), keep asking rather than lose the Status updates."""
    while True:
        try:
            tmp = path.with_suffix(path.suffix + ".tmp")
            with open(tmp, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
                w.writeheader()
                w.writerows(rows)
            tmp.replace(path)
            return
        except PermissionError:
            input(f"\n{path.name} is locked (open in Excel?). Close it, then press Enter to save the progress... ")


def append_log(path: Path, stamp: str, email: str, subject: str, result: str) -> None:
    log = path.with_name(path.stem + "_sent_log.csv")
    new = not log.exists()
    with open(log, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["time", "email", "subject", "result"])
        w.writerow([stamp, email, subject, result])


# --- sign-in ---------------------------------------------------------------------------------------------------------
def pasted_token() -> tuple[str, float]:
    tok = os.environ.get("OUTREACH_TOKEN", "").strip()
    if not tok:
        die("no token. Copy one from Graph Explorer (Access token tab), then run:\n"
            "    read -rs OUTREACH_TOKEN && export OUTREACH_TOKEN      (macOS/Linux)\n"
            "    $env:OUTREACH_TOKEN = Read-Host -MaskInput            (Windows PowerShell 7)")
    if tok.lower().startswith("bearer "):
        tok = tok[7:].strip()
    try:  # read the scopes and expiry from the token's payload; the token itself is never printed
        part = tok.split(".")[1]
        claims = json.loads(base64.urlsafe_b64decode(part + "=" * (-len(part) % 4)))
    except Exception:
        die(f"OUTREACH_TOKEN doesn't look like an access token ({len(tok)} characters; a real one has 1,500 or more). "
            "If it's short, the paste didn't take: copy it again from Graph Explorer.")
    if "Mail.Send" not in claims.get("scp", "").split():
        die("this token has no Mail.Send permission. In Graph Explorer: profile picture → Consent to permissions → "
            "Mail.Send → Consent, then run any query and copy a fresh token.")
    exp = float(claims.get("exp", 0))
    if exp - time.time() < 120:
        die("this token has expired. Copy a fresh one from Graph Explorer (tokens last about an hour).")
    return tok, exp


def app_token(client_id: str, tenant: str) -> tuple[str, float]:
    try:
        import msal
    except ImportError:
        die("signing in with --client-id needs msal:  pip install msal")
    cache_file = Path.home() / ".config" / "graph_outreach" / "msal_cache.json"
    cache_file.parent.mkdir(parents=True, exist_ok=True)
    cache = msal.SerializableTokenCache()
    if cache_file.exists():
        cache.deserialize(cache_file.read_text())
    app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant}", token_cache=cache)
    scopes = ["Mail.Send", "User.Read"]
    accounts = app.get_accounts()
    result = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            die(f"could not start the sign-in: {flow.get('error_description', flow)}. "
                "Check the client id and that 'Allow public client flows' is on.")
        print(flow["message"], flush=True)
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        die(f"sign-in failed: {result.get('error')}: {result.get('error_description')}")
    if cache.has_state_changed:
        cache_file.write_text(cache.serialize())
        cache_file.chmod(0o600)
    return result["access_token"], time.time() + float(result.get("expires_in", 3600))


# --- Graph calls -----------------------------------------------------------------------------------------------------
def graph(method: str, url: str, tok: str, **kw) -> requests.Response:
    """One Graph request, retried on throttling, server errors and dropped connections."""
    last = None
    for attempt in range(6):
        try:
            r = requests.request(method, GRAPH + url, headers={"Authorization": f"Bearer {tok}"}, timeout=60, **kw)
        except (requests.ConnectionError, requests.Timeout) as e:
            last = e
            wait = min(300, 15 * 2 ** attempt)
            print(f"  network problem ({type(e).__name__}); retrying in {wait} s...", flush=True)
            time.sleep(wait)
            continue
        if r.status_code == 429 or r.status_code >= 500:
            wait = int(r.headers.get("Retry-After", str(min(300, 15 * 2 ** attempt))))
            print(f"  Microsoft returned {r.status_code}; retrying in {wait} s...", flush=True)
            time.sleep(wait)
            continue
        return r
    raise AuthError(f"Microsoft Graph could not be reached after several tries ({last or 'repeated throttling'}). "
                    "Check your internet connection and run the same command again.")


def explain(r: requests.Response) -> str:
    try:
        err = r.json().get("error", {})
        return f"{err.get('code', '')}: {err.get('message', '')}".strip(": ")
    except ValueError:
        return r.text[:200]


def whoami(tok: str) -> dict:
    r = graph("GET", "/me?$select=displayName,mail,userPrincipalName", tok)
    if r.status_code == 401:
        raise AuthError("the token was rejected (401). Copy a fresh one from Graph Explorer.")
    if r.status_code != 200:
        raise AuthError(f"could not read your profile ({r.status_code} {explain(r)}).")
    return r.json()


def send(tok: str, to: str, subject: str, body: str, cc: str = "") -> None:
    msg = {"subject": subject, "body": {"contentType": "Text", "content": body},
           "toRecipients": [{"emailAddress": {"address": to}}]}
    cc_list = [a.strip() for a in re.split(r"[;,]", cc or "") if a.strip()]
    if cc_list:
        msg["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc_list]
    r = graph("POST", "/me/sendMail", tok, json={"message": msg, "saveToSentItems": True})
    if r.status_code == 202:
        return
    if r.status_code == 401:
        raise AuthError("the token expired or was rejected (401). Copy a fresh one and run the same command again; "
                        "rows already sent are marked and will be skipped.")
    if r.status_code == 403:
        raise AuthError(f"your mailbox isn't allowed to send this way (403 {explain(r)}). If the message mentions "
                        "'quota' or 'restricted', your account may be blocked from sending: stop and contact IT.")
    if r.status_code in (400, 404):
        raise RecipientError(f"{r.status_code} {explain(r)}")
    raise AuthError(f"unexpected response {r.status_code}: {explain(r)}")


# --- main ------------------------------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Send personalized emails through Microsoft Graph. See README.md.",
                                 formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    ap.add_argument("--csv", default=str(HERE / "recipients.csv"), help="recipient list")
    ap.add_argument("--template", default=str(HERE / "email_template.md"), help="email template")
    ap.add_argument("--var", action="append", default=[], metavar="NAME=VALUE",
                    help="a placeholder that is the same for everyone, e.g. --var SenderName='Jane Doe' (repeatable)")
    ap.add_argument("--limit", type=int, default=10, help="most emails to send in this run")
    ap.add_argument("--delay", type=float, default=90, help="average seconds between emails (randomized, min 20)")
    ap.add_argument("--only", metavar="ADDRESS", help="send only to this recipient (must be a sendable row in the CSV)")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true", help="print the emails; no sign-in, nothing sent")
    mode.add_argument("--check", action="store_true", help="check the token, mailbox, template and CSV; nothing sent")
    mode.add_argument("--test-self", action="store_true", help="send the next email to your own mailbox, marked [TEST]")
    mode.add_argument("--test-to", metavar="ADDRESS", help="send the next email to ADDRESS, marked [TEST]")
    ap.add_argument("--yes", "-y", action="store_true", help="don't ask for confirmation before a real batch")
    ap.add_argument("--max-errors", type=int, default=3, help="stop after this many rejected recipients in one run")
    ap.add_argument("--client-id", default=os.environ.get("OUTREACH_CLIENT_ID"), help="app registration client id (optional)")
    ap.add_argument("--tenant", default=os.environ.get("OUTREACH_TENANT", "organizations"), help="tenant id, with --client-id")
    args = ap.parse_args()
    if args.limit < 1:
        die("--limit must be at least 1")

    shared = {}
    for v in args.var:
        if "=" not in v:
            die(f"--var needs NAME=VALUE, got {v!r}")
        k, val = v.split("=", 1)
        shared[k.strip()] = val
    subj_t, body_t = load_template(Path(args.template))
    path = Path(args.csv)
    cols, rows = load_csv(path)

    # --- choose the rows ---------------------------------------------------------------------------------------------
    used = set(PLACEHOLDER.findall(subj_t + body_t))
    nowhere = sorted(u for u in used if u not in cols and u not in shared)
    if nowhere:
        print(f"Note: placeholders with no CSV column and no --var: {', '.join(nowhere)}")
    test_mode = args.test_self or args.test_to
    want = 1 if test_mode else args.limit
    seen, picked, skipped = set(), [], []
    for i, r in enumerate(rows):
        email = (r.get("Email") or "").strip()
        status = (r.get("Status") or "").strip()
        if email.lower() in seen and email:
            if status.lower() in SENDABLE:
                skipped.append(f"row {i + 2} ({email}): duplicate of an earlier row")
            continue
        seen.add(email.lower())
        if status.lower() not in SENDABLE:
            continue
        if args.only and email.lower() != args.only.strip().lower():
            continue
        if not EMAIL_RE.match(email):
            skipped.append(f"row {i + 2} ({email or 'no address'}): not a valid email address")
            continue
        values = {**shared, **{k: v for k, v in r.items() if k}}
        subject, m1 = render(subj_t, values)
        body, m2 = render(body_t, values)
        if m1 or m2:
            skipped.append(f"row {i + 2} ({email}): no value for {', '.join(sorted(set(m1 + m2)))}")
            continue
        if len(picked) < want:
            picked.append((i, subject, body))
    for s in skipped:
        print("SKIPPED", s)
    sendable = len(picked)
    if not picked:
        die("nothing to send. Rows are sent only when Status is empty or 'Not contacted', the address is valid and "
            "every placeholder has a value" + (f"; no sendable row has the address {args.only}" if args.only else "") + ".")

    if args.dry_run:
        for i, subject, body in picked:
            print(f"\n--- to {rows[i]['Email'].strip()}  (row {i + 2})\nSubject: {subject}\n\n{body}")
        print(f"\n{len(picked)} email(s) would be sent. Nothing was sent.")
        return

    # --- sign in ------------------------------------------------------------------------------------------------------
    try:
        tok, exp = app_token(args.client_id, args.tenant) if args.client_id else pasted_token()
        me = whoami(tok)
    except AuthError as e:
        die(str(e))
    mine = me.get("mail") or me.get("userPrincipalName")
    print(f"Signed in as {me.get('displayName')} <{mine}>; token valid for about {(exp - time.time()) / 60:.0f} min")

    if args.check:
        print(f"Template OK ({len(used)} placeholders). CSV OK: {len(rows)} rows, "
              f"next {sendable} sendable (checked up to --limit {args.limit}), {len(skipped)} skipped.")
        print("Everything needed to send is in place. Nothing was sent.")
        return

    if test_mode:
        to = mine if args.test_self else args.test_to
        if not to or not EMAIL_RE.match(to):
            die(f"can't send a test to {to!r}")
        i, subject, body = picked[0]
        try:
            send(tok, to, "[TEST] " + subject, f"(test of the email to {rows[i]['Email'].strip()}, row {i + 2})\n\n" + body)
        except (AuthError, RecipientError) as e:
            die(f"test email failed: {e}")
        print(f"Test email sent to {to}. Check your inbox (and junk folder). The CSV was not changed.")
        return

    # --- the batch ----------------------------------------------------------------------------------------------------
    minutes = len(picked) * args.delay / 60
    print(f"\nAbout to send {len(picked)} email(s) from {mine}, about {args.delay:.0f} s apart (~{minutes:.0f} min).")
    if minutes > (exp - time.time()) / 60 - 2:
        print("Warning: the token will probably expire before the batch ends. The run stops cleanly when it does;"
              " copy a fresh token and run the same command to continue.")
    if not args.yes:
        if input("Send them now? [y/N] ").strip().lower() not in ("y", "yes"):
            print("Cancelled. Nothing was sent.")
            return

    sent = errors = 0
    try:
        for n, (i, subject, body) in enumerate(picked, 1):
            r = rows[i]
            email = r["Email"].strip()
            if exp - time.time() < 60:
                print("The token is about to expire: stopping here. Copy a fresh token and run the same command to continue.")
                break
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            try:
                send(tok, email, subject, body, r.get("Cc", ""))
            except RecipientError as e:
                errors += 1
                r["Status"] = f"Error: {e}"[:200]
                save_csv(path, cols, rows)
                append_log(path, stamp, email, subject, f"error: {e}")
                print(f"FAILED {n}/{len(picked)}: {email}: {e} (marked in the CSV, continuing)", flush=True)
                if errors >= args.max_errors:
                    print(f"Stopping: {errors} recipients were rejected in this run. Check the addresses marked 'Error'.")
                    break
                continue
            r["Status"] = f"Sent {stamp}"
            save_csv(path, cols, rows)
            append_log(path, stamp, email, subject, "sent")
            sent += 1
            print(f"sent {n}/{len(picked)}: {email}", flush=True)
            if n < len(picked):
                time.sleep(max(20.0, random.gauss(args.delay, args.delay / 4)))
    except AuthError as e:
        print(f"\nSTOPPED: {e}", file=sys.stderr)
        print(f"{sent} sent before stopping; they are marked in {path.name}.")
        sys.exit(1)
    except KeyboardInterrupt:
        print(f"\nInterrupted. {sent} sent; they are marked in {path.name}. Run the same command to continue.")
        sys.exit(130)
    print(f"Done: {sent} sent, {errors} rejected. Progress is saved in {path.name}.")


if __name__ == "__main__":
    main()
