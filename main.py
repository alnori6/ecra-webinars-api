import os
import re
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Query

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

GRAPH = "https://graph.microsoft.com/v1.0"

app = FastAPI()

# Regex to grab the first URL from HTML body if needed
URL_RE = re.compile(r"(https?://[^\s\"<>]+)")
# Used to detect long runs of common mojibake markers
MOJIBAKE_RUN = re.compile(r"[ÃÂâÙØ]{2,}[\x00-\xFF]{2,}")

# Bump this string anytime you redeploy so you can verify Render updated
APP_VERSION = "2026-03-02-1"


@app.get("/")
def root() -> Dict[str, Any]:
    return {"ok": True, "message": "API is running. Use /webinars or /debug/events", "version": APP_VERSION}


@app.get("/version")
def version() -> Dict[str, Any]:
    # Simple endpoint so you can confirm Render is running the latest code
    return {"version": APP_VERSION}


def get_app_token() -> str:
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError("Missing TENANT_ID/CLIENT_ID/CLIENT_SECRET in environment variables")

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    r = requests.post(token_url, data=data, timeout=20)
    if r.status_code != 200:
        raise HTTPException(status_code=502, detail=f"Token error: {r.text}")

    return r.json()["access_token"]


def parse_graph_datetime(dt_str: str) -> datetime:
    """
    Graph often returns: "2026-02-16T14:00:00.0000000"
    - If it ends with Z, parse as UTC
    - If it has >6 microseconds digits, trim to 6
    - If timezone is missing, treat it as UTC
    """
    if dt_str.endswith("Z"):
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))

    if "." in dt_str:
        head, tail = dt_str.split(".", 1)
        digits = "".join(ch for ch in tail if ch.isdigit())[:6]
        dt_str = f"{head}.{digits}" if digits else head

    return datetime.fromisoformat(dt_str).replace(tzinfo=timezone.utc)


def fetch_all_pages(
    url: str,
    headers: Dict[str, str],
    params: Optional[Dict[str, str]] = None,
) -> List[Dict[str, Any]]:
    """
    Follow @odata.nextLink pagination automatically.
    """
    items: List[Dict[str, Any]] = []
    next_url: Optional[str] = url
    next_params = params

    while next_url:
        if next_params is not None:
            r = requests.get(next_url, headers=headers, params=next_params, timeout=25)
            next_params = None  # only for first request
        else:
            r = requests.get(next_url, headers=headers, timeout=25)

        if r.status_code != 200:
            raise HTTPException(status_code=502, detail=f"Graph error: {r.text}")

        body = r.json()
        items.extend(body.get("value", []))
        next_url = body.get("@odata.nextLink")

    return items


def extract_first_url(html: str) -> Optional[str]:
    if not html:
        return None
    m = URL_RE.search(html)
    return m.group(1) if m else None


def split_title(text: str) -> Dict[str, str]:
    """
    Your subject looks like:
    English title | Arabic title
    """
    parts = [p.strip() for p in (text or "").split("|")]
    if len(parts) >= 2:
        return {"title_en": parts[0], "title_ar": parts[1]}
    return {"title_en": text or "", "title_ar": ""}


def fix_mojibake(text: str) -> str:
    """
    Attempt to repair Arabic text that appears as mojibake (Ù… Ø… etc)
    without breaking already-correct Arabic.
    """
    if not text:
        return text

    # If it already contains real Arabic characters, keep it
    if any("\u0600" <= ch <= "\u06FF" for ch in text):
        return text

    markers = ("Ù", "Ø", "Ã", "â", "Â")
    if not any(m in text for m in markers):
        return text

    # 1) Try fixing the whole string
    for enc in ("latin1", "cp1252"):
        try:
            fixed = text.encode(enc).decode("utf-8")
            # If Arabic appears, it's likely repaired
            if any("\u0600" <= ch <= "\u06FF" for ch in fixed):
                return fixed
            # Even if Arabic doesn't appear, it might fix punctuation like “–”
            if fixed != text:
                return fixed
        except Exception:
            pass

    # 2) Try repairing only mojibake runs
    def _repair_match(m: re.Match) -> str:
        chunk = m.group(0)
        for enc in ("latin1", "cp1252"):
            try:
                return chunk.encode(enc).decode("utf-8")
            except Exception:
                continue
        try:
            return chunk.encode("latin1", errors="ignore").decode("utf-8", errors="ignore")
        except Exception:
            return chunk

    fixed2 = MOJIBAKE_RUN.sub(_repair_match, text)
    return fixed2 if fixed2 else text


@app.get("/debug/events")
def debug_events(
    user_principal_name: str = Query(...),
    days_back: int = Query(30, ge=0, le=1460),
    days_ahead: int = Query(30, ge=1, le=1460),
    limit: int = Query(30, ge=1, le=200),
) -> Dict[str, Any]:
    """
    Debug endpoint to confirm Graph is returning what you expect
    for the time window you choose.
    """
    token = get_app_token()
    headers = {"Authorization": f"Bearer {token}"}

    now = datetime.now(timezone.utc)
    start = now - timedelta(days=days_back)
    end = now + timedelta(days=days_ahead)

    base = f"{GRAPH}/users/{user_principal_name}/calendarView"
    params = {
        "startDateTime": start.isoformat(),
        "endDateTime": end.isoformat(),
        "$select": "id,subject,start,end,isCancelled",
        "$orderby": "start/dateTime",
        "$top": str(limit),
    }

    events = fetch_all_pages(base, headers, params=params)
    sample = events[:limit]

    return {
        "version": APP_VERSION,
        "rangeUTC": {"start": start.isoformat(), "end": end.isoformat()},
        "fetched": len(events),
        "sample": [
            {
                "subject_raw": (e.get("subject") or ""),
                "subject_fixed": fix_mojibake((e.get("subject") or "")),
                "start": (e.get("start") or {}).get("dateTime"),
                "timeZone": (e.get("start") or {}).get("timeZone"),
                "isCancelled": e.get("isCancelled", False),
            }
            for e in sample
        ],
    }


@app.get("/webinars")
def webinars(
    user_principal_name: str = Query(..., description="e.g. noora.amer@arabianchild.org"),
    subject_contains: str = Query(..., description="text to match in subject"),
    upcoming_only: bool = Query(False),
    days_back: int = Query(365, ge=0, le=1460, description="How many days in the past to include"),
    days_ahead: int = Query(365, ge=1, le=1460, description="How many days in the future to include"),
) -> Dict[str, Any]:
    """
    Returns webinars matching subject_contains within a time window that includes:
    - Past: now - days_back
    - Future: now + days_ahead
    """
    token = get_app_token()
    headers = {"Authorization": f"Bearer {token}"}

    now = datetime.now(timezone.utc)

    # ✅ IMPORTANT: include past + future in the Graph query window
    start_window = now - timedelta(days=days_back)
    end_window = now + timedelta(days=days_ahead)

    base = f"{GRAPH}/users/{user_principal_name}/calendarView"
    params = {
        "startDateTime": start_window.isoformat(),
        "endDateTime": end_window.isoformat(),
        "$select": "id,subject,start,end,iCalUId,webLink,onlineMeeting,isCancelled,body",
        "$orderby": "start/dateTime",
        "$top": "200",
    }

    events = fetch_all_pages(base, headers, params=params)

    items: List[Dict[str, Any]] = []
    needle = subject_contains.lower().strip()

    for e in events:
        raw_subject = (e.get("subject") or "").strip()
        subject = fix_mojibake(raw_subject)

        if not subject:
            continue

        # Case-insensitive subject contains
        if needle and needle not in subject.lower():
            continue

        start_str = (e.get("start") or {}).get("dateTime")
        end_str = (e.get("end") or {}).get("dateTime")
        if not start_str:
            continue

        start_dt = parse_graph_datetime(start_str)
        end_dt = parse_graph_datetime(end_str) if end_str else None

        is_cancelled = bool(e.get("isCancelled", False))

        # Determine status
        if is_cancelled:
            status = "cancelled"
        else:
            if end_dt and end_dt <= now:
                status = "past"
            elif start_dt > now:
                status = "upcoming"
            else:
                # event currently in progress -> treat as upcoming (so user sees Join)
                status = "upcoming"

        if upcoming_only and status != "upcoming":
            continue

        # Prefer Teams join URL
        join_url = (e.get("onlineMeeting") or {}).get("joinUrl")

        # Sometimes the join link is in the body
        body_html = (e.get("body") or {}).get("content") or ""
        body_url = extract_first_url(body_html)

        register_url = join_url or body_url or e.get("webLink")

        # Split EN|AR title if possible
        titles = split_title(subject)
        title_en = fix_mojibake(titles["title_en"])
        title_ar = fix_mojibake(titles["title_ar"])

        items.append(
            {
                "id": e.get("id"),
                "iCalUId": e.get("iCalUId"),
                "title": subject,
                "title_en": title_en,
                "title_ar": title_ar,
                "startUTC": start_dt.isoformat(),
                "endUTC": end_dt.isoformat() if end_dt else None,
                "status": status,
                "registerUrl": register_url,
            }
        )

    # Order: upcoming asc, past desc, cancelled desc
    upcoming = sorted((x for x in items if x["status"] == "upcoming"), key=lambda x: x["startUTC"])
    past = sorted((x for x in items if x["status"] == "past"), key=lambda x: x["startUTC"], reverse=True)
    cancelled = sorted((x for x in items if x["status"] == "cancelled"), key=lambda x: x["startUTC"], reverse=True)

    ordered = upcoming + past + cancelled
    return {"version": APP_VERSION, "count": len(ordered), "items": ordered}
