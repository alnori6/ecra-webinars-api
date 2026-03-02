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

URL_RE = re.compile(r"(https?://[^\s\"<>]+)")
MOJIBAKE_RUN = re.compile(r"[ÃÂâÙØ]{2,}[\x00-\xFF]{2,}")


@app.get("/")
def root() -> Dict[str, Any]:
    return {"ok": True, "message": "API is running. Use /webinars or /debug/events"}


def get_app_token() -> str:
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError("Missing TENANT_ID/CLIENT_ID/CLIENT_SECRET in .env")

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
    if dt_str.endswith("Z"):
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))

    if "." in dt_str:
        head, tail = dt_str.split(".", 1)
        digits = "".join(ch for ch in tail if ch.isdigit())[:6]
        dt_str = f"{head}.{digits}" if digits else head

    # If tz is missing, treat as UTC
    return datetime.fromisoformat(dt_str).replace(tzinfo=timezone.utc)


def fetch_all_pages(
    url: str,
    headers: Dict[str, str],
    params: Optional[Dict[str, str]] = None,
) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    next_url: Optional[str] = url
    next_params = params

    while next_url:
        if next_params is not None:
            r = requests.get(next_url, headers=headers, params=next_params, timeout=25)
            next_params = None
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
    parts = [p.strip() for p in (text or "").split("|")]
    if len(parts) >= 2:
        return {"title_en": parts[0], "title_ar": parts[1]}
    return {"title_en": text or "", "title_ar": ""}


def fix_mojibake(text: str) -> str:
    if not text:
        return text

    # If it already contains real Arabic, keep as-is
    if any("\u0600" <= ch <= "\u06FF" for ch in text):
        return text

    markers = ("Ù", "Ø", "Ã", "â", "Â")
    if not any(m in text for m in markers):
        return text

    # 1) Try repairing the whole string (best case)
    for enc in ("latin1", "cp1252"):
        try:
            fixed = text.encode(enc).decode("utf-8")
            if any("\u0600" <= ch <= "\u06FF" for ch in fixed):
                return fixed
        except Exception:
            pass

    # 2) Repair only the mojibake parts (robust case)
    def _repair_match(m: re.Match) -> str:
        chunk = m.group(0)
        for enc in ("latin1", "cp1252"):
            try:
                repaired = chunk.encode(enc).decode("utf-8")
                if any("\u0600" <= ch <= "\u06FF" for ch in repaired):
                    return repaired
                return repaired
            except Exception:
                continue
        # last resort: try ignoring invalid bytes
        try:
            return chunk.encode("latin1", errors="ignore").decode("utf-8", errors="ignore")
        except Exception:
            return chunk

    fixed2 = MOJIBAKE_RUN.sub(_repair_match, text)

    # Return repaired if it improved (Arabic appeared or markers reduced)
    if any("\u0600" <= ch <= "\u06FF" for ch in fixed2):
        return fixed2
    if ("Ù" not in fixed2 and "Ø" not in fixed2) or fixed2 != text:
        return fixed2

    return text


@app.get("/debug/events")
def debug_events(
    user_principal_name: str = Query(...),
    days_ahead: int = Query(365, ge=1, le=1460),
    limit: int = Query(30, ge=1, le=200),
) -> Dict[str, Any]:
    token = get_app_token()
    headers = {"Authorization": f"Bearer {token}"}

    start = datetime.now(timezone.utc)
    end = start + timedelta(days=days_ahead)

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
    days_ahead: int = Query(365, ge=1, le=1460),
) -> Dict[str, Any]:
    token = get_app_token()
    headers = {"Authorization": f"Bearer {token}"}

    now = datetime.now(timezone.utc)
    end_window = now + timedelta(days=days_ahead)

    base = f"{GRAPH}/users/{user_principal_name}/calendarView"
    params = {
        "startDateTime": now.isoformat(),
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

        titles = split_title(subject)
        title_en = fix_mojibake(titles["title_en"])
        title_ar = fix_mojibake(titles["title_ar"])
        if not subject:
            continue

        if needle and needle not in subject.lower():
            continue

        start_str = (e.get("start") or {}).get("dateTime")
        end_str = (e.get("end") or {}).get("dateTime")
        if not start_str:
            continue

        start_dt = parse_graph_datetime(start_str)
        end_dt = parse_graph_datetime(end_str) if end_str else None

        is_cancelled = bool(e.get("isCancelled", False))

        if is_cancelled:
            status = "cancelled"
        else:
            if end_dt and end_dt <= now:
                status = "past"
            elif start_dt > now:
                status = "upcoming"
            else:
                status = "upcoming"

        if upcoming_only and status != "upcoming":
            continue

        join_url = (e.get("onlineMeeting") or {}).get("joinUrl")
        body_html = (e.get("body") or {}).get("content") or ""
        body_url = extract_first_url(body_html)

        # Prefer join/registration URL. Fallback to calendar item link.
        register_url = join_url or body_url or e.get("webLink")

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

    upcoming = sorted((x for x in items if x["status"] == "upcoming"), key=lambda x: x["startUTC"])
    past = sorted((x for x in items if x["status"] == "past"), key=lambda x: x["startUTC"], reverse=True)
    cancelled = sorted((x for x in items if x["status"] == "cancelled"), key=lambda x: x["startUTC"], reverse=True)

    ordered = upcoming + past + cancelled
    return {"count": len(ordered), "items": ordered}