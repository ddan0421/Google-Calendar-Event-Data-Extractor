"""
Decode Google Calendar eventedit URL segment (or full URL) into:
  - Event ID (e.g. 364133hm7ifsli038ctptr28t6_20260327T183000Z)
  - Calendar ID (often an email)

Usage:
  python decode_gcal_event_link.py
      → prompts you to paste the URL or encoded segment, then Enter.

  python decode_gcal_event_link.py "https://...eventedit/..."
  echo URL | python decode_gcal_event_link.py
"""

from __future__ import annotations

import base64
import re
import sys
import urllib.parse


def extract_encoded_segment(text: str) -> str:
    text = text.strip()
    # Full URL: .../eventedit/SEGMENT or .../eventedit/SEGMENT?...
    m = re.search(r"/eventedit/([^/?#]+)", text, re.I)
    if m:
        return urllib.parse.unquote(m.group(1))
    # Sometimes eid= in query (different encoding; this script targets eventedit path)
    if "eventedit/" in text:
        part = text.split("eventedit/", 1)[1]
        return urllib.parse.unquote(part.split("?")[0].split("#")[0])
    return text.strip()


def decode_segment(encoded: str) -> tuple[str, str]:
    encoded = encoded.strip()
    pad = "=" * ((4 - len(encoded) % 4) % 4)
    raw = base64.urlsafe_b64decode(encoded + pad).decode("utf-8", errors="replace")
    raw = raw.strip()
    if " " not in raw:
        return raw, ""
    event_id, calendar_id = raw.rsplit(" ", 1)
    # Google uses: "<eventId> <calendarId>" — calendar id may contain spaces? rare; rsplit once is usual
    return event_id.strip(), calendar_id.strip()


def main() -> int:
    if len(sys.argv) > 1:
        blob = " ".join(sys.argv[1:])
    elif not sys.stdin.isatty():
        blob = sys.stdin.read()
    else:
        print("Paste your Google Calendar event URL (or the encoded part after eventedit/), then press Enter:")
        try:
            blob = input()
        except EOFError:
            blob = ""

    if not blob.strip():
        print("Nothing to decode.", file=sys.stderr)
        return 1

    segment = extract_encoded_segment(blob)
    try:
        event_id, calendar_id = decode_segment(segment)
    except Exception as e:
        print("Decode failed:", e, file=sys.stderr)
        return 1

    print("Event ID:     ", event_id)
    print("Calendar ID:  ", calendar_id or "(none — check your calendar settings)")
    print()
    print("For Apps Script CALENDAR_ID constant, use the Calendar ID line (or 'primary' if it is your main calendar).")
    print("For the prompt Event ID, use the Event ID line.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
