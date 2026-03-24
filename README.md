# Google Calendar Event Data Extractor

A collection of scripts designed to extract, transform, and analyze event data from Google Calendar. This repository serves as a central home for reusable tools and utilities that support Google-related workflows, automation, and data processing.

## `01-gcal-event-link-decoder.py`

**Purpose**  
Turns a Google Calendar **event edit** URL (or the raw base64 segment after `/eventedit/`) into the **Event ID** and **Calendar ID** strings the Calendar API and the spreadsheet script expect.

**How it works**  
- **`extract_encoded_segment(text)`** — From a full URL, finds the path segment after `eventedit/`, URL-decodes it, and strips query/fragment; if there is no URL pattern, treats the whole pasted string as the encoded segment.  
- **`decode_segment(encoded)`** — Pads and **URL-safe base64-decodes** the segment, decodes UTF-8, then splits on the last space into **`event_id`** and **`calendar_id`** (Google’s usual payload: `"<eventId> <calendarId>"`).

**Usage**

- Interactive: `python 01-gcal-event-link-decoder.py` — paste URL or segment, press Enter.  
- CLI arg: `python 01-gcal-event-link-decoder.py "https://...eventedit/..."`  
- Stdin: `echo "URL" | python 01-gcal-event-link-decoder.py`

**Output**  
Prints **Event ID** and **Calendar ID** (or notes if calendar id is missing), plus short reminders: use **Calendar ID** for the Apps Script `CALENDAR_ID` constant (or `primary` for the main calendar), and **Event ID** for the extractor prompt.

**Dependencies**  
Python 3 standard library only (`base64`, `re`, `sys`, `urllib.parse`).

## `02-google-calendar-extractor.gs`

**Purpose**  
Spreadsheet-bound Apps Script that loads **one** Google Calendar event by ID (via the **Advanced Calendar service** `Calendar.Events.get`), turns **human attendees** into rows, and writes them to a tab named **`Event guests`**. It does **not** scan a date range.

### Installing and using in Google Sheets (Apps Script)

1. **Open the script editor** — In the spreadsheet, go to **Extensions → Apps Script**. A new tab opens with the Apps Script project tied to this file.

2. **Paste the code** — Remove any default `myFunction` boilerplate if present. Copy the full contents of `02-google-calendar-extractor.gs` from this repo and paste into the editor (one file, e.g. `Code.gs`, is enough).

3. **Enable the Calendar advanced service** — In the Apps Script editor, click **Services** (or **+** next to “Services” in the left sidebar). Find **Google Calendar API**, click **Add**. The script calls `Calendar.Events.get`, which requires this service—not the basic CalendarApp alone.

4. **Save** — **Ctrl+S** / **Cmd+S**, or **File → Save**. Name the project if prompted (the name is only for you in script.google.com).

5. **Authorize the first run** — Run **`extractSingleEventGuests`** once from the toolbar (function dropdown → Run), or reload the sheet and use the menu (next step). Google will ask you to **Review permissions**: pick your account and allow. If you see “Google hasn’t verified this app,” open **Advanced** and choose the option to continue, then allow. The script needs access to **this spreadsheet** and **Google Calendar** (to read events your account can open).

6. **Configure `CALENDAR_ID` (optional)** — At the top of the script, set `CALENDAR_ID` to `'primary'` for your main calendar, or paste the **Calendar ID** from Google Calendar (**Settings for my calendars → your calendar → Integrate calendar → Calendar ID**). Use the value from `01-gcal-event-link-decoder.py` when the event lives on a specific calendar. Wrong ID usually means “Could not load this event.”

7. **Reload the spreadsheet** — Close the sheet tab and reopen it, or refresh the page, so **`onOpen`** runs and the **Calendar** menu appears. (If the menu is missing, run any function once from the editor, then refresh again.)

8. **Get the Event ID** — From an event’s browser **event edit** URL, run `01-gcal-event-link-decoder.py` and copy the **Event ID** line. You can also copy the ID from some Calendar API / developer flows; it must match the event on the calendar you set in `CALENDAR_ID`.

9. **Extract guests** — In the sheet: **Calendar → Extract guests from one event…**. Paste the **Event ID** in the dialog and confirm. Results go to the tab named in **`OUTPUT_SHEET_NAME`** (default **`Event guests`**), replacing previous content on that tab.

**Setup (top of file)**  
- **`CALENDAR_ID`** — Which calendar to read (`'primary'` or a shared calendar ID from Calendar settings).  
- **`OUTPUT_SHEET_NAME`** — Sheet tab to create/clear and fill.

### `onOpen()`

When the spreadsheet opens, adds a custom menu **Calendar → Extract guests from one event…**, which runs the main flow.

### `extractSingleEventGuests()`

The workflow: prompt for event ID → fetch event → normalize attendees → skip anyone without email or marked as a **resource** (room) → build rows with headers → write/clear **`Event guests`** → style the sheet → show a completion alert.

### Helpers (data shaping)

| Function | Role |
|----------|------|
| **`normalizeAttendees_(attendees)`** | Ensures attendees are always an array (handles missing or single object). |
| **`rsvpLabel_(status)`** | Maps API response status to readable labels (Accepted, Declined, Tentative, No response, etc.). |
| **`guestNameParts_(att, emailLower)`** | First/last name: use Calendar `displayName` if parseable, else guess from the email local part. |
| **`namePartsFromEmailLocal_(emailLower)`** | Strip `+tag`, replace `.` `_` `-` with spaces, title-case, then split via `parseDisplayName`. |
| **`titleCaseWord_(w)`** | Simple title case for one word. |
| **`parseDisplayName(displayName)`** | Strip `<email>` fragments and common titles/suffixes; first token → first name, rest → last name. |
| **`guessCompanyName(domain)`** | From email domain: `(Personal)` for common webmail, else a rough company label from the domain (e.g. `acme` from `mail.acme.com`). |
| **`formatGuestSheet_(sheet, numRows, numCols)`** | Bold blue header row, white text, freeze row 1, auto-resize columns, add filter when there is more than a header. |

### Dependencies

Same as **step 3** above: add the **Google Calendar API** service in Apps Script so **`Calendar.Events.get`** works (in addition to **`SpreadsheetApp`**).
