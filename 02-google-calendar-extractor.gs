/**
 * One calendar event only: prompts for Event ID, loads the event, writes guest
 * people info to a sheet. No date-range scan.
 *
 * Apps Script → Services (+) → enable "Google Calendar API".
 *
 * If the event is on a shared calendar, set CALENDAR_ID to that calendar’s ID
 * (Calendar settings → Integrate calendar), not 'primary'.
 */

const CALENDAR_ID = 'primary';
const OUTPUT_SHEET_NAME = 'Event guests';

function extractSingleEventGuests() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    'Event ID',
    'Paste the Google Calendar event ID for this one meeting.',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const eventId = String(r.getResponseText() || '').trim();
  if (!eventId) {
    ui.alert('No event ID entered.');
    return;
  }

  let event;
  try {
    event = Calendar.Events.get(CALENDAR_ID, eventId);
  } catch (e) {
    ui.alert(
      'Could not load this event.',
      String(e.message || e),
      ui.ButtonSet.OK
    );
    return;
  }

  const attendees = normalizeAttendees_(event.attendees);
  const headers = [
    'First name',
    'Last name',
    'Email',
    'Company (from domain)',
    'RSVP',
    'Organizer',
    'Optional guest'
  ];

  const rows = [headers];
  attendees.forEach(function (att) {
    if (!att.email || att.resource) return;
    const email = att.email.toLowerCase().trim();
    const domain = email.split('@')[1] || '';
    const nameParts = guestNameParts_(att, email);
    rows.push([
      nameParts.firstName,
      nameParts.lastName,
      email,
      guessCompanyName(domain),
      rsvpLabel_(att.responseStatus),
      att.organizer ? 'Yes' : '',
      att.optional === true ? 'Yes' : att.optional === false ? 'No' : ''
    ]);
  });

  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(OUTPUT_SHEET_NAME);

  let outRows = rows.length;
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1).setValue('No guests with an email on this event.');
    outRows = 2;
  }

  formatGuestSheet_(sheet, outRows, headers.length);

  ui.alert(
    'Done',
    rows.length > 1
      ? 'Wrote ' + (rows.length - 1) + ' guest row(s) to "' + OUTPUT_SHEET_NAME + '".'
      : 'No guest rows to write (sheet has headers only).',
    ui.ButtonSet.OK
  );
}

function normalizeAttendees_(attendees) {
  if (!attendees) return [];
  if (Array.isArray(attendees)) return attendees;
  return [attendees];
}

function rsvpLabel_(status) {
  const s = String(status || '').toLowerCase();
  if (s === 'accepted') return 'Accepted';
  if (s === 'declined') return 'Declined';
  if (s === 'tentative') return 'Tentative';
  if (s === 'needsaction') return 'No response';
  return status ? String(status) : '';
}

/**
 * Calendar often sends no displayName — only email. Then we guess from the
 * address (e.g. john.doe@co.com → John / Doe). Not always right for jsmith@…
 */
function guestNameParts_(att, emailLower) {
  const fromCalendar = parseDisplayName(att.displayName);
  if (fromCalendar.firstName || fromCalendar.lastName) {
    return fromCalendar;
  }
  return namePartsFromEmailLocal_(emailLower);
}

/** Part before @: drop +tag, turn . _ - into spaces, title-case, then split. */
function namePartsFromEmailLocal_(emailLower) {
  let local = (emailLower.split('@')[0] || '').trim();
  if (!local) return { firstName: '', lastName: '' };
  local = local.split('+')[0];
  const spaced = local.replace(/[._-]+/g, ' ').replace(/\s+/g, ' ').trim();
  const titled = spaced.split(' ').filter(Boolean).map(titleCaseWord_).join(' ');
  return parseDisplayName(titled);
}

function titleCaseWord_(w) {
  if (!w) return '';
  return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
}

/**
 * Splits one string: first whitespace-separated word → first name, rest → last.
 * Strips "Name <email>" style and common titles.
 */
function parseDisplayName(displayName) {
  if (!displayName || !displayName.trim()) {
    return { firstName: '', lastName: '' };
  }

  const name = displayName.trim();
  const cleanName = name.replace(/<[^>]+>/g, '').trim();
  const withoutTitles = cleanName
    .replace(/^(Dr\.|Mr\.|Mrs\.|Ms\.|Prof\.)\s*/i, '')
    .replace(/\s+(PhD|MD|Jr\.|Sr\.)$/i, '')
    .trim();

  const parts = withoutTitles.split(/\s+/);

  if (parts.length === 0) {
    return { firstName: '', lastName: '' };
  }
  if (parts.length === 1) {
    return { firstName: parts[0], lastName: '' };
  }
  return {
    firstName: parts[0],
    lastName: parts.slice(1).join(' ')
  };
}

function guessCompanyName(domain) {
  if (!domain) return '';

  const genericDomains = [
    'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com',
    'live.com', 'icloud.com', 'me.com', 'msn.com',
    'googlemail.com', 'protonmail.com', 'proton.me'
  ];

  if (genericDomains.includes(domain.toLowerCase())) {
    return '(Personal)';
  }

  const parts = domain.split('.');
  if (parts.length >= 2) {
    const mainPart = parts[parts.length - 2];
    return mainPart.charAt(0).toUpperCase() + mainPart.slice(1);
  }

  return domain;
}

function formatGuestSheet_(sheet, numRows, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  sheet.setFrozenRows(1);
  for (let i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
  }
  if (numRows > 1) {
    sheet.getRange(1, 1, numRows, numCols).createFilter();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar')
    .addItem('Extract guests from one event…', 'extractSingleEventGuests')
    .addToUi();
}
