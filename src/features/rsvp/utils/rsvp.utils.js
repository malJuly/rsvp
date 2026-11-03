/* Small pure helpers. No React, no network, no state. */
import { parseWhen, shortDate } from '../../../lib/dates.js';
import { ATTENDING, ID_PARAM } from '../constants/rsvp.constants.js';
import { STRINGS } from '../strings/rsvp.strings.js';

/** The personal guest ID from a query string such as "?id=ABCD2345", or null for a shared link. */
export function readIdFromSearch(search) {
  const raw = new URLSearchParams(search || '').get(ID_PARAM);
  const id = raw ? raw.trim() : '';
  return id || null;
}

/** Fills {key} placeholders. Unknown keys are left as they are. */
export function fillTemplate(template, vars) {
  return template.replace(/\{(\w+)\}/g, (match, key) => (vars[key] !== undefined ? String(vars[key]) : match));
}

/** "1 person" / "3 people". */
export function peopleText(n) {
  return n === 1 ? STRINGS.personOne : fillTemplate(STRINGS.personMany, { n });
}

/** "2026-11-06" -> "6th November 2026"; anything else -> "". */
export function formatDeadline(ymd) {
  const when = parseWhen(ymd);
  return when ? shortDate(when) : '';
}

/** The lead sentence under the title. hasDetails: whether the event date or venue is known. */
export function leadText(hasDetails, deadline) {
  const lead = hasDetails ? STRINGS.leadWithDetails : STRINGS.leadNoDetails;
  const date = formatDeadline(deadline);
  return date ? `${lead} ${fillTemplate(STRINGS.leadDeadline, { date })}` : lead;
}

/** Thank-you sentence after sending. name: the Sheet's name (personal) or the typed one (shared). */
export function thanksMessage(replied, name) {
  const template = replied.attending === ATTENDING.YES ? STRINGS.thanksYes : STRINGS.thanksNo;
  return fillTemplate(template, { name, people: peopleText(replied.pax) });
}

/** "You replied: Yes, 2 people" / "You replied: Unable to attend". */
export function repliedMessage(replied) {
  return replied.attending === ATTENDING.YES ? fillTemplate(STRINGS.repliedYes, { people: peopleText(replied.pax) }) : STRINGS.repliedNo;
}

export function countDigits(text) {
  return String(text).replace(/\D/g, '').length;
}

export function clamp(n, min, max) {
  return Math.min(Math.max(n, min), max);
}

/** Only plain https:// links may become a clickable href (the value comes from the Sheet). */
export function isSafeHttpsUrl(value) {
  return typeof value === 'string' && /^https:\/\/[^\s"'<>\\]+$/i.test(value);
}

/** Today's date as "YYYY-MM-DD" in Malaysia time (used by the mock, matching the script's deadline rule). */
export function todayInMalaysia(now = new Date()) {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kuala_Lumpur', year: 'numeric', month: '2-digit', day: '2-digit' }).format(now);
}
