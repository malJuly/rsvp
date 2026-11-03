const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// Values are written as Malaysia wall-clock time, so read them without any time zone conversion.
export function parseWhen(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/.exec(value || '');
  if (!m) return null;
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  return { y, mo, d, h: m[4] ? Number(m[4]) : null, mi: m[5] ? Number(m[5]) : 0, dow: new Date(Date.UTC(y, mo - 1, d)).getUTCDay() };
}

/** 1 -> "1st", 2 -> "2nd", 3 -> "3rd", 11 to 13 -> "th", 21 -> "21st". */
export function ordinal(n) {
  if (n % 100 >= 11 && n % 100 <= 13) return `${n}th`;
  return `${n}${['th', 'st', 'nd', 'rd'][n % 10] || 'th'}`;
}

export const longDate = (p) => `${DAYS[p.dow]}, ${ordinal(p.d)} ${MONTHS[p.mo - 1]} ${p.y}`;
export const shortDate = (p) => `${ordinal(p.d)} ${MONTHS[p.mo - 1]} ${p.y}`;
export const timeOf = (p) => `${((p.h + 11) % 12) + 1}:${String(p.mi).padStart(2, '0')} ${p.h < 12 ? 'AM' : 'PM'}`;

/** The moment to count down to, in Malaysia time. A date with no time counts to the start of that day. */
export function targetOf(p) {
  const pad = (n) => String(n).padStart(2, '0');
  const time = p.h === null ? '00:00' : `${pad(p.h)}:${pad(p.mi)}`;
  return new Date(`${p.y}-${pad(p.mo)}-${pad(p.d)}T${time}:00+08:00`).getTime();
}

export const placeOf = (ev) => [ev.venue_name, ev.venue_address].filter(Boolean).join(', ');

export function mapHref(ev) {
  if (ev.map_url && /^https?:\/\//i.test(ev.map_url)) return ev.map_url;
  const place = placeOf(ev);
  return place ? `https://www.google.com/maps/dir/?api=1&travelmode=driving&destination=${encodeURIComponent(place)}` : '';
}

export const wazeHref = (ev) => `https://www.waze.com/ul?q=${encodeURIComponent(placeOf(ev))}&navigate=yes`;
