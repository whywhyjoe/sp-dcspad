// The web's own date conventions, for WRITING DateTime FieldValues.
//
// ValidateUpdateListItem (like AddValidateUpdateItemUsingPath) parses a
// DateTime FieldValue in the WEB's regional settings: the locale's date order
// and separators, read as a wall-clock time in the web's time zone. SharePoint
// Online refuses every ISO 8601 form ('2026-10-06T16:00:00Z', with
// milliseconds, without the Z) with "You must specify a valid date within the
// range of 1/1/1900 and 12/31/8900." Verified live, 2026-10-06, on a
// Pacific-time web with LocaleId 1033: '10/6/2026 9:00 AM' was accepted and
// stored as 16:00 UTC.
//
// So a writer needs two things from the web: its format (RegionalSettings)
// and its UTC offset AT THE INSTANT being written (TimeZone/utcToLocalTime).
// A single flat Bias would be an hour off for half the year.
// field-editor.js toFormValue does the pure formatting; this module does the
// reads it can't. The list-data import (list-data-apply.js) shares the offset
// lookups below.

import { SpFileError } from '../sp-odata.js';

const ORDERS = { 0: 'mdy', 1: 'dmy', 2: 'ymd' };

const notWritten = (what) => new SpFileError(
  `The site’s ${what} could not be read, so the date was not written.`,
  { code: 'write' },
);

// RegionalSettings → the format half of toFormValue's `webDate`.
// DateFormat: 0 = month/day/year, 1 = day/month/year, 2 = year/month/day.
// TimeMarkerPosition: 0 = marker after the time ('9:00 AM'), 1 = before it
// ('午前 9:00'). Every value a date string is built from must be present and
// well-formed — an incomplete answer is refused, never filled in with US
// defaults (a day-first web would otherwise store June 10 for October 6).
// A 12-hour web also needs AM, PM and TimeMarkerPosition (0 or 1); a 24-hour
// web never shows a marker, so it needs none of the three.
export function webDateFormatOf(settings) {
  const s = settings || {};
  const order = ORDERS[s.DateFormat];
  const time24 = s.Time24;
  const ok = order
    && typeof s.DateSeparator === 'string' && s.DateSeparator
    && typeof s.TimeSeparator === 'string' && s.TimeSeparator
    && typeof time24 === 'boolean'
    && (time24 || (typeof s.AM === 'string' && s.AM && typeof s.PM === 'string' && s.PM
      && (s.TimeMarkerPosition === 0 || s.TimeMarkerPosition === 1)));
  if (!ok) throw notWritten('regional date format');
  return {
    order,
    sep: s.DateSeparator,
    timeSep: s.TimeSeparator,
    time24,
    am: s.AM || '',
    pm: s.PM || '',
    markerFirst: s.TimeMarkerPosition === 1,
  };
}

// UTC offset (ms) of the web's time zone at one instant, as the server
// computes it. A reply that doesn't parse to a whole-minute offset within
// ±14h is refused rather than turned into a guess.
export async function utcOffsetAtMs(client, isoInstant) {
  const data = await client.get(`web/RegionalSettings/TimeZone/utcToLocalTime(@d)?@d='${isoInstant}'`);
  const local = data?.value;
  const offsetMs = local
    ? new Date(`${local}Z`).getTime() - new Date(isoInstant).getTime()
    : NaN;
  if (!Number.isFinite(offsetMs) || offsetMs % 60000 !== 0 || Math.abs(offsetMs) > 14 * 3600000) {
    throw notWritten('time zone');
  }
  return offsetMs;
}

// Per-UTC-day offset, cached — mirrors SPUtils' tzOffsetByDay/tzOffsetExact:
// a day whose noon offset disagrees with its neighbours is a DST transition,
// so that one day gets an exact per-instant lookup instead of the cached noon
// value.
export async function webLocalOffsetMinutes(cache, client, utc) {
  const dayOf = (d) => d.toISOString().slice(0, 10);
  const offsetForDay = async (day) => {
    if (!cache.has(day)) cache.set(day, await utcOffsetAtMs(client, `${day}T12:00:00Z`));
    return cache.get(day);
  };
  const day = dayOf(utc);
  const here = await offsetForDay(day);
  const prev = await offsetForDay(dayOf(new Date(utc.getTime() - 86400000)));
  const next = await offsetForDay(dayOf(new Date(utc.getTime() + 86400000)));
  let offsetMs = here;
  if (here !== prev || here !== next) {
    offsetMs = await utcOffsetAtMs(client, utc.toISOString());
  }
  return offsetMs / 60000;
}

// For createFieldEditorForm's `webDateFor`: (instant: Date) → the webDate
// toFormValue needs for that instant ({ ...format, offsetMinutes, ambiguous }).
// RegionalSettings is read once; offsets are cached per day.
//
// RegionalSettings is read WITHOUT $select: naming a property this tenant's
// RegionalSettings doesn't expose would 400 the whole read, and the entity is
// small. A property it lacks then fails webDateFormatOf instead — a refused
// save, never a guessed format.
//
// `ambiguous`: the instant falls in the hour the web's clock repeats when
// daylight saving ends, so its wall-clock string also names a second instant.
// Found by asking whether the offset in force 3h either side, applied to this
// wall time, lands on an instant that really has that offset. A gap (clocks
// going forward) never needs this: an instant always maps to a real wall
// time. The ±3h window covers every daylight-saving shift in use (1h, or 30
// min on Lord Howe); only a one-off rollback longer than 3h (Kwajalein's 23h
// in 1969) could slip past it, which is accepted.
export function createWebDateResolver(client) {
  let formatRead = null;   // one in-flight read, dropped on failure so a retry reads again
  const offsets = new Map();
  const offsetAt = (ms) => webLocalOffsetMinutes(offsets, client, new Date(ms));
  return async function webDateFor(instant) {
    formatRead ??= client.get('web/RegionalSettings').then(webDateFormatOf)
      .catch((err) => { formatRead = null; throw err; });
    const format = await formatRead;
    const t = instant.getTime();
    const offsetMinutes = await offsetAt(t);
    let ambiguous = false;
    for (const around of [t - 3 * 3600000, t + 3 * 3600000]) {
      const other = await offsetAt(around);
      if (other === offsetMinutes) continue;
      const twin = t + (offsetMinutes - other) * 60000;
      if (twin !== t && await offsetAt(twin) === other) ambiguous = true;
    }
    return { ...format, offsetMinutes, ambiguous };
  };
}
