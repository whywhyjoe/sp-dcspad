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

const REGIONAL_SELECT = ['DateFormat', 'DateSeparator', 'TimeSeparator', 'Time24', 'AM', 'PM'];

// RegionalSettings → the format half of toFormValue's `webDate`.
// DateFormat: 0 = month/day/year, 1 = day/month/year, 2 = year/month/day.
export function webDateFormatOf(settings) {
  return {
    order: settings?.DateFormat === 2 ? 'ymd' : settings?.DateFormat === 1 ? 'dmy' : 'mdy',
    sep: settings?.DateSeparator || '/',
    timeSep: settings?.TimeSeparator || ':',
    time24: settings?.Time24 === true,
    am: settings?.AM || 'AM',
    pm: settings?.PM || 'PM',
  };
}

// UTC offset (ms) of the web's time zone at one instant, as the server
// computes it.
export async function utcOffsetAtMs(client, isoInstant) {
  const data = await client.get(`web/RegionalSettings/TimeZone/utcToLocalTime(@d)?@d='${isoInstant}'`);
  const local = data?.value;
  if (!local) {
    throw new SpFileError('Web time zone could not be read; date not written.', { code: 'write' });
  }
  return new Date(`${local}Z`).getTime() - new Date(isoInstant).getTime();
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
// toFormValue needs for that instant ({ ...format, offsetMinutes }).
// RegionalSettings is read once; offsets are cached per day.
export function createWebDateResolver(client) {
  let format = null;
  const offsets = new Map();
  return async function webDateFor(instant) {
    if (!format) {
      format = webDateFormatOf(await client.get('web/RegionalSettings', { select: REGIONAL_SELECT }));
    }
    return { ...format, offsetMinutes: await webLocalOffsetMinutes(offsets, client, instant) };
  };
}
