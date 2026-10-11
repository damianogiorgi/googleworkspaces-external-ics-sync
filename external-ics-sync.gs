/*
 * live-to-work-calendar-sync
 * Copyright (C) 2026 Damiano Giorgi
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 *
 * SPDX-License-Identifier: GPL-3.0-or-later
 */

/**
 * Live -> Work busy blocks (direct fetch)
 *
 * Reads your published Outlook.com (Live) calendar straight from its ICS link
 * and copies the events into your primary Google Workspace calendar as private
 * blocks, so colleagues see you as busy. Skips Google's slow "From URL"
 * polling: the only delay left is Microsoft's own feed caching.
 * One-way: the script only creates, updates and deletes its own copies.
 *
 * Setup:
 * 1. Outlook.com > Settings > Calendar > Shared calendars > Publish a calendar.
 *    "Can view titles and locations" is enough. Copy the ICS link.
 * 2. In the Apps Script project, add a second script file named "ical" and
 *    paste the content of ical.gs into it (the iCal parser).
 * 3. Project Settings (gear) > Script Properties > Add property:
 *    LIVE_ICS_URL = <the ICS link>. Keeping it out of the code means it
 *    doesn't leak if you share or paste the script.
 * 4. Services (+) > "Google Calendar API" must be enabled (identifier "Calendar").
 * 5. Run installTrigger() once and authorize (it now also asks to
 *    "connect to an external service": that's the ICS fetch).
 *
 * Once it runs fine you can remove the old "From URL" subscription in Google.
 * To undo everything: run removeAllCopies() and delete the trigger.
 */

const CONFIG = {
  TARGET_CALENDAR_ID: 'primary',
  DAYS_BACK: 1,
  DAYS_AHEAD: 60,
  COPY_TITLE: true,       // copies are private, colleagues only see "busy" anyway
  FALLBACK_TITLE: 'Busy',
  TITLE_PREFIX: '🏠 ',    // visible marker for you; set '' to disable
  COLOR_ID: '10',         // 1 Lavender, 2 Sage, 3 Grape, 4 Flamingo, 5 Banana, 6 Tangerine,
                          // 7 Peacock, 8 Graphite, 9 Blueberry, 10 Basil, 11 Tomato; '' = default
  COPY_LOCATION: true,    // lets iOS compute "Time to Leave" on the copies
  REMINDERS: 'default',   // 'default' = your calendar's default notifications,
                          // [] = none, or e.g. [{ method: 'popup', minutes: 30 }]
  SKIP_ALL_DAY: true,     // birthdays and "pay the bill" shouldn't block a whole day
  SKIP_FREE: true,        // respect events marked as "free" in Live
  SYNC_EVERY_MINUTES: 15, // allowed: 1, 5, 10, 15, 30
};

const TAG = 'liveSync';
const TAG_ID = 'liveSyncId';
const MAX_OCCURRENCES_PER_SERIES = 20000; // safety net for runaway recurrences

function sync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return; // another run is in progress

  try {
    const now = Date.now();
    const winStart = new Date(now - CONFIG.DAYS_BACK * 864e5);
    const winEnd = new Date(now + CONFIG.DAYS_AHEAD * 864e5);

    // What should exist: one copy per source occurrence.
    // If the fetch or parse fails this throws before touching anything,
    // so a Microsoft hiccup never wipes your copies.
    const wanted = new Map();
    expandIcs(fetchIcs(), winStart, winEnd)
      .filter(e => !(CONFIG.SKIP_ALL_DAY && e.allDay))
      .filter(e => !(CONFIG.SKIP_FREE && e.free))
      .forEach(e => wanted.set(e.key, e));

    // What exists: only events this script created
    const existing = new Map();
    listAll(CONFIG.TARGET_CALENDAR_ID, {
      timeMin: winStart.toISOString(),
      timeMax: winEnd.toISOString(),
      singleEvents: true,
      privateExtendedProperty: `${TAG}=1`,
    }).forEach(e => {
      const key = e.extendedProperties.private[TAG_ID];
      if (existing.has(key)) {
        Calendar.Events.remove(CONFIG.TARGET_CALENDAR_ID, e.id); // stray duplicate
      } else {
        existing.set(key, e);
      }
    });

    // Create or update
    for (const [key, src] of wanted) {
      const body = buildCopy(key, src);
      const copy = existing.get(key);
      if (!copy) {
        Calendar.Events.insert(body, CONFIG.TARGET_CALENDAR_ID);
      } else if (changed(copy, body)) {
        Calendar.Events.patch(body, CONFIG.TARGET_CALENDAR_ID, copy.id);
      }
      existing.delete(key);
    }

    // Whatever is left has no source anymore (deleted or moved in Live)
    for (const copy of existing.values()) {
      Calendar.Events.remove(CONFIG.TARGET_CALENDAR_ID, copy.id);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ICS fetch and expansion ---------- */

function fetchIcs() {
  const url = PropertiesService.getScriptProperties().getProperty('LIVE_ICS_URL');
  if (!url) throw new Error('Set the LIVE_ICS_URL script property first.');
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  const text = res.getContentText();
  if (res.getResponseCode() !== 200 || text.indexOf('BEGIN:VCALENDAR') === -1) {
    throw new Error(`ICS fetch failed (HTTP ${res.getResponseCode()}), skipping this run.`);
  }
  return text;
}

/**
 * Turns ICS text into concrete occurrences overlapping [winStart, winEnd):
 * recurring series expanded, EXDATEs dropped, moved or edited instances
 * (RECURRENCE-ID) applied, cancelled ones skipped, time zones resolved.
 */
function expandIcs(text, winStart, winEnd) {
  const vcal = new ICAL.Component(ICAL.parse(text));

  // Outlook embeds its own time zone definitions (e.g. "W. Europe Standard Time")
  vcal.getAllSubcomponents('vtimezone').forEach(tz => {
    const id = tz.getFirstPropertyValue('tzid');
    if (id && !ICAL.TimezoneService.has(id)) ICAL.TimezoneService.register(tz);
  });

  // Group by UID: the series master plus its edited instances
  const byUid = new Map();
  vcal.getAllSubcomponents('vevent').forEach(ve => {
    const uid = ve.getFirstPropertyValue('uid');
    if (!uid) return;
    if (!byUid.has(uid)) byUid.set(uid, { master: null, exceptions: [] });
    const g = byUid.get(uid);
    if (ve.hasProperty('recurrence-id')) g.exceptions.push(ve);
    else g.master = ve;
  });

  const winStartMs = winStart.getTime();
  const winEndMs = winEnd.getTime();
  const out = [];

  const push = (key, item, startT, endT) => {
    if (isCancelled(item.component)) return;
    const startMs = startT.toJSDate().getTime();
    const endMs = (endT || startT).toJSDate().getTime();
    if (endMs <= winStartMs || startMs >= winEndMs) return;
    out.push(toOccurrence(key, item, startT, endT || startT));
  };

  for (const [uid, g] of byUid) {
    if (!g.master) {
      // Edited instances whose series isn't in the feed: treat as one-offs
      g.exceptions.forEach(ex => {
        const ev = new ICAL.Event(ex);
        push(`${uid}|${ev.recurrenceId.toString()}`, ev, ev.startDate, ev.endDate);
      });
      continue;
    }

    const event = new ICAL.Event(g.master);
    g.exceptions.forEach(ex => event.relateException(ex));

    if (!event.isRecurring()) {
      push(uid, event, event.startDate, event.endDate);
      continue;
    }

    const it = event.iterator();
    let next;
    let count = 0;
    while ((next = it.next()) && count++ < MAX_OCCURRENCES_PER_SERIES) {
      if (next.toJSDate().getTime() >= winEndMs) break;
      const d = event.getOccurrenceDetails(next);
      push(`${uid}|${d.recurrenceId.toString()}`, d.item, d.startDate, d.endDate);
    }
  }
  return out;
}

function isCancelled(comp) {
  return String(comp.getFirstPropertyValue('status') || '').toUpperCase() === 'CANCELLED';
}

function toOccurrence(key, item, startT, endT) {
  const comp = item.component;
  const transp = String(comp.getFirstPropertyValue('transp') || '').toUpperCase();
  const busy = String(comp.getFirstPropertyValue('x-microsoft-cdo-busystatus') || '').toUpperCase();
  const allDay = startT.isDate;
  return {
    key,
    summary: item.summary || '',
    location: item.location || '',
    allDay,
    free: transp === 'TRANSPARENT' || busy === 'FREE',
    start: allDay ? { date: startT.toString() } : { dateTime: startT.toJSDate().toISOString() },
    end: allDay ? { date: endT.toString() } : { dateTime: endT.toJSDate().toISOString() },
  };
}

/* ---------- Google Calendar side ---------- */

function buildCopy(key, src) {
  const title = CONFIG.COPY_TITLE && src.summary ? src.summary : CONFIG.FALLBACK_TITLE;
  const body = {
    summary: CONFIG.TITLE_PREFIX + title,
    start: src.start,
    end: src.end,
    visibility: 'private',                            // colleagues see "busy" only
    transparency: 'opaque',                           // actually blocks scheduling
    location: CONFIG.COPY_LOCATION ? src.location : '',
    reminders: CONFIG.REMINDERS === 'default'
      ? { useDefault: true }
      : { useDefault: false, overrides: CONFIG.REMINDERS },
    extendedProperties: { private: { [TAG]: '1', [TAG_ID]: key } },
  };
  if (CONFIG.COLOR_ID) body.colorId = CONFIG.COLOR_ID;
  return body;
}

function changed(copy, body) {
  return copy.summary !== body.summary ||
    (copy.colorId || '') !== (body.colorId || '') ||
    (copy.location || '') !== body.location ||
    reminderKey(copy.reminders) !== reminderKey(body.reminders) ||
    timeKey(copy.start) !== timeKey(body.start) ||
    timeKey(copy.end) !== timeKey(body.end);
}

function reminderKey(r) {
  r = r || {};
  if (r.useDefault) return 'default';
  return (r.overrides || []).map(o => `${o.method}:${o.minutes}`).sort().join(',');
}

function timeKey(t) {
  return t.date ? t.date : String(new Date(t.dateTime).getTime());
}

function listAll(calendarId, params) {
  const items = [];
  let pageToken = null;
  do {
    const query = Object.assign({ maxResults: 250 }, params);
    if (pageToken) query.pageToken = pageToken;
    const res = Calendar.Events.list(calendarId, query);
    items.push(...(res.items || []));
    pageToken = res.nextPageToken;
  } while (pageToken);
  return items;
}

/* ---------- Helpers you run by hand ---------- */

function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sync')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(CONFIG.SYNC_EVERY_MINUTES).create();
  sync();
}

/** Dry run: logs what the feed contains in the sync window, changes nothing. */
function preview() {
  const now = Date.now();
  expandIcs(fetchIcs(), new Date(now - CONFIG.DAYS_BACK * 864e5), new Date(now + CONFIG.DAYS_AHEAD * 864e5))
    .forEach(e => console.log(`${e.start.dateTime || e.start.date}  ${e.summary}` +
      `${e.allDay ? '  [all-day]' : ''}${e.free ? '  [free]' : ''}`));
}

function removeAllCopies() {
  listAll(CONFIG.TARGET_CALENDAR_ID, { privateExtendedProperty: `${TAG}=1` })
    .forEach(e => Calendar.Events.remove(CONFIG.TARGET_CALENDAR_ID, e.id));
}
