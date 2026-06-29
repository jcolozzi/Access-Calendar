/* ============================================================================
 * calendar.ics.js — CalendarIcs
 * ----------------------------------------------------------------------------
 * Builds RFC 5545 (.ics) content from a calendar event object, using the
 * vendored ics() library (js/vendor/ics.js) as the underlying VEVENT writer.
 *
 * Responsibilities (single-purpose, no DOM/state coupling):
 *   - RFC 5545 text escaping for SUMMARY / DESCRIPTION / LOCATION.
 *   - Mapping the app's recurrence fields to a full "RRULE:..." line.
 *   - Correct all-day handling (VALUE=DATE with an exclusive end date).
 *   - UTF-8 safe base64 encoding for transport through the (naive) JSON bridge.
 *   - A standalone browser-download fallback (used when not hosted in Access).
 *
 * Loaded with `defer` BEFORE calendar.js, so `CalendarIcs` is ready by the time
 * the user can click the "Export .ics" button.
 * ==========================================================================*/
var CalendarIcs = (function () {
  'use strict';

  // ics weekday tokens, indexed 0=Sunday .. 6=Saturday (matches app DOW values)
  var DOW = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

  // ── RFC 5545 text escaping ────────────────────────────────────────────────
  // Order matters: backslash first, then the structural delimiters, then EOL.
  function icsEscape(text) {
    return String(text == null ? '' : text)
      .replace(/\\/g, '\\\\')
      .replace(/;/g, '\\;')
      .replace(/,/g, '\\,')
      .replace(/\r\n|\r|\n/g, '\\n');
  }

  // 'YYYY-MM-DD' (+ optional 'HH:MM') → a LOCAL Date, optionally shifted N days.
  // Building from numeric components avoids the "YYYY-MM-DD parses as UTC" trap
  // that would shift all-day events by a day in negative time zones.
  function toDate(ymd, hm, addDays) {
    if (!ymd) { return null; }
    var p = String(ymd).split('-');
    var y = parseInt(p[0], 10), mo = parseInt(p[1], 10) - 1, d = parseInt(p[2], 10);
    var hh = 0, mm = 0;
    if (hm) {
      var t = String(hm).split(':');
      hh = parseInt(t[0], 10) || 0;
      mm = parseInt(t[1], 10) || 0;
    }
    var dt = new Date(y, mo, d, hh, mm, 0, 0);
    if (addDays) { dt.setDate(dt.getDate() + addDays); }
    return dt;
  }

  // 'YYYY-MM-DD' → 'YYYYMMDD' (for UNTIL values)
  function ymdCompact(ymd) {
    return String(ymd || '').replace(/-/g, '');
  }

  // ── Recurrence → RRULE ────────────────────────────────────────────────────
  // Returns a full "RRULE:FREQ=..." line, or null when not recurring.
  // We build the whole line ourselves and hand it to ics.js via {rrule:...},
  // which bypasses ics.js's auto-builder (it emits a lowercase "rrule:" prefix).
  function buildRRule(ev) {
    var type = ev.recurType;
    if (!type || type === 'none') { return null; }

    var interval = parseInt(ev.recurInterval, 10) || 1;
    var parts;

    switch (type) {
      case 'daily':
        parts = ['FREQ=DAILY', 'INTERVAL=' + interval];
        break;

      case 'weekly': {
        parts = ['FREQ=WEEKLY', 'INTERVAL=' + interval];
        var days = String(ev.recurDaysOfWeek || '')
          .split(',')
          .map(function (s) { return s.trim(); })
          .filter(function (s) { return s !== ''; })
          .map(function (s) { return DOW[parseInt(s, 10)]; })
          .filter(Boolean);
        if (days.length) { parts.push('BYDAY=' + days.join(',')); }
        break;
      }

      case 'monthly':
        parts = ['FREQ=MONTHLY', 'INTERVAL=' + interval];
        if ((ev.recurMonthlyMode || 'day') === 'weekday') {
          // e.g. "the Last Friday" → BYDAY=FR;BYSETPOS=-1
          var dow = DOW[parseInt(ev.recurMonthDOW, 10) || 0];
          var week = parseInt(ev.recurMonthWeek, 10) || 1; // 1..4, or -1 for "Last"
          parts.push('BYDAY=' + dow);
          parts.push('BYSETPOS=' + week);
        } else {
          parts.push('BYMONTHDAY=' + (parseInt(ev.recurMonthDay, 10) || 1));
        }
        break;

      case 'yearly':
        // Anchors on DTSTART's month/day automatically.
        parts = ['FREQ=YEARLY', 'INTERVAL=' + interval];
        break;

      default:
        return null;
    }

    // End condition — value type of UNTIL must match DTSTART (DATE vs DATE-TIME)
    var endType = ev.recurEndType || 'never';
    if (endType === 'count') {
      var count = parseInt(ev.recurCount, 10) || 0;
      if (count > 0) { parts.push('COUNT=' + count); }
    } else if (endType === 'date' && ev.recurEndDate) {
      if (ev.allDay) {
        parts.push('UNTIL=' + ymdCompact(ev.recurEndDate));
      } else {
        parts.push('UNTIL=' + ymdCompact(ev.recurEndDate) + 'T235959Z');
      }
    }

    return 'RRULE:' + parts.join(';');
  }

  // ── Build the VCALENDAR string for a single event (or null) ───────────────
  function buildIcsString(ev) {
    if (!ev || !ev.date) { return null; }
    if (typeof ics === 'undefined') { return null; }

    var begin, stop;
    if (ev.allDay) {
      begin = toDate(ev.date, null, 0);
      stop  = toDate(ev.endDate || ev.date, null, 1); // exclusive end (next midnight)
    } else {
      begin = toDate(ev.date, ev.start || '00:00', 0);
      stop  = toDate(ev.endDate || ev.date, ev.end || ev.start || '00:00', 0);
      if (!stop || stop <= begin) {                   // guard zero / negative spans
        stop = new Date(begin.getTime() + 60 * 60 * 1000);
      }
    }
    if (!begin || !stop) { return null; }

    var rrule = buildRRule(ev);
    var cal = ics(Date.now() + '.access-calendar', 'Access Calendar');

    // Pass Date objects (not strings): ics.js reads them with local getters,
    // matching the local components we constructed above.
    var added = cal.addEvent(
      icsEscape(ev.title || 'Untitled'),
      icsEscape(ev.notes || ''),
      icsEscape(ev.location || ''),
      begin,
      stop,
      rrule ? { rrule: rrule } : undefined
    );
    if (added === false) { return null; }

    var out = cal.build();
    if (!out) { return null; }

    // ics.js writes ";VALUE=DATE-TIME:YYYYMMDD" (no time part) for all-day
    // events. Re-tag those as VALUE=DATE so clients treat them as true all-day
    // entries. DTSTAMP always carries a time, so it is never matched.
    out = out.replace(/;VALUE=DATE-TIME:(\d{8})(?!T)/g, ';VALUE=DATE:$1');
    return out;
  }

  // ── Encoding / transport ──────────────────────────────────────────────────
  // UTF-8 safe base64 so the ICS text survives the naive JSON bridge intact.
  function toBase64(str) {
    try {
      return btoa(unescape(encodeURIComponent(str)));
    } catch (e) {
      return null;
    }
  }

  function buildBase64(ev) {
    var s = buildIcsString(ev);
    return s ? toBase64(s) : null;
  }

  // ── Standalone fallback: trigger a browser download (no Access bridge) ─────
  function download(filename, base64) {
    try {
      var bin = atob(base64);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) { bytes[i] = bin.charCodeAt(i); }
      var blob = new Blob([bytes], { type: 'text/calendar;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = filename || 'event.ics';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    } catch (e) { /* no-op */ }
  }

  // Public surface
  return {
    escape: icsEscape,
    buildRRule: buildRRule,
    buildIcsString: buildIcsString,
    buildBase64: buildBase64,
    download: download
  };
})();

window.CalendarIcs = CalendarIcs;
