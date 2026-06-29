# Access Calendar

![Platform](https://img.shields.io/badge/Platform-Microsoft%20Access-00599C)
![UI](https://img.shields.io/badge/UI-Calendar%20WebView2-2563EB)
![Features](https://img.shields.io/badge/Features-Recurring%20Events%20%7C%20Reminders-7C3AED)
![Status](https://img.shields.io/badge/Status-Active-22C55E)

Access Calendar is a modern calendar experience for Microsoft Access, powered by a browser-based UI and VBA backend logic.

It brings a polished scheduling interface to Access while keeping your data and automation where your team already works.

## Highlights

- Multi-calendar scheduling with a primary calendar context
- Month, week, 5-day week, and day views
- Recurring events with occurrence-level operations
- Calendar groups and color-based organization
- Reminder popups and theme preferences
- Add events to your Outlook calendar (single or recurring)
- Export events to `.ics` files and attach them to a new Outlook email
- Optional location field on events
- Fast drag/drop and resize interactions
- Live cross-session sync — changes by one user appear in other open sessions within ~3 seconds

## Screenshots

### Month view

![Calendar week view placeholder](https://github.com/jcolozzi/Access-Calendar/blob/main/images/month-view.png)

### Week view

![Calendar week view placeholder](https://github.com/jcolozzi/Access-Calendar/blob/main/images/week-view.png)

### Day view

![Calendar day view placeholder](https://github.com/jcolozzi/Access-Calendar/blob/main/images/week-view.png)

### Settings

![Calendar settings view placeholder](https://github.com/jcolozzi/Access-Calendar/blob/main/images/settings-view.png)

## Quick Start

1. Open `Calendar.accdb` in Microsoft Access on Windows.
2. Open the calendar form.
3. Start creating calendars and appointments.

## Project Layout

- `calendar.html`: calendar UI markup
- `css/calendar.css`: calendar styles
- `js/`: calendar UI logic (`calendar.js`), ICS export (`calendar.ics.js`), and the vendored `vendor/ics.js`
- `cls/`: VBA classes for bridge, command routing, repos, recurrence, reminders, and Outlook/ICS export
- `bas/`: Standard VBA modules (change log for cross-session sync)
- `ARCHITECTURE.md`: deeper architecture and integration notes

## Documentation

- Technical/developer README: [README.dev.md](README.dev.md)
