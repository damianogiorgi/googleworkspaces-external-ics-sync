# live-to-work-calendar-sync

A one-way sync from a personal Outlook.com (Live) calendar to a Google Workspace calendar, written in Google Apps Script.

Personal events become **private busy blocks** on your work calendar, so colleagues don't book calls while you're at the mechanic. Nothing travels in the other direction: no work data leaves the work account, and no third-party service gets access to it.

📖 The full story, with the motivation and the design choices: [Busy, Not Nosy: a One-Way Sync from My Personal Calendar to Google Workspace](ARTICLE_URL)

## How it works

Every 15 minutes the script downloads the ICS feed of your published Outlook.com calendar, expands recurring events (exceptions, cancellations and time zones included) with [ical.js](https://github.com/kewisch/ical.js), and keeps a copy of each event on your primary Google calendar. Copies are private and block scheduling, and they're tagged so the script only ever touches events it created. If the feed can't be downloaded, the run is skipped and nothing changes.

Options in `CONFIG` let you choose whether to copy titles and locations, a title prefix and color for personal events, reminders, the sync window, and whether to skip all-day and "free" events.

## Files

| File | What it is |
|------|------------|
| `live-to-work-sync.gs` | The sync script |
| `ical.gs` | ical.js 1.5.0 (minified), the iCalendar parser. Do not edit |

## Setup

1. On Outlook.com: Settings > Calendar > Shared calendars > publish your calendar ("Can view titles and locations" is enough) and copy the ICS link.
2. On [script.google.com](https://script.google.com), with your work account, create a project and paste `live-to-work-sync.gs`.
3. Add a second script file named `ical` and paste `ical.gs`.
4. Services (+) > add **Google Calendar API**, keeping the identifier `Calendar`.
5. Project Settings > Script Properties > add `LIVE_ICS_URL` with the ICS link.
6. Run `preview()` to check what the script sees in the feed (it changes nothing) and authorize it.
7. Run `installTrigger()` to start syncing.

To remove everything the script created, run `removeAllCopies()` and delete the trigger.

## Before you use it

- **Ask first.** This runs in your work account and needs permission to fetch an external URL. Check your company's policy and talk to your Workspace admin.
- **Keep the ICS link secret.** Anyone with it can read your personal calendar. It lives in Script Properties for this reason: never commit it.
- **Expect some delay.** Outlook.com may cache the published feed, so new events can take a while to show up.

## License

`live-to-work-sync.gs` is released under the [GNU General Public License v3.0 or later](LICENSE).

`ical.gs` is an unmodified copy of [ical.js](https://github.com/kewisch/ical.js) by Philipp Kewisch and contributors, and remains under the [Mozilla Public License 2.0](https://mozilla.org/MPL/2.0/), as stated in its header. If you modify that file, your changes to it are covered by the MPL as well.
