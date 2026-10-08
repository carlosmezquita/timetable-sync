# Timetable Sync

Convert a university iCalendar feed into useful Outlook appointments with module names, rooms and your own Free/Busy preferences.

Working CLI and reconciliation engine; Microsoft sign-in and live deployment require account setup. Outlook's primary calendar is the destination. Google Calendar copying is not included: showing Outlook on your phone is a separate integration to verify.

For a Microsoft 365-only route, see the [standard-connector Power Automate implementation](power-automate/README.md). It is supplied as a manual preview. A successful fresh feed download is required before applying or scheduling calendar changes.

## Use your local environment

~~~powershell
.\.venv\Scripts\timetable-sync.exe sync
.\.venv\Scripts\timetable-sync.exe rule "activity:Self Directive Studies" free
.\.venv\Scripts\timetable-sync.exe rule "module:M34521" busy
.\.venv\Scripts\timetable-sync.exe rules
~~~

A normal sync is an **offline preview**: source data is read, Outlook is neither checked nor changed. After authentication, sync --outlook previews against live Outlook; sync --apply writes it.

Fresh checkout (Python 3.10+):

~~~powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.lock
.\.venv\Scripts\python.exe -m pip install --no-deps -e .
.\.venv\Scripts\timetable-sync.exe init
~~~

Private .env contains the feed URL, client ID and encryption key. It is excluded from Git. Your school password is entered only on Microsoft's sign-in page.

## Microsoft setup

School credentials identify you. An Entra app registration identifies the application.

1. Open [Microsoft Entra](https://entra.microsoft.com/), select **App registrations**, then **New registration**. Use a single-tenant registration if your school permits it.
2. Set TSYNC_CLIENT_ID to Application (client) ID and TSYNC_TENANT_ID to Directory (tenant) ID in .env.
3. **Authentication**, then **Advanced settings**: enable **Allow public client flows** for device-code sign-in. No client secret is needed.
4. API permissions: add Microsoft Graph **delegated Calendars.ReadWrite**. School administrator consent may be required.
5. Sign in, then review a live preview before applying.

~~~powershell
.\.venv\Scripts\timetable-sync.exe auth
.\.venv\Scripts\timetable-sync.exe sync --outlook
.\.venv\Scripts\timetable-sync.exe sync --apply
~~~

If registration is disabled, request an approved public-client registration from university IT. A registration in another tenant needs the correct multitenant audience and school consent; it does not bypass school policy.

See [Microsoft device flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code) and [Graph event creation](https://learn.microsoft.com/en-us/graph/api/calendar-post-events?view=graph-rest-1.0).

## Availability

Tutorials, workshops and unknown types default to Busy. Self Directive Studies / Self Directed Studies default to Free.

Rule priority: occurrence, recurrence family, module-and-activity, module, activity, default. Selector names are case-insensitive.

~~~powershell
.\.venv\Scripts\timetable-sync.exe rule "activity:Workshop" busy
.\.venv\Scripts\timetable-sync.exe rule "module-activity:M33132|Tutorial" free
.\.venv\Scripts\timetable-sync.exe rule "event:FULL_KEY_FROM_PREVIEW" free
.\.venv\Scripts\timetable-sync.exe rule "activity:Workshop" default
~~~

Change a managed appointment's Show As to Free or Busy in Outlook: the next successful sync captures it as an occurrence preference. Explicit occurrence rules take priority. Removing an occurrence rule resets its captured preference. Source time, title, room and description are refreshed when source data changes.

## Reliability

- Full HTTPS download each run; invalid, failed or oversized feeds abort before writes.
- Recurrences, EXDATE, RECURRENCE-ID exceptions and London daylight saving are expanded into individual appointments.
- CMIS recurrence segment IDs are normalised; same-day time edits preserve identity.
- Missing classes require two successful applied checks at least five minutes apart. Preview does not count.
- Explicit cancellations can be immediate, subject to deletion safety limits.
- More than ten deletions or 25% of stored upcoming appointments are held for review.
- Ambiguous moves to a different day are held. After checking the source, run link-move NEW_KEY EXISTING_KEY and preview again.
- Historical and out-of-window appointments remain.
- Only marked events are managed. ETag version checks protect concurrent edits; meetings are rejected.
- Persistent creation transaction IDs and ownership discovery recover interrupted creates.
- Process locking prevents overlapping runs; OAuth cache is encrypted.

The five-minute timer cannot guarantee instant updates. Publication delays, source availability and sign-in policy affect freshness. Completely empty feeds are conservatively rejected. Keep the original subscription during initial verification, then hide it to avoid duplicate display.

## Health and development

~~~powershell
.\.venv\Scripts\timetable-sync.exe status
.\.venv\Scripts\timetable-sync.exe health
.\.venv\Scripts\timetable-sync.exe backup
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m ruff check src tests
.\.venv\Scripts\python.exe -m ruff format --check src tests
~~~

health fails when the last applied sync is older than 15 minutes or has warnings/held changes. Optional HTTPS TSYNC_HEARTBEAT_URL is called after an applied run; use a dead-man monitor to detect a stopped timer. Inspect health/status for held changes.

[Micro deployment](docs/deployment.md) | [Architecture](docs/architecture.md)

All test/demo events are synthetic. Feed URLs, tokens, state, backups and real timetable data must stay outside Git.
