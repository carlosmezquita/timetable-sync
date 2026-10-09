# Power Automate synchroniser

This implementation runs in Microsoft 365 using three **standard** connectors: OneDrive for Business, Excel Online (Business), and Office 365 Outlook. It needs no VPS, Premium trial, or custom Entra application.

The Office Script and package builder are implemented and tested. The package defaults to a **manual preview** with `ApplyChanges: false`. A successful, fresh feed download and Microsoft flow validation are deployment requirements. A broken download is not a working synchroniser: it stops the run before calendar changes.

## Components

- `TimetableSync.ts`: an Office Script that validates and expands the feed, compares owned appointments, renders the `Preview` worksheet, and commits acknowledged state.
- `build-flow.py`: generates a cloud-flow package from an exported seed containing your connected standard actions. The seed is treated as data.
- `sample.zip`: an unconfigured, sanitised package. Its feed, workbook, script and calendar placeholders must be replaced before use.
- `test-planner.cjs` and `test_flow.py`: recurrence, reconciliation, and workflow safety checks using synthetic events. An optional private-feed parity check can compare expansion with the Python implementation.

## Install and connect

1. Create an Excel workbook in your school OneDrive, such as **Timetable Sync Control.xlsx**. In **Automate > New Script > Create in Code Editor**, paste `TimetableSync.ts` and name it **Timetable Sync Planner**. The `Preview` worksheet is reserved for this script.
2. Create a temporary manually triggered seed flow containing **Upload file from URL**, **Run script**, and **Get calendar view of events (V3)**. Connect each to your own school account. Select the workbook, planner script, and your main Outlook calendar. The seed must contain no calendar write actions.
3. Export the seed as a package through **Export > Package (.zip)**. Keep this export private: it contains your feed URL and account configuration.
4. Build a configured package locally:

   ```text
   python power-automate/build-flow.py private-seed.zip private-standard-flow.zip
   ```

5. In Power Automate, choose **My flows > Import > Import Package (Legacy)**, upload the generated package, create it as a new flow, and select your three existing school connections.
6. Run the imported flow manually. Check the workbook's `Preview` worksheet and flow history. Verify titles, London times, locations, Free/Busy choices, and held changes against the source timetable. **Do not enable ApplyChanges until the fresh-download preview succeeds.**

You can extract a private configuration file or build a shareable, unconfigured sample:

```text
python power-automate/build-flow.py private-seed.zip private-config.json --extract-config
python power-automate/build-flow.py private-seed.zip private-standard-flow.zip --config private-config.json
python power-automate/build-flow.py private-seed.zip sample.zip --public
```

Keep private exports, configuration, state, packages, and real timetable files outside Git.

## Behaviour

Each download uses a unique OneDrive filename with overwrite disabled. The flow verifies that a new file exists, reads its contents, and checks freshness and full iCalendar boundaries. It reads Outlook in explicit pages of 256 events; an unfinished or oversized scan cannot be applied.

The script supports bounded daily and weekly recurrences with COUNT, UNTIL, INTERVAL, BYDAY, WKST, RDATE, EXDATE, and individual RECURRENCE-ID overrides. It handles UTC and Europe/London timestamps, including daylight saving. Unsupported rules, ambiguous transition times, conflicting identities, and orphan exceptions stop the run rather than silently dropping classes.

CMIS recurrence segments use their embedded event identity and original London date. Explicit moved exceptions retain that original identity. Unexplained cross-date changes are held for review instead of being guessed.

Only appointments carrying the source's `Timetable Sync [v1:...]` ownership marker are managed. Personal events are ignored; owned meetings, all-day events, and recurring series are rejected. Lost create checkpoints are recovered from ownership markers.

Self Directive Studies / Self Directed Studies default to Free; other activities default to Busy. Outlook Web **Show as** edits are captured as occurrence preferences. Optional `Rules` in the flow's `Config` variable use lowercase selectors such as `activity:workshop`, `module:m12345`, or `module-activity:m12345|tutorial`, with values `free` or `busy`. An explicit `event:KEY` rule wins over a captured preference.

Set `Rules.calendarCategory` to an existing Outlook category name, such as `Timetable`, to apply it to new and upcoming managed appointments. Outlook supplies the category's colour. Existing labels and Free/Busy choices remain; a missing required category triggers an update. Updates merge the required category with the latest labels read from Outlook, preserving category edits made after planning. Historical appointments outside the scan are left unchanged.

Missing or cancelled appointments require two successful applied observations at least five minutes apart. More than ten deletions or 25% of the upcoming stored appointments are held. Historical appointments remain. Preview runs do not advance deletion confirmations.

## Applying and scheduling

After a successful reviewed preview, change `Config.ApplyChanges` to true and run manually first. Every update/deletion re-reads the target, verifies ownership and absence of attendees/recurrence, and then acts. Updates preserve current reminders/categories and check for Free/Busy edits made since planning. Operations run serially; a failure stops further mutations. Create retries are disabled to prevent duplicate requests.

The script commits state only after all requested operations have acknowledgements. The flow then saves that state in your private OneDrive. Use only **one active flow per source/state file**. The standard Outlook connector does not expose ETag conditional writes, so a very small race between its final read and update remains; the Python Graph implementation offers stronger concurrency protection.

Once manual runs are verified, use the designer to replace the button trigger with Recurrence and configure the connectors to use the owner's connections. Start with **15 minutes**, check the account's action allowance and actual run analytics, and adjust only within those limits. A 15-minute interval can take about 30 minutes to confirm a disappeared class. Feed publication and service delays still affect freshness. Keep the existing subscription visible during verification, then hide it to avoid duplicate display.

Source/scan failures refresh the Preview status without attempting calendar operations. Apply failures explicitly report that some operations may have completed; inspect the flow run before retrying. A source endpoint that the OneDrive downloader cannot retrieve must be repaired or replaced with a supported standard download path. Office Scripts cannot fetch external feeds when called by Power Automate.

## Verification

Use Python 3.10+ and Node 24 (no Node packages are required):

```text
node power-automate/test-planner.cjs
python power-automate/test_flow.py
```

After editing the script, also save and run a synthetic preview in Excel to check Microsoft's Office Scripts compiler/runtime. The local Node checks do not replace that validation.

References: [OneDrive connector and download verification](https://learn.microsoft.com/en-us/connectors/onedriveforbusinessconnector/), [Outlook connector and paging/update behaviour](https://learn.microsoft.com/en-us/connectors/office365connector/), [Office Script external-call restrictions](https://learn.microsoft.com/en-us/office/dev/scripts/develop/external-calls), [Microsoft 365 Power Automate licensing](https://learn.microsoft.com/en-us/power-platform/admin/power-automate-licensing/faqs).
