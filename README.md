# ScriptBuilderSheets

Utilities for orchestrating Google Apps Script automations that sync Google Sheets data into Google Slides.

## Slide sync workflow

The Apps Script code in [`src/SlideSync.gs`](src/SlideSync.gs) exposes a `refreshAllSlides()` controller that reads configuration rows and pushes spreadsheet tables into slide tables. The script is intended to be pasted into an Apps Script project that has access to both the Sheets and Slides that you want to manage.

### Configuration sheet

Create a sheet named **`Slide Sync Config`** inside the spreadsheet that will host the Apps Script project. Give it the following headers (row 1):

| Column | Description |
| --- | --- |
| `enabled` | Optional flag (`TRUE/YES/1`) to toggle the row. Leave blank to run. |
| `spreadsheet id` | ID of the source spreadsheet containing the data. |
| `sheet name` | Sheet within the spreadsheet to pull data from. |
| `range` | Optional A1 range for the data. Leave blank to use the entire sheet range. |
| `skip header rows` | Optional number of rows to skip from the top before syncing. |
| `max rows` | Optional cap on the number of rows that will be written to the slide table. |
| `remove blank rows` | Optional flag to drop rows that are entirely blank (defaults to `TRUE`). |
| `schedule enabled` | Optional flag to let this job run via the background scheduler. |
| `cron expression` | Cron timing (`minute hour day month weekday`) that defines when the scheduler should refresh the job. |
| `slide id` | ID of the destination slide deck. |
| `page index` | Zero-based index of the slide that contains the table. |
| `match text` | Text used to identify the correct table on the slide (matching the first cell). |
| `template` | Optional template key (`weekly-report`, `monthly-dashboard`, `daily-metrics`) that pre-fills defaults. |
| `validation rules` | Optional JSON rules that enforce data checks before syncing. |

Each populated row defines one sync job. When `refreshAllSlides()` runs it loads each enabled row and calls `runParse()` to mirror the spreadsheet data into the matching slide table.

If the configuration sheet cannot be found, the script falls back to the hard-coded jobs defined inside `getFallbackUpdates()` so teams can keep existing workflows while migrating to a fully configurable setup.

### Deploying

1. Open the spreadsheet that should control the sync and launch the **Extensions → Apps Script** editor.
2. Paste the contents of [`src/SlideSync.gs`](src/SlideSync.gs) into a script file (or import via clasp).
3. Adjust the fallback configuration (optional) or rely solely on the rows in `Slide Sync Config`.
4. Run `refreshAllSlides()` and grant the required permissions the first time it is executed.
5. (Optional) Open the **Slide Sync Manager** sidebar via `showSlideSyncSidebar()` (or assign it to a custom menu) to preview jobs, toggle schedules, and run individual syncs.
6. Use `enableScheduledRefresh()` to install the minute-level polling trigger when cron-based schedules are configured.

### Extending

The helper functions exposed in the script are modular so you can build additional scripts:

- Call `runParse(updateConfig)` with a single configuration object to refresh a single slide.
- Adjust `normalizeDataToTable` if you need to format numbers or apply currency symbols before writing to Slides.
- Expand the configuration sheet with new columns (e.g., custom filters) and extend `getConfiguredUpdates()` to read them.
- Define new templates in `SLIDE_SYNC_TEMPLATES` to offer one-click defaults for common media reporting layouts.
- Extend `validateDataSet` if you require additional business rules (e.g., allowed value lists, date validation).

### Preview and validation

- `previewAllSlides()` returns diff metadata for every configured job without touching Slides.
- `previewUpdate(updateConfig)` is used by the sidebar to show row-level changes before import.
- Validation rules accept JSON such as `{ "requiredColumns": ["campaign"], "numericColumns": ["impressions"], "minRows": 1 }`.

### Scheduling

- Add cron expressions (e.g., `0 9 * * 1`) to configuration rows and call `enableScheduledRefresh()` once to install the poller.
- The `pollScheduledSlides()` trigger checks cron expressions every five minutes and calls `runParse()` only when due.
- Use `disableScheduledRefresh()` to remove the polling trigger.

### Requirements

- Google Workspace account with access to the target Sheets and Slides.
- Apps Script project with the **Google Sheets** and **Google Slides** advanced services enabled (they are standard services).

