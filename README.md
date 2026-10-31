# Google Drive Folder Inventory

This package creates a browsable, sortable Google Sheets inventory of any Google Drive folder and its descendants. It works with PDFs, MP3s, images, videos, archives, Google documents, and other file types. Use the same script in a separate blank spreadsheet for each collection.

The inventory reads Drive metadata. It does not download, inspect, rename, move, or delete the source files. A refresh rebuilds the generated inventory after the new scan completes. The last successfully published inventory remains available while the next scan runs.

## Files in this package

| File | Purpose |
| --- | --- |
| `DriveInventory.gs` | Apps Script source |
| `appsscript.json` | Manifest, including the Advanced Drive v3 and Advanced Sheets v4 services |
| `README.md` | Setup, usage, metadata limits, and recovery |
| `test_inventory.cjs` | Optional offline developer test harness (not part of Apps Script installation) |

Install the supplied source and manifest together.

## Installation

1. Create a new, blank Google Sheet in the Google account that can access your collection. Give it a useful name, such as "Gaming collection inventory" or "Old-time radio inventory." Start with a new workbook: the script owns and rebuilds its generated tabs.

2. In the spreadsheet, select **Extensions → Apps Script**. This creates or opens a script bound to that spreadsheet. Give the script project a recognizable name.

3. Open the default `Code.gs` file in the Apps Script editor. Replace its entire contents with the contents of the supplied `DriveInventory.gs` file. Keeping the editor filename `Code.gs` is fine, or you can rename it `DriveInventory.gs`. Install the source once, not in both files. Do not paste the source into a sheet cell or add Markdown code-fence markers around it.

4. Open **Project Settings** in the Apps Script editor. Enable **Show appsscript.json manifest file in editor.** Return to the editor, open `appsscript.json`, and replace its entire contents with the supplied `appsscript.json` file. Keep the manifest as valid JSON, without comments or surrounding code fences.

5. Save the project. If the function selector has not refreshed, reload the Apps Script editor. Select `setupInventory` in the function selector and click **Run**. Complete Google's authorization flow using the account that owns this script and can access the collection. Review the permissions before granting them.

6. Dismiss the setup confirmation, return to the spreadsheet, and reload it. A **Drive Inventory** menu should appear. The setup function prepares Status and internal state, installs the menu, and records the identity of this workbook. It does not choose a root or start a scan.

7. From **Drive Inventory**, choose **Choose root folder and scan…**. Paste the full Google Drive folder URL or just its folder ID, then confirm. This starts the first scan. A URL containing a `resourcekey` parameter is accepted; keep that parameter when pasting such a link. The root must be an accessible, non-trashed folder itself, not a file or a shortcut. A resource key supplements a shared link; it does not bypass the owner's sharing settings.

8. Watch the **Status** tab for progress, or use **Drive Inventory → Show status**. Large scans continue in background executions. The default worker aims for about three minutes of work per execution, and a temporary clock trigger runs at five-minute intervals to continue it. You can close the tab while it runs. Google controls trigger scheduling, so continuation may not be immediate. Do not select **Refresh** merely to check progress: that starts over.

No web app deployment, public URL, service account, or external server is needed. With the default Cloud project created by Apps Script, adding the advanced services normally enables the corresponding APIs automatically. If you have deliberately linked the script to a standard Google Cloud project, enable both the Google Drive API and Google Sheets API in that project as well.

## What the sheets contain

### Inventory

Contains one row per non-folder item, with a clickable file link and its folder context. It includes the useful metadata Drive returns: names, paths, file types, sizes, dates, descriptions, ownership information where available, checksums where available, and applicable image, video, or shortcut details.

- The complete inventory has **32 columns**; the compact view shows its first **12**.
- Owner and last-modifier fields include the names and email addresses Drive makes available. Those people are file users, not necessarily authors or performers.
- Empty metadata remains empty instead of being guessed.

**Additional metadata (JSON)** retains available image and video metadata, SHA-1, revision and ownership information, full extensions, and custom properties. `appProperties` includes only private properties visible to this application; it does not reveal other applications' private metadata. These extra values are retained as text rather than expanded into dozens of mostly empty columns. Very long text is capped below the cell limit and marked as truncated, with a cell note explaining the truncation.

### Folders

Contains one row per folder, including the selected root. It records folder identities, paths, and counts of immediate contents.

- Counts are **direct** counts: a folder's file count does not include files inside its subfolders. The same applies to its direct file-size total.
- The folder table has **16 columns**; the compact view shows its first **11**.
- The root uses a folder path of `.` and the category `(root)`. Other categories come from the first folder level below the chosen root. A category is a folder label, not an inferred genre or game system.

### Status

Identifies the chosen root and reports the current scan state, progress, completion information, or the error that caused a pause. Use it to distinguish a running or paused scan from a completed inventory.

The script also uses hidden state and staging sheets. Leave those sheets and their names alone. They allow the scan to resume and keep unfinished results separate from the published inventory.

## Browsing, searching, and sorting

- Use the filter controls in the header row to sort or filter the inventory. The folder path gives each row its context, so sorting by filename, size, date, or type does not separate a file from its location. A file ID is its stable identity; names and paths are not guaranteed to be unique.
- Use the normal Google Sheets **Find** command to search the list. Searchable text includes the filenames, folder paths, and any metadata text included in the sheet. This is a search of the inventory, not a search inside the books or audio.
- Use **Drive Inventory → Show all metadata columns** to expose the technical columns. Use **Drive Inventory → Show compact columns** to return to the main browsing columns. Hiding a column changes the view; it does not remove its data. These commands operate on the published inventory and require a completed scan.
- Dates in the inventory are native spreadsheet date/time values displayed as UTC, so they can be sorted and filtered as dates. Numeric sizes are numbers rather than strings. MiB means 1,048,576 bytes; **Size (bytes)** retains the source value.
- Click a file or folder link to open it in Google Drive. Opening the link still requires the viewer to have access to that item. Sharing the spreadsheet does not grant access to the files it lists.

## PDF, MP3, and other metadata limits

PDF and MP3 collections are both supported as file inventories. Their embedded tags are a different layer of metadata from the information supplied by Drive.

A normal Drive listing does **not** provide:

- PDF author, publisher, page count, ISBN, edition, or publication year
- MP3 ID3 fields such as artist, album, track title, track number, genre, or broadcast date

This package does not parse those fields, even when they are embedded in the source files. Extracting them would require an additional process that reads file contents.

Image and video metadata is included where Drive supplies it. The presence of video-related metadata does not imply that equivalent fields are available for audio files. Fields that do not apply to an item stay blank.

Drive descriptions and custom file information may be useful when someone has already populated them. They are not automatically generated book abstracts or episode descriptions. File creation and modification timestamps describe the Drive file; they are not publication dates, purchase dates, or broadcast dates.

An archive such as a ZIP file appears as one file. Its internal files are not enumerated. Google-native documents and shortcuts may lack size or checksum fields that are available for ordinary stored files. Checksums can help compare files, but this package does not decide which copies to keep or delete.

## Shortcuts and collection scope

The selected root and its ordinary descendant folders define the inventory. Folder shortcuts are recorded as shortcuts and are not followed recursively. This prevents a shortcut from silently expanding the scan into another part of Drive. Available shortcut target information is retained as metadata.

The scan includes all file types that it can list within the selected tree. Trashed items are excluded. Access to shared-drive folders depends on the account's permissions and any organization policies affecting Drive or Apps Script.

The code restricts traversal to the chosen root. Google's metadata authorization is not a grant limited to that single folder; the consent screen describes the actual scope. The manifest requests Drive metadata access, Sheets access, and permissions needed for the menu, background triggers, and HTTP requests. The HTTP permission supports Drive requests that need a resource-key header. The code writes only its selected output workbook and does not edit source files.

## Refresh, pause, resume, and changing collections

| Menu command | Behavior |
| --- | --- |
| **Refresh — start a new scan** | Begins a fresh scan of the chosen root. It does not append another copy of the list. New file and folder tables are published together only after the scan and final preparation succeed. If no root has been chosen yet, this opens the root prompt. |
| **Pause scanning** | Stops continuation and retains saved progress. Does not interrupt a batch that already holds the worker lock. |
| **Resume / run next batch** | Continues that scan and runs a batch immediately. Also available when you want to run the next batch manually instead of waiting for the clock trigger. Use after a recoverable error or an intentional pause. |
| **Choose root folder and scan…** | Changes the target and starts a new scan. Prior published output remains available until the new target has been scanned and published. |

Use **Refresh** when you want to discard unfinished work and start from the beginning.

Only one batch or control operation runs at a time. If a command reports that a scan batch is running, try again after the current batch finishes.

The worker saves progress in small batches. A batch's rows and its checkpoint are committed together. If a folder's pagination token can no longer be used, the scanner rolls back the staged progress for that folder and lists it again. It does not need to discard the entire collection scan for that condition.

For separate persistent inventories, create a separate blank workbook for each root and install this same package in each one.

**Inventory**, **Folders**, and **Status** are generated output. Do not put personal notes or other irreplaceable edits in them. Keep notes in a separate sheet or workbook, keyed by File ID if you want to associate them with inventory records. The script does not merge arbitrary manual edits back into regenerated output. Publication replaces the generated sheet objects as well as their contents, so saved views, custom formatting, or formulas elsewhere that refer directly to those sheet objects may need to be recreated. This version is a generated list, not an editable annotation database.

Drive listings are not an instantaneous snapshot of a whole collection. If you reorganize folders during a long scan, run another refresh afterward to reconcile the inventory with the new organization.

## Large collections and troubleshooting

Apps Script has execution and daily quotas. The scanner splits work across executions, but that does not bypass daily limits. A very large collection can take multiple executions and may require resuming after a quota becomes available again. Different inventory workbooks and other scripts on the same account can consume the same account-level quotas.

The default automatic-work allowance is deliberately conservative: **45 minutes** per 24-hour budget window, reserving **four minutes** for each background execution. The worker targets roughly **three minutes**, leaving room for service calls and overhead. When the reserved budget is exhausted, Status shows `WAITING FOR AUTOMATIC RUNTIME BUDGET` and the next eligible time. The clock trigger stays installed and can continue automatically when that window resets.

**Resume / run next batch** can run a manual batch while the automatic allowance is waiting. It does not bypass Google's actual quotas, and it does not reset the reserved automatic-work budget. The allowance is a guard in this script, not a measurement of every script's actual runtime across your Google account.

Google currently lists a 20-million-cell limit for a spreadsheet. This script uses a configurable **19-million-cell** ceiling to leave headroom. The check counts allocated cells across the workbook, not just cells containing values. Both the published file/folder tables and the staged replacement consume cells during a refresh, as do Status, internal state, and any other sheets.

The script checks cell capacity; it does not estimate the workbook's byte size or guarantee that every other platform/API limit will be satisfied. Those errors also pause the scan. If a collection cannot fit, use separate inventories for subfolders or consider a CSV/database-based inventory. Reducing visible columns with **Show compact columns** does not reduce the underlying cell count.

### Configuration

The `DI` configuration near the top of `DriveInventory.gs` contains:

| Constant | Default | Notes |
| --- | --- | --- |
| `PAGE_SIZE` | `200` | |
| `WORK_MS` | `180000` | |
| `TRIGGER_MINUTES` | `5` | |
| `AUTO_BUDGET_MS` | | |
| `AUTO_RESERVATION_MS` | | |
| `MAX_CELLS` | `19000000` | |
| `EXCLUDED_FOLDER_IDS` | `[]` (empty) | Adding descendant-folder IDs deliberately excludes those subtrees. Excluded folders remain visible in Folders with an Excluded status; their content counts are not measured. The selected root itself is always scanned. |

The defaults are intended to be usable without tuning.

### Common issues

| Symptom | What to check |
| --- | --- |
| **Drive Inventory** menu is missing | Reload the spreadsheet and check that `setupInventory` ran successfully in the script bound to this workbook. |
| Drive or Sheets reported as undefined | Confirm that the supplied manifest was installed, saved, and loaded. Verify that Drive v3 and Sheets v4 appear in the Apps Script project's advanced services. For a standard Cloud project, verify that both corresponding APIs are enabled there. |
| Authorization fails | Verify the signed-in account and the requested folder's access. Work or school administrators can restrict Apps Script or Drive API use. An organization restriction must be resolved through the account or its administrator; deploying this script as a web app is not a substitute. |
| Scan pauses | Read the error on Status. Resolve the reported problem and use **Resume / run next batch**. A paused or failed scan leaves the previous published Inventory and Folders available. During the first-ever scan there is no earlier complete inventory to fall back to. |
| Trigger seems slow | Allow for Google's scheduling delay and inspect Status. You can use **Resume / run next batch** to continue work. Avoid repeatedly starting new refreshes: each fresh scan discards unfinished progress and starts over. |
| Metadata is blank | First check whether Drive actually exposes that field for the file type. Missing PDF bibliographic fields or MP3 ID3 tags are expected with this metadata-only approach; rescanning will not extract them from the content. |

## Validation status

The supplied source passed a JavaScript syntax check and 14 offline integration tests using mocked Google services. The tests cover complete and empty scans, mixed file types, nested folders, literal text, pagination, duplicate IDs, interrupted commits, expired tokens, failed refreshes, publication recovery, exclusions, capacity limits, automatic runtime budgets, and root URL parsing.

To run these developer checks locally:

```bash
node test_inventory.cjs
```

These checks are not a live integration run against your Google account. The script has not been authorized or executed against your actual collection in this conversation. Your first run verifies account permissions, API availability, and the behavior of the real collection. A small test folder is useful before starting a very large inventory.

## Official references

- [Apps Script advanced Drive service](https://developers.google.com/apps-script/advanced/drive)
- [Apps Script advanced Sheets service](https://developers.google.com/apps-script/advanced/sheets)
- [Enable advanced services and corresponding APIs](https://developers.google.com/apps-script/guides/services/advanced)
- [Drive file metadata fields](https://developers.google.com/workspace/drive/api/reference/rest/v3/files)
- [Search and list files and folders](https://developers.google.com/workspace/drive/api/guides/search-files)
- [Drive API authorization scopes](https://developers.google.com/workspace/drive/api/guides/api-specific-auth)
- [Apps Script quotas and execution limits](https://developers.google.com/apps-script/guides/services/quotas)
- [Apps Script installable and time-driven triggers](https://developers.google.com/apps-script/guides/triggers/installable)
- [Google Drive storage and document size limits](https://support.google.com/drive/answer/37603)
