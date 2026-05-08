# Clasp_ToDoList Agent Notes

This file governs the whole project root.

## Project Shape

- This is a Google Apps Script project managed by `clasp`.
- Main files:
  - `程式碼.js`: backend Apps Script logic, webhook handling, spreadsheet IO.
  - `index.html`: HtmlService frontend.
  - `appsscript.json`: Apps Script manifest.
- The deployed app is sensitive to HtmlService frontend parsing errors. A small JS syntax issue can make the whole app fail to load.

## Frontend Rules

- Prefer conservative browser-compatible JavaScript in `index.html`.
- Be extra careful with large inline script changes. HtmlService can be less forgiving than local Node syntax checks.
- Avoid adding complex inline template literal blocks unless there is a strong reason.
- Prefer simple string concatenation or small helper functions over deeply nested dynamic HTML.
- When editing event handlers embedded in HTML strings, escape values carefully.
- If the page stops loading and shows `userCodeAppPanel` syntax errors, suspect recent `index.html` edits first.
- Keep first-load logic minimal. Initial render should only fetch what is needed for the first screen.
- Do not let secondary preload work block user switching. Switching users should prioritize loading that user's tasks; meeting preload can run in the background.
- When adding async frontend calls that depend on `currentUser`, capture the requested user and ignore stale responses if `currentUser` changed before the callback returns.
- Meetings are cached client-side after `getMeetings()` because that call already includes the detail fields used by the meeting detail panel. Prefer rendering from `meetingDetailsById` before calling `getMeetingById()`.
- When navigating from a task card to its related meeting, pass the target meeting id through `switchView('meetings', meetingId)` / `fetchMeetings(meetingId)` so a later meeting-list response does not auto-select the first meeting and overwrite the intended detail.
- Only force-refresh meetings after explicit refresh/sync actions or when cache loading failed. Normal tab switches and task-to-meeting jumps should use cached data for responsiveness.
- Task search and sorting are frontend display concerns. Keep filtering in `getVisibleTasks()` and ordering in `sortTasksForDisplay()` so refresh/search/sort paths stay consistent without changing Sheet row order.
- Task memo is stored in the `memo` column at the end of `TASK_HEADERS` to avoid shifting existing task data. It should be visible on task cards and in "派出追蹤" because that view reads the assignee's source task row.
- The "派出追蹤" tab reuses the Kanban board UI but loads `getAssignedTasks(currentUser)`. It should show tasks where `assigner` is the current user and `assignee` is another member, reading from the assignee's sheet so status changes made by the assignee are reflected. Keep the assignee filter frontend-only with All/member options, and never include the current user in that filter.
- Keep "派出追蹤" read-only for task cards/status changes. Do not allow drag-to-status updates or full edit modal opens from that view unless the product decision changes explicitly. Memo is the exception: both the assigner and assignee may update the shared task memo through the memo-only editor.
- Keep operational/diagnostic controls out of the header. Webhook logs live inside the settings modal; do not re-add standalone header buttons for reload or webhook logs unless explicitly requested.
- The settings modal should stay focused on persistent configuration. Do not re-add webhook logs or one-off maintenance actions such as Test API, Sync Hedy, or Create Webhook buttons without a clear user request.
- Header action buttons must use non-wrapping text classes such as `whitespace-nowrap` and `shrink-0`; avoid adding long button labels that cause the header to wrap or split Chinese text vertically.

## GAS / Spreadsheet Rules

- Do not perform expensive Drive scans or schema checks on every read path.
- Read only the needed sheet range, not full `getDataRange()` by default.
- Avoid putting `ensureMemberSchema` on hot read paths unless strictly necessary.
- Cache lightweight lookups within a single execution when possible.
- Be careful with webhook handlers: they should do the minimum work needed and avoid repeated spreadsheet opens.

## Webhook Rules

- Treat Hedy events separately:
  - `todo.exported`: creates a ToDo task.
  - `session.created`: ignore or log only; do not create/update meeting records from it.
  - `session.ended`: stores completed meeting/session information.
  - `session.exported`: stores meeting/session information, not tasks.
  - other events: log or store only if needed.
- Do not infer Hedy todos from `session.exported` unless explicitly requested.
- `todo.exported` payloads are single-item events. Do not assume one webhook contains all todos from a session.
- Use the Hedy todo `id` as a dedupe key when importing tasks.

## Hedy API Sync

- The header "同步會議" action uses one button and a fixed batch size of 5. If no Hedy sessions cursor exists for the user, sync the latest 5 sessions; otherwise use the saved cursor to sync the next older 5 sessions.
- Store Hedy sessions pagination cursor per user in `PropertiesService` so each member can continue historical backfill independently.
- When Hedy reports no more older sessions, clear the saved cursor so a future click starts from the latest sessions again instead of repeating the final page.
- Time-driven Hedy triggers should call `scheduledSyncHedyForConfiguredUser()`, which reads `HEDY_SYNC_USER` and syncs only the latest 5 sessions. Do not use the cursor-based header sync for scheduled automation.
- Do not use account-wide `/todos` for the header sync; it can scan too much historical data. Use `/sessions/{sessionId}/todos` for the selected recent sessions.
- Batch Hedy sync must build a one-time Sheet index for Tasks, Meetings, and Highlights before processing multiple sessions/todos. Avoid calling row-finder helpers inside per-item loops because that causes repeated full-sheet scans.
- Preserve each todo's `sessionId` when importing so `createTaskFromTodoExport_` can link the task to the corresponding meeting via `externalSessionId`.
- Task cards should keep `meetingId` and `meetingTitle` so clicking the meeting link opens the related meeting detail.
- The settings modal "同步 Hedy" may remain a broader/manual sync path, but header sync should stay lightweight.

## Push / Deploy Workflow

- `clasp push` is required to update the Apps Script project source.
- `clasp push` may say `Skipping push.` even when local debugging is still in progress. If the user expects the remote source to match local changes, verify and use `clasp push -f` when appropriate.
- `clasp push` does not deploy a new web app version.
- Only run deploy commands if the user explicitly asks for deployment.
- After major fixes, confirm whether the user is testing:
  - latest project code in the Apps Script editor, or
  - an older deployed web app URL.

## Apps Script Authorization

- When adding code that uses a new authorized service, update `appsscript.json` `oauthScopes` explicitly before testing.
- `UrlFetchApp.fetch()` requires `https://www.googleapis.com/auth/script.external_request`.
- Spreadsheet writes require `https://www.googleapis.com/auth/spreadsheets`; Drive file access requires `https://www.googleapis.com/auth/drive`.
- Adding scopes is not enough by itself. The deploying Google account must manually run a function that actually touches the authorized service so Apps Script can show the consent dialog.
- Do not use `doGet` as the authorization trigger unless it reaches the new service. It usually only renders HtmlService and will not prompt for `UrlFetchApp` authorization.
- For Hedy API authorization, run `authorizeHedyApiAccess()` from the Apps Script editor function dropdown. Refresh the editor if the function is not visible.
- After authorization, if testing the production Web App `/exec` URL, deploy a new Web App version. `clasp push` alone only updates editor source and `/dev` behavior.
- Simple triggers such as `onEdit(e)` cannot reliably call services that need authorization. Use installable triggers for authorized work.

## Verification Checklist

- Before claiming frontend fixes:
  - run a local syntax check on the inline script extracted from `index.html`
  - sanity check `程式碼.js` syntax
- Before claiming webhook fixes:
  - verify event routing logic in `doPost`
  - confirm the intended event type, especially `todo.exported` vs `session.exported`
- Before claiming remote fixes:
  - ensure `clasp push` actually updated the remote project

## Safe Change Strategy

- For risky frontend changes, prefer small increments over bundled UI rewrites.
- If a new version breaks loading, revert toward the last known-good version first, then reintroduce changes one group at a time.
- When performance regresses, inspect initial load path separately from interaction path.
