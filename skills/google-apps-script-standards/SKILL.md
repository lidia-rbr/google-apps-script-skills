---
name: google-apps-script-standards
description: Write, refactor, and review Google Apps Script for Google Workspace. Use when building Sheets or Docs automation, scheduled jobs, triggers, or HTML Service web apps, especially to batch API calls, map Sheets columns by header, or reduce google.script.run round trips.
---

# Google Apps Script Development Standards

Use V8 JavaScript. Separate orchestration, Workspace API calls, data transformations, and UI code. Preserve existing public entry points and behavior when refactoring.

## Config.js and entry points

- Keep one global `CONFIG` object in `Config.js` with a stable `CONFIG.APP_NAME`, `CONFIG.ENV`, and the spreadsheet and document IDs for each environment. Include tab names, header mappings, folder IDs, and batch limits when the feature needs them. With the Apps Script editor, use the equivalent `Config.gs` file.
- Default new samples to `DEV`. Validate the selected environment and required IDs before opening files or writing data. Keep credentials outside source and never send `CONFIG`, credentials, or internal file IDs to the browser as part of the page payload.
- Give each feature a `Main_` orchestrator with a capital `M` and underscore-separated words, such as `Main_Process_Orders`. It should read its configuration, fetch inputs, call focused helpers, and apply one planned set of changes.
- Keep platform entry points such as `doGet`, `doPost`, `onOpen`, and `onEdit` under their required names. Use a thin wrapper when the entry point delegates to a feature's `Main`. Preserve the names of installed trigger handlers, menu callbacks, and HTML-callable functions.
- End internal utility names with `_`, such as `requireColumnIndex0Based_`, to keep them out of the editor's Run dropdown and unavailable to `google.script.run`. Use top-level public functions for editor entry points, triggers, menu actions, and browser calls.

## NASA/JPL-inspired JavaScript style

- Use `let` for bindings that change and `const` for bindings that stay fixed. Never use `var`.
- Use short inline comments in each `Main_` orchestrator to mark the major workflow steps and explain decisions that affect behavior. Use function-level JSDoc to describe what callers can expect.
- Keep control flow direct. Prefer early validation and guard clauses. Give loops an explicit bound, avoid recursion, and advance large jobs across bounded executions.
- Keep functions focused and aim for at most 60 lines. Extract a named helper at a meaningful read, validate, transform, or write boundary. Keep local variables in the narrowest scope.
- Validate assumptions that matter, including required headers, output shape, and selected environment. Check relevant service results. Log the operation and a useful error, then propagate unexpected failures so the Apps Script Executions view shows them. Keep secrets and full records out of logs and internal error details out of browser responses.

## JSDoc and editor completion

- Document public functions and typed helpers with `@param`, `@returns`, and `@throws` when callers need to know about a failure. Describe units, index bases, and side effects where relevant. Use `@typedef` for shared record shapes and `@type` for API payloads or values the editor cannot infer.
- Install `@types/google-apps-script` as a local development dependency with `npm install --save-dev @types/google-apps-script`. Enable `// @ts-check` in local JavaScript files when you want type diagnostics in addition to IntelliSense. If your project limits ambient types, include `google-apps-script` in `compilerOptions.types`.
- Use `SpreadsheetApp.Sheet` as a local JSDoc alias for a native sheet tab. Derive it from `SpreadsheetApp.getActiveSheet` so the editor completes Sheet methods. The alias is only for type annotations. At runtime, obtain the tab from `SpreadsheetApp` or a spreadsheet object. Check for `null` when using `getSheetByName(name)`. Use `GoogleAppsScript.Sheets.Schema.ValueRange` for a Sheets API values payload.

Example for an existing integration that takes a native spreadsheet object. Use the advanced Sheets service for batch data operations.

```javascript
/** @typedef {ReturnType<typeof SpreadsheetApp.getActiveSheet>} SpreadsheetApp.Sheet */

/**
 * Resolves a required tab in an already-open spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Open spreadsheet.
 * @param {string} sheetName Exact tab name.
 * @returns {SpreadsheetApp.Sheet} The tab, with Sheet methods available to callers.
 * @throws {Error} If the tab does not exist.
 */
function requireSheet_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Required sheet is missing: ${sheetName}`);
  return sheet;
}

/**
 * Returns the occupied range of a native tab.
 * @param {SpreadsheetApp.Sheet} sheet Tab to inspect.
 * @returns {ReturnType<SpreadsheetApp.Sheet['getDataRange']>} Data range.
 */
function getDataRange_(sheet) {
  return sheet.getDataRange();
}
```

## Workspace APIs and batch operations

- Use the Google Workspace APIs for data operations. In Apps Script, use the corresponding advanced service (`Sheets`, `Docs`, or `Drive`) when available. Use REST when an endpoint lacks an advanced wrapper. Enable the required services and scopes in `appsscript.json` and enable the APIs in the associated Cloud project when required.
- Use built-in Apps Script facilities for platform-specific tasks such as HTML rendering, trigger setup, configuration storage, and locking. Keep API calls and Apps Script service calls outside `for`, `while`, and array iteration callbacks, including calls hidden inside helpers. Iterate only over data already fetched into memory.
- Read all necessary ranges together with `Sheets.Spreadsheets.Values.batchGet` when reading multiple ranges in one spreadsheet. Build changes in memory, then use `Sheets.Spreadsheets.Values.batchUpdate` for values or `Sheets.Spreadsheets.batchUpdate` for grid operations such as formatting or inserting rows. Use a Docs API batch update for multiple document edits.
- Write only the intended ranges and fields. Use `RAW` for literal values and `USER_ENTERED` when spreadsheet parsing is intentional. Preserve formulas, headers, and untouched rows. Skip empty update requests.
- For large workloads, cap the amount of data fetched and written per execution. Persist the cursor after a successful write and continue in a later execution. Make retries safe by using stable record IDs and deterministic writes. Guard shared state when runs can overlap, and stop visibly on repeated failures.

## Heavy processes and quota recovery

Run one bounded batch per execution. Set batch size, runtime headroom, and retry limits in `Config.js`. Store a stable job ID and the last confirmed cursor in `PropertiesService` or a job table. Advance the cursor only after confirming the batch write. Use an installable time-driven trigger to resume and release any lock before sleeping or deferring. A hard execution timeout can stop the script before `catch` or `finally` runs, so keep the checkpoint current.

Catch failures around the smallest operation safe to repeat, such as a batch read or a write with a stable target and deterministic values. A timeout can leave a write's outcome unknown. Reconcile that batch before retrying it if repeating it could create duplicate rows, files, or emails.

### Recognize common quota errors

For REST calls, prefer the HTTP status and the API's documented reason codes. If a service exposes only an exception message, use narrow, case-insensitive fragments and leave unknown errors unclassified. Google's quota messages are examples and can vary by service and locale.

| Signal | Action |
| --- | --- |
| HTTP `429`, `rateLimitExceeded`, `userRateLimitExceeded`, or `in a short time:` | Sleep briefly with jitter and retry the safe operation if the execution has time left. |
| Apps Script exception mentioning `per second` in a script invocation limit or `running simultaneously` | Defer the batch and reduce trigger frequency or overlapping runs. A rejected invocation may never enter your code. |
| `dailyLimitExceeded`, `invoked too many times:`, or `computer time for one day` | Keep the cursor and resume after the applicable quota resets. Apps Script daily quotas generally reset 24 hours after the first request. Check the service's reset policy. |
| `Limit exceeded:` for operation size or `Exceeded maximum execution time` | Reduce the payload or next batch size. Preserve the checkpoint and resume in a fresh execution if the run timed out. |
| Unclassified failure, including a generic HTTP `403` | Propagate it for diagnosis. A `403` may be a permissions failure. |

### Retry a safe operation once in the current run

Normalize a REST failure to an error carrying `httpStatus`, `reasons` from the response body, and `retryAfterMs` if a `Retry-After` header is present. Those fields belong to your REST adapter. Inside Apps Script, use the documented message fragments as the fallback when the advanced service exposes no structured details.

```javascript
/**
 * @typedef {Object} QuotaFailure
 * @property {string} [message] Apps Script exception message.
 * @property {number} [httpStatus] HTTP status supplied by the REST adapter.
 * @property {string[]} [reasons] Structured Google API error reasons.
 * @property {number} [retryAfterMs] Server-requested delay in milliseconds.
 */

/**
 * Classifies a documented quota signal and retains a server-provided delay.
 * @param {unknown} error Thrown Apps Script exception or normalized REST failure.
 * @returns {{kind: string, retryAfterMs: number}} Quota category and minimum wait.
 */
function classifyQuotaError_(error) {
  const details = error && typeof error === 'object'
    ? /** @type {QuotaFailure} */ (error)
    : { message: typeof error === 'string' ? error : '' };
  const { message = '', httpStatus, reasons = [], retryAfterMs = 0 } = details;
  if (reasons.includes('dailyLimitExceeded')) return { kind: 'daily', retryAfterMs };
  if (httpStatus === 429 || reasons.includes('rateLimitExceeded') || reasons.includes('userRateLimitExceeded')) {
    return { kind: 'throttle', retryAfterMs };
  }
  if (httpStatus !== undefined) return { kind: 'other', retryAfterMs };
  if (/\bin a short time\s*:/i.test(message)) return { kind: 'throttle', retryAfterMs };
  if (/\binvoked too many times\s*:|computer time for one day/i.test(message)) return { kind: 'daily', retryAfterMs };
  if (/\bscript invoked .*\bper second\b|\brunning simultaneously\b/i.test(message)) {
    return { kind: 'concurrency', retryAfterMs };
  }
  if (/\bexceeded maximum execution time\b/i.test(message)) return { kind: 'runtime', retryAfterMs };
  if (/^\s*(?:exception:\s*)?limit exceeded\s*:/i.test(message)) return { kind: 'hard_limit', retryAfterMs };
  return { kind: 'other', retryAfterMs };
}

/**
 * Retries a replay-safe operation once when a short throttle allows it.
 * @template T
 * @param {() => T} operation Replay-safe batch operation.
 * @param {number} latestRetryStartMs Latest epoch time to retry, with time reserved for the request and checkpoint.
 * @param {number} maxSleepMs Maximum configured in-run delay in milliseconds.
 * @returns {T} Result from the original call or its retry.
 * @throws {Error} If the error is ineligible for retry or the retry fails.
 */
function retryOnceOnQuota_(operation, latestRetryStartMs, maxSleepMs) {
  try {
    return operation();
  } catch (error) {
    const { kind, retryAfterMs } = classifyQuotaError_(error);
    if (kind !== 'throttle') throw error;

    // Wait long enough to respect the server without using up this execution.
    const delayMs = Math.ceil(Math.max(1000 + Math.random() * 1000, retryAfterMs));
    if (!Number.isFinite(delayMs) || !Number.isFinite(latestRetryStartMs) || !Number.isFinite(maxSleepMs)
        || delayMs > Math.min(maxSleepMs, 300000) || Date.now() + delayMs >= latestRetryStartMs) {
      throw error;
    }
    Utilities.sleep(delayMs);
    if (Date.now() >= latestRetryStartMs) throw error;
    return operation();
  }
}
```

Call the helper outside your data loops. If it throws again, the `Main_` orchestrator classifies the latest failure and keeps the cursor for a deferred or failed batch. For throttled batches resumed by a trigger, persist the retry count and next eligible run time with capped exponential backoff and jitter. Stop when the configured attempt limit is reached. Log one `[APP_NAME]` summary with the job ID, outcome, quota category, retry count, and next eligible run time when applicable.

[Apps Script quota exceptions](https://developers.google.com/apps-script/guides/services/quotas#exception_messages) · [Apps Script timeouts](https://developers.google.com/apps-script/guides/sheets/functions#return_values) · [Sheets API backoff](https://developers.google.com/workspace/sheets/api/limits#resolve_time-based_quota_errors) · [Drive API error reasons](https://developers.google.com/workspace/drive/api/guides/handle-errors) · [Utilities.sleep](https://developers.google.com/apps-script/reference/utilities/utilities#sleep(milliseconds))

## Sheets headers and index bases

- Resolve columns by header name before processing rows. Validate required headers and reject duplicates or ambiguities produced by normalization. Treat missing rows, blank rows, and ragged API results explicitly.
- Never address a data column with a numeric literal such as `row[3]`. Use `row[statusColumnIndex0Based]` or map the rows to named fields. Use a separate mapping for each tab with a different schema.
- Name sheet indices by base, such as `rowIndex1Based`, `columnIndex0Based`, and `endRowIndex0BasedExclusive`. A1 row numbers are one-based. Sheets API `GridRange` indices are zero-based, with exclusive end indices. Convert at the API boundary and account for the starting row and column when a fetched range starts away from `A1`.

### Example: mark active orders in a single write request

Use a small DEV sheet with `Status` and `Processing Status` headers in row 1. Keep `Processing Status` as a writable status column. Enable the advanced Sheets service before running `Main_Process_Orders`.

```javascript
// Config.js
const CONFIG = {
  APP_NAME: 'orders-automation',
  ENV: 'DEV',
  DEV: { ORDERS_SPREADSHEET_ID: 'DEV_SHEET_ID', REPORT_DOCUMENT_ID: 'DEV_DOC_ID' },
  PROD: { ORDERS_SPREADSHEET_ID: 'PROD_SHEET_ID', REPORT_DOCUMENT_ID: 'PROD_DOC_ID' },
  ORDERS: {
    SHEET_NAME: 'Orders',
    HEADERS: { STATUS: 'Status', PROCESSING_STATUS: 'Processing Status' },
    MAX_DATA_ROWS: 500,
  },
};
```

```javascript
/**
 * Runs the configured order batch and records its outcome in Cloud Logging.
 * @returns {number} Number of cells updated.
 * @throws {Error} If the configuration, sheet schema, or API update is invalid.
 */
function Main_Process_Orders() {
  const startedAtMs = Date.now();
  try {
    // Read and validate the sheet, then batch any pending updates.
    const updatedCells = processOrders_();
    // Record whether the run changed any cells.
    logOrdersRun_({
      outcome: updatedCells === 0 ? 'skipped' : 'completed',
      durationMs: Date.now() - startedAtMs,
      updatedCells,
    });
    return updatedCells;
  } catch (error) {
    // Record safe diagnostic details for Logs Explorer.
    logOrdersRun_({
      outcome: 'failed',
      durationMs: Date.now() - startedAtMs,
      errorType: error instanceof Error ? error.name : 'UnknownError',
    });
    // Rethrow so Apps Script marks the execution as failed.
    throw error;
  }
}

/**
 * Marks active orders with a single Sheets API batch write.
 * @returns {number} Number of cells updated, or zero if no write is needed.
 */
function processOrders_() {
  const environment = CONFIG.ENV;
  if (environment !== 'DEV' && environment !== 'PROD') {
    throw new Error('Set CONFIG.ENV to DEV or PROD.');
  }
  const spreadsheetId = CONFIG[environment].ORDERS_SPREADSHEET_ID;
  if (!spreadsheetId) throw new Error('Configure the orders spreadsheet ID.');
  if (typeof Sheets === 'undefined') throw new Error('Enable the advanced Sheets service.');

  const { SHEET_NAME, HEADERS, MAX_DATA_ROWS } = CONFIG.ORDERS;
  const sheetA1 = "'" + SHEET_NAME.replace(/'/g, "''") + "'";
  const response = Sheets.Spreadsheets.Values.get(spreadsheetId, sheetA1);
  const [headers, ...rows] = response.values ?? [];
  if (rows.length > MAX_DATA_ROWS) {
    throw new Error('The orders sheet exceeds the demo limit. Use a bounded read and a checkpoint.');
  }

  const statusColumnIndex0Based = requireColumnIndex0Based_(headers, HEADERS.STATUS);
  const outputColumnIndex0Based = requireColumnIndex0Based_(headers, HEADERS.PROCESSING_STATUS);
  const outputColumnA1 = toColumnA1_(outputColumnIndex0Based);
  const firstDataRowIndex1Based = 2;
  /** @type {GoogleAppsScript.Sheets.Schema.ValueRange[]} */
  const updates = [];

  rows.forEach((row, rowIndex0Based) => {
    if (row[statusColumnIndex0Based] !== 'Active' || row[outputColumnIndex0Based] === 'Processed') return;
    const rowIndex1Based = firstDataRowIndex1Based + rowIndex0Based;
    updates.push({ range: `${sheetA1}!${outputColumnA1}${rowIndex1Based}`, values: [['Processed']] });
  });

  if (updates.length === 0) return 0;
  const result = Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'RAW', data: updates }, spreadsheetId);
  if (result.totalUpdatedCells !== updates.length) {
    throw new Error('Unexpected update count. Check the destination before retrying.');
  }
  return result.totalUpdatedCells;
}

/**
 * Returns the unique column index after trimming header whitespace.
 * @param {unknown[] | undefined} headers Header row returned by the Sheets API.
 * @param {string} requiredHeader Expected header text.
 * @returns {number} Zero-based column index within the fetched range.
 * @throws {Error} If the header is missing or duplicated after trimming.
 */
function requireColumnIndex0Based_(headers, requiredHeader) {
  const normalizedHeaders = (headers ?? []).map((header) => String(header ?? '').trim());
  const columnIndex0Based = normalizedHeaders.indexOf(requiredHeader);
  if (columnIndex0Based === -1) throw new Error(`Missing header: ${requiredHeader}`);
  if (columnIndex0Based !== normalizedHeaders.lastIndexOf(requiredHeader)) {
    throw new Error(`Duplicate header: ${requiredHeader}`);
  }
  return columnIndex0Based;
}

/**
 * Converts a spreadsheet column index to an A1 column label.
 * @param {number} columnIndex0Based Zero-based column index.
 * @returns {string} A1 column letters, such as A or AA.
 */
function toColumnA1_(columnIndex0Based) {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let columnIndex1Based = columnIndex0Based + 1;
  let columnA1 = '';
  while (columnIndex1Based > 0) {
    const digitIndex0Based = (columnIndex1Based - 1) % alphabet.length;
    columnA1 = alphabet.charAt(digitIndex0Based) + columnA1;
    columnIndex1Based = Math.floor((columnIndex1Based - 1) / alphabet.length);
  }
  return columnA1;
}

/**
 * @typedef {Object} OrdersRunLog
 * @property {'completed' | 'skipped' | 'failed'} outcome Run outcome.
 * @property {number} durationMs Elapsed time in milliseconds.
 * @property {number} [updatedCells] Confirmed count for a completed or skipped run.
 * @property {string} [errorType] Exception type for a failed run.
 */

/**
 * Writes an order processing summary to the linked Cloud project.
 * @param {OrdersRunLog} summary Outcome and safe diagnostic fields.
 * @returns {void}
 */
function logOrdersRun_(summary) {
  Logger.log({
    message: `[${CONFIG.APP_NAME}] Main_Process_Orders ${summary.outcome}`,
    appName: CONFIG.APP_NAME,
    operation: 'Main_Process_Orders',
    env: CONFIG.ENV,
    ...summary,
  });
}
```

## Cloud Logging

- Assume the script is linked to a standard Google Cloud project. Use `Logger.log({...})` with a `message` and JSON-serializable fields, so Logs Explorer can filter `jsonPayload`. Start each message with `[APP_NAME]`, taking the name from `CONFIG.APP_NAME`. Keep the app name stable across environments and log the environment separately.
- Emit one structured summary per run or batch with the operation, outcome, duration in milliseconds, and relevant counts. Add a run ID when you need to correlate multiple entries. Keep logging outside loops and avoid logging records, credentials, document contents, or personal data.
- Log failures once at the entry point with safe context. Rethrow unexpected errors from jobs and triggers so Executions and Cloud Error Reporting can record the failure. At browser-facing boundaries, expose a user-safe error and keep diagnostic details server-side.

```text
resource.type="app_script_function"
jsonPayload.appName="orders-automation"
jsonPayload.operation="Main_Process_Orders"
```

Add `jsonPayload.outcome="failed"` to find structured failure summaries. The app name is a message prefix and a payload field. Apps Script supplies the Cloud Logging `logName`.

## HTML Service web apps

- Keep routing and template evaluation in `webapp.gs` (or `webapp.js` with clasp). For larger UIs, use `Index.html` for markup and split out `Stylesheet.html` and `JavaScript.html` when it improves maintainability.
- In `doGet`, assign a small, authorized first-render payload to the template, then call `evaluate()`. Use `<?= ... ?>` for contextually escaped values. Reserve `<?!= ... ?>` for trusted static includes.
- If the complete first-view data is small and quick to fetch, render it through the template. If fetching it would delay the initial page, render the shell and a minimal payload, then request the remaining first-view data once after the page loads. Paginate or defer data for later views.
- Group related reads and writes in one `google.script.run` call per user action where practical. Use a public function such as `Main_Load_Details` and attach both success and failure handlers. Update the UI from the response. Avoid a request per control, row, or render.
- Enforce access and validate submitted data on the server. Templates and `google.script.run` responses both deliver data to the browser. Send only the fields the current user is authorized to see. Keep the default framing behavior unless embedding is an explicit requirement with appropriate protection.

## Dos and don'ts

| Do | Avoid |
| --- | --- |
| Resolve `Status` from the header row before processing data. | Read `row[3]` because the status happens to be in column D. |
| Read in bulk, build a write plan, and send one batch per execution. | Call an API, `SpreadsheetApp`, or another Apps Script service inside a loop or its helper. |
| Use `Main_Process_Orders` and a private helper such as `buildUpdates_`. | Use a private helper as a trigger, menu callback, or browser endpoint. |
| Render a small initial payload on the server and load heavier data in one deferred call. | Send a separate `google.script.run` call for each table row. |
| Log one searchable summary with `CONFIG.APP_NAME`, outcome, duration, and counts. | Log a line per row or include raw data in a log entry. |
| Start in DEV and verify formulas outside the target ranges. | Run a new batch on production first or write an entire table back from value-only reads. |

## Before delivering a change

1. Check the intended entry point, its execution identity, required APIs and permissions, and the active environment.
2. Test on a DEV copy with reordered, missing, and duplicate headers, an empty or short data set, and formulas outside the output range. Run the same operation twice and check for duplicate effects.
3. Confirm successful and failed runs can be found in Logs Explorer with the app and operation filters.
4. Check that all service calls sit outside iteration and the batch touches only intended ranges. For a web app, verify initial rendering, deferred loading, and the failure state as an intended user.

## References

- [NASA/JPL Power of Ten](https://spinroot.com/gerard/pdf/P10.pdf)
- [Apps Script advanced Sheets service](https://developers.google.com/apps-script/advanced/sheets)
- [Sheets API batch requests](https://developers.google.com/workspace/sheets/api/guides/batch)
- [Apps Script HTML templates](https://developers.google.com/apps-script/guides/html/templates)
- [Apps Script client-server calls](https://developers.google.com/apps-script/guides/html/communication)
- [Apps Script development type definitions](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/google-apps-script)
- [Apps Script Logger](https://developers.google.com/apps-script/reference/base/logger)
- [Cloud Logging query language](https://docs.cloud.google.com/logging/docs/view/logging-query-language)
