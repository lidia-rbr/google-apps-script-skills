---
name: google-apps-script-standards
description: Guideline and instructions to develop premium, performance-optimized, and standard-compliant Google Apps Script applications. Trigger this when generating, modifying, or reviewing Google Apps Script code.
---

# Google Apps Script Development Standards

This skill guides you through the creation, modification, and optimization of Google Apps Script (GAS) code, ensuring high performance, clean architecture, and adherence to professional development standards.

## 1. Core Principles & Architecture

- **V8 Runtime Compliance**: Write modern ES6+ JavaScript (let/const, arrow functions, destructuring, template literals, promises).
- **Clean Architecture & Separation of Concerns**: Keep business logic, database (Spreadsheets) interaction, and UI (HTML/CSS) separate.
- **Quota & Limitation Awareness**: Design with GAS execution limits in mind (6-minute limit, URL Fetch quotas, etc.).

### 🏗️ Orchestrator Pattern (Feature Main)
- **One Feature = One Main Orchestrator**: Every major feature must have a single "main" orchestrator function that coordinates the process end-to-end.
- **Strict Function Length Limit (< 60 Lines)**: No function (including the main orchestrator) should exceed **60 lines of code**. This keeps code modular, extremely readable, and easy to test.
- **Single Responsibility Helpers**: The main orchestrator function should delegate tasks to small, specialized helper functions that do one thing and return the result back to the main function.

#### Pattern Example:
```javascript
/**
 * Main orchestrator to process the weekly store sales report.
 * Coordinates fetching, parsing, and emailing.
 */
function main_processWeeklySalesReport() {
  const folderId = "FOLDER_ID_123";
  
  // 1. Get the latest file
  const lastFile = getLastFileInFolder(folderId);
  if (!lastFile) {
    console.warn("No files found to process.");
    return;
  }
  
  // 2. Extract data
  const data = extractSalesData(lastFile);
  
  // 3. Process and format
  const processedData = formatSalesData(data);
  
  // 4. Save to spreadsheet
  updateSalesSheet(processedData);
  
  console.log("Weekly report processed successfully.");
}

/**
 * Helper: Retrieve the latest file in a specific folder.
 * @param {string} folderId 
 * @returns {GoogleAppsScript.Drive.File|null}
 */
function getLastFileInFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  let latestFile = null;
  
  while (files.hasNext()) {
    const file = files.next();
    if (!latestFile || file.getLastUpdated() > latestFile.getLastUpdated()) {
      latestFile = file;
    }
  }
  return latestFile;
}
```

---

## 2. Standard Project File Structure

Every Google Apps Script project must adhere to this structured file organization:

| File Name | Extension | Purpose / Content |
| :--- | :--- | :--- |
| **`config.gs`** | `.gs` | Holds the unique `CONFIG` namespace object with environments (DEV, PROD), versioning, and static IDs. |
| **`webapp.gs`** | `.gs` | Handles the webapp server-side routing (e.g., `doGet()`, `doPost()`) and page builders (e.g., `include()`). |
| **`utils.gs`** | `.gs` | Regroups all independent, reusable utility helper functions (e.g., date formatters, basic math, array helper wrappers). |
| **`Index.html`** | `.html` | Core HTML structural mockup of the Web App. No inline CSS or JavaScript. |
| **`Stylesheet.html`** | `.html` | Main CSS styling enclosed strictly within `<style></style>` tags. |
| **`JavaScript.html`** | `.html` | Client-side JavaScript logic enclosed strictly within `<script></script>` tags. |

---

## 3. Code Quality & Formatting

### JSDoc Documentation
Every function must be documented using JSDoc. Define parameters, types, and return values clearly.
```javascript
/**
 * Fetches user records from the main sheet filtered by department.
 * @param {string} department - The department name to filter by.
 * @returns {Array<Object>} List of employee objects.
 */
function getEmployeesByDepartment(department) {
  // Logic here
}
```

### Parameter Minimization (Clean Code)
- **Pass the Minimum Number of Parameters**: Functions must accept the bare minimum number of arguments. Avoid bloated parameter lists (ideally **2 parameters maximum**).
- **Group Related Arguments**: If multiple arguments are related, group them into a single, cohesive configuration object (Parameter Object pattern).
- **Leverage CONFIG**: Do not pass configuration variables (like Folder IDs, sheet names, etc.) as arguments; functions should retrieve them directly from `CONFIG.ENV`.

```javascript
// Avoid ❌
function generateReport(folderId, templateId, userName, date, format) { ... }

// Recommended ✅ - Pass a single structured object
function generateReport({ userName, date, format }) {
  // Access CONFIG.ENV.TEMPLATE_ID and CONFIG.ENV.DEST_FOLDER_ID directly!
}
```

### Naming & Global Variables Conventions
- **Variables & Functions**: `camelCase` (e.g., `activeSheet`, `fetchData()`).
- **Internal / Private Helper Functions**: Prefix with an underscore `_` (e.g., `_formatDate(date)`).

- **Global Configuration Object (`config.gs`)**: 
  All global configurations, metadata, environment toggles, and sheet IDs must be grouped within a single namespace object named **`CONFIG`** located inside a dedicated file named **`config.gs`**.
  
  The configuration object must support environment separation (`DEV` vs `PROD`) for key assets like template IDs and destination folders:
  
  ```javascript
  // In config.gs ✅
  const CONFIG = {
    env: 'PROD', // 'DEV' or 'PROD'
    version: '1.26.87',
    author: 'Software Factory',
    DEV: {
      TEMPLATE_ID: 'xxx_dev_template_id_xxx',
      DEST_FOLDER_ID: 'xxx_dev_folder_id_xxx'
    },
    PROD: {
      TEMPLATE_ID: 'xxx_prod_template_id_xxx',
      DEST_FOLDER_ID: 'xxx_prod_folder_id_xxx'
    }
  };

  // Bind active environment config to CONFIG.ENV for supreme readability
  CONFIG.ENV = CONFIG[CONFIG.env];
  ```

  Use this object inside your application code to dynamically fetch environment-specific variables:
  ```javascript
  // Example usage:
  const templateFile = DriveApp.getFileById(CONFIG.ENV.TEMPLATE_ID);
  ```

  Never declare unstructured variables in the global scope to avoid name collisions.

### Visual Code Style & Iteration
- **Prefer Array `.forEach()` over standard `for` loops**: To make the code highly visual, descriptive, and clean, prefer using the `.forEach()` method (or `.map()`, `.filter()`) for iterating over lists.
  ```javascript
  // Recommended visual style ✅
  storeList.forEach(store => {
    store.ean = "3700000000000";
  });
  ```


---

## 4. Spreadsheet Performance Standards (Critical)

> [!IMPORTANT]
> The single most common cause of slow Apps Script execution is reading/writing to spreadsheets cell-by-cell.

### Anti-Pattern ❌
```javascript
// Extremely slow and API heavy!
for (let i = 1; i <= lastRow; i++) {
  let val = sheet.getRange(i, 1).getValue();
  if (val === 'Active') {
    sheet.getRange(i, 2).setValue('Processed');
  }
}
```

### Standard Pattern (Batch Operations) ✅
```javascript
// Fast: Single read, single write, and visual iteration
const range = sheet.getRange(1, 1, lastRow, 2);
const values = range.getValues();

values.forEach(row => {
  if (row[0] === 'Active') {
    row[1] = 'Processed';
  }
});
range.setValues(values);
```

---

### Robust Spreadsheet Mapping (Column-Agnostic) ✅

To prevent the code from breaking when users reorder columns in a spreadsheet, never hardcode column indices (e.g., `row[0]`, `row[1]`). Instead, define a mapping configuration object and resolve column indices dynamically by matching headers.

#### 1. Configuration Mapping Object (e.g., inside `config.gs`)
```javascript
const CONFIG_EXECUTIONS = {
  sheetName: "Executions",
  headerRowIndex: 1,
  mapping: {
    "executionId": "Id execution",
    "status": "Statut",
    "module": "module",
    "matrixName": "matrice",
    "date": "date",
    "time": "heure",
    "link": "link"
  }
};
```

#### 2. Standard Mapping Helper (e.g., inside `utils.gs`)
```javascript
/**
 * Loads and maps spreadsheet data into an array of objects dynamically based on header names.
 * Ensures the code is robust against column reordering.
 * 
 * @param {Object} config - Config object containing sheetName, headerRowIndex, and mapping.
 * @returns {Array<Object>} List of mapped objects.
 */
function loadAndMapSheetData(config) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
  if (!sheet) throw new Error(`Sheet "${config.sheetName}" not found.`);
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= config.headerRowIndex) return [];
  
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = values[config.headerRowIndex - 1];
  
  // Resolve column index for each mapping key
  const columnIndexMap = {};
  Object.keys(config.mapping).forEach(key => {
    const headerName = config.mapping[key];
    const index = headers.indexOf(headerName);
    if (index === -1) {
      throw new Error(`Required header "${headerName}" not found in sheet "${config.sheetName}".`);
    }
    columnIndexMap[key] = index;
  });
  
  // Map rows to objects starting below headers
  const mappedData = [];
  for (let i = config.headerRowIndex; i < values.length; i++) {
    const row = values[i];
    const item = {};
    Object.keys(columnIndexMap).forEach(key => {
      item[key] = row[columnIndexMap[key]];
    });
    mappedData.push(item);
  }
  
  return mappedData;
}
```

---

## 5. Error Handling & Logging

- Use `try...catch` blocks at the entry points of your script (triggers, API web apps, menu items) to catch runtime errors gracefully.
- Use `console.log()`, `console.warn()`, and `console.error()` for rich cloud logging and debugging rather than outdated `Logger.log()`.
- Return structured JSON responses with `success: true/false` and an `error` key for custom APIs or HTML interactions.

```javascript
function doGet(e) {
  try {
    const data = processRequest(e);
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Critical doGet Failure: " + error.toString(), error.stack);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
```

---

## 6. Web App & UI Standards (HTML Service)

- **Strict 3-File Separation**: Every HTML interface must be cleanly split into exactly 3 dedicated files to keep code highly maintainable:
  1. `Index.html` - Contains ONLY the pure HTML structure.
  2. `Stylesheet.html` - Contains ONLY the CSS style enclosed in `<style>` tags.
  3. `JavaScript.html` - Contains ONLY the client-side JavaScript enclosed in `<script>` tags.

- **Web App Logic Placement (`webapp.gs`)**:
  All server-side routing entrypoints (`doGet`, `doPost`), template injection helpers (like `include()`), and utility functions that are strictly serving the Web App must be placed in a dedicated file named **`webapp.gs`**.
  
  ```javascript
  // In webapp.gs ✅
  
  /**
   * Main entrypoint for serving the Web App HTML.
   */
  function doGet(e) {
    return HtmlService.createTemplateFromFile('Index')
                      .evaluate()
                      .setTitle('My Apps Script App')
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  /**
   * Helper to include external HTML files (CSS or JS) dynamically.
   */
  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
  ```

- **Integration Structure**:
  Inside your main `Index.html`, inject the stylesheet in the `<head>` and the javascript at the bottom of the `<body>` using server-side template tags:
  ```html
  <!DOCTYPE html>
  <html>
    <head>
      <base target="_top">
      <!-- Import CSS -->
      <?!= include('Stylesheet'); ?>
    </head>
    <body>
      <div class="container">
        <h1>My Apps Script App</h1>
      </div>

      <!-- Import JS Client-Side -->
      <?!= include('JavaScript'); ?>
    </body>
  </html>
  ```

- **Asynchronous Client-Server Communication**: Always use `google.script.run` inside `JavaScript.html` with `.withSuccessHandler()` and `.withFailureHandler()`. Never block the UI thread.

