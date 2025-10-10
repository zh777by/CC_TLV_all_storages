# Google Sheets Monthly Summary & Pivot Builder

This Google Apps Script generates **daily and monthly summary tables** and an **annual pivot dashboard** for order and item statistics across multiple monthly sheets.
It works on sheets named in the `MM/YYYY` format (e.g., `09/2025`) and creates both per-month visual summaries and a consolidated “PIVOT” overview.

---

## ✳️ Features

* **Automatic monthly summary table**

  * Detects daily data in the active monthly sheet.
  * Groups and totals **Orders** and **Items** per day.
  * Calculates monthly totals and averages.
  * Generates a colored summary table below existing data.
  * Builds a **column chart** titled “ORDERS and ITEMS — MM/YYYY”.
* **Global PIVOT summary**

  * Scans all monthly sheets (`MM/YYYY`).
  * Aggregates totals and averages for each month.
  * Creates or updates a unified summary sheet named **`PIVOT`**.
  * Adds a **bar chart** “ORDERS and ITEMS — all months”.
* **Smart detection**

  * Automatically detects the month and year based on sheet name and date columns.
  * Recognizes Hebrew column headers (`נפתחה`, `מספר פריטים`).
* **Polished formatting**

  * Alternating row colors.
  * Bold headers and totals.
  * Compact number formatting for readability.
* **Backwards compatibility**

  * Recognizes and reuses old summary sheet names like `'סיכום חודשי'`.

---

## 🧩 Sheet Structure Requirements

Each monthly sheet must:

* Be named in **`MM/YYYY`** format (e.g. `09/2025`).
* Contain columns labeled:

  * **`נפתחה`** → *Date Opened*
  * **`מספר פריטים`** → *Number of Items*
* Contain rows with daily order entries (including numeric item counts).

---

## 🧮 Summary Table Logic

When you run **`buildMonthlySummaryTable()`**, the script:

1. Detects all unique dates for the current month.
2. Aggregates data per date:

   * **Orders** = number of rows or highest “A” number.
   * **Items** = sum of all item counts for that day.
3. Generates a summary table:

   ```
   ORDERS | ITEMS | DATE | DAYS
   ```
4. Adds totals and daily averages:

   * Total orders, total items, total days.
   * Average orders/day, average items/day.
5. Applies color formatting:

   * Header row → `HEADER_BG`
   * Alternating rows → `ODD_BG`, `EVEN_BG`
6. Adds a visual chart “ORDERS and ITEMS — MM/YYYY”.

---

## 📊 Global Pivot Table

Running **`buildAllMonthsSummary()`** will:

1. Scan all sheets named `MM/YYYY`.
2. Read their summary tables (ORDERS, ITEMS, DATE, DAYS).
3. Collect totals and averages per month.
4. Write results into a single sheet:

   ```
   MONTH | DAYS | ORDERS (total) | ITEMS (total) | ORDERS (avg/day) | ITEMS (avg/day)
   ```
5. Apply alternating colors and numeric formats.
6. Create a bar chart “ORDERS and ITEMS — all months”.

If a sheet named `PIVOT` doesn’t exist, it will:

* Rename an old sheet `'סיכום חודשי'` if found, or
* Create a new sheet named `PIVOT`.

---

## 🎨 Colors & Styles

| Element           | Variable      | Default Value | Description             |
| ----------------- | ------------- | ------------- | ----------------------- |
| Header background | `HEADER_BG`   | `#63D297`     | Green header background |
| Odd rows          | `ODD_BG`      | `#E8F8F2`     | Light teal background   |
| Even rows         | `EVEN_BG`     | `#F4FBF8`     | Pale green background   |
| Header font color | `HEADER_FONT` | `#000000`     | Black text for headers  |

Charts use:

* **Orders** → Blue (`#007BFF`)
* **Items** → Red (`#D9534F`)
* **Trendline** → Soft pink (`#F5A5A5`)

---

## ⚙️ Public Functions

### 1. `buildMonthlySummaryTable()`

Creates or updates the summary table and chart for the **active monthly sheet**.
Must be run from a sheet named `MM/YYYY`.

### 2. `buildAllMonthsSummary()`

Generates or updates the **PIVOT** sheet with aggregated data across all monthly sheets.

### 3. `buildJuneSummaryTable()`

Alias for backward compatibility with old buttons (calls `buildMonthlySummaryTable()`).

---

## 🛠️ Helper Functions (Internal)

* `_monthNumFromSheetName(name)` – Validates `MM/YYYY` format.
* `_findCols(sheet)` – Detects key columns (`נפתחה`, `מספר פריטים`).
* `_detectYearMonth()` – Determines correct month/year from data.
* `_findMonthlyTableRegion()` – Locates existing monthly summary tables.
* `_removeExistingChartsByTitlePrefix(sheet, prefix)` – Avoids duplicate charts.
* `_ymdFromValueOrDisplay()` – Extracts Y/M/D from date values or strings.
* `_toNumber()`, `_pad2()` – Utility functions for parsing and formatting.

---

## 🚀 Usage Instructions

1. Open your spreadsheet.
2. Go to **Extensions → Apps Script**.
3. Paste this code into the editor.
4. Save and authorize the script.
5. Open a monthly sheet (e.g. `09/2025`) and run:

   * `buildMonthlySummaryTable()` — for that month.
   * `buildAllMonthsSummary()` — to refresh the PIVOT sheet.

---

## 🧾 Example Output

**Monthly summary (on sheet `09/2025`):**

| ORDERS    | ITEMS      | DATE                      | DAYS       |
| --------- | ---------- | ------------------------- | ---------- |
| 45        | 280        | 01/09/2025                | 1          |
| 38        | 240        | 02/09/2025                | 2          |
| ...       | ...        | ...                       | ...        |
| **280**   | **1520**   | **09/2025 (total)**       | Total days |
| **46.67** | **253.33** | **09/2025 (average/day)** | 6          |

**Global Pivot:**

| MONTH   | DAYS | ORDERS (total) | ITEMS (total) | ORDERS (avg/day) | ITEMS (avg/day) |
| ------- | ---- | -------------- | ------------- | ---------------- | --------------- |
| 07/2025 | 25   | 960            | 5400          | 38.4             | 216             |
| 08/2025 | 27   | 1045           | 6120          | 38.7             | 226.7           |
| 09/2025 | 30   | 1250           | 7200          | 41.6             | 240             |
| ...     | ...  | ...            | ...           | ...              | ...             |

---

## 🔁 Backward Compatibility

* Old sheets titled **‘סיכום חודשי’** are automatically renamed to `PIVOT`.
* Legacy trigger buttons calling `buildJuneSummaryTable()` are still supported.

---

## ⚠️ Common Errors

| Message                                      | Meaning                                   | Fix                                                    |
| -------------------------------------------- | ----------------------------------------- | ------------------------------------------------------ |
| `Откройте месячный лист…`                    | Active sheet name not in `MM/YYYY` format | Rename the sheet correctly                             |
| `Не нашёл заголовки "נפתחה" и "מספר פריטים"` | Missing headers                           | Ensure correct Hebrew column names                     |
| `Нет строк за MM/YYYY`                       | No data found for that month              | Verify that daily rows contain valid dates and numbers |

---

## 🗂️ Files and Sheet Naming

| Type          | Example   | Purpose                       |
| ------------- | --------- | ----------------------------- |
| Monthly sheet | `09/2025` | Contains daily order records  |
| Summary sheet | `PIVOT`   | Consolidated monthly overview |

---

## 📈 Charts

* **Per month:**
  Title → `ORDERS and ITEMS — MM/YYYY`
  Type → Column chart

* **Global summary (PIVOT):**
  Title → `ORDERS and ITEMS — all months`
  Type → Column chart with color-coded bars

---

## 🏷️ Version

**v1.0 — Initial release**

* Added monthly summaries and charts
* Added PIVOT summary across months
* Added backward compatibility and styling
