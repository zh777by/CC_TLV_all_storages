# 📦 Google Sheets Inventory Aggregator — “Total now all”

This Google Apps Script builds a consolidated inventory sheet named **`Total now all`** by reading the **Total Now** (pcs) values from four source spreadsheets (BISQUE, YU, BY, HH).
It normalizes item names, merges duplicates case-insensitively, pulls **SKU CC#** from YU/BY/HH (or falls back to BISQUE’s ITEM), and applies formatting and thresholds.

---

## ✳️ What it Produces

A sheet called **`Total now all`** with the columns:

| Col | Header                    | Description                                                         |
| --- | ------------------------- | ------------------------------------------------------------------- |
| A   | `ITEM / DESCRIPTION`      | Title-cased item/description; duplicates merged (case-insensitive). |
| B   | `SKU CC# as in cataloque` | SKU from YU → else BY → else HH → else BISQUE’s `item`.             |
| C   | `BISQUE_IL (pcs)`         | Quantity from BISQUE source.                                        |
| D   | `YU_Storage (pcs)`        | Quantity from YU source.                                            |
| E   | `BY_Storage (pcs)`        | Quantity from BY source.                                            |
| F   | `HH_Storage (pcs)`        | Quantity from HH source.                                            |
| G   | `TOTAL NOW ALL (pcs)`     | Formula `=C+D+E+F` (bold, centered, with conditional coloring).     |

**Conditional text color on G:**

* Orange when `10 ≤ G < 20`
* Red when `G < 10`

Rows are banded with alternating backgrounds; header row is frozen.

---

## 🧰 Main Actions

The script adds a custom menu **`TotalNow`**:

* **Collect Data** → runs `buildTotalNowAll()` to rebuild the sheet from sources
* **Sort by TOTAL (desc)** → sorts by column **G** descending
* **Sort by TOTAL (asc)** → sorts by column **G** ascending

---

## 🔗 Data Sources (edit these to match your files)

Inside the script, update `SRC` if needed:

```js
const SRC = {
  BISQUE: { id: 'https://docs.google.com/spreadsheets/d/<ID1>/edit#gid=1644122405', sheetName: 'חדש' },
  YU:     { id: 'https://docs.google.com/spreadsheets/d/<ID2>/edit#gid=1764326434', sheetName: 'STORAGE' },
  BY:     { id: 'https://docs.google.com/spreadsheets/d/<ID3>/edit#gid=1764326434', sheetName: 'STORAGE' },
  HH:     { id: 'https://docs.google.com/spreadsheets/d/<ID4>/edit#gid=1764326434', sheetName: 'STORAGE' },
};
```

You can paste full URLs or just the spreadsheet IDs — the code extracts the ID either way.

---

## 🧭 How the Script Finds Columns

The script scans header rows (first up to 10 rows) and matches (case-insensitive, trimmed):

* **Description / Item key**

  * Exact description: `description`, `תיאור`
  * Or any of: `item`, `item name`, `name`, `sku`, `description`, `שם`, `פריט`, `תיאור`
* **Total** (numeric quantity column):

  * Any of: `total now (pcs)`, `total now`, `total`, `total_now`, `totalnow`, `total now all`, `סהכ`, `סה״כ`, `סה"כ`
* **SKU CC#** (only for YU/BY/HH):

  * Any of: `sku cc# as in cataloque/catalog/catalogue` (with/without `#`)
* **ITEM fallback** (only for BISQUE):

  * `item`

> If a required header is not found, that source is skipped gracefully.

---

## 🧮 Merging & Normalizing

* **Keying / dedupe:** item labels are normalized with Unicode NFKC, trimmed, spaces collapsed, lowercased → items are merged case-insensitively.
* **Display name:** first seen label (prefer BISQUE) → converted to **Title Case** (word-wise, hyphens preserved).
* **Numbers:** parsed from numbers or strings; spaces removed, `,` treated as decimal separator.

**SKU selection order:** YU → BY → HH → BISQUE’s `item` (as last resort).

---

## 🚀 Installation & Use

1. Open the destination spreadsheet (where you want `Total now all`).
2. Go to **Extensions → Apps Script**.
3. Paste this script (replace prior version if needed).
4. Update `SRC` IDs and `sheetName` values to match your sources.
5. **Save** and **authorize** when prompted.
6. Reload the spreadsheet; use **Menu → TotalNow → Collect Data**.

---

## 🧷 Sorting

Use the custom menu or call directly:

```js
sortTotalNowAllDesc(); // by TOTAL (G) high → low
sortTotalNowAllAsc();  // by TOTAL (G) low → high
```

---

## ⚙️ Formatting Applied

* Header row bold, centered, background `#AFC4E2`.
* Data rows A:G alternating backgrounds (`#EEF4FB` / `#FFFFFF`).
* Columns **C:F** number format `'0'`.
* Column **G**:

  * Formula `=SUM(RC[-4]:RC[-1])`
  * Bold, centered, number format `'0'`
  * **Conditional font color**:

    * Orange `#FFA500` if `10 ≤ G < 20`
    * Red `#FF0000` if `G < 10`
* First row frozen.

---

## 🧩 Error Handling & Skips

* Missing sheet in a source → error thrown (with sheet name).
* Missing required headers in a source → that source contributes nothing (others still load).
* No rows found after headers → that source contributes nothing.
* The target sheet is cleared (values & formats) before writing fresh data.

---

## 🧪 Troubleshooting

| Symptom                               | Likely Cause                           | Fix                                                                  |
| ------------------------------------- | -------------------------------------- | -------------------------------------------------------------------- |
| `Лист "<name>" не найден`             | Wrong `sheetName` in `SRC`             | Set the correct tab name (e.g., `STORAGE`, `חדש`)                    |
| Items duplicated with different cases | Labels differ (spacing/Unicode)        | The script already merges in a normalized way; check source spelling |
| `SKU CC#` empty                       | Not present in YU/BY/HH                | Ensure the SKU column header matches the accepted variants           |
| G (TOTAL) not colored                 | Too few rows or other rules overlapped | The script filters previous rules in G2:G; re-run “Collect Data”     |

---

## 🔒 Permissions

The script opens external spreadsheets by ID → you’ll be asked to grant access.
All source files must be accessible to the script’s executing account.

---

## 🏷️ Version

**v1.0 — Initial release**

* Multi-source merge (BISQUE, YU, BY, HH)
* Case-insensitive item dedupe + title-case display
* SKU merge with fallback
* Conditional formatting on totals
* Custom menu & sort helpers

---

---

If you want, I can add a tiny “Usage GIF” section with screenshots showing **Collect Data → result → Sort** to make the README more recruiter-friendly.
