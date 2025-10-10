# CC_TLV_all_storages

This Google Apps Script streamlines inventory tracking on **COLORS** and **STORAGE** sheets by adding dated “daily blocks” (IN / OUT / out to floor), recalculating per-block **Total**, and keeping the **TOTAL NOW** column up-to-date. It also standardizes formulas and validations so numeric input and formulas work smoothly.

## Features

* **Daily blocks** with automatic headers (date in row 1, labels in row 2)

  * 2-column: `<IN|OUT|out to floor> | Total`
  * 3-column: `IN | out to floor | Total`
  * 4-column: `IN | OUT | out to floor | Total`
* **Automatic Total formulas** per block (see logic below)
* **TOTAL NOW** always pulls the value from the **rightmost** “Total” column
* **Action dropdown** in cell `A1` on each working sheet
* **Numeric formats & validations**: formulas allowed, soft numeric checks for manual input
* **Conditional formatting** on `TOTAL NOW`: `<20` orange; `<10` red
* Temporarily **removes protections** to safely update headers/formulas

## Sheet Assumptions

* Processed sheets are listed in `VALID_SHEETS` (default: **COLORS**, **STORAGE**).
* **Row 1**: date (merged across the block) when a new block is added.
* **Row 2**: block labels (`IN`, `OUT`, `out to floor`, `Total`).
* Data starts from `DATA_START_ROW` (default **3**).
* A **`TOTAL NOW`** column exists in **row 1** (the script auto-detects it).

## Calculation Logic

Let `N(x) = IFERROR(VALUE(TRIM(x)), 0)`

* **4-column block**
  `Total = N(IN) + N(OUT) − N(out to floor) − N(prevTotal)`
* **3-column block**
  `Total = N(IN) + N(out to floor) − N(prevTotal)`
* **2-column block**
  `Total = N(prevTotal) ± N(value)`
  (sign from left label: `IN` → `+`, `OUT`/`out to floor` → `−`)

`prevTotal` is the **Total** of the previous block (immediately to the left).

### `TOTAL NOW` Formula

For each data row, `TOTAL NOW` reads from the **rightmost** “Total” in row 2:

```excel
=IFERROR(
  INDEX(ROW:ROW,
    MAX(FILTER(COLUMN(FirstBlockStart$2:Last$2),
               FirstBlockStart$2:Last$2="Total"))
  ),
"")
```

`FirstBlockStart` is the first column of the first block; `Last` is the last column on the sheet.

## Usage

### Action Dropdown (A1)

Each sheet in `VALID_SHEETS` gets a dropdown in **A1**:

* `➕ Add new day block (IN)` → adds **IN | Total**
* `➕ Add new day block (OUT)` → adds **OUT | Total**
* `➕ Add new day block (out to floor)` → adds **out to floor | Total**
* `🔁 Update new day block` → recalculates the newest block’s **Total**, fixes all Totals, and refreshes **TOTAL NOW**

Pick an action; the script runs and then clears A1.

### Data Entry

* Editable input columns are those labeled `IN`, `OUT`, `out to floor`.
* **Manual input** must be numeric (supports `.` or `,` as decimal separator).
* **Formulas are allowed**; results are formatted as numeric.

## What the Script Does Automatically

* Writes **date** for a new block (row 1, merged across the block)
* Formats block headers (row 2: centered, wrapped)
* Applies numeric format `0.############` to inputs and Totals
* Adds **soft numeric validation** (warns but does not block formulas)
* Repairs **all** “Total” formulas when needed
* Keeps **TOTAL NOW** in sync with the rightmost “Total”
* Adds **conditional formatting** on `TOTAL NOW`
  (`<20` → orange text, `<10` → red text)

## Public Functions

* `addNewDayBlock(mode, sheetName)`
  Adds a new 2-column block `<mode> | Total`.
  `mode`: `"IN"` | `"OUT"` | `"out to floor"` (case-insensitive)
  `sheetName`: a sheet in `VALID_SHEETS`.

* `updateDayBlock(sheetName)`
  Recomputes the newest block’s **Total**, then ensures all Totals and updates `TOTAL NOW`.

* `fixAllTotalsAndTotalNow(sheetName)`
  One-off repair of all **Total** formulas and `TOTAL NOW`.

* `onceAllowFormulasEverywhere()`
  Softens validations in all input columns across `VALID_SHEETS` (useful after updates).

* `removeInstallableOnEditTriggers()`
  Removes legacy installable `OnEdit` triggers (cleanup utility).

## Configuration

* `VALID_SHEETS`: list of sheet names to manage
* `DATA_START_ROW`: first data row (default `3`)

## Changelog (placeholder)

* **v1.0** — Initial release: daily blocks, Total formulas, TOTAL NOW, action dropdown, validations, conditional formatting.



