# Workbook Inspection & Modification Guide

`xlvbatools` provides powerful inspection (workbook dumper) and modification (cell modifier) utilities to read, visual-test, and update spreadsheet contents programmatically.

---

## Workbook Inspection (Dumper)

The workbook dumper extracts cell data, interactive shapes, and named ranges to structured files.

### 1. Cell-Level JSON Model
Instead of simple raw 2D lists, cells are returned using a detailed address-mapped model:

```json
{
  "A1": {
    "row": 1,
    "col": 1,
    "value": 42.0,
    "text": "$42.00",
    "formula": null,
    "is_error": false,
    "error_type": null
  }
}
```

* **`text` (Formatted Text):** Captures the cell's locale-formatted display text (e.g. currency, formatted dates) using Excel's `.Text` property instead of just raw values.
* **`formula`:** Holds the cell's formula string (if applicable).
* **`is_error` / `error_type`:** Detects and flags cell error values (e.g., `#DIV/0!`, `#VALUE!`, `#REF!`).

### 2. Interactive Shapes & Named Ranges
* **Interactive Shapes:** Documents elements with click handlers or macro links (e.g., Form Control buttons, ActiveX elements, linked macros under `OnAction`).
* **Named Ranges:** Extracts named range definitions alongside their evaluated values and scope (Workbook vs. Sheet-level).

### 3. Pillow-Composited Screenshots
The screenshot generator exports PNG images of target worksheets.

Rendering occurs in an isolated owned Excel process. The source workbook is opened read-only with macros, events, and link updates disabled. Excel renders the original range directly; only its bitmap is transferred to a blank temporary chart workbook, so worksheet VBA is never copied or compiled. Data and screenshot requests share the same session.

Only visible worksheet tabs are rendered by default. Pass `include_hidden_sheets=True` in Python or `--include-hidden-sheets` on the CLI to explicitly include Hidden and VeryHidden worksheets.
* **Headers:** Inserts column letters (A, B, C...) and row numbers (1, 2, 3...) on the top and left margins using Pillow.
* **Gridlines:** Forces Excel gridlines to be visible and overlays grid dividers on the screenshot.
* **Retry Loop:** Automatically retries range copying and chart pasting up to 5 times (with backoff) to eliminate transient COM clipboard locks.

### 4. Session Safety
* **Graceful Targeted Closure**: The dumper utilizes `ExcelSession` to open the workbook. This uses ROT and window handle (Hwnd) tracking to close only the target workbook and its instance, ensuring other open Excel processes and user workbooks are left untouched.

---

## Workbook Modification

The modifier allows you to update Excel cells, write formulas, and manage named ranges:

```python
from xlvbatools.workbook.modifier import modify_cell

# Set a cell value
modify_cell("workbook.xlsm", sheet="Sheet1", cell="C3", value=12.5)

# Set a cell formula (excel.Calculate() is triggered automatically on success)
modify_cell("workbook.xlsm", sheet="Sheet1", cell="C4", formula="=C3*2")

# Create a named range
modify_cell("workbook.xlsm", name="TaxRate", refers_to="=Sheet1!$C$3")

# Delete a named range
modify_cell("workbook.xlsm", name="TaxRate", delete_name=True)
```

### Session Safety
* **Graceful Targeted Closure**: The cell modifier runs within `ExcelSession` context, which uses ROT and Hwnd tracking to close only the target workbook being updated. Unrelated open user workbooks and their Excel processes remain running and completely unaffected.
