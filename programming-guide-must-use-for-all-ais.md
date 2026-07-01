# Programming Guide for AI Assistants — AHK v2 + SAP SM30 Bulk Loader

This document captures hard-won lessons from building the SM30 bulk import tooling in this repo. **Future AI agents should read this before changing AHK or SAP scripting code here.**

---

## Project conventions (non-negotiable)

| Rule | Detail |
|------|--------|
| Runnable scripts | Use `.ahk2` extension |
| Libraries / includes | Use `.ahk` under `src/` |
| Header | `#Requires AutoHotkey v2.0` on every file |
| No lambdas | Do **not** use `() => expr` or arrow functions |
| No exception variables | Use `try { } catch { }` only — **not** `catch err`, `catch as e`, etc. |
| Error detail | Build messages from known context + `A_LastError`; never `e.Message` |
| Logs | Written to `logs/` (gitignored) via `SapFileLogger` |
| Git branches | Use prefix `cursor/` and suffix `-7156` when creating branches |

---

## AHK v2 pitfalls (counterintuitive)

### 1. Class methods share local variable names

**Symptom:** `Error: This local variable has not been assigned a value` on a line that clearly assigns it (e.g. `gui := Gui(...)`).

**Cause:** In AHK v2, local variable names are effectively shared across all methods of a class. If one method declares a parameter named `gui` (e.g. `_OnExcelResize(gui, minMax, ...)`), another method cannot use `gui` as a local variable — the runtime treats it as the same symbol, still unassigned in that method.

**Fix:** Never reuse generic names like `gui`, `control`, `item` across class methods. Use distinct names: `excelWin`, `runWin`, `senderGui`, etc.

```ahk
; BAD — _OnExcelResize(gui, ...) elsewhere breaks this
_BuildWindow() {
    gui := Gui()
}

; GOOD
_BuildWindow() {
    excelWin := Gui()
    this.excelGui := excelWin
}
_OnResize(senderGui, minMax, width, height, *) {
}
```

This was the hardest GUI bug to diagnose because the error points at the assignment line, not the other method.

---

### 2. Plain objects do not have `.Has()`

Use `columnDef.HasOwnProp("field")` instead of `columnDef.Has("field")`.

Avoid calling bare `HasOwnProp("x")` at global scope — it can trigger VarUnset warnings. Always call on an object: `obj.HasOwnProp("x")`.

---

### 3. Arrays have no `.Join()`

Build strings manually or use a small helper (`JoinLines()`). Do not assume JavaScript-style array methods exist.

---

### 4. `Map` vs object access

`_EnsurePageReadyForFill` returns `Map("visibleStart", x, "rowsOnPage", y)`. Access with `pagePlan["visibleStart"]`, not dot notation for dynamic keys unless using a plain object literal.

---

### 5. Progress / callbacks without lambdas

Pass bound methods: `ObjBindMethod(this, "_OnImportProgress")`. Do not use inline arrow callbacks.

---

## SAP GUI scripting — mental model

### Object hierarchy for session picking

```
GuiApplication  (from ComObjGet("SAPGUI").GetScriptingEngine)
  └── Children[i]   (= GuiConnection)
        └── Children[j]   (= GuiSession)
```

- `Sm30BulkLoader.Attach()` always takes **first connection, first session** — wrong when multiple SAP windows are open.
- Use `Sm30BulkLoader.FromSession(chosenSession, policy)` for UI session pickers.
- Label sessions via `session.Info`: `SystemName`, `Client`, `User`, `Transaction`, `SessionNumber`.

---

### SM30 table cell paths

Comes from the **SAP Script Recorder**, not guessed:

```ahk
{ index: 0, kind: "Text", prefix: "ctxt", field: "WUE/PFEPRUNTYPE-VKORG" }
```

Built path:

```text
{tablePath}/{prefix}/{field}[{columnIndex},{visibleRowIndex}]
```

Kinds: `Text` → `.Text`, `Key` → `.Key` (dropdown), `Selected` → checkbox.

---

### Absolute row vs visible row vs scroll

| Concept | Meaning |
|---------|---------|
| `absoluteRow` | Logical row index in the full table (what the loader tracks) |
| `scrollPos` | `VerticalScrollbar.Position` — absolute index of the **top** visible row |
| `visibleRow` | Row on screen: `absoluteRow - scrollPos` |
| `VisibleRowCount` | Rows currently visible (changes if user resizes window) |
| `RowCount` | **Not** the count of filled rows — includes **preallocated empty rows** |

**Scroll target for a row:** put it on screen with:

```ahk
scrollPos := absoluteRowIndex
if (scrollPos > maxScroll)
    scrollPos := maxScroll
if (absoluteRowIndex - scrollPos >= visibleCount)
    scrollPos := absoluteRowIndex - visibleCount + 1
```

Always re-read `VisibleRowCount` and scroll metrics after scroll, Enter, skip recovery, or window resize.

---

### Stale COM references

After `SendVKey`, skip button presses, or error recovery, **re-fetch the table**:

```ahk
this.table := this.session.FindById(this.tableFindId)
```

Do not keep using an old `GuiTableControl` reference across SAP round trips.

---

## Error recovery (SV009 duplicate key) — what was hard

### What we tried that failed

1. **Status bar snapshot comparison** — After each skip press, SAP often shows the **same** SV009 text ("Es ist schon ein Eintrag mit gleichem Schlüssel vorhanden") even when the duplicate row **was** removed. Stopping when `snapshotAfter = snapshotBefore` exits after one skip and leaves ~19 duplicates.

2. **Resync `absoluteRow := table.RowCount` after skips** — `RowCount` jumped to 77 (preallocated slots). Cursor landed at row 77, `_EnsurePhysicalRow` looped forever because Enter did not increase `RowCount`.

### What works

```text
while status bar shows error (MessageType E or A, non-empty text):
    press wnd[0]/tbar[1]/btn[20] (skip)
    wait until session not Busy
until no error
press Enter once to continue
```

After a page commit with skips:

```ahk
dataIndex += rowsOnPage                    ; always consume input rows attempted
absoluteRow += rowsOnPage - skipped        ; only advance by rows actually kept
```

**Never** set `absoluteRow := table.RowCount` after skips.

---

## Scroll after partial page failure — what was hard

**Scenario:** Page of 39 rows, 19 duplicate skips → only 20 rows saved. Loader advanced to `absoluteRow=39` and scrolled to `scrollPos=39`. User expected ~row 20 / scroll ~20.

**Fix:** Subtract skipped count from absolute row advancement (see above). Next fill starts at the correct logical position.

**Last-line defence (writable check):** Before filling, validate cells are `Changeable`. If the **bottom** rows of the planned page are not writable:

1. Count consecutive non-writable rows from the bottom upward.
2. Scroll up by that many (`scrollPos -= count`).
3. Retry (re-read `VisibleRowCount` each attempt).

Do **not** fall back to `table.RowCount` as the next write target.

---

## Row creation guard

When `absoluteRow >= rowCount`, create rows by focus + Enter on last visible cell of last column.

If `RowCount` does **not** increase after Enter, **fail fast** — do not loop 1000 times. That situation means the cursor is on a preallocated empty slot, not a true "create new row" boundary.

---

## Throughput tuning

- Default fill mode: **page** (column-major, one Enter per visible page).
- Row mode: debugging only (verbose, slow).
- `OpenView()` → `Maximize()` + `ResizeWorkingPaneEx(width, height, false)` — default height **2000** for more visible rows per page.
- Quiet COM logging during page fills (`policy.quiet := true`).

---

## Excel import GUI — architecture notes

Two-step flow works well:

1. **Load Excel** — file, worksheet (enable dropdown only if `sheetNames.Length > 1`), **SAP session dropdown + refresh**, table catalog, preview, test one row (selected session, no save).
2. **Run** — session shown in summary (chosen on step 1), auto-save off by default, progress bar, log buttons.

Excel reading requires **Microsoft Excel installed** (COM: `Excel.Application`). Regenerate sample: `python scripts/generate_sample_excel.py` → `examples/data/pfepruntype_sample.xlsx`.

Add new customizing tables via JSON in `config/tables/*.json` (see `examples/sm30/config/tables/pfepruntype.json`). Do not hardcode tables in `Sm30TableCatalog.ahk`. Use the built-in `Sm30JsonParser` — do **not** require AHK v2.1 `JsonLoad` (many users run v2.0 or v2.1-alpha where `#Requires AutoHotkey v2.1` fails).

---

## Excel import GUI — freeze fix

Opening Excel via COM (`Excel.Application`) can take several seconds and **blocks the AHK thread**. If run directly in a button handler, the GUI appears frozen with no feedback.

**Fix:**
1. Show loading state immediately (`Reading Excel...`, disable buttons).
2. Defer COM work with `SetTimer(ObjBindMethod(this, "_ProcessExcelLoad"), -1)` so the GUI repaints first.
3. Open Excel only once per operation (`LoadWorkbook` lists sheets and reads rows in one session).
4. Set `excel.ScreenUpdating := false` and `excel.EnableEvents := false`.

---

## Compiled exe layout

When the launcher is compiled, **`A_ScriptDir` is the folder containing the exe**. Ship these next to the exe:

```text
MySm30Tool.exe
config/tables/*.json
config/app.json
data/
logs/            (created at runtime)
```

Libraries in `src/` are compiled into the exe. Config and data are **not** — they must sit beside the exe. Use `Sm30AppPaths.BaseDir()` / `TablesDir()` / `DataDir()` / `LogsDir()` — never hardcode `..\..\src` for runtime data.

Launcher path: `examples/sm30/sm30_bulk_import_gui.ahk2`

---

## Counterintuitive SAP behaviors (checklist)

| Behavior | Intuition | Reality |
|----------|-----------|---------|
| Skip worked? | Status bar text changes | Same SV009 text can repeat for the next duplicate |
| Next row after skips? | Use `RowCount` | `RowCount` includes empty preallocated rows |
| Row saved count | `rowsOnPage` advanced | Must subtract `skipped` from absolute cursor |
| Table reference | Keep object from start | Refresh via `FindById` after every commit/recovery |
| Visible rows | Fixed | User resize changes `VisibleRowCount` — re-read often |
| Test vs bulk session | Same picker | Test can use `ActiveSession`; bulk needs explicit picker |

---

## Key files

| File | Purpose |
|------|---------|
| `src/Sm30BulkLoader.ahk` | Core bulk fill, scroll, error recovery |
| `src/Sm30BulkImportGui.ahk` | Two-step import UI |
| `src/Sm30TableCatalog.ahk` | View/table/column definitions |
| `src/Sm30ExcelImport.ahk` | Excel COM reader |
| `src/Sm30SapSessions.ahk` | List SAP sessions |
| `src/core/SapFileLogger.ahk` | File logging + `LoggingSapHookPolicy` |
| `examples/sm30/sm30_bulk_import_gui.ahk2` | Launch GUI |
| `examples/sm30/config/tables/*.json` | SM30 table definitions |
| `src/Sm30AppPaths.ahk` | Resolves paths from `A_ScriptDir` (script or compiled exe) |
| `src/Sm30JsonConfig.ahk` | Built-in JSON parser for table configs (no JsonLoad / v2.1 needed) |
| `examples/sm30_dummy_fill_test.ahk2` | Stress test with intentional duplicates |

---

## What to do before proposing "fixes"

1. Read the latest log in `logs/` — status bar text, `scrollPos`, `absoluteRow`, `RowCount`, `visibleCount` together tell the story.
2. Distinguish **input row index** (`dataIndex`) from **SAP absolute row cursor** — they diverge after skips.
3. Do not add stall detection based on status bar text equality.
4. Do not use `table.RowCount` as write cursor after error recovery.
5. Avoid variable names in classes that might collide across methods (`gui`, `item`, `control`, `data`).
6. Match existing code style: minimal diff, no over-abstraction, no unnecessary tests unless requested.

---

## Session summary (chronological)

1. Built SM30 bulk loader with page mode, scroll sync, CSV load.
2. Fixed `.Has()` → `HasOwnProp`, stale table refs, missing `Join`.
3. Error recovery: infinite skip loop → then stopped after 1 skip (snapshot logic) → final fix: loop until no error, no snapshot.
4. Infinite row creation at row 77 → caused by wrong `absoluteRow` after skips + `RowCount` resync.
5. Scroll landed at 39 instead of 20 → fixed with `absoluteRow += rowsOnPage - skipped`.
6. Added trailing non-writable scroll-up fallback and working pane height 2000.
7. Built Excel import GUI + sample xlsx; fixed AHK class `gui` variable name collision.

**Hardest problems:** (1) SAP keeping identical error text after successful skip, (2) `RowCount` semantics vs actual filled rows, (3) AHK v2 class-wide local variable sharing for `gui`.
