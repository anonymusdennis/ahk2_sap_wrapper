# AHK2 SAP GUI Scripting Wrapper

AutoHotkey v2 COM wrapper/proxy for SAP GUI Scripting with hookable call pipeline.

**AI / contributors:** read [`programming-guide-must-use-for-all-ais.md`](programming-guide-must-use-for-all-ais.md) before changing code. This repo is intended to hold multiple SAP automation tools over time, not only SM30.

## Usage

```ahk
#Requires AutoHotkey v2.0
#Include src/SapWrapper.ahk

policy := SapHookPolicy()
SapGuiAuto := ComObjGet("SAPGUI")
app := GuiApplication(SapGuiAuto.GetScriptingEngine, policy)

ses := app.Children[0].Children[0]
ses.FindById("wnd[0]/tbar[0]/okcd").Text := "/nSE16"
ses.FindById("wnd[0]").SendVKey(0)
MsgBox(ses.Info.SystemName)
```

See `/examples/demo_rot_attach.ahk2` for a runnable example.

Runnable scripts use the `.ahk2` extension so they launch with AutoHotkey v2. Library/includes stay as `.ahk`.

## v2 API

`src/SapWrapper2.ahk` layers additional features on top of the same core. **v1 scripts keep working unchanged** — including `src/SapWrapper.ahk` gives you exactly the v1 behavior. Including `src/SapWrapper2.ahk` adds:

- **Typed wrapper classes for the full SAP GUI Scripting object model** (`src/generated/TypedWrappers.ahk`, ~60 classes such as `GuiTableControl`, `GuiGridView`, `GuiMainWindow`, `GuiTextField`, …) with explicit properties/methods for autocomplete. Wrapped COM return values automatically use them.
- **Enum constants** (`GuiEventType`, `GuiMessageBoxType`, `GuiScrollbarType`, …) plus a `TypeAsNumber` fallback for type detection when `.Type` is unavailable.
- **`Sap.Attach()` / `Sap.App()`** convenience roots:

```ahk
#Requires AutoHotkey v2.0
#Include src/SapWrapper2.ahk

ses := Sap.Attach()            ; active session
ses := Sap.Attach(policy, 0, 1) ; connection 0 / session 1
```

- **Opt-in stale-reference recovery** — `proxy.SetStaleRecovery(true)` makes proxies re-resolve themselves via their `FindById` origin and retry once when a call fails after a SAP round trip.
- **Richer error context** — `FindById` failures are probed with `FindById(id, false)` and classified as `not-found` / `transient` / `com-error` in the error message.
- **`SapHookPolicy2`** — adds `On_Retry`, `On_ErrorEx(info)` (structured error Map), and an `On_Popup` hook point. All policies are duck-typed; plain v1 `SapHookPolicy` objects remain accepted everywhere.
- **Collections are iterable** — `for item in session.Children` and `for index, item in collection` (0-based) work, and indexing falls back from `Item()` to `ElementAt()` without firing error hooks.

### Regenerating the generated files

`src/generated/Allowlists.ahk`, `src/generated/TypeNumbers.ahk` and `src/generated/TypedWrappers.ahk` are produced from the condensed API index:

```text
python scripts/generate_allowlists.py
```

The generator handles doc quirks (the `theGuiVContainer` typo, multi-base "Inherits members from" comments) and synthesizes `GuiTextField`, which is missing from the condensed index. CI fails if the generated files are stale.

## Tests and CI

Headless unit tests (no SAP GUI required — COM objects are faked) live in `tests/run_tests.ahk2`:

```text
AutoHotkey64.exe /ErrorStdOut tests\run_tests.ahk2
```

GitHub Actions (`.github/workflows/ci.yml`) syntax-checks all entry scripts, verifies `src/generated/` is up to date, and runs the tests on a Windows runner.

## SM30 bulk table fill

Use `Sm30BulkLoader` to insert many rows into any SM30 maintenance view. By default it fills **one visible page at a time** (column-major, one Enter per page) for much better speed than row-by-row mode.

```ahk
loader.SetFillMode("page")   ; default, fast
loader.SetFillMode("row")    ; slow, verbose, useful for debugging
```

```ahk
#Requires AutoHotkey v2.0
#Include src/Sm30BulkLoader.ahk

loader := Sm30BulkLoader.Attach()

columns := [
    { index: 0, kind: "Text", prefix: "ctxt", field: "WUE/PFEPRUNTYPE-VKORG" },
    { index: 1, kind: "Text", prefix: "txt", field: "WUE/PFEPRUNTYPE-MATNR_V" },
    { index: 2, kind: "Text", prefix: "txt", field: "WUE/PFEPRUNTYPE-MATNR_B" },
    { index: 3, kind: "Key", prefix: "cmb", field: "WUE/PFEPRUNTYPE-/WUE/PFEP_RUN_TYPE" }
]

rows := loader.LoadCsv(A_ScriptDir "\data\my_table.csv", columns, true)

loader
    .OpenView("/WUE/PFEPRUNTYPE")
    .EnterMaintenance()
    .NewEntries()
    .UseTable()                  ; auto-detect first table on screen
    .FillRows(columns, rows)
    .Save()
```

Runnable example: `/examples/sm30_bulk_fill.ahk2`

Dummy data test (generates 60 rows by default, then uploads to `/WUE/PFEPRUNTYPE`):

```text
examples/sm30_dummy_fill_test.ahk2
examples/sm30_dummy_fill_test.ahk2 120
```

Logs are written to `logs/` (gitignored). On success or failure the dialog shows the log file path.

Column definitions need the SAP field metadata from your script recorder:

```ahk
{ index: 0, kind: "Text", prefix: "ctxt", field: "WUE/PFEPRUNTYPE-VKORG" }
{ index: 3, kind: "Key", prefix: "cmb", field: "WUE/PFEPRUNTYPE-/WUE/PFEP_RUN_TYPE" }
```

`prefix` is the control type folder in the table path (`ctxt`, `txt`, `cmb`, ...).
`field` is the field name segment from the recorded `findById` path.

Column `kind` values:
- `Text` — standard input fields (`Text` property)
- `Key` — dropdown/combobox fields (`Key` property)
- `Selected` — checkbox fields

For tables with a non-standard toolbar button layout, pass explicit button IDs to `EnterMaintenance()`, `NewEntries()`, or `Save()`.

Duplicate/error recovery during save uses the status bar message (`wnd[0]/sbar`) and skips invalid rows via `wnd[0]/tbar[1]/btn[20]` until no error remains. Skipped rows are subtracted from the absolute row cursor so scroll lands on the next real row (not preallocated empty slots). After skips, Enter is pressed once to continue.

Before each page/row fill, target cells are checked with `Changeable`. If the bottom of the page is not writable, the loader counts trailing blocked rows and scrolls up by that amount (handles window resize via fresh `VisibleRowCount` reads).

By default `OpenView()` maximizes the SAP window and expands the working pane to height 2000 (`ResizeWorkingPaneEx`) for more visible rows per page. Override with `SetExpandWindowForThroughput(true, 0, 2000)` or disable with `SetExpandWindowForThroughput(false)`.

```ahk
loader.SetErrorRecovery(true, "wnd[0]/tbar[1]/btn[20]")
```

## Excel import GUI

Graphical tool for loading Excel data, testing one row in SAP, then running the bulk import with session selection and progress. Works with **AutoHotkey v2.0+** (includes alpha builds).

```text
examples/sm30/sm30_bulk_import_gui.ahk2
```

Folder layout (same layout next to the script **or** compiled `.exe`):

```text
examples/sm30/
  sm30_bulk_import_gui.ahk2
  config/
    app.json
    tables/
      pfepruntype.json
  data/
    pfepruntype_sample.xlsx
  logs/                  (created at runtime)
```

Sample workbook:

```text
examples/sm30/data/pfepruntype_sample.xlsx
```

Regenerate the sample file (requires Python + openpyxl):

```text
python scripts/generate_sample_excel.py
```

### Add more SM30 tables

Add a JSON file under `config/tables/` (one file per view). Example: `config/tables/pfepruntype.json`

```json
{
  "id": "pfepruntype",
  "label": "PFEP Run Type (/WUE/PFEPRUNTYPE)",
  "viewName": "/WUE/PFEPRUNTYPE",
  "tableId": "wnd[0]/usr/tbl/...",
  "columns": [
    { "index": 0, "kind": "Text", "prefix": "ctxt", "field": "...", "name": "VKORG" }
  ]
}
```

Column metadata comes from the SAP Script Recorder `findById` paths. Reload the GUI to pick up new files.

**Step 1 — Load Excel**
- Browse for `.xlsx` / `.xlsm` / `.xls` (requires Microsoft Excel installed; shows loading status while Excel opens)
- Choose worksheet when the workbook has multiple sheets
- Choose SAP session from dropdown (Refresh list)
- Choose the SM30 customizing table (from JSON configs)
- Preview loaded rows
- **Test first row in SAP** writes row 1 using the selected session (no save)

**Step 2 — Run**
- Shows the session chosen on step 1
- **Auto-save when finished** is off by default
- Progress bar during import
- Open log file / logs folder buttons
