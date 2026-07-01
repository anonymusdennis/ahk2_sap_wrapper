# AHK2 SAP GUI Scripting Wrapper

AutoHotkey v2 COM wrapper/proxy for SAP GUI Scripting with hookable call pipeline.

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

Duplicate/error recovery during save uses the status bar message (`wnd[0]/sbar`) and skips invalid rows via `wnd[0]/tbar[1]/btn[20]` until no error remains:

```ahk
loader.SetErrorRecovery(true, "wnd[0]/tbar[1]/btn[20]")
```
