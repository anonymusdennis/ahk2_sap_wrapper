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

See `/examples/demo_rot_attach.ahk` for a runnable example.

## SM30 bulk table fill

Use `Sm30BulkLoader` to insert many rows into any SM30 maintenance view. The loader uses `GuiTableControl.GetCell()` and automatically scrolls the vertical scrollbar before writing rows that would fall outside the visible window.

```ahk
#Requires AutoHotkey v2.0
#Include src/Sm30BulkLoader.ahk

loader := Sm30BulkLoader.Attach()

columns := [
    { index: 0, kind: "Text" },
    { index: 1, kind: "Text" },
    { index: 3, kind: "Key" }   ; dropdown/combobox
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

Runnable example: `/examples/sm30_bulk_fill.ahk`

Column `kind` values:
- `Text` — standard input fields (`Text` property)
- `Key` — dropdown/combobox fields (`Key` property)
- `Selected` — checkbox fields

For tables with a non-standard toolbar button layout, pass explicit button IDs to `EnterMaintenance()`, `NewEntries()`, or `Save()`.
