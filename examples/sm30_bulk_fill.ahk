#Requires AutoHotkey v2.0
#Include ../src/Sm30BulkLoader.ahk

; Example: bulk-fill the /WUE/PFEPRUNTYPE SM30 view.
; Adapt columns + rows (or CSV) for any other SM30 table.

policy := SapHookPolicy()

try {
    loader := Sm30BulkLoader.Attach(policy)

    ; Column definitions map CSV/array positions to SAP table column indices.
    ; kind: "Text" for input fields, "Key" for dropdown/combobox values.
    columns := [
        { index: 0, kind: "Text" },  ; VKORG
        { index: 1, kind: "Text" },  ; MATNR_V
        { index: 2, kind: "Text" },  ; MATNR_B
        { index: 3, kind: "Key" }    ; PFEP_RUN_TYPE
    ]

    ; Sample rows (replace with thousands of rows or load from CSV below).
    rows := [
        ["E001", "test_matnr", "test_matnr2", "K"],
        ["E001", "test_matnr2", "test_matnr3", "A"],
        ["E001", "test_matnr3", "test_matnr4", "S1"],
        ["E001", "test_matnr4", "test_matnr5", "K"]
    ]

    ; Optional CSV input:
    ; csvPath := A_ScriptDir "\data\pfepruntype.csv"
    ; rows := loader.LoadCsv(csvPath, columns, true)

    loader
        .OpenView("/WUE/PFEPRUNTYPE")
        .EnterMaintenance()
        .NewEntries()
        .UseTable("wnd[0]/usr/tbl/WUE/SAPLMMC_PFEPTCTRL_/WUE/PFEPRUNTYPE")

    filledCount := loader.FillRows(columns, rows)
    loader.Save()

    MsgBox("Filled " filledCount " rows into /WUE/PFEPRUNTYPE.")
} catch {
    MsgBox("SM30 bulk fill failed. Ensure SAP GUI is open, scripting is enabled, and you are logged in.")
}
