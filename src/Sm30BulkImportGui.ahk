#Requires AutoHotkey v2.0

#Include Sm30BulkLoader.ahk
#Include Sm30TableCatalog.ahk
#Include Sm30ExcelImport.ahk
#Include Sm30SapSessions.ahk

; Two-step GUI for Excel-driven SM30 bulk import.
class Sm30BulkImportGui {
    __New() {
        this.policy := SapHookPolicy()
        this.logger := SapFileLogger()
        this.logPath := this.logger.logPath
        this.excelPath := ""
        this.sheetNames := []
        this.selectedSheet := ""
        this.tableDef := Sm30TableCatalog.GetByIndex(1)
        this.rows := []
        this.sessionEntries := []
        this.importRunning := false
        this._BuildExcelWindow()
        this._BuildRunWindow()
    }

    Show() {
        this._ResetExcelWindow()
        this.excelGui.Show()
    }

    _BuildExcelWindow() {
        gui := Gui("+Resize", "SM30 Bulk Import — Load Excel")
        gui.SetFont("s10", "Segoe UI")
        gui.OnEvent("Close", ObjBindMethod(this, "_OnExcelClose"))
        gui.OnEvent("Size", ObjBindMethod(this, "_OnExcelResize"))

        gui.Add("Text", "w620", "Import customizing data from Excel into SAP SM30 maintenance views.")
        gui.Add("Text", "w620 cGray", "Select a workbook, choose the worksheet and target table, preview the data,"
            . " then test one row in SAP before starting the full import.")

        gui.Add("GroupBox", "xm w640 h72 Section", "Excel file")
        gui.Add("Text", "xs+20 ys+20 w70", "File:")
        this.excelPathEdit := gui.Add("Edit", "x+0 w470 ReadOnly", "")
        browseBtn := gui.Add("Button", "x+8 w80", "Browse...")
        browseBtn.OnEvent("Click", ObjBindMethod(this, "_BrowseExcelFile"))

        gui.Add("GroupBox", "xm w640 h120 Section", "Import mapping")
        gui.Add("Text", "xs+20 ys+20 w110", "Worksheet:")
        this.worksheetCombo := gui.Add("DropDownList", "x+0 w500 Choose1 Disabled", ["(select a file first)"])
        this.worksheetCombo.OnEvent("Change", ObjBindMethod(this, "_OnWorksheetChanged"))

        gui.Add("Text", "xs+20 y+12 w110", "SM30 table:")
        tableLabels := Sm30TableCatalog.GetLabels()
        this.tableCombo := gui.Add("DropDownList", "x+0 w500 Choose1", tableLabels)
        this.tableCombo.OnEvent("Change", ObjBindMethod(this, "_OnTableChanged"))

        gui.Add("GroupBox", "xm w640 h220 Section", "Preview")
        this.rowCountText := gui.Add("Text", "xs+20 ys+20 w600", "Rows loaded: 0")
        this.previewEdit := gui.Add("Edit", "xs w600 h150 ReadOnly -VScroll", "(no data loaded)")

        gui.Add("Text", "xm w640 cGray", "The test write uses the active SAP GUI session and does not save.")
        testBtn := gui.Add("Button", "xm w200 h32", "Test first row in SAP")
        testBtn.OnEvent("Click", ObjBindMethod(this, "_TestFirstRow"))

        okBtn := gui.Add("Button", "x+240 w120 h32 Default", "OK")
        okBtn.OnEvent("Click", ObjBindMethod(this, "_OnExcelOk"))
        cancelBtn := gui.Add("Button", "x+8 w120 h32", "Cancel")
        cancelBtn.OnEvent("Click", ObjBindMethod(this, "_OnExcelClose"))

        this.excelGui := gui
    }

    _BuildRunWindow() {
        gui := Gui("+Resize +MinSize640x420", "SM30 Bulk Import — Run")
        gui.SetFont("s10", "Segoe UI")
        gui.OnEvent("Close", ObjBindMethod(this, "_OnRunClose"))

        gui.Add("Text", "w620", "Upload loaded Excel rows into the selected SAP session.")
        this.runSummaryText := gui.Add("Text", "w620", "")

        gui.Add("GroupBox", "xm w640 h90 Section", "SAP session")
        gui.Add("Text", "xs+20 ys+24 w110", "Session:")
        this.sessionCombo := gui.Add("DropDownList", "x+0 w500 Choose1", ["(no SAP sessions found)"])
        refreshBtn := gui.Add("Button", "xs+130 y+10 w120", "Refresh list")
        refreshBtn.OnEvent("Click", ObjBindMethod(this, "_RefreshSessions"))

        gui.Add("GroupBox", "xm w640 h70 Section", "Options")
        this.autoSaveCheck := gui.Add("CheckBox", "xs+20 ys+26 Checked0", "Press Save in SAP when import finishes")

        gui.Add("GroupBox", "xm w640 h120 Section", "Progress")
        this.progressBar := gui.Add("Progress", "xs+20 ys+26 w600", 0)
        this.statusText := gui.Add("Text", "xs w600 h50 +Wrap", "Ready.")

        gui.Add("GroupBox", "xm w640 h90 Section", "Tools")
        openLogBtn := gui.Add("Button", "xs+20 ys+24 w140", "Open log file")
        openLogBtn.OnEvent("Click", ObjBindMethod(this, "_OpenLogFile"))
        openLogFolderBtn := gui.Add("Button", "x+8 w140", "Open logs folder")
        openLogFolderBtn.OnEvent("Click", ObjBindMethod(this, "_OpenLogFolder"))
        backBtn := gui.Add("Button", "x+8 w140", "Back to Excel")
        backBtn.OnEvent("Click", ObjBindMethod(this, "_BackToExcel"))

        this.startBtn := gui.Add("Button", "xm w160 h36 Default", "Start import")
        this.startBtn.OnEvent("Click", ObjBindMethod(this, "_StartImport"))
        runCancelBtn := gui.Add("Button", "x+8 w160 h36", "Close")
        runCancelBtn.OnEvent("Click", ObjBindMethod(this, "_OnRunClose"))

        this.runGui := gui
    }

    _ResetExcelWindow() {
        this.excelPath := ""
        this.sheetNames := []
        this.selectedSheet := ""
        this.rows := []
        this.excelPathEdit.Value := ""
        this._SetWorksheetOptions(["(select a file first)"], true)
        this.tableCombo.Choose(1)
        this.tableDef := Sm30TableCatalog.GetByIndex(1)
        this.rowCountText.Value := "Rows loaded: 0"
        this.previewEdit.Value := "(no data loaded)"
    }

    _BrowseExcelFile(*) {
        selectedPath := FileSelect("1", A_ScriptDir "\data", "Select Excel file", "Excel (*.xlsx; *.xlsm; *.xls)")
        if (selectedPath = "") {
            return
        }
        this.excelPath := selectedPath
        this.excelPathEdit.Value := selectedPath
        try {
            this.sheetNames := Sm30ExcelImport.ListSheetNames(selectedPath)
        } catch {
            MsgBox("Could not read the selected Excel file.`n`nEnsure Microsoft Excel is installed.", "Excel import", "Icon!")
            return
        }

        if (this.sheetNames.Length = 0) {
            MsgBox("The selected workbook has no worksheets.", "Excel import", "Icon!")
            return
        }

        this.selectedSheet := this.sheetNames[1]
        this._SetWorksheetOptions(this.sheetNames, this.sheetNames.Length <= 1)
        this._ReloadRows()
    }

    _SetWorksheetOptions(names, disabled := false) {
        this.worksheetCombo.Delete()
        for sheetName in names {
            this.worksheetCombo.Add([sheetName])
        }
        this.worksheetCombo.Choose(1)
        if (disabled) {
            this.worksheetCombo.Enabled := false
        } else {
            this.worksheetCombo.Enabled := true
        }
    }

    _OnWorksheetChanged(*) {
        if (this.excelPath = "") {
            return
        }
        sheetName := this.worksheetCombo.Text
        if (sheetName = "" || sheetName = "(select a file first)") {
            return
        }
        this.selectedSheet := sheetName
        this._ReloadRows()
    }

    _OnTableChanged(*) {
        tableIndex := this.tableCombo.Value
        this.tableDef := Sm30TableCatalog.GetByIndex(tableIndex)
        if (this.excelPath != "") {
            this._ReloadRows()
        }
    }

    _ReloadRows() {
        if (this.excelPath = "" || !IsObject(this.tableDef)) {
            return
        }
        columnCount := this.tableDef.columns.Length
        try {
            this.rows := Sm30ExcelImport.ReadRows(
                this.excelPath,
                this.selectedSheet,
                columnCount,
                true
            )
        } catch {
            this.rows := []
            this.rowCountText.Value := "Rows loaded: 0"
            this.previewEdit.Value := "Could not read worksheet data."
            MsgBox("Could not read rows from the selected worksheet.", "Excel import", "Icon!")
            return
        }

        this.rowCountText.Value := "Rows loaded: " this.rows.Length
        if (this.rows.Length = 0) {
            this.previewEdit.Value := "No data rows found (header row is skipped automatically)."
            return
        }
        this.previewEdit.Value := Sm30ExcelImport.PreviewRows(this.rows, 8)
    }

    _ValidateExcelStep() {
        if (this.excelPath = "") {
            MsgBox("Select an Excel file first.", "Excel import", "Icon!")
            return false
        }
        if (!IsObject(this.tableDef)) {
            MsgBox("Select a customizing table.", "Excel import", "Icon!")
            return false
        }
        if (this.rows.Length = 0) {
            MsgBox("No data rows were loaded from Excel.", "Excel import", "Icon!")
            return false
        }
        return true
    }

    _OnExcelOk(*) {
        if (!this._ValidateExcelStep()) {
            return
        }
        this.excelGui.Hide()
        this._OpenRunWindow()
    }

    _OpenRunWindow() {
        this._RefreshSessions()
        summary := "File: " this.excelPath "`n"
            . "Worksheet: " this.selectedSheet "`n"
            . "Table: " this.tableDef.label "`n"
            . "Rows to import: " this.rows.Length "`n"
            . "Log file: " this.logPath
        this.runSummaryText.Value := summary
        this.progressBar.Value := 0
        this.statusText.Value := "Ready to import " this.rows.Length " rows."
        this.startBtn.Enabled := true
        this.runGui.Show()
    }

    _RefreshSessions(*) {
        this.sessionEntries := Sm30SapSessions.List(this.policy)
        labels := Sm30SapSessions.GetLabels(this.sessionEntries)
        this.sessionCombo.Delete()
        if (labels.Length = 0) {
            this.sessionCombo.Add(["(no SAP sessions found — open SAP GUI and log in)"])
            this.sessionCombo.Choose(1)
            this.sessionCombo.Enabled := false
            return
        }
        for label in labels {
            this.sessionCombo.Add([label])
        }
        this.sessionCombo.Choose(1)
        this.sessionCombo.Enabled := true
    }

    _TestFirstRow(*) {
        if (!this._ValidateExcelStep()) {
            return
        }

        session := Sm30SapSessions.GetActiveSession(this.policy)
        if (!IsObject(session)) {
            MsgBox("No active SAP GUI session found.`n`nOpen SAP, log in, and make the target session active.",
                "Test row", "Icon!")
            return
        }

        testRow := [this.rows[1].Clone()]
        hookPolicy := LoggingSapHookPolicy(this.logger)
        loader := ""
        try {
            loader := Sm30BulkLoader.FromSession(session, hookPolicy)
            loader.SetLogger(this.logger)
            loader.SetFillMode("row")
            loader.SetQuietComLogging(false)
            loader
                .OpenView(this.tableDef.viewName)
                .EnterMaintenance()
                .NewEntries()
                .UseTable(this.tableDef.tableId)
            loader.FillRows(this.tableDef.columns, testRow)
            this.logger.Info("Test row write succeeded for table " this.tableDef.viewName)
            MsgBox("Test row written successfully.`n`n"
                . "Table: " this.tableDef.label "`n"
                . "Values: " Sm30ExcelImport._JoinFields(testRow[1], " | ") "`n`n"
                . "Nothing was saved. Check SAP, then continue with OK if the row looks correct.`n`n"
                . "Log: " this.logPath,
                "Test row", "Iconi")
        } catch {
            detail := IsObject(loader) && loader.lastFailure != "" ? loader.lastFailure : "SAP test write failed."
            this.logger.Error("Test row failed: " detail)
            MsgBox("Test row failed.`n`n" detail "`n`nLog: " this.logPath, "Test row", "Icon!")
        }
    }

    _StartImport(*) {
        if (this.importRunning) {
            return
        }
        if (this.sessionEntries.Length = 0) {
            MsgBox("No SAP sessions available. Open SAP GUI and click Refresh list.", "Import", "Icon!")
            return
        }

        sessionIndex := this.sessionCombo.Value
        if (sessionIndex < 1 || sessionIndex > this.sessionEntries.Length) {
            MsgBox("Select a SAP session.", "Import", "Icon!")
            return
        }

        this.importRunning := true
        this.startBtn.Enabled := false
        this.progressBar.Value := 0
        this.statusText.Value := "Starting import..."

        sessionEntry := this.sessionEntries[sessionIndex]
        hookPolicy := LoggingSapHookPolicy(this.logger)
        loader := ""
        filledCount := 0
        autoSave := this.autoSaveCheck.Value

        try {
            loader := Sm30BulkLoader.FromSession(sessionEntry.session, hookPolicy)
            loader.SetLogger(this.logger)
            loader.SetFillMode("page")
            loader.SetQuietComLogging(true)
            loader.SetProgressCallback(ObjBindMethod(this, "_OnImportProgress"))
            loader
                .OpenView(this.tableDef.viewName)
                .EnterMaintenance()
                .NewEntries()
                .UseTable(this.tableDef.tableId)
            filledCount := loader.FillRows(this.tableDef.columns, this.rows)
            if (autoSave) {
                this.statusText.Value := "Saving in SAP..."
                loader.Save()
            }
            this.progressBar.Value := 100
            this.statusText.Value := "Import finished. Rows processed: " filledCount
                . (autoSave ? " (saved)" : " (not saved — review in SAP first)")
            this.logger.Info("Bulk import finished rows=" filledCount " autoSave=" autoSave)
            MsgBox("Import finished.`n`nRows processed: " filledCount "`n"
                . (autoSave ? "SAP Save was executed." : "SAP Save was NOT executed.") "`n`n"
                . "Log: " this.logPath,
                "Import complete", "Iconi")
        } catch {
            detail := IsObject(loader) && loader.lastFailure != "" ? loader.lastFailure : "Import failed."
            this.statusText.Value := "Import failed."
            this.logger.Error("Bulk import failed: " detail)
            MsgBox("Import failed.`n`n" detail "`n`nLog: " this.logPath, "Import", "Icon!")
        } finally {
            this.importRunning := false
            this.startBtn.Enabled := true
        }
    }

    _OnImportProgress(completed, total, message := "") {
        percent := total > 0 ? Round((completed / total) * 100) : 0
        if (percent > 100) {
            percent := 100
        }
        if (percent < 0) {
            percent := 0
        }
        this.progressBar.Value := percent
        statusLine := "Progress: " completed " / " total " (" percent "%)"
        if (message != "") {
            statusLine .= "`n" message
        }
        this.statusText.Value := statusLine
    }

    _OpenLogFile(*) {
        if (this.logPath != "" && FileExist(this.logPath)) {
            Run(this.logPath)
            return
        }
        MsgBox("Log file not found yet.`n`n" this.logPath, "Logs", "Icon!")
    }

    _OpenLogFolder(*) {
        logDir := this.logger.logDir
        if (DirExist(logDir)) {
            Run('explorer.exe "' logDir '"')
            return
        }
        MsgBox("Logs folder not found.`n`n" logDir, "Logs", "Icon!")
    }

    _BackToExcel(*) {
        if (this.importRunning) {
            MsgBox("Import is still running.", "Import", "Icon!")
            return
        }
        this.runGui.Hide()
        this.excelGui.Show()
    }

    _OnExcelClose(*) {
        if (this.importRunning) {
            return
        }
        this.excelGui.Destroy()
        if (IsObject(this.runGui)) {
            this.runGui.Destroy()
        }
        ExitApp()
    }

    _OnRunClose(*) {
        if (this.importRunning) {
            MsgBox("Wait until the import finishes.", "Import", "Icon!")
            return
        }
        this.runGui.Hide()
        this.excelGui.Show()
    }

    _OnExcelResize(gui, minMax, width, height, *) {
        if (minMax = -1) {
            return
        }
    }
}
