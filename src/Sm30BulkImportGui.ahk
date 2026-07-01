#Requires AutoHotkey v2.0

#Include Sm30BulkLoader.ahk
#Include Sm30AppPaths.ahk
#Include Sm30TableCatalog.ahk
#Include Sm30ExcelImport.ahk
#Include Sm30SapSessions.ahk

; Two-step GUI for Excel-driven SM30 bulk import.
class Sm30BulkImportGui {
    __New() {
        this.policy := SapHookPolicy()
        logPath := Sm30AppPaths.LogsDir() "\sm30_" FormatTime(, "yyyyMMdd_HHmmss") ".log"
        this.logger := SapFileLogger(logPath)
        this.logPath := this.logger.logPath
        this.excelPath := ""
        this.sheetNames := []
        this.selectedSheet := ""
        this.tableDef := Sm30TableCatalog.GetByIndex(1)
        this.rows := []
        this.sessionEntries := []
        this.selectedSessionEntry := ""
        this.importRunning := false
        this.excelLoading := false
        this.pendingExcelPath := ""
        this._BuildExcelWindow()
        this._BuildRunWindow()
    }

    Show() {
        this._ResetExcelWindow()
        this._RefreshSessions()
        this.excelGui.Show()
    }

    _BuildExcelWindow() {
        excelWin := Gui("+Resize", "SM30 Bulk Import — Load Excel")
        excelWin.SetFont("s10", "Segoe UI")
        excelWin.OnEvent("Close", ObjBindMethod(this, "_OnExcelClose"))
        excelWin.OnEvent("Size", ObjBindMethod(this, "_OnExcelResize"))

        excelWin.Add("Text", "w620", "Import customizing data from Excel into SAP SM30 maintenance views.")
        excelWin.Add("Text", "w620 cGray", "Table definitions live in config/tables/*.json next to this script or exe.")

        excelWin.Add("GroupBox", "xm w640 h72 Section", "Excel file")
        excelWin.Add("Text", "xs+20 ys+20 w70", "File:")
        this.excelPathEdit := excelWin.Add("Edit", "x+0 w470 ReadOnly", "")
        this.browseBtn := excelWin.Add("Button", "x+8 w80", "Browse...")
        this.browseBtn.OnEvent("Click", ObjBindMethod(this, "_BrowseExcelFile"))

        excelWin.Add("GroupBox", "xm w640 h120 Section", "Import mapping")
        excelWin.Add("Text", "xs+20 ys+20 w110", "Worksheet:")
        this.worksheetCombo := excelWin.Add("DropDownList", "x+0 w500 Choose1 Disabled", ["(select a file first)"])
        this.worksheetCombo.OnEvent("Change", ObjBindMethod(this, "_OnWorksheetChanged"))

        excelWin.Add("Text", "xs+20 y+12 w110", "SM30 table:")
        tableLabels := Sm30TableCatalog.GetLabels()
        this.tableCombo := excelWin.Add("DropDownList", "x+0 w500 Choose1", tableLabels)
        this.tableCombo.OnEvent("Change", ObjBindMethod(this, "_OnTableChanged"))

        excelWin.Add("GroupBox", "xm w640 h90 Section", "SAP session")
        excelWin.Add("Text", "xs+20 ys+24 w110", "Session:")
        this.sessionCombo := excelWin.Add("DropDownList", "x+0 w500 Choose1", ["(no SAP sessions found)"])
        this.refreshSessionsBtn := excelWin.Add("Button", "xs+130 y+10 w120", "Refresh list")
        this.refreshSessionsBtn.OnEvent("Click", ObjBindMethod(this, "_RefreshSessions"))

        excelWin.Add("GroupBox", "xm w640 h220 Section", "Preview")
        this.rowCountText := excelWin.Add("Text", "xs+20 ys+20 w600", "Rows loaded: 0")
        this.previewEdit := excelWin.Add("Edit", "xs w600 h150 ReadOnly -VScroll", "(no data loaded)")

        excelWin.Add("Text", "xm w640 cGray", "Test write uses the selected SAP session above and does not save.")
        this.testBtn := excelWin.Add("Button", "xm w200 h32", "Test first row in SAP")
        this.testBtn.OnEvent("Click", ObjBindMethod(this, "_TestFirstRow"))

        this.okBtn := excelWin.Add("Button", "x+240 w120 h32 Default", "OK")
        this.okBtn.OnEvent("Click", ObjBindMethod(this, "_OnExcelOk"))
        cancelBtn := excelWin.Add("Button", "x+8 w120 h32", "Cancel")
        cancelBtn.OnEvent("Click", ObjBindMethod(this, "_OnExcelClose"))

        this.excelGui := excelWin
    }

    _BuildRunWindow() {
        runWin := Gui("+Resize +MinSize640x420", "SM30 Bulk Import — Run")
        runWin.SetFont("s10", "Segoe UI")
        runWin.OnEvent("Close", ObjBindMethod(this, "_OnRunClose"))

        runWin.Add("Text", "w620", "Upload loaded Excel rows into the selected SAP session.")
        this.runSummaryText := runWin.Add("Text", "w620", "")

        runWin.Add("GroupBox", "xm w640 h70 Section", "Options")
        this.autoSaveCheck := runWin.Add("CheckBox", "xs+20 ys+26 Checked0", "Press Save in SAP when import finishes")

        runWin.Add("GroupBox", "xm w640 h120 Section", "Progress")
        this.progressBar := runWin.Add("Progress", "xs+20 ys+26 w600", 0)
        this.statusText := runWin.Add("Text", "xs w600 h50 +Wrap", "Ready.")

        runWin.Add("GroupBox", "xm w640 h90 Section", "Tools")
        openLogBtn := runWin.Add("Button", "xs+20 ys+24 w140", "Open log file")
        openLogBtn.OnEvent("Click", ObjBindMethod(this, "_OpenLogFile"))
        openLogFolderBtn := runWin.Add("Button", "x+8 w140", "Open logs folder")
        openLogFolderBtn.OnEvent("Click", ObjBindMethod(this, "_OpenLogFolder"))
        backBtn := runWin.Add("Button", "x+8 w140", "Back to Excel")
        backBtn.OnEvent("Click", ObjBindMethod(this, "_BackToExcel"))

        this.startBtn := runWin.Add("Button", "xm w160 h36 Default", "Start import")
        this.startBtn.OnEvent("Click", ObjBindMethod(this, "_StartImport"))
        runCancelBtn := runWin.Add("Button", "x+8 w160 h36", "Close")
        runCancelBtn.OnEvent("Click", ObjBindMethod(this, "_OnRunClose"))

        this.runGui := runWin
    }

    _SetExcelLoading(isLoading, message := "") {
        this.excelLoading := isLoading
        this.browseBtn.Enabled := !isLoading
        this.testBtn.Enabled := !isLoading
        this.okBtn.Enabled := !isLoading
        this.tableCombo.Enabled := !isLoading
        this.refreshSessionsBtn.Enabled := !isLoading
        if (!isLoading && this.sheetNames.Length > 1) {
            this.worksheetCombo.Enabled := true
        } else if (isLoading) {
            this.worksheetCombo.Enabled := false
        }
        if (!isLoading && this.sessionEntries.Length > 0) {
            this.sessionCombo.Enabled := true
        } else if (isLoading) {
            this.sessionCombo.Enabled := false
        }
        if (isLoading) {
            this.rowCountText.Value := message != "" ? message : "Loading..."
            this.previewEdit.Value := "Reading Excel... please wait."
        }
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
        if (this.excelLoading) {
            return
        }
        dataDir := Sm30AppPaths.DataDir()
        if (!DirExist(dataDir)) {
            DirCreate(dataDir)
        }
        selectedPath := FileSelect("1", dataDir, "Select Excel file", "Excel (*.xlsx; *.xlsm; *.xls)")
        if (selectedPath = "") {
            return
        }
        this.pendingExcelPath := selectedPath
        this._SetExcelLoading(true, "Opening Excel workbook...")
        SetTimer(ObjBindMethod(this, "_ProcessExcelFileLoad"), -1)
    }

    _ProcessExcelFileLoad() {
        selectedPath := this.pendingExcelPath
        this.pendingExcelPath := ""
        if (selectedPath = "") {
            this._SetExcelLoading(false)
            return
        }

        columnCount := this.tableDef.columns.Length
        try {
            result := Sm30ExcelImport.LoadWorkbook(selectedPath, "", columnCount, true)
            this.excelPath := selectedPath
            this.excelPathEdit.Value := selectedPath
            this.sheetNames := result.sheetNames
            this.selectedSheet := result.selectedSheet
            this.rows := result.rows
            this._SetWorksheetOptions(this.sheetNames, this.sheetNames.Length <= 1)
            this._UpdatePreviewFromRows()
        } catch {
            this.excelPath := ""
            this.rows := []
            this.excelPathEdit.Value := selectedPath
            this.rowCountText.Value := "Rows loaded: 0"
            this.previewEdit.Value := "Could not read Excel file."
            MsgBox("Could not read the selected Excel file.`n`nEnsure Microsoft Excel is installed.", "Excel import", "Icon!")
        } finally {
            this._SetExcelLoading(false)
        }
    }

    _SetWorksheetOptions(names, disabled := false) {
        this.worksheetCombo.Delete()
        for sheetName in names {
            this.worksheetCombo.Add([sheetName])
        }
        chooseIndex := 1
        if (this.selectedSheet != "") {
            loop names.Length {
                if (names[A_Index] = this.selectedSheet) {
                    chooseIndex := A_Index
                    break
                }
            }
        }
        this.worksheetCombo.Choose(chooseIndex)
        this.worksheetCombo.Enabled := !disabled && !this.excelLoading
    }

    _OnWorksheetChanged(*) {
        if (this.excelPath = "" || this.excelLoading) {
            return
        }
        sheetName := this.worksheetCombo.Text
        if (sheetName = "" || sheetName = "(select a file first)") {
            return
        }
        this.selectedSheet := sheetName
        this._ScheduleRowReload()
    }

    _OnTableChanged(*) {
        tableIndex := this.tableCombo.Value
        this.tableDef := Sm30TableCatalog.GetByIndex(tableIndex)
        if (this.excelPath != "" && !this.excelLoading) {
            this._ScheduleRowReload()
        }
    }

    _ScheduleRowReload() {
        if (this.excelPath = "") {
            return
        }
        this._SetExcelLoading(true, "Reading worksheet...")
        SetTimer(ObjBindMethod(this, "_ProcessRowReload"), -1)
    }

    _ProcessRowReload() {
        try {
            this._ReloadRowsNow()
        } catch {
            this.rows := []
            this.rowCountText.Value := "Rows loaded: 0"
            this.previewEdit.Value := "Could not read worksheet data."
            MsgBox("Could not read rows from the selected worksheet.", "Excel import", "Icon!")
        } finally {
            this._SetExcelLoading(false)
        }
    }

    _ReloadRowsNow() {
        if (this.excelPath = "" || !IsObject(this.tableDef)) {
            return
        }
        columnCount := this.tableDef.columns.Length
        this.rows := Sm30ExcelImport.ReadRows(
            this.excelPath,
            this.selectedSheet,
            columnCount,
            true
        )
        this._UpdatePreviewFromRows()
    }

    _UpdatePreviewFromRows() {
        this.rowCountText.Value := "Rows loaded: " this.rows.Length
        if (this.rows.Length = 0) {
            this.previewEdit.Value := "No data rows found (header row is skipped automatically)."
            return
        }
        this.previewEdit.Value := Sm30ExcelImport.PreviewRows(this.rows, 8)
    }

    _ValidateExcelStep() {
        if (this.excelLoading) {
            MsgBox("Excel file is still loading. Please wait.", "Excel import", "Icon!")
            return false
        }
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

    _ValidateSessionSelected(showMessage := true) {
        if (this.sessionEntries.Length = 0) {
            if (showMessage) {
                MsgBox("No SAP sessions available.`n`nOpen SAP GUI, log in, then click Refresh list.",
                    "SAP session", "Icon!")
            }
            return false
        }
        sessionIndex := this.sessionCombo.Value
        if (sessionIndex < 1 || sessionIndex > this.sessionEntries.Length) {
            if (showMessage) {
                MsgBox("Select a SAP session.", "SAP session", "Icon!")
            }
            return false
        }
        return true
    }

    _GetSelectedSessionEntry() {
        if (!this._ValidateSessionSelected(false)) {
            return ""
        }
        return this.sessionEntries[this.sessionCombo.Value]
    }

    _OnExcelOk(*) {
        if (!this._ValidateExcelStep()) {
            return
        }
        if (!this._ValidateSessionSelected()) {
            return
        }
        this.selectedSessionEntry := this._GetSelectedSessionEntry()
        this.excelGui.Hide()
        this._OpenRunWindow()
    }

    _OpenRunWindow() {
        sessionLabel := this.selectedSessionEntry.label
        summary := "File: " this.excelPath "`n"
            . "Worksheet: " this.selectedSheet "`n"
            . "Table: " this.tableDef.label "`n"
            . "SAP session: " sessionLabel "`n"
            . "Rows to import: " this.rows.Length "`n"
            . "Log file: " this.logPath
        this.runSummaryText.Value := summary
        this.progressBar.Value := 0
        this.statusText.Value := "Ready to import " this.rows.Length " rows."
        this.startBtn.Enabled := true
        this.runGui.Show()
    }

    _RefreshSessions(*) {
        if (this.excelLoading) {
            return
        }
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
        this.sessionCombo.Enabled := !this.excelLoading
    }

    _TestFirstRow(*) {
        if (!this._ValidateExcelStep()) {
            return
        }
        if (!this._ValidateSessionSelected()) {
            return
        }

        sessionEntry := this._GetSelectedSessionEntry()
        testRow := [this.rows[1].Clone()]
        hookPolicy := LoggingSapHookPolicy(this.logger)
        loader := ""
        try {
            loader := Sm30BulkLoader.FromSession(sessionEntry.session, hookPolicy)
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
                . "Session: " sessionEntry.label "`n"
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
        if (!IsObject(this.selectedSessionEntry)) {
            MsgBox("No SAP session selected. Go back and choose a session.", "Import", "Icon!")
            return
        }

        this.importRunning := true
        this.startBtn.Enabled := false
        this.progressBar.Value := 0
        this.statusText.Value := "Starting import..."

        sessionEntry := this.selectedSessionEntry
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

    _OnExcelResize(senderGui, minMax, width, height, *) {
        if (minMax = -1) {
            return
        }
    }
}
