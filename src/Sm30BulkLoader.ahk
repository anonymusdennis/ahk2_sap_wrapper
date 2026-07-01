#Requires AutoHotkey v2.0

#Include SapWrapper.ahk
#Include core/SapFileLogger.ahk

; Bulk-load rows into any SM30 maintenance view using SAP GUI Scripting.
; Handles vertical scrolling so writes never target rows outside the visible window.
class Sm30BulkLoader {
    __New(session, policy := "") {
        this.session := session
        this.policy := IsObject(policy) ? policy : SapHookPolicy()
        this.table := ""
        this.tableFindId := ""
        this.tablePath := ""
        this.columns := []
        this.logger := ""
        this.logPath := ""
        this.lastFailure := ""
        this.lastSkipCount := 0
        this.progressCallback := ""
        this.fillMode := "page"
        this.verboseCellLogging := false
        this.autoRecoverErrors := true
        this.skipErrorButtonId := "wnd[0]/tbar[1]/btn[20]"
        this.scrollPauseMs := 30
        this.rowCreatePauseMs := 80
        this.expandWindowForThroughput := true
        this.workingPaneWidth := 0
        this.workingPaneHeight := 2000
    }

    static Attach(policy := "") {
        hookPolicy := IsObject(policy) ? policy : SapHookPolicy()
        sapGuiAuto := ComObjGet("SAPGUI")
        appCom := sapGuiAuto.GetScriptingEngine
        app := GuiApplication(appCom, hookPolicy)
        session := app.Children[0].Children[0]
        loader := Sm30BulkLoader(session, hookPolicy)
        loader._Log("INFO", "Attached to SAP session")
        return loader
    }

    static FromSession(session, policy := "") {
        return Sm30BulkLoader(session, policy)
    }

    EnableLogging(logPath := "") {
        if (IsObject(this.logger)) {
            this._Log("INFO", "Logging already enabled: " this.logPath)
            return this
        }
        this.logger := SapFileLogger(logPath)
        this.logPath := this.logger.logPath
        this._Log("INFO", "Logging enabled")
        return this
    }

    SetLogger(logger) {
        if (IsObject(logger)) {
            this.logger := logger
            this.logPath := logger.logPath
        }
        return this
    }

    SetFillMode(mode) {
        if (mode != "page" && mode != "row") {
            throw Error('Fill mode must be "page" or "row".')
        }
        this.fillMode := mode
        this.verboseCellLogging := (mode = "row")
        return this
    }

    SetQuietComLogging(quiet := true) {
        if (this.policy.HasOwnProp("quiet")) {
            this.policy.quiet := quiet
        }
        return this
    }

    SetErrorRecovery(enabled := true, skipButtonId := "wnd[0]/tbar[1]/btn[20]") {
        this.autoRecoverErrors := enabled
        this.skipErrorButtonId := skipButtonId
        return this
    }

    SetExpandWindowForThroughput(enabled := true, width := 0, height := 0) {
        this.expandWindowForThroughput := enabled
        this.workingPaneWidth := width
        this.workingPaneHeight := height
        return this
    }

    SetProgressCallback(callback) {
        this.progressCallback := callback
        return this
    }

    OpenView(viewName) {
        this._Log("INFO", "OpenView start viewName=" viewName)
        wnd := this.session.FindById("wnd[0]")
        this._ExpandWindowForThroughput(wnd)
        this.session.FindById("wnd[0]/tbar[0]/okcd").Text := "/nsm30"
        wnd.SendVKey(0)
        this._WaitNotBusy()

        this.session.FindById("wnd[0]/usr/ctxtVIEWNAME").Text := viewName
        wnd.SendVKey(0)
        this._WaitNotBusy()
        this._Log("INFO", "OpenView done viewName=" viewName)
        return this
    }

    EnterMaintenance(updateButtonId := "wnd[0]/usr/btnUPDATE_PUSH") {
        this._Log("INFO", "EnterMaintenance button=" updateButtonId)
        this.session.FindById(updateButtonId).Press()
        this._WaitNotBusy()
        return this
    }

    NewEntries(newEntriesButtonId := "wnd[0]/tbar[1]/btn[5]") {
        this._Log("INFO", "NewEntries button=" newEntriesButtonId)
        this.session.FindById(newEntriesButtonId).Press()
        this._WaitNotBusy()
        return this
    }

    UseTable(tableId := "") {
        this._Log("INFO", "UseTable tableId=" tableId)
        this.tableFindId := tableId
        this.table := this._FindTableControl(tableId)
        if (this.tableFindId = "") {
            this.tableFindId := this._RelativeFindById(this.table.Id)
        }
        this.tablePath := this.tableFindId
        this._LogTableState("UseTable ready")
        return this
    }

    FillRows(columns, rows, startAbsoluteRow := 0) {
        if (!IsObject(this.table)) {
            throw Error("Call UseTable() before FillRows().")
        }
        if (columns.Length = 0) {
            throw Error("At least one column definition is required.")
        }
        if (rows.Length = 0) {
            return 0
        }

        this.columns := columns
        this._Log("INFO", "FillRows start rows=" rows.Length " mode=" this.fillMode " startAbsoluteRow=" startAbsoluteRow)
        this._LogColumns(columns)
        this._ReportProgress(0, rows.Length, "Starting fill")

        wasQuiet := false
        if (this.fillMode = "page" && this.policy.HasOwnProp("quiet")) {
            wasQuiet := this.policy.quiet
            this.policy.quiet := true
        }

        try {
            if (this.fillMode = "row") {
                filledCount := this._FillRowsByRow(columns, rows, startAbsoluteRow)
            } else {
                filledCount := this._FillRowsByPage(columns, rows, startAbsoluteRow)
            }
        } catch {
            if (this.policy.HasOwnProp("quiet")) {
                this.policy.quiet := wasQuiet
            }
            throw
        }

        if (this.policy.HasOwnProp("quiet")) {
            this.policy.quiet := wasQuiet
        }

        this._Log("INFO", "FillRows done rows=" filledCount)
        this._ReportProgress(filledCount, rows.Length, "Fill complete")
        return filledCount
    }

    _FillRowsByRow(columns, rows, startAbsoluteRow) {
        absoluteRow := startAbsoluteRow
        for rowValues in rows {
            this._RefreshTable()
            this._Log("INFO", "Row start absoluteRow=" absoluteRow " values=" SapLogFormat.Args(rowValues))
            this._EnsurePhysicalRow(absoluteRow)
            visibleRow := this._EnsureRowReadyForFill(absoluteRow, columns)
            this._Log("INFO", "Row mapped absoluteRow=" absoluteRow " visibleRow=" visibleRow)
            this._WriteRow(visibleRow, absoluteRow, columns, rowValues)
            skipped := this.lastSkipCount
            absoluteRow += 1 - skipped
            this._ReportProgress(A_Index, rows.Length, "Row " A_Index " committed")
        }
        return rows.Length
    }

    _FillRowsByPage(columns, rows, startAbsoluteRow) {
        dataIndex := 1
        absoluteRow := startAbsoluteRow
        this.lastSkipCount := 0
        totalRows := rows.Length

        while (dataIndex <= totalRows) {
            this._RefreshTable()
            this._EnsurePhysicalRow(absoluteRow)

            remaining := totalRows - dataIndex + 1
            pagePlan := this._EnsurePageReadyForFill(absoluteRow, remaining, columns)
            visibleStart := pagePlan["visibleStart"]
            rowsOnPage := pagePlan["rowsOnPage"]

            absoluteEnd := absoluteRow + rowsOnPage - 1
            this._Log("INFO", "Fill page absoluteRows=" absoluteRow "-" absoluteEnd
                . " visibleStart=" visibleStart " count=" rowsOnPage)

            this._FillPageColumnMajor(visibleStart, rowsOnPage, columns, rows, dataIndex, absoluteRow)
            lastVisibleRow := visibleStart + rowsOnPage - 1
            skipped := this._CommitPage(lastVisibleRow, columns, absoluteEnd)

            dataIndex += rowsOnPage
            absoluteRow += rowsOnPage - skipped
            if (skipped > 0) {
                this._Log("INFO", "Skipped " skipped " error entries; continue at absoluteRow="
                    . absoluteRow " dataIndex=" dataIndex)
            }
            this._ReportProgress(dataIndex - 1, totalRows, "Page committed")
        }

        return totalRows
    }

    _FillPageColumnMajor(visibleStartRow, rowCount, columns, rows, dataStartIndex, absoluteStartRow) {
        for colArrayIdx, columnDef in columns {
            loop rowCount {
                visibleRow := visibleStartRow + A_Index - 1
                dataIdx := dataStartIndex + A_Index - 1
                absoluteRowIndex := absoluteStartRow + A_Index - 1
                rowValues := rows[dataIdx]
                if (colArrayIdx > rowValues.Length) {
                    continue
                }
                value := rowValues[colArrayIdx]
                this._WriteCellFast(visibleRow, absoluteRowIndex, columnDef, value)
            }
        }
    }

    _WriteCellFast(visibleRowIndex, absoluteRowIndex, columnDef, value) {
        cellPath := columnDef.HasOwnProp("field") ? this._BuildCellPath(visibleRowIndex, columnDef) : "<GetCell>"
        try {
            if (this.verboseCellLogging) {
                this._Log("INFO", "Write cell absoluteRow=" absoluteRowIndex
                    . " visibleRow=" visibleRowIndex
                    . " column=" columnDef.index
                    . " path=" cellPath
                    . " value=" value)
            }
            cell := this._ResolveCell(visibleRowIndex, columnDef)
            if (!this._IsCellChangeable(cell)) {
                throw Error("Cell is not open for input")
            }
            this._SetCellValue(cell, value, columnDef)
            if (this.verboseCellLogging) {
                this._LogReadBack(cell, columnDef, cellPath)
            }
        } catch {
            this.lastFailure := "Failed writing absolute row " absoluteRowIndex
                . ", visible row " visibleRowIndex
                . ", column " columnDef.index
                . ", path " cellPath
                . ", value " value
                . ", LastError=" A_LastError
            this._Log("ERROR", this.lastFailure)
            throw Error(this.lastFailure)
        }
    }

    _SyncScrollForAbsoluteRow(absoluteRowIndex) {
        this._RefreshTable()
        table := this.table
        metrics := this._ReadTableScrollMetrics(table)
        visibleCount := metrics["visibleCount"]
        maxScroll := metrics["maxScroll"]
        scrollPos := metrics["scrollPos"]

        desiredScrollPos := absoluteRowIndex
        if (desiredScrollPos > maxScroll) {
            desiredScrollPos := maxScroll
        }
        if (absoluteRowIndex - desiredScrollPos >= visibleCount) {
            desiredScrollPos := absoluteRowIndex - visibleCount + 1
            if (desiredScrollPos < 0) {
                desiredScrollPos := 0
            }
            if (desiredScrollPos > maxScroll) {
                desiredScrollPos := maxScroll
            }
        }

        if (desiredScrollPos != scrollPos) {
            this._Log("INFO", "Scroll to absoluteRow=" absoluteRowIndex " scrollPos=" desiredScrollPos)
            table.VerticalScrollbar.Position := desiredScrollPos
            this._WaitNotBusy()
            Sleep(this.scrollPauseMs)
            this._RefreshTable()
            table := this.table
            metrics := this._ReadTableScrollMetrics(table)
            scrollPos := metrics["scrollPos"]
            visibleCount := metrics["visibleCount"]
        }

        visibleRow := absoluteRowIndex - scrollPos
        if (visibleRow < 0 || visibleRow >= visibleCount) {
            this.lastFailure := "Visible row " visibleRow " out of bounds after scroll sync"
                . " absoluteRow=" absoluteRowIndex
                . " scrollPos=" scrollPos
                . " visibleCount=" visibleCount
            this._Log("ERROR", this.lastFailure)
            throw Error(this.lastFailure)
        }

        this._LogTableState("SyncScroll absoluteRow=" absoluteRowIndex " visibleRow=" visibleRow)
        return visibleRow
    }

    _ReadTableScrollMetrics(table := "") {
        if (!IsObject(table)) {
            this._RefreshTable()
            table := this.table
        }
        scrollPos := 0
        maxScroll := 0
        visibleCount := 0
        try {
            scrollPos := table.VerticalScrollbar.Position
            maxScroll := table.VerticalScrollbar.Maximum
            visibleCount := table.VisibleRowCount
        } catch {
        }
        return Map(
            "scrollPos", scrollPos,
            "maxScroll", maxScroll,
            "visibleCount", visibleCount
        )
    }

    _ScrollTableBy(deltaRows) {
        if (deltaRows = 0) {
            return 0
        }
        this._RefreshTable()
        table := this.table
        metrics := this._ReadTableScrollMetrics(table)
        scrollPos := metrics["scrollPos"]
        maxScroll := metrics["maxScroll"]
        newPos := scrollPos + deltaRows
        if (newPos < 0) {
            newPos := 0
        }
        if (newPos > maxScroll) {
            newPos := maxScroll
        }
        if (newPos = scrollPos) {
            return 0
        }
        this._Log("INFO", "Adjust scroll by " deltaRows " from " scrollPos " to " newPos)
        table.VerticalScrollbar.Position := newPos
        this._WaitNotBusy()
        Sleep(this.scrollPauseMs)
        this._RefreshTable()
        return newPos - scrollPos
    }

    _IsVisibleRowChangeable(visibleRowIndex, columns, absoluteRowIndex := "") {
        for columnDef in columns {
            cellPath := this._BuildCellPath(visibleRowIndex, columnDef)
            try {
                cell := this._ResolveCell(visibleRowIndex, columnDef)
                if (!this._IsCellChangeable(cell)) {
                    if (absoluteRowIndex != "") {
                        this._Log("WARN", "Cell not changeable absoluteRow=" absoluteRowIndex
                            . " visibleRow=" visibleRowIndex
                            . " column=" columnDef.index
                            . " path=" cellPath)
                    }
                    return false
                }
            } catch {
                if (absoluteRowIndex != "") {
                    this._Log("WARN", "Cell not reachable absoluteRow=" absoluteRowIndex
                        . " visibleRow=" visibleRowIndex
                        . " column=" columnDef.index
                        . " path=" cellPath)
                }
                return false
            }
        }
        return true
    }

    _CountTrailingNonWritableRows(visibleStartRow, rowCount, columns, absoluteStartRow) {
        trailing := 0
        loop rowCount {
            visibleRow := visibleStartRow + rowCount - A_Index
            absoluteRowIndex := absoluteStartRow + rowCount - A_Index
            if (this._IsVisibleRowChangeable(visibleRow, columns, absoluteRowIndex)) {
                break
            }
            trailing += 1
        }
        return trailing
    }

    _EnsurePageReadyForFill(absoluteRow, remaining, columns) {
        attempt := 0
        targetAbsoluteRow := absoluteRow
        while (attempt < 8) {
            attempt += 1
            visibleStart := this._SyncScrollForAbsoluteRow(targetAbsoluteRow)
            visibleCount := this._ReadTableScrollMetrics()["visibleCount"]
            rowsOnPage := Min(visibleCount - visibleStart, remaining)
            if (rowsOnPage <= 0) {
                throw Error("Could not fit any rows on screen at absolute row " targetAbsoluteRow ".")
            }
            if (visibleStart + rowsOnPage > visibleCount) {
                rowsOnPage := visibleCount - visibleStart
            }

            if (this._ValidatePageChangeable(visibleStart, rowsOnPage, columns, targetAbsoluteRow)) {
                return Map("visibleStart", visibleStart, "rowsOnPage", rowsOnPage)
            }

            trailingBad := this._CountTrailingNonWritableRows(
                visibleStart, rowsOnPage, columns, targetAbsoluteRow)
            if (trailingBad > 0) {
                this._Log("WARN", "Trailing non-writable rows=" trailingBad
                    . " at absoluteRow=" targetAbsoluteRow " attempt=" attempt)
                this._ScrollTableBy(-trailingBad)
                continue
            }

            this._Log("WARN", "Page cells not writable at absoluteRow=" targetAbsoluteRow
                . " attempt=" attempt " visibleCount=" visibleCount)
            this._ScrollTableBy(-1)
        }

        this.lastFailure := "Target page cells are not open for input at absolute row " absoluteRow
        throw Error(this.lastFailure)
    }

    _EnsureRowReadyForFill(absoluteRow, columns) {
        attempt := 0
        targetAbsoluteRow := absoluteRow
        while (attempt < 8) {
            attempt += 1
            visibleRow := this._SyncScrollForAbsoluteRow(targetAbsoluteRow)
            if (this._IsVisibleRowChangeable(visibleRow, columns, targetAbsoluteRow)) {
                return visibleRow
            }
            this._Log("WARN", "Row cells not writable at absoluteRow=" targetAbsoluteRow
                . " attempt=" attempt)
            this._ScrollTableBy(-1)
        }

        this.lastFailure := "Target row cells are not open for input at absolute row " absoluteRow
        throw Error(this.lastFailure)
    }

    _ValidatePageChangeable(visibleStartRow, rowCount, columns, absoluteStartRow) {
        loop rowCount {
            visibleRow := visibleStartRow + A_Index - 1
            absoluteRowIndex := absoluteStartRow + A_Index - 1
            if (!this._IsVisibleRowChangeable(visibleRow, columns, absoluteRowIndex)) {
                return false
            }
        }
        return true
    }

    _IsCellChangeable(cell) {
        try {
            changeable := cell.Changeable
            return (changeable = true || changeable = 1 || changeable = -1)
        } catch {
            return false
        }
    }

    _PressEnterAfterRecovery(visibleRowIndex := -1) {
        this._Log("INFO", "Enter after error recovery")
        if (this.columns.Length > 0 && visibleRowIndex >= 0) {
            try {
                lastColumnDef := this.columns[this.columns.Length]
                cell := this._ResolveCell(visibleRowIndex, lastColumnDef)
                cell.SetFocus()
            } catch {
            }
        }
        this.session.FindById("wnd[0]").SendVKey(0)
        this._WaitNotBusy()
        this._RefreshTable()
    }

    _CommitPage(visibleRowIndex, columns, absoluteRowIndex := "") {
        lastColumnDef := columns[columns.Length]
        cellPath := this._BuildCellPath(visibleRowIndex, lastColumnDef)
        this._Log("INFO", "Commit page visibleRow=" visibleRowIndex
            . " absoluteRow=" absoluteRowIndex
            . " path=" cellPath)
        cell := this._ResolveCell(visibleRowIndex, lastColumnDef)
        cell.SetFocus()
        this.session.FindById("wnd[0]").SendVKey(0)
        this._WaitNotBusy()
        this._RefreshTable()
        skipped := this._RecoverFromSapErrors()
        this.lastSkipCount := skipped
        if (skipped > 0) {
            this._PressEnterAfterRecovery(visibleRowIndex)
        }
        return skipped
    }

    Save(saveButtonId := "wnd[0]/tbar[0]/btn[11]") {
        this._Log("INFO", "Save button=" saveButtonId)
        this.session.FindById(saveButtonId).Press()
        this._WaitNotBusy()
        skipped := this._RecoverFromSapErrors()
        if (skipped > 0) {
            this._PressEnterAfterRecovery()
        }
        this._Log("INFO", "Save done skippedErrors=" skipped)
        return this
    }

    LoadCsv(csvPath, columns, hasHeader := true, delimiter := ",") {
        rows := []
        lineNumber := 0
        for line in StrSplit(FileRead(csvPath), "`n", "`r") {
            lineNumber += 1
            trimmed := Trim(line)
            if (trimmed = "") {
                continue
            }
            if (hasHeader && lineNumber = 1) {
                continue
            }

            fields := this._ParseCsvLine(trimmed, delimiter)
            rowValues := []
            for columnDef in columns {
                sourceIndex := columnDef.HasOwnProp("csvIndex")
                    ? columnDef.csvIndex
                    : columnDef.index
                if (sourceIndex >= fields.Length) {
                    rowValues.Push("")
                } else {
                    rowValues.Push(fields[sourceIndex + 1])
                }
            }
            rows.Push(rowValues)
        }
        this._Log("INFO", "LoadCsv path=" csvPath " rows=" rows.Length)
        return rows
    }

    _ExpandWindowForThroughput(wnd) {
        if (!this.expandWindowForThroughput) {
            try {
                wnd.Maximize()
            } catch {
            }
            return
        }

        try {
            wnd.Maximize()
        } catch {
        }

        width := this.workingPaneWidth
        height := this.workingPaneHeight
        if (width <= 0) {
            width := 9999
        }
        if (height <= 0) {
            height := 2000
        }

        try {
            wnd.ResizeWorkingPaneEx(width, height, false)
            this._Log("INFO", "Expanded working pane width=" width " height=" height)
        } catch {
            try {
                wnd.ResizeWorkingPane(width, height, false)
                this._Log("INFO", "Expanded working pane (legacy) width=" width " height=" height)
            } catch {
                this._Log("WARN", "Could not expand working pane; using maximize only")
            }
        }
    }

    _RelativeFindById(fullId) {
        pos := InStr(fullId, "wnd[")
        if (pos) {
            return SubStr(fullId, pos)
        }
        return fullId
    }

    _RefreshTable() {
        if (this.tableFindId = "") {
            throw Error("No table id stored; call UseTable() first.")
        }
        this.table := this.session.FindById(this.tableFindId)
        if (this.table.Type != "GuiTableControl") {
            throw Error("Refreshed control is not a GuiTableControl: " this.tableFindId)
        }
        this._Log("INFO", "Refreshed table reference id=" this.tableFindId)
    }

    _FindTableControl(tableId) {
        if (tableId != "") {
            table := this.session.FindById(tableId)
            if (table.Type != "GuiTableControl") {
                this._Log("ERROR", "Control is not GuiTableControl type=" table.Type " id=" tableId)
                throw Error("Control is not a GuiTableControl: " tableId)
            }
            this._Log("INFO", "Found table by id=" tableId " resolvedId=" table.Id)
            return table
        }

        usr := this.session.FindById("wnd[0]/usr")
        childCount := usr.Children.Length
        loop childCount {
            child := usr.Children[A_Index - 1]
            this._Log("INFO", "Scan usr child type=" child.Type " id=" child.Id)
            if (child.Type = "GuiTableControl") {
                return child
            }
        }
        this._Log("ERROR", "No GuiTableControl found under wnd[0]/usr")
        throw Error("No GuiTableControl found under wnd[0]/usr.")
    }

    _EnsureVisibleRow(absoluteRowIndex) {
        return this._SyncScrollForAbsoluteRow(absoluteRowIndex)
    }

    _EnsurePhysicalRow(absoluteRowIndex) {
        this._RefreshTable()
        rowCount := this.table.RowCount
        if (absoluteRowIndex = 0 && rowCount <= 0) {
            this._Log("WARN", "RowCount <= 0 for first row; writing into visible slot without create")
            return
        }

        guard := 0
        while (absoluteRowIndex >= rowCount) {
            guard += 1
            if (guard > 1000) {
                this._Log("ERROR", "Row creation guard exceeded absoluteRow=" absoluteRowIndex
                    . " rowCount=" rowCount)
                throw Error("Could not create enough table rows for index " absoluteRowIndex ".")
            }
            this._Log("INFO", "Create physical row absoluteRow=" absoluteRowIndex
                . " rowCount=" rowCount)
            rowCountBefore := rowCount
            this._CreateNewTableRow()
            this._RefreshTable()
            rowCount := this.table.RowCount
            if (rowCount <= rowCountBefore) {
                this._Log("ERROR", "RowCount did not increase after Enter absoluteRow="
                    . absoluteRowIndex " rowCount=" rowCount)
                throw Error("Could not create table row at index " absoluteRowIndex
                    . "; RowCount stayed at " rowCount ".")
            }
        }
    }

    _CreateNewTableRow() {
        table := this.table
        lastVisibleRow := table.VisibleRowCount - 1
        if (lastVisibleRow < 0) {
            this._Log("ERROR", "CreateNewTableRow failed: no visible rows")
            throw Error("Table has no visible rows.")
        }

        if (this.columns.Length = 0) {
            this._Log("ERROR", "CreateNewTableRow failed: no column definitions")
            throw Error("No column definitions available for row creation.")
        }

        lastColumnDef := this.columns[this.columns.Length]
        cellPath := this._BuildCellPath(lastVisibleRow, lastColumnDef)
        this._Log("INFO", "CreateNewTableRow focus+enter path=" cellPath)
        cell := this._ResolveCell(lastVisibleRow, lastColumnDef)
        cell.SetFocus()
        this.session.FindById("wnd[0]").SendVKey(0)
        this._WaitNotBusy()
        Sleep(this.rowCreatePauseMs)
        this._RefreshTable()
        this._LogTableState("CreateNewTableRow done")
    }

    _WriteRow(visibleRowIndex, absoluteRowIndex, columns, rowValues) {
        columnCount := Min(columns.Length, rowValues.Length)
        loop columnCount {
            columnDef := columns[A_Index]
            value := rowValues[A_Index]
            cellPath := columnDef.HasOwnProp("field") ? this._BuildCellPath(visibleRowIndex, columnDef) : "<GetCell>"
            try {
                this._Log("INFO", "Write cell absoluteRow=" absoluteRowIndex
                    . " visibleRow=" visibleRowIndex
                    . " column=" columnDef.index
                    . " kind=" (columnDef.HasOwnProp("kind") ? columnDef.kind : "Text")
                    . " path=" cellPath
                    . " value=" value)
                cell := this._ResolveCell(visibleRowIndex, columnDef)
                if (!this._IsCellChangeable(cell)) {
                    throw Error("Cell is not open for input")
                }
                cell.SetFocus()
                this._SetCellValue(cell, value, columnDef)
                this._LogReadBack(cell, columnDef, cellPath)
            } catch {
                this.lastFailure := "Failed writing absolute row " absoluteRowIndex
                    . ", visible row " visibleRowIndex
                    . ", column " columnDef.index
                    . ", path " cellPath
                    . ", value " value
                    . ", LastError=" A_LastError
                this._Log("ERROR", this.lastFailure)
                throw Error(this.lastFailure)
            }
        }
        this._CommitRow(visibleRowIndex, columns)
    }

    _ResolveCell(visibleRowIndex, columnDef) {
        if (columnDef.HasOwnProp("field") && columnDef.HasOwnProp("prefix")) {
            return this.session.FindById(this._BuildCellPath(visibleRowIndex, columnDef))
        }
        return this.table.GetCell(visibleRowIndex, columnDef.index)
    }

    _BuildCellPath(visibleRowIndex, columnDef) {
        columnIndex := columnDef.index
        prefix := columnDef.prefix
        field := columnDef.field
        return this.tablePath "/" prefix "/" field "[" columnIndex "," visibleRowIndex "]"
    }

    _CommitRow(visibleRowIndex, columns) {
        this._CommitPage(visibleRowIndex, columns)
    }

    _SetCellValue(cell, value, columnDef) {
        kind := columnDef.HasOwnProp("kind") ? columnDef.kind : "Text"
        if (kind = "Key") {
            cell.Key := value
            return
        }
        if (kind = "Selected") {
            cell.Selected := (value = true || value = 1 || value = "X" || value = "x")
            return
        }
        cell.Text := value
    }

    _LogReadBack(cell, columnDef, cellPath) {
        kind := columnDef.HasOwnProp("kind") ? columnDef.kind : "Text"
        readBack := "<unavailable>"
        try {
            if (kind = "Key") {
                readBack := cell.Key
            } else if (kind = "Selected") {
                readBack := cell.Selected
            } else {
                readBack := cell.Text
            }
        } catch {
            readBack := "<readback-failed LastError=" A_LastError ">"
        }
        this._Log("INFO", "Read back path=" cellPath " value=" readBack)
    }

    _LogTableState(context) {
        table := this.table
        scrollPos := ""
        scrollMax := ""
        visibleCount := ""
        rowCount := ""
        try {
            scrollPos := table.VerticalScrollbar.Position
            scrollMax := table.VerticalScrollbar.Maximum
            visibleCount := table.VisibleRowCount
            rowCount := table.RowCount
        } catch {
            scrollPos := "<err>"
        }
        this._Log("INFO", context
            . " tablePath=" this.tablePath
            . " rowCount=" rowCount
            . " visibleCount=" visibleCount
            . " scrollPos=" scrollPos
            . " scrollMax=" scrollMax)
    }

    _LogColumns(columns) {
        for columnDef in columns {
            name := columnDef.HasOwnProp("name") ? columnDef.name : ""
            field := columnDef.HasOwnProp("field") ? columnDef.field : ""
            prefix := columnDef.HasOwnProp("prefix") ? columnDef.prefix : ""
            kind := columnDef.HasOwnProp("kind") ? columnDef.kind : "Text"
            this._Log("INFO", "Column index=" columnDef.index
                . " kind=" kind
                . " prefix=" prefix
                . " field=" field
                . " name=" name)
        }
    }

    _Log(level, message) {
        if (!IsObject(this.logger)) {
            return
        }
        if (level = "ERROR") {
            this.logger.Error(message)
            return
        }
        if (level = "WARN") {
            this.logger.Warn(message)
            return
        }
        this.logger.Info(message)
    }

    _ReportProgress(completed, total, message := "") {
        if (!IsObject(this.progressCallback)) {
            return
        }
        try {
            this.progressCallback.Call(completed, total, message)
        } catch {
        }
    }

    _WaitNotBusy() {
        guard := 0
        while (this.session.Busy) {
            guard += 1
            if (guard > 10000) {
                this._Log("ERROR", "Timed out waiting for SAP session idle")
                throw Error("Timed out waiting for SAP session to become idle.")
            }
            Sleep(10)
        }
    }

    _HasSapError() {
        try {
            sbar := this.session.FindById("wnd[0]/sbar")
            msgType := sbar.MessageType
            if (msgType != "E" && msgType != "A") {
                return false
            }
            msgText := Trim(sbar.Text)
            if (msgText = "") {
                return false
            }
            return true
        } catch {
        }
        return false
    }

    _IsSkipButtonAvailable(skipButtonId) {
        try {
            btn := this.session.FindById(skipButtonId)
            if (btn.Type != "GuiButton") {
                return false
            }
            try {
                if (btn.Changeable = 0 || btn.Changeable = false) {
                    return false
                }
            } catch {
            }
            return true
        } catch {
            return false
        }
    }

    _LogSapMessage(context := "SAP message") {
        try {
            sbar := this.session.FindById("wnd[0]/sbar")
            this._Log("WARN", context
                . " type=" sbar.MessageType
                . " id=" sbar.MessageId
                . " number=" sbar.MessageNumber
                . " text=" sbar.Text)
        } catch {
            this._Log("WARN", context " (status bar unavailable)")
        }
    }

    _RecoverFromSapErrors() {
        if (!this.autoRecoverErrors) {
            return 0
        }

        skipCount := 0
        guard := 0
        while (this._HasSapError() && guard < 1000) {
            guard += 1
            if (!this._IsSkipButtonAvailable(this.skipErrorButtonId)) {
                this._LogSapMessage("SAP error but skip button unavailable")
                this.lastFailure := "SAP error shown but skip button is unavailable"
                throw Error(this.lastFailure)
            }

            this._LogSapMessage("Recovering SAP error skip=" skipCount + 1)
            this.session.FindById(this.skipErrorButtonId).Press()
            this._WaitNotBusy()
            Sleep(100)
            skipCount += 1
        }

        if (this._HasSapError()) {
            this._LogSapMessage("SAP error remains after skip recovery")
            this.lastFailure := "SAP error remains after " skipCount " skip presses"
            throw Error(this.lastFailure)
        }

        if (skipCount > 0) {
            this._Log("INFO", "Recovered from " skipCount " SAP error entries via skip button")
            if (this.tableFindId != "") {
                this._RefreshTable()
            }
        }
        return skipCount
    }

    _ParseCsvLine(line, delimiter) {
        fields := []
        current := ""
        inQuotes := false
        lineLen := StrLen(line)
        loop lineLen {
            char := SubStr(line, A_Index, 1)
            if (char = '"') {
                if (inQuotes && A_Index < lineLen && SubStr(line, A_Index + 1, 1) = '"') {
                    current .= '"'
                    continue
                }
                inQuotes := !inQuotes
                continue
            }
            if (!inQuotes && char = delimiter) {
                fields.Push(current)
                current := ""
                continue
            }
            current .= char
        }
        fields.Push(current)
        return fields
    }
}
