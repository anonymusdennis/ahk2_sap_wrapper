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
        this.tablePath := ""
        this.columns := []
        this.logger := ""
        this.logPath := ""
        this.lastFailure := ""
        this.scrollPauseMs := 30
        this.rowCreatePauseMs := 80
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

    OpenView(viewName) {
        this._Log("INFO", "OpenView start viewName=" viewName)
        wnd := this.session.FindById("wnd[0]")
        wnd.Maximize()
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
        this.table := this._FindTableControl(tableId)
        this.tablePath := tableId != "" ? tableId : this.table.Id
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
        this._Log("INFO", "FillRows start rows=" rows.Length " startAbsoluteRow=" startAbsoluteRow)
        this._LogColumns(columns)

        absoluteRow := startAbsoluteRow
        for rowValues in rows {
            this._Log("INFO", "Row start absoluteRow=" absoluteRow " values=" SapLogFormat.Args(rowValues))
            this._EnsurePhysicalRow(absoluteRow)
            visibleRow := this._EnsureVisibleRow(absoluteRow)
            this._Log("INFO", "Row mapped absoluteRow=" absoluteRow " visibleRow=" visibleRow)
            this._WriteRow(visibleRow, absoluteRow, columns, rowValues)
            absoluteRow += 1
        }

        this._Log("INFO", "FillRows done rows=" rows.Length)
        return rows.Length
    }

    Save(saveButtonId := "wnd[0]/tbar[0]/btn[11]") {
        this._Log("INFO", "Save button=" saveButtonId)
        this.session.FindById(saveButtonId).Press()
        this._WaitNotBusy()
        this._Log("INFO", "Save done")
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
                sourceIndex := columnDef.Has("csvIndex")
                    ? columnDef["csvIndex"]
                    : columnDef["index"]
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
        table := this.table
        scrollPos := table.VerticalScrollbar.Position
        visibleCount := table.VisibleRowCount

        while (absoluteRowIndex < scrollPos) {
            this._Log("INFO", "Scroll up absoluteRow=" absoluteRowIndex " scrollPos=" scrollPos)
            scrollPos -= 1
            table.VerticalScrollbar.Position := scrollPos
            this._WaitNotBusy()
            Sleep(this.scrollPauseMs)
        }

        while (absoluteRowIndex >= scrollPos + visibleCount) {
            nextPos := scrollPos + 1
            if (nextPos > table.VerticalScrollbar.Maximum) {
                this._Log("WARN", "Reached scroll max, creating row absoluteRow=" absoluteRowIndex
                    . " scrollPos=" scrollPos " max=" table.VerticalScrollbar.Maximum)
                this._CreateNewTableRow()
                scrollPos := table.VerticalScrollbar.Position
                visibleCount := table.VisibleRowCount
                continue
            }
            this._Log("INFO", "Scroll down absoluteRow=" absoluteRowIndex
                . " scrollPos=" scrollPos " -> " nextPos " visibleCount=" visibleCount)
            table.VerticalScrollbar.Position := nextPos
            scrollPos := nextPos
            this._WaitNotBusy()
            Sleep(this.scrollPauseMs)
            visibleCount := table.VisibleRowCount
        }

        visibleRow := absoluteRowIndex - scrollPos
        this._LogTableState("EnsureVisibleRow absoluteRow=" absoluteRowIndex " visibleRow=" visibleRow)
        return visibleRow
    }

    _EnsurePhysicalRow(absoluteRowIndex) {
        table := this.table
        if (absoluteRowIndex = 0 && table.RowCount <= 0) {
            this._Log("WARN", "RowCount <= 0 for first row; writing into visible slot without create")
            return
        }

        guard := 0
        while (absoluteRowIndex >= table.RowCount) {
            guard += 1
            if (guard > 1000) {
                this._Log("ERROR", "Row creation guard exceeded absoluteRow=" absoluteRowIndex
                    . " rowCount=" table.RowCount)
                throw Error("Could not create enough table rows for index " absoluteRowIndex ".")
            }
            this._Log("INFO", "Create physical row absoluteRow=" absoluteRowIndex
                . " rowCount=" table.RowCount)
            this._CreateNewTableRow()
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
        this._LogTableState("CreateNewTableRow done")
    }

    _WriteRow(visibleRowIndex, absoluteRowIndex, columns, rowValues) {
        columnCount := Min(columns.Length, rowValues.Length)
        loop columnCount {
            columnDef := columns[A_Index]
            value := rowValues[A_Index]
            cellPath := columnDef.Has("field") ? this._BuildCellPath(visibleRowIndex, columnDef) : "<GetCell>"
            try {
                this._Log("INFO", "Write cell absoluteRow=" absoluteRowIndex
                    . " visibleRow=" visibleRowIndex
                    . " column=" columnDef["index"]
                    . " kind=" (columnDef.Has("kind") ? columnDef["kind"] : "Text")
                    . " path=" cellPath
                    . " value=" value)
                cell := this._ResolveCell(visibleRowIndex, columnDef)
                cell.SetFocus()
                this._SetCellValue(cell, value, columnDef)
                this._LogReadBack(cell, columnDef, cellPath)
            } catch {
                this.lastFailure := "Failed writing absolute row " absoluteRowIndex
                    . ", visible row " visibleRowIndex
                    . ", column " columnDef["index"]
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
        if (columnDef.Has("field") && columnDef.Has("prefix")) {
            return this.session.FindById(this._BuildCellPath(visibleRowIndex, columnDef))
        }
        return this.table.GetCell(visibleRowIndex, columnDef["index"])
    }

    _BuildCellPath(visibleRowIndex, columnDef) {
        columnIndex := columnDef["index"]
        prefix := columnDef["prefix"]
        field := columnDef["field"]
        return this.tablePath "/" prefix "/" field "[" columnIndex "," visibleRowIndex "]"
    }

    _CommitRow(visibleRowIndex, columns) {
        lastColumnDef := columns[columns.Length]
        cellPath := this._BuildCellPath(visibleRowIndex, lastColumnDef)
        this._Log("INFO", "Commit row visibleRow=" visibleRowIndex " path=" cellPath)
        cell := this._ResolveCell(visibleRowIndex, lastColumnDef)
        cell.SetFocus()
        this.session.FindById("wnd[0]").SendVKey(0)
        this._WaitNotBusy()
    }

    _SetCellValue(cell, value, columnDef) {
        kind := columnDef.Has("kind") ? columnDef["kind"] : "Text"
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
        kind := columnDef.Has("kind") ? columnDef["kind"] : "Text"
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
            name := columnDef.Has("name") ? columnDef["name"] : ""
            field := columnDef.Has("field") ? columnDef["field"] : ""
            prefix := columnDef.Has("prefix") ? columnDef["prefix"] : ""
            kind := columnDef.Has("kind") ? columnDef["kind"] : "Text"
            this._Log("INFO", "Column index=" columnDef["index"]
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
