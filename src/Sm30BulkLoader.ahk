#Requires AutoHotkey v2.0

#Include SapWrapper.ahk

; Bulk-load rows into any SM30 maintenance view using SAP GUI Scripting.
; Handles vertical scrolling so writes never target rows outside the visible window.
class Sm30BulkLoader {
    __New(session, policy := "") {
        this.session := session
        this.policy := IsObject(policy) ? policy : SapHookPolicy()
        this.table := ""
        this.tablePath := ""
        this.columns := []
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
        return Sm30BulkLoader(session, hookPolicy)
    }

    static FromSession(session, policy := "") {
        return Sm30BulkLoader(session, policy)
    }

    OpenView(viewName) {
        wnd := this.session.FindById("wnd[0]")
        wnd.Maximize()
        this.session.FindById("wnd[0]/tbar[0]/okcd").Text := "/nsm30"
        wnd.SendVKey(0)
        this._WaitNotBusy()

        this.session.FindById("wnd[0]/usr/ctxtVIEWNAME").Text := viewName
        wnd.SendVKey(0)
        this._WaitNotBusy()
        return this
    }

    EnterMaintenance(updateButtonId := "wnd[0]/usr/btnUPDATE_PUSH") {
        this.session.FindById(updateButtonId).Press()
        this._WaitNotBusy()
        return this
    }

    NewEntries(newEntriesButtonId := "wnd[0]/tbar[1]/btn[5]") {
        this.session.FindById(newEntriesButtonId).Press()
        this._WaitNotBusy()
        return this
    }

    UseTable(tableId := "") {
        this.table := this._FindTableControl(tableId)
        this.tablePath := tableId != "" ? tableId : this.table.Id
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
        absoluteRow := startAbsoluteRow
        for rowValues in rows {
            this._EnsurePhysicalRow(absoluteRow)
            visibleRow := this._EnsureVisibleRow(absoluteRow)
            this._WriteRow(visibleRow, absoluteRow, columns, rowValues)
            absoluteRow += 1
        }
        return rows.Length
    }

    Save(saveButtonId := "wnd[0]/tbar[0]/btn[11]") {
        this.session.FindById(saveButtonId).Press()
        this._WaitNotBusy()
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
        return rows
    }

    _FindTableControl(tableId) {
        if (tableId != "") {
            table := this.session.FindById(tableId)
            if (table.Type != "GuiTableControl") {
                throw Error("Control is not a GuiTableControl: " tableId)
            }
            return table
        }

        usr := this.session.FindById("wnd[0]/usr")
        childCount := usr.Children.Length
        loop childCount {
            child := usr.Children[A_Index - 1]
            if (child.Type = "GuiTableControl") {
                return child
            }
        }
        throw Error("No GuiTableControl found under wnd[0]/usr.")
    }

    _EnsureVisibleRow(absoluteRowIndex) {
        table := this.table
        scrollPos := table.VerticalScrollbar.Position
        visibleCount := table.VisibleRowCount

        while (absoluteRowIndex < scrollPos) {
            scrollPos -= 1
            table.VerticalScrollbar.Position := scrollPos
            this._WaitNotBusy()
            Sleep(this.scrollPauseMs)
        }

        while (absoluteRowIndex >= scrollPos + visibleCount) {
            nextPos := scrollPos + 1
            if (nextPos > table.VerticalScrollbar.Maximum) {
                this._CreateNewTableRow()
                scrollPos := table.VerticalScrollbar.Position
                visibleCount := table.VisibleRowCount
                continue
            }
            table.VerticalScrollbar.Position := nextPos
            scrollPos := nextPos
            this._WaitNotBusy()
            Sleep(this.scrollPauseMs)
            visibleCount := table.VisibleRowCount
        }

        return absoluteRowIndex - scrollPos
    }

    _EnsurePhysicalRow(absoluteRowIndex) {
        table := this.table
        if (absoluteRowIndex = 0 && table.RowCount <= 0) {
            return
        }

        guard := 0
        while (absoluteRowIndex >= table.RowCount) {
            guard += 1
            if (guard > 1000) {
                throw Error("Could not create enough table rows for index " absoluteRowIndex ".")
            }
            this._CreateNewTableRow()
        }
    }

    _CreateNewTableRow() {
        table := this.table
        lastVisibleRow := table.VisibleRowCount - 1
        if (lastVisibleRow < 0) {
            throw Error("Table has no visible rows.")
        }

        if (this.columns.Length = 0) {
            throw Error("No column definitions available for row creation.")
        }

        lastColumnDef := this.columns[this.columns.Length]
        cell := this._ResolveCell(lastVisibleRow, lastColumnDef)
        cell.SetFocus()
        this.session.FindById("wnd[0]").SendVKey(0)
        this._WaitNotBusy()
        Sleep(this.rowCreatePauseMs)
    }

    _WriteRow(visibleRowIndex, absoluteRowIndex, columns, rowValues) {
        columnCount := Min(columns.Length, rowValues.Length)
        loop columnCount {
            columnDef := columns[A_Index]
            value := rowValues[A_Index]
            try {
                cell := this._ResolveCell(visibleRowIndex, columnDef)
                cell.SetFocus()
                this._SetCellValue(cell, value, columnDef)
            } catch {
                this.lastFailure := "Failed writing absolute row " absoluteRowIndex
                    . ", visible row " visibleRowIndex
                    . ", column " columnDef["index"]
                    . ", path " this._BuildCellPath(visibleRowIndex, columnDef)
                    . ", value " value
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

    _WaitNotBusy() {
        guard := 0
        while (this.session.Busy) {
            guard += 1
            if (guard > 10000) {
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
