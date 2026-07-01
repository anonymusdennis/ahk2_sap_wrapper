#Requires AutoHotkey v2.0

; Read Excel workbooks for SM30 bulk import via Excel COM automation.
class Sm30ExcelImport {
    static ListSheetNames(excelPath) {
        excel := ""
        workbook := ""
        names := []
        try {
            excel := ComObject("Excel.Application")
            excel.Visible := false
            excel.DisplayAlerts := false
            workbook := excel.Workbooks.Open(excelPath, , true)
            loop workbook.Worksheets.Count {
                names.Push(workbook.Worksheets(A_Index).Name)
            }
        } catch {
            throw Error("Could not open Excel file: " excelPath)
        } finally {
            Sm30ExcelImport._CloseWorkbook(workbook)
            Sm30ExcelImport._QuitExcel(excel)
        }
        return names
    }

    static ReadRows(excelPath, sheetName, columnCount, hasHeader := true) {
        excel := ""
        workbook := ""
        rows := []
        try {
            excel := ComObject("Excel.Application")
            excel.Visible := false
            excel.DisplayAlerts := false
            workbook := excel.Workbooks.Open(excelPath, , true)
            worksheet := Sm30ExcelImport._FindWorksheet(workbook, sheetName)
            usedRows := worksheet.UsedRange.Rows.Count
            startRow := hasHeader ? 2 : 1
            if (usedRows < startRow) {
                return rows
            }

            loop usedRows - startRow + 1 {
                rowNumber := startRow + A_Index - 1
                rowValues := []
                hasData := false
                loop columnCount {
                    cellValue := worksheet.Cells(rowNumber, A_Index).Text
                    cellText := Trim(String(cellValue))
                    if (cellText != "") {
                        hasData := true
                    }
                    rowValues.Push(cellText)
                }
                if (hasData) {
                    rows.Push(rowValues)
                }
            }
        } catch {
            throw Error("Could not read Excel sheet '" sheetName "' from: " excelPath)
        } finally {
            Sm30ExcelImport._CloseWorkbook(workbook)
            Sm30ExcelImport._QuitExcel(excel)
        }
        return rows
    }

    static PreviewRows(rows, maxLines := 5) {
        lines := []
        previewCount := Min(maxLines, rows.Length)
        loop previewCount {
            row := rows[A_Index]
            lines.Push(A_Index ".  " Sm30ExcelImport._JoinFields(row, " | "))
        }
        if (rows.Length > previewCount) {
            lines.Push("... (" (rows.Length - previewCount) " more rows)")
        }
        return Sm30ExcelImport._JoinFields(lines, "`n")
    }

    static _FindWorksheet(workbook, sheetName) {
        if (sheetName = "") {
            return workbook.Worksheets(1)
        }
        loop workbook.Worksheets.Count {
            sheet := workbook.Worksheets(A_Index)
            if (sheet.Name = sheetName) {
                return sheet
            }
        }
        throw Error("Worksheet not found: " sheetName)
    }

    static _CloseWorkbook(workbook) {
        if (!IsObject(workbook)) {
            return
        }
        try {
            workbook.Close(false)
        } catch {
        }
    }

    static _QuitExcel(excel) {
        if (!IsObject(excel)) {
            return
        }
        try {
            excel.Quit()
        } catch {
        }
    }

    static _JoinFields(fields, delimiter := ", ") {
        output := ""
        for field in fields {
            if (output != "") {
                output .= delimiter
            }
            output .= field
        }
        return output
    }
}
