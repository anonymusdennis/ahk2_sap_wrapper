#Requires AutoHotkey v2.0

; Read Excel workbooks for SM30 bulk import via Excel COM automation.
class Sm30ExcelImport {
    static LoadWorkbook(excelPath, sheetName := "", columnCount := 0, hasHeader := true) {
        excel := ""
        workbook := ""
        sheetNames := []
        rows := []
        selectedSheet := sheetName
        try {
            excel := Sm30ExcelImport._CreateExcelApp()
            workbook := excel.Workbooks.Open(excelPath, , true)
            loop workbook.Worksheets.Count {
                sheetNames.Push(workbook.Worksheets(A_Index).Name)
            }
            if (selectedSheet = "") {
                selectedSheet := sheetNames[1]
            }
            if (columnCount > 0) {
                worksheet := Sm30ExcelImport._FindWorksheet(workbook, selectedSheet)
                rows := Sm30ExcelImport._ReadWorksheetRows(worksheet, columnCount, hasHeader)
            }
        } catch {
            throw Error("Could not open Excel file: " excelPath)
        } finally {
            Sm30ExcelImport._CloseWorkbook(workbook)
            Sm30ExcelImport._QuitExcel(excel)
        }
        return {
            sheetNames: sheetNames,
            selectedSheet: selectedSheet,
            rows: rows
        }
    }

    static ListSheetNames(excelPath) {
        result := Sm30ExcelImport.LoadWorkbook(excelPath)
        return result.sheetNames
    }

    static ReadRows(excelPath, sheetName, columnCount, hasHeader := true) {
        result := Sm30ExcelImport.LoadWorkbook(excelPath, sheetName, columnCount, hasHeader)
        return result.rows
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

    static _CreateExcelApp() {
        excel := ComObject("Excel.Application")
        excel.Visible := false
        excel.DisplayAlerts := false
        excel.ScreenUpdating := false
        excel.EnableEvents := false
        return excel
    }

    static _ReadWorksheetRows(worksheet, columnCount, hasHeader) {
        rows := []
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
        return rows
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
