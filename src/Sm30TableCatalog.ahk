#Requires AutoHotkey v2.1

#Include Sm30AppPaths.ahk
#Include Sm30JsonConfig.ahk

; Load SM30 table definitions from config/tables/*.json
class Sm30TableCatalog {
    static _cache := ""

    static GetAll() {
        if (IsObject(Sm30TableCatalog._cache)) {
            return Sm30TableCatalog._cache
        }
        tablesDir := Sm30AppPaths.TablesDir()
        Sm30TableCatalog._cache := Sm30JsonConfig.LoadAllFromDir(tablesDir)
        return Sm30TableCatalog._cache
    }

    static Reload() {
        Sm30TableCatalog._cache := ""
        return Sm30TableCatalog.GetAll()
    }

    static GetLabels() {
        labels := []
        for tableDef in Sm30TableCatalog.GetAll() {
            labels.Push(tableDef.label)
        }
        return labels
    }

    static GetByIndex(index) {
        tables := Sm30TableCatalog.GetAll()
        if (index < 1 || index > tables.Length) {
            return ""
        }
        return tables[index]
    }

    static FindByViewName(viewName) {
        for tableDef in Sm30TableCatalog.GetAll() {
            if (tableDef.viewName = viewName) {
                return tableDef
            }
        }
        return ""
    }

    static FindById(tableId) {
        for tableDef in Sm30TableCatalog.GetAll() {
            if (tableDef.HasOwnProp("id") && tableDef.id = tableId) {
                return tableDef
            }
        }
        return ""
    }
}
