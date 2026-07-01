#Requires AutoHotkey v2.1

; Load JSON table definitions from config/tables/*.json
class Sm30JsonConfig {
    static LoadText(jsonText) {
        if (!IsSet(JsonLoad)) {
            throw Error("JsonLoad requires AutoHotkey v2.1 or newer.")
        }
        parsed := JsonLoad(jsonText)
        return Sm30JsonConfig._NormalizeValue(parsed)
    }

    static LoadFile(jsonPath) {
        if (!FileExist(jsonPath)) {
            throw Error("JSON file not found: " jsonPath)
        }
        return Sm30JsonConfig.LoadText(FileRead(jsonPath, "UTF-8"))
    }

    static LoadAllFromDir(tablesDir) {
        tables := []
        if (!DirExist(tablesDir)) {
            return tables
        }
        loop files tablesDir "\*.json", "F" {
            try {
                tableDef := Sm30JsonConfig.LoadFile(A_LoopFileFullPath)
                if (!tableDef.HasOwnProp("id")) {
                    tableDef.id := A_LoopFileName
                }
                tables.Push(tableDef)
            } catch {
                throw Error("Failed to load table config: " A_LoopFileFullPath)
            }
        }
        if (tables.Length = 0) {
            throw Error("No table configs found in: " tablesDir)
        }
        return tables
    }

    static _NormalizeValue(value) {
        valueType := Type(value)
        if (valueType = "Array") {
            normalized := []
            for item in value {
                normalized.Push(Sm30JsonConfig._NormalizeValue(item))
            }
            return normalized
        }
        if (valueType = "Map") {
            normalized := {}
            for key, item in value {
                normalized.%key% := Sm30JsonConfig._NormalizeValue(item)
            }
            return normalized
        }
        return value
    }
}
