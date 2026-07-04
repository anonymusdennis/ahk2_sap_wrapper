#Requires AutoHotkey v2.0

; Load JSON table definitions from config/tables/*.json (no JsonLoad / v2.1 required).
class Sm30JsonConfig {
    static LoadText(jsonText) {
        jsonText := Trim(jsonText, "`r`n `t")
        if (SubStr(jsonText, 1, 1) = Chr(0xFEFF)) {
            jsonText := SubStr(jsonText, 2)
        }
        if (jsonText = "") {
            throw Error("JSON text is empty.")
        }
        parser := Sm30JsonParser(jsonText)
        parsed := parser.Parse()
        return Sm30JsonConfig._ToPlainObject(parsed)
    }

    static LoadFile(jsonPath) {
        if (!FileExist(jsonPath)) {
            throw Error("JSON file not found: " jsonPath)
        }
        jsonText := Sm30JsonConfig._ReadFileText(jsonPath)
        tableDef := Sm30JsonConfig.LoadText(jsonText)
        Sm30JsonConfig._ValidateTableDef(tableDef, jsonPath)
        return tableDef
    }

    static LoadAllFromDir(tablesDir) {
        tables := []
        if (!DirExist(tablesDir)) {
            throw Error("Table config folder not found: " tablesDir)
        }
        loop files tablesDir "\*.json", "F" {
            tableDef := Sm30JsonConfig.LoadFile(A_LoopFileFullPath)
            if (!Sm30JsonConfig._HasProp(tableDef, "id") || tableDef.id = "") {
                SplitPath(A_LoopFileFullPath, &baseName, , &ext)
                tableDef.id := baseName
            }
            tables.Push(tableDef)
        }
        if (tables.Length = 0) {
            throw Error("No table configs found in: " tablesDir)
        }
        return tables
    }

    static _ReadFileText(path) {
        if (FileGetSize(path) = 0) {
            throw Error("Config file is empty: " path)
        }
        file := ""
        try {
            file := FileOpen(path, "r", "UTF-8")
        } catch {
        }
        if (!IsObject(file)) {
            try {
                file := FileOpen(path, "r")
            } catch {
            }
        }
        if (!IsObject(file)) {
            throw Error("Could not open config file: " path)
        }
        content := file.Read()
        file.Close()
        if (content = "") {
            throw Error("Config file read returned no text: " path)
        }
        return content
    }

    static _ValidateTableDef(tableDef, jsonPath) {
        valueType := Type(tableDef)
        if (valueType != "Object") {
            throw Error("Root JSON value must be an object, got '" valueType "': " jsonPath)
        }
        required := ["label", "viewName", "tableId", "columns"]
        for fieldName in required {
            if (!Sm30JsonConfig._HasProp(tableDef, fieldName)) {
                throw Error("Missing required field '" fieldName "' in " jsonPath)
            }
        }
        if (Type(tableDef.columns) != "Array" || tableDef.columns.Length = 0) {
            throw Error("'columns' must be a non-empty array in " jsonPath)
        }
    }

    static _HasProp(obj, name) {
        valueType := Type(obj)
        if (valueType = "Map") {
            return obj.Has(name)
        }
        try {
            return obj.HasOwnProp(name)
        } catch {
            return false
        }
    }

    static _ToPlainObject(value) {
        valueType := Type(value)
        if (valueType = "Array") {
            converted := []
            for item in value {
                converted.Push(Sm30JsonConfig._ToPlainObject(item))
            }
            return converted
        }
        if (valueType = "Object") {
            ; Parser output is already a plain object (_SetProperty). On v2.1-alpha,
            ; for-in over Object() throws "Value not enumerable".
            return value
        }
        if (valueType = "Map") {
            converted := Object()
            for key, item in value {
                Sm30JsonConfig._SetProperty(converted, key, Sm30JsonConfig._ToPlainObject(item))
            }
            return converted
        }
        return value
    }

    static _SetProperty(obj, key, value) {
        obj.%Sm30JsonConfig._SanitizeKey(key)% := value
    }

    ; JSON keys are used as AHK property names. Keys that are not valid
    ; identifiers (e.g. SAP paths like "/WUE/FIELD-NAME") are sanitized
    ; instead of aborting the whole config load: invalid characters become
    ; "_" and a leading digit gets a "_" prefix.
    static _SanitizeKey(key) {
        if (RegExMatch(key, "^[A-Za-z_]\w*$")) {
            return key
        }
        sanitized := RegExReplace(key, "\W", "_")
        if (RegExMatch(sanitized, "^\d")) {
            sanitized := "_" sanitized
        }
        if (sanitized = "") {
            sanitized := "_"
        }
        return sanitized
    }
}

class Sm30JsonParser {
    static lastError := ""

    __New(text) {
        this.text := text
        this.len := StrLen(text)
        this.pos := 1
    }

    Parse() {
        Sm30JsonParser.lastError := ""
        this._SkipWhitespace()
        value := this._ReadValue()
        this._SkipWhitespace()
        if (this.pos <= this.len) {
            this._Fail("Unexpected JSON text at position " this.pos)
        }
        return value
    }

    _Fail(message) {
        snippet := Sm30JsonParser._Snippet(this.text, this.pos)
        Sm30JsonParser.lastError := message " Context: " snippet
        throw Error(Sm30JsonParser.lastError)
    }

    static _Snippet(text, pos) {
        start := Max(1, pos - 24)
        return "..." SubStr(text, start, 48) "..."
    }

    _SkipWhitespace() {
        while (this.pos <= this.len) {
            ch := SubStr(this.text, this.pos, 1)
            if (ch != " " && ch != "`t" && ch != "`r" && ch != "`n") {
                break
            }
            this.pos += 1
        }
    }

    _PeekChar() {
        this._SkipWhitespace()
        if (this.pos > this.len) {
            return ""
        }
        return SubStr(this.text, this.pos, 1)
    }

    _ReadValue() {
        ch := this._PeekChar()
        if (ch = "{") {
            return this._ReadObject()
        }
        if (ch = "[") {
            return this._ReadArray()
        }
        if (ch = '"') {
            return this._ReadString()
        }
        if (ch = "t" || ch = "f") {
            return this._ReadBool()
        }
        if (ch = "n") {
            return this._ReadNull()
        }
        if (ch = "-" || this._IsDigit(ch)) {
            return this._ReadNumber()
        }
        this._Fail("Invalid JSON value at position " this.pos)
    }

    _ReadObject() {
        obj := Object()
        this.pos += 1
        this._SkipWhitespace()
        if (this._PeekChar() = "}") {
            this.pos += 1
            return obj
        }
        loop {
            this._SkipWhitespace()
            key := this._ReadString()
            this._SkipWhitespace()
            if (SubStr(this.text, this.pos, 1) != ":") {
                this._Fail("Expected ':' in JSON object at position " this.pos)
            }
            this.pos += 1
            Sm30JsonConfig._SetProperty(obj, key, this._ReadValue())
            this._SkipWhitespace()
            ch := SubStr(this.text, this.pos, 1)
            if (ch = "}") {
                this.pos += 1
                break
            }
            if (ch != ",") {
                this._Fail("Expected ',' or '}' in JSON object at position " this.pos)
            }
            this.pos += 1
        }
        return obj
    }

    _ReadArray() {
        arr := []
        this.pos += 1
        this._SkipWhitespace()
        if (this._PeekChar() = "]") {
            this.pos += 1
            return arr
        }
        loop {
            arr.Push(this._ReadValue())
            this._SkipWhitespace()
            ch := SubStr(this.text, this.pos, 1)
            if (ch = "]") {
                this.pos += 1
                break
            }
            if (ch != ",") {
                this._Fail("Expected ',' or ']' in JSON array at position " this.pos)
            }
            this.pos += 1
        }
        return arr
    }

    _ReadString() {
        if (SubStr(this.text, this.pos, 1) != '"') {
            this._Fail("Expected JSON string at position " this.pos)
        }
        this.pos += 1
        result := ""
        while (this.pos <= this.len) {
            ch := SubStr(this.text, this.pos, 1)
            if (ch = '"') {
                this.pos += 1
                return result
            }
            if (ch = "\") {
                this.pos += 1
                if (this.pos > this.len) {
                    this._Fail("Invalid JSON string escape at position " this.pos)
                }
                esc := SubStr(this.text, this.pos, 1)
                if (esc = "n") {
                    result .= "`n"
                } else if (esc = "r") {
                    result .= "`r"
                } else if (esc = "t") {
                    result .= "`t"
                } else if (esc = '"') {
                    result .= '"'
                } else if (esc = "\") {
                    result .= "\"
                } else if (esc = "/") {
                    result .= "/"
                } else if (esc = "u") {
                    hex := SubStr(this.text, this.pos + 1, 4)
                    result .= Chr("0x" hex)
                    this.pos += 4
                } else {
                    result .= esc
                }
                this.pos += 1
                continue
            }
            result .= ch
            this.pos += 1
        }
        this._Fail("Unterminated JSON string")
    }

    _ReadBool() {
        if (SubStr(this.text, this.pos, 4) = "true") {
            this.pos += 4
            return true
        }
        if (SubStr(this.text, this.pos, 5) = "false") {
            this.pos += 5
            return false
        }
        this._Fail("Invalid JSON boolean at position " this.pos)
    }

    _ReadNull() {
        if (SubStr(this.text, this.pos, 4) = "null") {
            this.pos += 4
            return ""
        }
        this._Fail("Invalid JSON null at position " this.pos)
    }

    _IsDigit(ch) {
        return ch != "" && InStr("0123456789", ch)
    }

    _ReadNumber() {
        start := this.pos
        if (SubStr(this.text, this.pos, 1) = "-") {
            this.pos += 1
        }
        while (this.pos <= this.len) {
            ch := SubStr(this.text, this.pos, 1)
            if (!this._IsDigit(ch)) {
                break
            }
            this.pos += 1
        }
        if (this.pos <= this.len && SubStr(this.text, this.pos, 1) = ".") {
            this.pos += 1
            while (this.pos <= this.len) {
                ch := SubStr(this.text, this.pos, 1)
                if (!this._IsDigit(ch)) {
                    break
                }
                this.pos += 1
            }
        }
        numText := SubStr(this.text, start, this.pos - start)
        if (numText = "" || numText = "-") {
            this._Fail("Invalid JSON number at position " start)
        }
        if (InStr(numText, ".")) {
            return Float(numText)
        }
        return Integer(numText)
    }
}
