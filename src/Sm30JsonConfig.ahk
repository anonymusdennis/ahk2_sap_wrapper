#Requires AutoHotkey v2.0

; Load JSON table definitions from config/tables/*.json (no JsonLoad / v2.1 required).
class Sm30JsonConfig {
    static LoadText(jsonText) {
        jsonText := Trim(jsonText, "`r`n `t")
        if (SubStr(jsonText, 1, 1) = Chr(0xFEFF)) {
            jsonText := SubStr(jsonText, 2)
        }
        parser := Sm30JsonParser(jsonText)
        return parser.Parse()
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
}

class Sm30JsonParser {
    __New(text) {
        this.text := text
        this.len := StrLen(text)
        this.pos := 1
    }

    Parse() {
        this._SkipWhitespace()
        value := this._ReadValue()
        this._SkipWhitespace()
        if (this.pos <= this.len) {
            throw Error("Unexpected JSON text at position " this.pos)
        }
        return value
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
        if (ch = "-" || (ch >= "0" && ch <= "9")) {
            return this._ReadNumber()
        }
        throw Error("Invalid JSON value at position " this.pos)
    }

    _ReadObject() {
        obj := {}
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
                throw Error("Expected ':' in JSON object at position " this.pos)
            }
            this.pos += 1
            obj[key] := this._ReadValue()
            this._SkipWhitespace()
            ch := SubStr(this.text, this.pos, 1)
            if (ch = "}") {
                this.pos += 1
                break
            }
            if (ch != ",") {
                throw Error("Expected ',' or '}' in JSON object at position " this.pos)
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
                throw Error("Expected ',' or ']' in JSON array at position " this.pos)
            }
            this.pos += 1
        }
        return arr
    }

    _ReadString() {
        if (SubStr(this.text, this.pos, 1) != '"') {
            throw Error("Expected JSON string at position " this.pos)
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
                    throw Error("Invalid JSON string escape at position " this.pos)
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
        throw Error("Unterminated JSON string")
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
        throw Error("Invalid JSON boolean at position " this.pos)
    }

    _ReadNull() {
        if (SubStr(this.text, this.pos, 4) = "null") {
            this.pos += 4
            return ""
        }
        throw Error("Invalid JSON null at position " this.pos)
    }

    _ReadNumber() {
        start := this.pos
        if (SubStr(this.text, this.pos, 1) = "-") {
            this.pos += 1
        }
        while (this.pos <= this.len) {
            ch := SubStr(this.text, this.pos, 1)
            if (ch < "0" || ch > "9") {
                break
            }
            this.pos += 1
        }
        if (this.pos <= this.len && SubStr(this.text, this.pos, 1) = ".") {
            this.pos += 1
            while (this.pos <= this.len) {
                ch := SubStr(this.text, this.pos, 1)
                if (ch < "0" || ch > "9") {
                    break
                }
                this.pos += 1
            }
        }
        numText := SubStr(this.text, start, this.pos - start)
        if (InStr(numText, ".")) {
            return Float(numText)
        }
        return Integer(numText)
    }
}
