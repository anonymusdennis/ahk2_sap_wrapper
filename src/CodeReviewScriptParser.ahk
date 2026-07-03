#Requires AutoHotkey v2.0

; Extract transport IDs from pasted SAP recorder scripts or plain lists.
class CodeReviewScriptParser {
    static ParseText(scriptText) {
        transports := []
        lines := StrSplit(scriptText, "`n", "`r")
        for line in lines {
            trimmed := Trim(line)
            if (trimmed = "") {
                continue
            }
            transportId := CodeReviewScriptParser._ExtractFromLine(trimmed)
            if (transportId != "" && !CodeReviewScriptParser._ListContains(transports, transportId)) {
                transports.Push(transportId)
            }
        }
        return {
            transports: transports
        }
    }

    static ParseFile(path) {
        content := ""
        try {
            file := FileOpen(path, "r", "UTF-8")
            content := file.Read()
            file.Close()
        } catch {
            try {
                content := FileRead(path)
            } catch {
                throw Error("Could not read file: " path)
            }
        }
        return CodeReviewScriptParser.ParseText(content)
    }

    static _ExtractFromLine(line) {
        if (RegExMatch(line, 'i)TR_TRKORR[^"]*"\)\.text\s*=\s*"([^"]+)"', &matchQuoted)) {
            return matchQuoted[1]
        }
        if (RegExMatch(line, 'i)TR_TRKORR[^"]*"\)\.text\s*=\s*([A-Z0-9]+)', &matchBare)) {
            return matchBare[1]
        }
        if (RegExMatch(line, "^[A-Z0-9]{8,}$")) {
            return line
        }
        return ""
    }

    static _ListContains(list, value) {
        for item in list {
            if (item = value) {
                return true
            }
        }
        return false
    }
}
