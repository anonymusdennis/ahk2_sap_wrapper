#Requires AutoHotkey v2.0

#Include CodeReviewStepClassifier.ahk

; Parse SAP GUI Script Recorder output (VBS or AHK style) into review steps.
class CodeReviewScriptParser {
    static ParseText(scriptText) {
        transports := []
        microSteps := []
        lines := StrSplit(scriptText, "`n", "`r")
        for line in lines {
            trimmed := Trim(line)
            if (trimmed = "" || CodeReviewScriptParser._IsCommentOrSetup(trimmed)) {
                continue
            }
            step := CodeReviewScriptParser._ParseLine(trimmed)
            if (!IsObject(step)) {
                continue
            }
            microSteps.Push(step)
            transportId := CodeReviewScriptParser._ExtractTransportId(step)
            if (transportId != "" && !CodeReviewScriptParser._ListContains(transports, transportId)) {
                transports.Push(transportId)
            }
        }
        if (transports.Length = 0 && microSteps.Length > 0) {
            transports.Push("(unknown)")
        }
        return {
            transports: transports,
            microSteps: microSteps
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
                throw Error("Could not read script file: " path)
            }
        }
        return CodeReviewScriptParser.ParseText(content)
    }

    static _IsCommentOrSetup(line) {
        if (SubStr(line, 1, 1) = "'") {
            return true
        }
        lower := StrLower(line)
        if (InStr(lower, "set sapguiauto") || InStr(lower, "set application")
            || InStr(lower, "set connection") || InStr(lower, "set session")
            || InStr(lower, "wscript.connectobject") || InStr(lower, "if not isobject")
            || InStr(lower, "if isobject")) {
            return true
        }
        return false
    }

    static _ParseLine(line) {
        id := ""
        member := ""
        args := ""
        value := ""
        op := "call"

        if (RegExMatch(line, 'i)findById\("([^"]+)"\)\.(\w+)\s*=\s*(.+)$', &matchAssign)) {
            id := matchAssign[1]
            member := matchAssign[2]
            value := CodeReviewScriptParser._TrimTrailingComment(matchAssign[3])
            op := "set"
        } else if (RegExMatch(line, 'i)findById\("([^"]+)"\)\.(\w+)\((.*)\)', &matchCall)) {
            id := matchCall[1]
            member := matchCall[2]
            args := Trim(matchCall[3])
            op := "call"
        } else if (RegExMatch(line, 'i)findById\("([^"]+)"\)\.(\w+)\s+(.+)$', &matchSpace)) {
            id := matchSpace[1]
            member := matchSpace[2]
            args := CodeReviewScriptParser._TrimTrailingComment(matchSpace[3])
            op := "call"
        } else {
            return ""
        }

        kind := CodeReviewStepClassifier.Classify(id, member, op, value, args)
        label := CodeReviewStepClassifier.BuildLabel(kind, id, member, op, value, args)
        return {
            raw: line,
            elementId: id,
            member: member,
            op: op,
            value: value,
            args: args,
            kind: kind,
            label: label
        }
    }

    static _ExtractTransportId(step) {
        if (InStr(step.elementId, "TR_TRKORR") && step.op = "set" && step.value != "") {
            return CodeReviewScriptParser._Unquote(step.value)
        }
        if (InStr(step.elementId, "TRKORR") && step.op = "set" && step.value != "") {
            return CodeReviewScriptParser._Unquote(step.value)
        }
        return ""
    }

    static _Unquote(text) {
        cleaned := Trim(text)
        if ((SubStr(cleaned, 1, 1) = '"' && SubStr(cleaned, -1) = '"')
            || (SubStr(cleaned, 1, 1) = "'" && SubStr(cleaned, -1) = "'")) {
            return SubStr(cleaned, 2, StrLen(cleaned) - 2)
        }
        return cleaned
    }

    static _TrimTrailingComment(text) {
        quote := ""
        out := ""
        loop ParseInteger(StrLen(text)) {
            ch := SubStr(text, A_Index, 1)
            if (quote = "") {
                if (ch = '"' || ch = "'") {
                    quote := ch
                } else if (ch = "'") {
                    break
                }
            } else if (ch = quote) {
                quote := ""
            }
            out .= ch
        }
        return Trim(out)
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
