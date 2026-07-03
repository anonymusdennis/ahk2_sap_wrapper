#Requires AutoHotkey v2.0

#Include CodeReviewScriptParser.ahk
#Include SapWrapper.ahk
#Include Sm30SapSessions.ahk

; Build macro groups, navigate transports, execute SAP recorder steps.
class CodeReviewStepEngine {
    __New(session := "", policy := "") {
        this.session := session
        this.policy := IsObject(policy) ? policy : SapHookPolicy()
        this.transports := []
        this.microSteps := []
        this.macroSteps := []
        this.transportIndex := 1
        this.macroIndex := 0
        this.microIndex := 0
        this.skippedMacroIndexes := Map()
    }

    LoadScriptText(scriptText) {
        parsed := CodeReviewScriptParser.ParseText(scriptText)
        this.transports := parsed.transports
        this.microSteps := parsed.microSteps
        this.macroSteps := CodeReviewStepEngine._BuildMacroSteps(this.microSteps)
        this.transportIndex := 1
        this.macroIndex := 0
        this.microIndex := 0
        this.skippedMacroIndexes := Map()
        return this.GetState()
    }

    SetSession(session) {
        this.session := session
    }

    GetState() {
        currentMacro := this.macroIndex >= 1 && this.macroIndex <= this.macroSteps.Length
            ? this.macroSteps[this.macroIndex] : ""
        nextMacros := []
        previewCount := 0
        idx := this.macroIndex + 1
        while (idx <= this.macroSteps.Length && previewCount < 4) {
            if (!this.skippedMacroIndexes.Has(idx)) {
                nextMacros.Push(this.macroSteps[idx])
                previewCount += 1
            }
            idx += 1
        }
        currentMicro := this._CurrentMicroStep()
        return {
            transportIndex: this.transportIndex,
            transportCount: this.transports.Length,
            transportId: this.transports.Length > 0 ? this.transports[this.transportIndex] : "",
            macroIndex: this.macroIndex,
            macroCount: this.macroSteps.Length,
            macroLabel: IsObject(currentMacro) ? currentMacro.label : "(not started)",
            macroGroup: IsObject(currentMacro) ? currentMacro.group : "",
            microIndex: this.microIndex,
            microCount: this.microSteps.Length,
            microLabel: IsObject(currentMicro) ? currentMicro.label : "",
            nextMacros: nextMacros,
            skippedCount: this.skippedMacroIndexes.Count
        }
    }

    StartReview() {
        this.macroIndex := 0
        this.microIndex := 0
        return this.GetState()
    }

    NextMacro() {
        if (this.macroSteps.Length = 0) {
            return this.GetState()
        }
        target := this._FindNextMacroIndex(this.macroIndex + 1, 1)
        if (target = 0) {
            return this.GetState()
        }
        this._ExecuteMacroRange(this.macroIndex + 1, target)
        this.macroIndex := target
        this.microIndex := this.macroSteps[target].endMicro
        return this.GetState()
    }

    PrevMacro() {
        if (this.macroSteps.Length = 0 || this.macroIndex <= 0) {
            return this.GetState()
        }
        target := this._FindNextMacroIndex(this.macroIndex - 1, -1)
        if (target = 0) {
            this.macroIndex := 0
            this.microIndex := 0
            return this.GetState()
        }
        this.macroIndex := target
        this.microIndex := this.macroSteps[target].startMicro - 1
        return this.GetState()
    }

    SkipNextMacro() {
        target := this._FindNextMacroIndex(this.macroIndex + 1, 1)
        if (target = 0) {
            return this.GetState()
        }
        this.skippedMacroIndexes[target] := true
        endMicro := this.macroSteps[target].endMicro
        this.macroIndex := target
        this.microIndex := endMicro
        return this.GetState()
    }

    NextMicro() {
        nextIndex := this.microIndex + 1
        if (nextIndex > this.microSteps.Length) {
            return this.GetState()
        }
        this._ExecuteMicroStep(nextIndex)
        this.microIndex := nextIndex
        this.macroIndex := this._MacroIndexForMicro(nextIndex)
        return this.GetState()
    }

    PrevMicro() {
        if (this.microIndex <= 0) {
            return this.GetState()
        }
        prevIndex := this.microIndex - 1
        this.microIndex := prevIndex
        this.macroIndex := this._MacroIndexForMicro(prevIndex)
        return this.GetState()
    }

    NextTransport() {
        if (this.transportIndex >= this.transports.Length) {
            return this.GetState()
        }
        this.transportIndex += 1
        this.macroIndex := 0
        this.microIndex := 0
        this.skippedMacroIndexes := Map()
        return this.GetState()
    }

    PrevTransport() {
        if (this.transportIndex <= 1) {
            return this.GetState()
        }
        this.transportIndex -= 1
        this.macroIndex := 0
        this.microIndex := 0
        this.skippedMacroIndexes := Map()
        return this.GetState()
    }

    _CurrentMicroStep() {
        if (this.microIndex >= 1 && this.microIndex <= this.microSteps.Length) {
            return this.microSteps[this.microIndex]
        }
        return ""
    }

    _MacroIndexForMicro(microOneBased) {
        loop this.macroSteps.Length {
            macro := this.macroSteps[A_Index]
            if (microOneBased >= macro.startMicro && microOneBased <= macro.endMicro) {
                return A_Index
            }
        }
        return 0
    }

    _FindNextMacroIndex(fromIndex, direction) {
        idx := fromIndex
        while (idx >= 1 && idx <= this.macroSteps.Length) {
            if (!this.skippedMacroIndexes.Has(idx)) {
                return idx
            }
            idx += direction
        }
        return 0
    }

    _ExecuteMacroRange(fromMacro, toMacro) {
        if (!IsObject(this.session)) {
            throw Error("No SAP session selected.")
        }
        startMicro := this.macroSteps[fromMacro].startMicro
        endMicro := this.macroSteps[toMacro].endMicro
        idx := startMicro
        while (idx <= endMicro) {
            this._ExecuteMicroStep(idx)
            idx += 1
        }
    }

    _ExecuteMicroStep(microOneBased) {
        if (!IsObject(this.session)) {
            throw Error("No SAP session selected.")
        }
        step := this.microSteps[microOneBased]
        CodeReviewStepExecutor.RunStep(this.session, step)
    }

    static _BuildMacroSteps(microSteps) {
        macroSteps := []
        if (microSteps.Length = 0) {
            return macroSteps
        }
        currentGroup := ""
        startMicro := 1
        labels := []
        loop microSteps.Length {
            step := microSteps[A_Index]
            group := CodeReviewStepClassifier.MacroGroupForKind(step.kind)
            if (currentGroup = "") {
                currentGroup := group
            }
            if (group != currentGroup) {
                macroSteps.Push(CodeReviewStepEngine._MakeMacro(currentGroup, startMicro, A_Index - 1, labels, microSteps))
                currentGroup := group
                startMicro := A_Index
                labels := []
            }
            labels.Push(step.label)
        }
        macroSteps.Push(CodeReviewStepEngine._MakeMacro(currentGroup, startMicro, microSteps.Length, labels, microSteps))
        return macroSteps
    }

    static _MakeMacro(group, startMicro, endMicro, labels, microSteps) {
        headline := labels.Length > 0 ? labels[1] : "Step"
        if (labels.Length > 1) {
            headline := headline " (+ " (labels.Length - 1) " actions)"
        }
        return {
            group: group,
            label: headline,
            startMicro: startMicro,
            endMicro: endMicro,
            stepCount: endMicro - startMicro + 1,
            preview: CodeReviewStepEngine._JoinLabels(labels, 3)
        }
    }

    static _JoinLabels(labels, maxCount) {
        out := ""
        count := 0
        for label in labels {
            count += 1
            if (count > maxCount) {
                out .= "`n..."
                break
            }
            if (out != "") {
                out .= "`n"
            }
            out .= "- " label
        }
        return out
    }
}

class CodeReviewStepExecutor {
    static NormalizeMember(member) {
        known := Map(
            "findbyid", "FindById",
            "sendvkey", "SendVKey",
            "setfocus", "SetFocus",
            "caretposition", "CaretPosition",
            "text", "Text",
            "press", "Press",
            "expandnode", "ExpandNode",
            "topnode", "TopNode",
            "selectednode", "SelectedNode",
            "doubleclicknode", "DoubleClickNode",
            "select", "Select",
            "currentcellcolumn", "CurrentCellColumn",
            "clickcurrentcell", "ClickCurrentCell",
            "maximize", "Maximize"
        )
        key := StrLower(member)
        if (known.Has(key)) {
            return known[key]
        }
        return member
    }

    static RunStep(session, step) {
        target := session.FindById(step.elementId)
        member := CodeReviewStepExecutor.NormalizeMember(step.member)
        if (step.op = "set") {
            value := CodeReviewStepExecutor._ParseValue(step.value)
            target.%member% := value
            CodeReviewStepExecutor._RunKindBehavior(session, target, step)
            return
        }
        args := CodeReviewStepExecutor._ParseArgs(step.args)
        if (args.Length = 0) {
            target.%member%()
        } else if (args.Length = 1) {
            target.%member%(args[1])
        } else if (args.Length = 2) {
            target.%member%(args[1], args[2])
        } else {
            target.%member%(args[1], args[2], args[3])
        }
        CodeReviewStepExecutor._RunKindBehavior(session, target, step)
    }

    static _RunKindBehavior(session, target, step) {
        switch step.kind {
            case "editor_view", "editor_scroll":
                CodeReviewStepExecutor._TryEditorScroll(target)
            case "code_history":
                CodeReviewStepExecutor._TrySendVKey(session, 86)
            case "view_name":
                CodeReviewStepExecutor._TrySendVKey(session, 71)
            default:
                return
        }
    }

    static _TryEditorScroll(target) {
        try {
            pos := target.VerticalScrollbar
            target.VerticalScrollbar := pos + 120
        } catch {
        }
    }

    static _TrySendVKey(session, keyCode) {
        try {
            session.FindById("wnd[0]").SendVKey(keyCode)
        } catch {
        }
    }

    static _ParseValue(rawValue) {
        text := Trim(rawValue)
        if ((SubStr(text, 1, 1) = '"' && SubStr(text, -1) = '"')
            || (SubStr(text, 1, 1) = "'" && SubStr(text, -1) = "'")) {
            return SubStr(text, 2, StrLen(text) - 2)
        }
        if (RegExMatch(text, "^\d+$")) {
            return Integer(text)
        }
        return text
    }

    static _ParseArgs(rawArgs) {
        args := []
        if (Trim(rawArgs) = "") {
            return args
        }
        text := Trim(rawArgs)
        if ((SubStr(text, 1, 1) = '"' && SubStr(text, -1) = '"')
            || (SubStr(text, 1, 1) = "'" && SubStr(text, -1) = "'")) {
            args.Push(SubStr(text, 2, StrLen(text) - 2))
            return args
        }
        if (RegExMatch(text, "^\d+$")) {
            args.Push(Integer(text))
            return args
        }
        args.Push(text)
        return args
    }
}
