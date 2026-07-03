#Requires AutoHotkey v2.0

#Include CodeReviewStepEngine.ahk
#Include Sm30AppPaths.ahk

; Step-through SAP transport code review with numpad navigation and preview UI.
class CodeReviewPlayerGui {
    __New() {
        this.policy := SapHookPolicy()
        this.engine := CodeReviewStepEngine("", this.policy)
        this.sessionEntries := []
        this.selectedSessionEntry := ""
        this.hotkeysActive := false
        this.lastError := ""
        this._BuildWindow()
    }

    Show() {
        this._RefreshSessions()
        this.mainWin.Show()
    }

    _BuildWindow() {
        mainWin := Gui("+Resize +MinSize760x640", "SAP Code Review Player")
        mainWin.SetFont("s10", "Segoe UI")
        mainWin.OnEvent("Close", ObjBindMethod(this, "_OnClose"))
        mainWin.OnEvent("Size", ObjBindMethod(this, "_OnResize"))

        mainWin.Add("Text", "w720", "Paste SAP Script Recorder output, then step through the review with the numpad.")
        mainWin.Add("Text", "w720 cGray", "Numpad + / - = next / previous step   |   / / * = next / previous transport")
        mainWin.Add("Text", "w720 cGray", "Numpad 9 / 3 = fine step forward / back   |   Numpad 6 = skip next step")

        mainWin.Add("GroupBox", "xm w740 h110 Section", "SAP session")
        mainWin.Add("Text", "xs+20 ys+24 w90", "Session:")
        this.sessionCombo := mainWin.Add("DropDownList", "x+0 w560 Choose1", ["(no SAP sessions found)"])
        this.refreshSessionsBtn := mainWin.Add("Button", "xs+130 y+10 w120", "Refresh list")
        this.refreshSessionsBtn.OnEvent("Click", ObjBindMethod(this, "_RefreshSessions"))

        mainWin.Add("GroupBox", "xm w740 h220 Section", "Script input")
        mainWin.Add("Text", "xs+20 ys+20 w700", "Paste recorder script (VBS or AHK style):")
        this.scriptEdit := mainWin.Add("Edit", "xs w700 h130 +VScroll", "")
        btnRowY := ""
        this.loadSampleBtn := mainWin.Add("Button", "xs w140", "Load sample")
        this.loadSampleBtn.OnEvent("Click", ObjBindMethod(this, "_LoadSampleScript"))
        this.parseBtn := mainWin.Add("Button", "x+8 w140", "Parse script")
        this.parseBtn.OnEvent("Click", ObjBindMethod(this, "_ParseScript"))
        this.startBtn := mainWin.Add("Button", "x+8 w140", "Start review")
        this.startBtn.OnEvent("Click", ObjBindMethod(this, "_StartReview"))

        mainWin.Add("GroupBox", "xm w740 h250 Section", "Current position")
        this.transportText := mainWin.Add("Text", "xs+20 ys+20 w700", "Transport: (none)")
        this.macroText := mainWin.Add("Text", "xs w700", "Step: (not started)")
        this.microText := mainWin.Add("Text", "xs w700", "Fine step: (none)")
        this.statusText := mainWin.Add("Text", "xs w700 h40 +Wrap", "Ready.")
        mainWin.Add("Text", "xs w700", "Coming next:")
        this.previewEdit := mainWin.Add("Edit", "xs w700 h110 ReadOnly -VScroll -HScroll", "(parse a script to see upcoming steps)")

        mainWin.Add("GroupBox", "xm w740 h70 Section", "Hotkeys")
        this.hotkeyStatusText := mainWin.Add("Text", "xs+20 ys+24 w700", "Hotkeys: inactive until review starts")
        this.toggleHotkeysBtn := mainWin.Add("Button", "xs w160", "Pause hotkeys")
        this.toggleHotkeysBtn.Enabled := false
        this.toggleHotkeysBtn.OnEvent("Click", ObjBindMethod(this, "_ToggleHotkeys"))

        this.mainWin := mainWin
    }

    _RefreshSessions(*) {
        this.sessionEntries := Sm30SapSessions.List(this.policy)
        labels := Sm30SapSessions.GetLabels(this.sessionEntries)
        if (labels.Length = 0) {
            labels := ["(no SAP sessions found)"]
        }
        this.sessionCombo.Delete()
        for label in labels {
            this.sessionCombo.Add([label])
        }
        this.sessionCombo.Choose(1)
        if (this.sessionEntries.Length > 0) {
            this.selectedSessionEntry := this.sessionEntries[1]
            this.engine.SetSession(this.selectedSessionEntry.session)
        } else {
            this.selectedSessionEntry := ""
            this.engine.SetSession("")
        }
    }

    _ParseScript(*) {
        this.lastError := ""
        try {
            if (this.sessionEntries.Length > 0) {
                chosen := this.sessionCombo.Text
                for entry in this.sessionEntries {
                    if (entry.label = chosen) {
                        this.selectedSessionEntry := entry
                        this.engine.SetSession(entry.session)
                        break
                    }
                }
            }
            state := this.engine.LoadScriptText(this.scriptEdit.Value)
            this._RenderState(state)
            this.statusText.Value := "Parsed " state.macroCount " review steps for transport "
                . state.transportId "."
        } catch {
            this.lastError := "Could not parse script."
            this.statusText.Value := this.lastError
        }
    }

    _StartReview(*) {
        this._ParseScript()
        this.engine.StartReview()
        this._EnableHotkeys(true)
        this.statusText.Value := "Review started. Use numpad keys to navigate."
        this._RenderState(this.engine.GetState())
    }

    _LoadSampleScript(*) {
        samplePath := Sm30AppPaths.DataDir() "\sample_se01_review.txt"
        try {
            parsed := CodeReviewScriptParser.ParseFile(samplePath)
            file := FileOpen(samplePath, "r", "UTF-8")
            this.scriptEdit.Value := file.Read()
            file.Close()
            this.statusText.Value := "Loaded sample script (" parsed.microSteps.Length " fine steps)."
        } catch {
            this.statusText.Value := "Sample file not found at " samplePath
        }
    }

    _RenderState(state) {
        transportLine := "Transport " state.transportIndex " / " state.transportCount
            . " — " state.transportId
        this.transportText.Value := transportLine
        macroLine := "Step " state.macroIndex " / " state.macroCount " — " state.macroLabel
        if (state.macroGroup != "") {
            macroLine .= "  [" state.macroGroup "]"
        }
        this.macroText.Value := macroLine
        microLine := "Fine step " state.microIndex " / " state.microCount
        if (state.microLabel != "") {
            microLine .= " — " state.microLabel
        }
        this.microText.Value := microLine

        previewText := ""
        if (state.nextMacros.Length = 0) {
            previewText := "(no further steps)"
        } else {
            idx := 0
            for macro in state.nextMacros {
                idx += 1
                previewText .= idx ". " macro.label "`n" macro.preview "`n`n"
            }
        }
        this.previewEdit.Value := previewText
    }

    _EnableHotkeys(enable) {
        if (enable && !this.hotkeysActive) {
            Hotkey("NumpadAdd", ObjBindMethod(this, "_HotkeyNextMacro"))
            Hotkey("NumpadSub", ObjBindMethod(this, "_HotkeyPrevMacro"))
            Hotkey("NumpadDiv", ObjBindMethod(this, "_HotkeyNextTransport"))
            Hotkey("NumpadMult", ObjBindMethod(this, "_HotkeyPrevTransport"))
            Hotkey("Numpad9", ObjBindMethod(this, "_HotkeyNextMicro"))
            Hotkey("Numpad6", ObjBindMethod(this, "_HotkeySkipMacro"))
            Hotkey("Numpad3", ObjBindMethod(this, "_HotkeyPrevMicro"))
            this.hotkeysActive := true
            this.hotkeyStatusText.Value := "Hotkeys: active (works while this tool is running)"
            this.toggleHotkeysBtn.Enabled := true
            this.toggleHotkeysBtn.Text := "Pause hotkeys"
            return
        }
        if (!enable && this.hotkeysActive) {
            Hotkey("NumpadAdd", "Off")
            Hotkey("NumpadSub", "Off")
            Hotkey("NumpadDiv", "Off")
            Hotkey("NumpadMult", "Off")
            Hotkey("Numpad9", "Off")
            Hotkey("Numpad6", "Off")
            Hotkey("Numpad3", "Off")
            this.hotkeysActive := false
            this.hotkeyStatusText.Value := "Hotkeys: paused"
            this.toggleHotkeysBtn.Text := "Resume hotkeys"
        }
    }

    _ToggleHotkeys(*) {
        if (this.hotkeysActive) {
            this._EnableHotkeys(false)
        } else {
            this._EnableHotkeys(true)
        }
    }

    _HotkeyNextMacro(*) {
        this._RunNavigationAction(ObjBindMethod(this.engine, "NextMacro"))
    }

    _HotkeyPrevMacro(*) {
        this._RunNavigationAction(ObjBindMethod(this.engine, "PrevMacro"))
    }

    _HotkeyNextTransport(*) {
        this._RunNavigationAction(ObjBindMethod(this.engine, "NextTransport"))
    }

    _HotkeyPrevTransport(*) {
        this._RunNavigationAction(ObjBindMethod(this.engine, "PrevTransport"))
    }

    _HotkeyNextMicro(*) {
        this._RunNavigationAction(ObjBindMethod(this.engine, "NextMicro"))
    }

    _HotkeyPrevMicro(*) {
        this._RunNavigationAction(ObjBindMethod(this.engine, "PrevMicro"))
    }

    _HotkeySkipMacro(*) {
        this._RunNavigationAction(ObjBindMethod(this.engine, "SkipNextMacro"))
    }

    _RunNavigationAction(actionFn) {
        this.lastError := ""
        try {
            state := actionFn.Call()
            this._RenderState(state)
            this.statusText.Value := "OK — " state.macroLabel
        } catch {
            this.lastError := "SAP step failed. Check session, screen state, and scripting."
            this.statusText.Value := this.lastError
        }
    }

    _OnClose(*) {
        this._EnableHotkeys(false)
        ExitApp()
    }

    _OnResize(senderGui, minMax, width, height, *) {
        if (minMax = -1) {
            return
        }
        margin := 20
        contentWidth := width - (margin * 2)
        if (contentWidth < 400) {
            contentWidth := 400
        }
    }
}
