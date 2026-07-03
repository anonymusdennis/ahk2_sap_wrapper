#Requires AutoHotkey v2.0

#Include CodeReviewRunner.ahk
#Include CodeReviewAppPaths.ahk

; Dedicated SE01 code review controller with checkpoint navigation.
class CodeReviewPlayerGui {
    __New() {
        this.policy := SapHookPolicy()
        this.runner := CodeReviewRunner("", "", this.policy)
        this.sessionEntries := []
        this.hotkeysActive := false
        this._BuildWindow()
    }

    Show() {
        this._RefreshSessions()
        this._LoadDefaultConfig()
        this.mainWin.Show()
    }

    _BuildWindow() {
        mainWin := Gui("+Resize +MinSize820x720", "SE01 Code Review")
        mainWin.SetFont("s10", "Segoe UI")
        mainWin.OnEvent("Close", ObjBindMethod(this, "_OnClose"))

        mainWin.Add("Text", "w780", "Hard-coded SE01 review flow from config/se01_review.json")
        mainWin.Add("Text", "w780 cGray", "Numpad 0 = checkpoint 0   |   Numpad 1 = checkpoint 1   |   Numpad 5 = action   |   Numpad + = continue")

        mainWin.Add("GroupBox", "xm w800 h100 Section", "SAP session")
        mainWin.Add("Text", "xs+20 ys+24 w90", "Session:")
        this.sessionCombo := mainWin.Add("DropDownList", "x+0 w620 Choose1", ["(no SAP sessions found)"])
        this.refreshSessionsBtn := mainWin.Add("Button", "xs+130 y+10 w120", "Refresh list")
        this.refreshSessionsBtn.OnEvent("Click", ObjBindMethod(this, "_RefreshSessions"))

        mainWin.Add("GroupBox", "xm w800 h170 Section", "Transport list")
        mainWin.Add("Text", "xs+20 ys+20 w760", "Paste transport IDs (recorder script or one ID per line):")
        this.transportEdit := mainWin.Add("Edit", "xs w760 h70 +VScroll", "")
        this.loadSampleBtn := mainWin.Add("Button", "xs w140", "Load sample")
        this.loadSampleBtn.OnEvent("Click", ObjBindMethod(this, "_LoadSampleTransports"))
        this.parseBtn := mainWin.Add("Button", "x+8 w140", "Load transports")
        this.parseBtn.OnEvent("Click", ObjBindMethod(this, "_LoadTransports"))

        mainWin.Add("GroupBox", "xm w800 h120 Section", "Active transport")
        this.transportList := mainWin.Add("ListBox", "xs+20 ys+24 w560 h70")
        this.transportList.OnEvent("Change", ObjBindMethod(this, "_OnTransportSelected"))
        this.prevTransportBtn := mainWin.Add("Button", "x+8 yp w100", "Previous")
        this.prevTransportBtn.OnEvent("Click", ObjBindMethod(this, "_PrevTransport"))
        this.nextTransportBtn := mainWin.Add("Button", "xp y+8 w100", "Next")
        this.nextTransportBtn.OnEvent("Click", ObjBindMethod(this, "_NextTransport"))
        this.transportSummaryText := mainWin.Add("Text", "xs w760", "Transports loaded: 0")

        mainWin.Add("GroupBox", "xm w800 h210 Section", "Review flow")
        this.stateText := mainWin.Add("Text", "xs+20 ys+24 w760", "State: Idle")
        this.statusText := mainWin.Add("Text", "xs w760 h36 +Wrap", "Ready.")
        this.continueHintText := mainWin.Add("Text", "xs w760 cNavy", "")
        this.actionHintText := mainWin.Add("Text", "xs w760 cTeal", "")
        this.checkpointHintText := mainWin.Add("Text", "xs w760 cGray", "")

        mainWin.Add("GroupBox", "xm w800 h120 Section", "Controls")
        cp0Btn := mainWin.Add("Button", "xs+20 ys+28 w150", "Checkpoint 0")
        cp0Btn.OnEvent("Click", ObjBindMethod(this, "_GoCheckpoint0"))
        cp1Btn := mainWin.Add("Button", "x+8 w150", "Checkpoint 1")
        cp1Btn.OnEvent("Click", ObjBindMethod(this, "_GoCheckpoint1"))
        actionBtn := mainWin.Add("Button", "x+8 w150", "Action (5)")
        actionBtn.OnEvent("Click", ObjBindMethod(this, "_ActionKey"))
        continueBtn := mainWin.Add("Button", "x+8 w150", "Continue (+)")
        continueBtn.OnEvent("Click", ObjBindMethod(this, "_ContinueKey"))
        this.startBtn := mainWin.Add("Button", "xs+20 y+12 w150", "Start hotkeys")
        this.startBtn.OnEvent("Click", ObjBindMethod(this, "_StartHotkeys"))

        this.mainWin := mainWin
    }

    _LoadDefaultConfig() {
        try {
            reviewDef := CodeReviewConfig.LoadDefault()
            this.runner.SetReviewDef(reviewDef)
            this.statusText.Value := "Loaded review config: " reviewDef.label
        } catch {
            this.statusText.Value := "Could not load " CodeReviewAppPaths.DefaultReviewConfigPath()
        }
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
        this._SyncSessionFromCombo()
    }

    _SyncSessionFromCombo() {
        if (this.sessionEntries.Length = 0) {
            this.runner.SetSession("")
            return
        }
        chosen := this.sessionCombo.Text
        for entry in this.sessionEntries {
            if (entry.label = chosen) {
                this.runner.SetSession(entry.session)
                return
            }
        }
        this.runner.SetSession(this.sessionEntries[1].session)
    }

    _LoadTransports(*) {
        this._SyncSessionFromCombo()
        try {
            state := this.runner.LoadTransportsFromText(this.transportEdit.Value)
            this._RenderTransportList()
            this._RenderState(state)
            this.statusText.Value := "Loaded " state.transportCount " transport(s)."
        } catch {
            this.statusText.Value := "Could not load transports."
        }
    }

    _LoadSampleTransports(*) {
        samplePath := CodeReviewAppPaths.DataDir() "\sample_transports.txt"
        try {
            file := FileOpen(samplePath, "r", "UTF-8")
            this.transportEdit.Value := file.Read()
            file.Close()
            this.statusText.Value := "Loaded sample transport list."
        } catch {
            this.transportEdit.Value := "W4DK930869"
            this.statusText.Value := "Using built-in sample transport ID."
        }
    }

    _RenderTransportList() {
        this.transportList.Delete()
        idx := 0
        for transportId in this.runner.transports {
            idx += 1
            marker := idx = this.runner.transportIndex ? ">> " : "   "
            this.transportList.Add([marker transportId])
        }
        if (this.runner.transports.Length > 0) {
            this.transportList.Choose(this.runner.transportIndex)
        }
        this.transportSummaryText.Value := "Transports loaded: " this.runner.transports.Length
            . "   Active: " this.runner.GetCurrentTransport()
    }

    _OnTransportSelected(*) {
        chosenIndex := this.transportList.Value
        if (chosenIndex >= 1) {
            this.runner.SetTransportIndex(chosenIndex)
            this._RenderTransportList()
            this._RenderState(this.runner.GetState())
        }
    }

    _PrevTransport(*) {
        this._RunAction(ObjBindMethod(this.runner, "PrevTransport"))
        this._RenderTransportList()
    }

    _NextTransport(*) {
        this._RunAction(ObjBindMethod(this.runner, "NextTransport"))
        this._RenderTransportList()
    }

    _GoCheckpoint0(*) {
        this._RunAction(ObjBindMethod(this.runner, "GoCheckpoint0"))
    }

    _GoCheckpoint1(*) {
        this._RunAction(ObjBindMethod(this.runner, "GoCheckpoint1"))
    }

    _ActionKey(*) {
        this._RunAction(ObjBindMethod(this.runner, "ActionKey"))
    }

    _ContinueKey(*) {
        this._RunAction(ObjBindMethod(this.runner, "ContinueKey"))
    }

    _StartHotkeys(*) {
        this._EnableHotkeys(true)
        this.statusText.Value := "Hotkeys active."
    }

    _EnableHotkeys(enable) {
        if (enable && !this.hotkeysActive) {
            Hotkey("Numpad0", ObjBindMethod(this, "_GoCheckpoint0"))
            Hotkey("Numpad1", ObjBindMethod(this, "_GoCheckpoint1"))
            Hotkey("Numpad5", ObjBindMethod(this, "_ActionKey"))
            Hotkey("NumpadAdd", ObjBindMethod(this, "_ContinueKey"))
            this.hotkeysActive := true
            this.startBtn.Text := "Hotkeys active"
            return
        }
        if (!enable && this.hotkeysActive) {
            Hotkey("Numpad0", "Off")
            Hotkey("Numpad1", "Off")
            Hotkey("Numpad5", "Off")
            Hotkey("NumpadAdd", "Off")
            this.hotkeysActive := false
            this.startBtn.Text := "Start hotkeys"
        }
    }

    _RunAction(actionFn) {
        this._SyncSessionFromCombo()
        try {
            state := actionFn.Call()
            this._RenderState(state)
            this.statusText.Value := state.lastMessage
        } catch {
            this.statusText.Value := "SAP action failed. Check session, screen, and scripting."
        }
    }

    _RenderState(state) {
        this.stateText.Value := "State: " state.stateLabel
        this.continueHintText.Value := state.continueHint
        this.actionHintText.Value := state.actionHint
        this.checkpointHintText.Value := state.checkpointHint
        if (state.transportCount > 0) {
            this.transportSummaryText.Value := "Transports: " state.transportIndex " / "
                state.transportCount "   Active: " state.transportId
        }
    }

    _OnClose(*) {
        this._EnableHotkeys(false)
        ExitApp()
    }
}
