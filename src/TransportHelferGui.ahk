#Requires AutoHotkey v2.0

#Include TransportHelfer.ahk

class TransportHelferGui {
    __New() {
        this.policy := SapHookPolicy()
        this.helper := TransportHelfer("", this.policy)
        this.sessionEntries := []
        this._BuildWindow()
        this._EnableHotkeys(true)
    }

    Show() {
        this._RefreshSessions()
        this.mainWin.Show()
    }

    _BuildWindow() {
        mainWin := Gui("+Resize +MinSize640x480", "Transport Helfer")
        mainWin.SetFont("s10", "Segoe UI")
        mainWin.OnEvent("Close", ObjBindMethod(this, "_OnClose"))

        mainWin.Add("Text", "w600", "Numpad 1 = open transport in SE01   |   Numpad 2 = release   |   Numpad 3 = next transport")

        mainWin.Add("GroupBox", "xm w620 h90 Section", "SAP session")
        mainWin.Add("Text", "xs+20 ys+24 w70", "Session:")
        this.sessionCombo := mainWin.Add("DropDownList", "x+0 w420 Choose1", ["(no SAP sessions found)"])
        refreshBtn := mainWin.Add("Button", "x+8 w100", "Refresh")
        refreshBtn.OnEvent("Click", ObjBindMethod(this, "_RefreshSessions"))

        mainWin.Add("GroupBox", "xm w620 h150 Section", "Transports")
        mainWin.Add("Text", "xs+20 ys+20 w580", "Paste IDs (one per line or recorder script):")
        this.transportEdit := mainWin.Add("Edit", "xs w580 h60 +VScroll", "")
        loadBtn := mainWin.Add("Button", "xs w120", "Load list")
        loadBtn.OnEvent("Click", ObjBindMethod(this, "_LoadTransports"))

        mainWin.Add("GroupBox", "xm w620 h130 Section", "Active")
        this.transportList := mainWin.Add("ListBox", "xs+20 ys+24 w580 h70")
        this.transportList.OnEvent("Change", ObjBindMethod(this, "_OnTransportSelected"))
        this.summaryText := mainWin.Add("Text", "xs w580", "No transports loaded.")
        this.statusText := mainWin.Add("Text", "xs w580 h40 +Wrap", "Ready.")

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
        this._SyncSession()
    }

    _SyncSession() {
        if (this.sessionEntries.Length = 0) {
            this.helper.SetSession("")
            return
        }
        chosen := this.sessionCombo.Text
        for entry in this.sessionEntries {
            if (entry.label = chosen) {
                this.helper.SetSession(entry.session)
                return
            }
        }
        this.helper.SetSession(this.sessionEntries[1].session)
    }

    _LoadTransports(*) {
        this._SyncSession()
        try {
            this.helper.LoadTransportsFromText(this.transportEdit.Value)
            this._RenderList()
            this.statusText.Value := "Loaded " this.helper.transports.Length " transport(s)."
        } catch {
            this.statusText.Value := "Could not load transports."
        }
    }

    _RenderList() {
        this.transportList.Delete()
        idx := 0
        for transportId in this.helper.transports {
            idx += 1
            marker := idx = this.helper.transportIndex ? ">> " : "   "
            this.transportList.Add([marker transportId])
        }
        if (this.helper.transportIndex >= 1) {
            this.transportList.Choose(this.helper.transportIndex)
        }
        state := this.helper.GetState()
        if (state.transportCount = 0) {
            this.summaryText.Value := "No transports loaded."
        } else {
            this.summaryText.Value := "Active: " state.transportIndex " / " state.transportCount
                . " — " state.transportId
        }
    }

    _OnTransportSelected(*) {
        chosenIndex := this.transportList.Value
        if (chosenIndex >= 1) {
            this.helper.SetTransportIndex(chosenIndex)
            this._RenderList()
        }
    }

    _OpenTransport(*) {
        this._RunSapAction(ObjBindMethod(this.helper, "OpenCurrentTransport"))
    }

    _ReleaseTransport(*) {
        this._RunSapAction(ObjBindMethod(this.helper, "ReleaseCurrentTransport"))
    }

    _NextTransport(*) {
        this.helper.NextTransport()
        this._RenderList()
        state := this.helper.GetState()
        this.statusText.Value := "Selected transport " state.transportId "."
    }

    _RunSapAction(actionFn) {
        this._SyncSession()
        try {
            result := actionFn.Call()
            this.statusText.Value := result.message
        } catch {
            this.statusText.Value := "SAP action failed. Check session and screen state."
        }
    }

    _EnableHotkeys(enable) {
        if (enable) {
            Hotkey("Numpad1", ObjBindMethod(this, "_OpenTransport"))
            Hotkey("Numpad2", ObjBindMethod(this, "_ReleaseTransport"))
            Hotkey("Numpad3", ObjBindMethod(this, "_NextTransport"))
            return
        }
        Hotkey("Numpad1", "Off")
        Hotkey("Numpad2", "Off")
        Hotkey("Numpad3", "Off")
    }

    _OnClose(*) {
        this._EnableHotkeys(false)
        ExitApp()
    }
}
