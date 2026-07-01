#Requires AutoHotkey v2.0

#Include SapWrapper.ahk

; Enumerate open SAP GUI sessions for picker UIs.
class Sm30SapSessions {
    static List(policy := "") {
        hookPolicy := IsObject(policy) ? policy : SapHookPolicy()
        entries := []
        try {
            sapGuiAuto := ComObjGet("SAPGUI")
            app := GuiApplication(sapGuiAuto.GetScriptingEngine, hookPolicy)
            connectionCount := app.Children.Length
            loop connectionCount {
                connection := app.Children[A_Index - 1]
                connectionIndex := A_Index
                sessionCount := connection.Children.Length
                loop sessionCount {
                    session := connection.Children[A_Index - 1]
                    sessionIndex := A_Index
                    label := Sm30SapSessions._BuildLabel(session, connectionIndex, sessionIndex)
                    entries.Push({
                        session: session,
                        label: label,
                        connectionIndex: connectionIndex,
                        sessionIndex: sessionIndex
                    })
                }
            }
        } catch {
        }
        return entries
    }

    static GetActiveSession(policy := "") {
        hookPolicy := IsObject(policy) ? policy : SapHookPolicy()
        try {
            sapGuiAuto := ComObjGet("SAPGUI")
            app := GuiApplication(sapGuiAuto.GetScriptingEngine, hookPolicy)
            return app.ActiveSession
        } catch {
        }
        return ""
    }

    static GetLabels(entries) {
        labels := []
        for entry in entries {
            labels.Push(entry.label)
        }
        return labels
    }

    static _BuildLabel(session, connectionIndex, sessionIndex) {
        infoText := "Connection " connectionIndex " / Session " sessionIndex
        try {
            info := session.Info
            infoText := info.SystemName " / Client " info.Client " / " info.User
                . " / " info.Transaction " (conn " connectionIndex ", ses " info.SessionNumber ")"
        } catch {
        }
        return infoText
    }
}
