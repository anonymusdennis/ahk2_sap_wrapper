#Requires AutoHotkey v2.0

#Include SapWrapper.ahk

; Enumerate open SAP GUI sessions for picker UIs.
class Sm30SapSessions {
    ; Human-readable description of the last List()/GetActiveSession()
    ; failure ("" when the last call succeeded). Lets UIs distinguish
    ; "no sessions open" from "could not attach to SAP GUI".
    static lastError := ""

    static List(policy := "") {
        hookPolicy := IsObject(policy) ? policy : SapHookPolicy()
        entries := []
        Sm30SapSessions.lastError := ""
        try {
            sapGuiAuto := ComObjGet("SAPGUI")
        } catch {
            Sm30SapSessions.lastError := "Could not attach to SAP GUI (SAPGUI object not running or scripting disabled)."
            return entries
        }
        try {
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
            Sm30SapSessions.lastError := "SAP GUI is running but sessions could not be enumerated (scripting engine error)."
        }
        return entries
    }

    static GetActiveSession(policy := "") {
        hookPolicy := IsObject(policy) ? policy : SapHookPolicy()
        Sm30SapSessions.lastError := ""
        try {
            sapGuiAuto := ComObjGet("SAPGUI")
            app := GuiApplication(sapGuiAuto.GetScriptingEngine, hookPolicy)
            return app.ActiveSession
        } catch {
            Sm30SapSessions.lastError := "Could not attach to SAP GUI (SAPGUI object not running or scripting disabled)."
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
