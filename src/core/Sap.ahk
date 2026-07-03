#Requires AutoHotkey v2.0

; Convenience entry points for attaching to a running SAP GUI.
;
;   ses := Sap.Attach()                 ; active session
;   ses := Sap.Attach(policy)           ; active session, custom hook policy
;   ses := Sap.Attach(policy, 0, 1)     ; connection 0 / session 1
;   app := Sap.App(policy)              ; wrapped GuiApplication root
class Sap {
    ; Returns the wrapped scripting engine root (GuiApplication).
    static App(policy := "", strict := false) {
        sapGuiAuto := ComObjGet("SAPGUI")
        return GuiApplication(sapGuiAuto.GetScriptingEngine, policy, strict)
    }

    ; Returns a wrapped GuiSession. Without indexes the currently active
    ; session is used (GuiApplication.ActiveSession), replacing the old
    ; hardcoded Children[0].Children[0] pattern.
    static Attach(policy := "", connectionIndex := -1, sessionIndex := -1) {
        app := Sap.App(policy)
        if (connectionIndex < 0 && sessionIndex < 0) {
            session := app.ActiveSession
            if (IsObject(session)) {
                return session
            }
            ; Fall back to the first session if none is marked active.
            connectionIndex := 0
            sessionIndex := 0
        }
        if (connectionIndex < 0) {
            connectionIndex := 0
        }
        if (sessionIndex < 0) {
            sessionIndex := 0
        }
        return app.Children[connectionIndex].Children[sessionIndex]
    }
}
