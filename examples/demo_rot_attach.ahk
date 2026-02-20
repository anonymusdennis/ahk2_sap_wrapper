#Requires AutoHotkey v2.0
#Include ../src/SapWrapper.ahk

policy := SapHookPolicy()
SapGuiAuto := ComObjGet("SAPGUI")
appCom := SapGuiAuto.GetScriptingEngine

app := GuiApplication(appCom, policy)
con := app.Children[0]
ses := con.Children[0]

try {
    ses.FindById("wnd[0]/tbar[0]/okcd").Text := "/nSE16"
    ses.FindById("wnd[0]").SendVKey(0)
    MsgBox("System name: " ses.Info.SystemName)
} catch {
    MsgBox("Demo failed while interacting with SAP GUI.")
}
