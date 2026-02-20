# AHK2 SAP GUI Scripting Wrapper

AutoHotkey v2 COM wrapper/proxy for SAP GUI Scripting with hookable call pipeline.

## Usage

```ahk
#Requires AutoHotkey v2.0
#Include src/SapWrapper.ahk

policy := SapHookPolicy()
SapGuiAuto := ComObjGet("SAPGUI")
app := GuiApplication(SapGuiAuto.GetScriptingEngine, policy)

ses := app.Children[0].Children[0]
ses.FindById("wnd[0]/tbar[0]/okcd").Text := "/nSE16"
ses.FindById("wnd[0]").SendVKey(0)
MsgBox(ses.Info.SystemName)
```

See `/examples/demo_rot_attach.ahk` for a runnable example.
