#Requires AutoHotkey v2.0

class GuiFrameWindow extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "") {
        super.__New(comObj, "GuiFrameWindow", path = "" ? "GuiFrameWindow" : path, policy, strict)
    }

    SendVKey(vkey) {
        return this.InvokeCall("SendVKey", vkey)
    }
}
