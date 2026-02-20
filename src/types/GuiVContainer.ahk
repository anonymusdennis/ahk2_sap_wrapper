#Requires AutoHotkey v2.0

class GuiVContainer extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "") {
        super.__New(comObj, "GuiVContainer", path = "" ? "GuiVContainer" : path, policy, strict)
    }

    Children {
        get {
            return this.InvokeGet("Children")
        }
    }
}
