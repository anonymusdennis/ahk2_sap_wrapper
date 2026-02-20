#Requires AutoHotkey v2.0

class GuiVComponent extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "") {
        super.__New(comObj, "GuiVComponent", path = "" ? "GuiVComponent" : path, policy, strict)
    }

    Text {
        get {
            return this.InvokeGet("Text")
        }
        set(value) {
            return this.InvokeSet("Text", value)
        }
    }
}
