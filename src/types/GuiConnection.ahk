#Requires AutoHotkey v2.0

class GuiConnection extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "/app/con[0]") {
        super.__New(comObj, "GuiConnection", path, policy, strict)
    }

    Children {
        get {
            return this.InvokeGet("Children")
        }
    }

    Sessions {
        get {
            return this.InvokeGet("Sessions")
        }
    }
}
