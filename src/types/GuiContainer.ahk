#Requires AutoHotkey v2.0

class GuiContainer extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "") {
        super.__New(comObj, "GuiContainer", path = "" ? "GuiContainer" : path, policy, strict)
    }

    Children {
        get {
            return this.InvokeGet("Children")
        }
    }

    FindById(id, raise := unset) {
        if IsSet(raise) {
            return this.InvokeCall("FindById", id, raise)
        }
        return this.InvokeCall("FindById", id)
    }
}
