#Requires AutoHotkey v2.0

class GuiSession extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "/app/con[0]/ses[0]") {
        super.__New(comObj, "GuiSession", path, policy, strict)
    }

    ActiveWindow {
        get {
            return this.InvokeGet("ActiveWindow")
        }
    }

    Info {
        get {
            return this.InvokeGet("Info")
        }
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

    SendVKey(vkey) {
        return this.InvokeCall("SendVKey", vkey)
    }

    StartTransaction(tcode) {
        return this.InvokeCall("StartTransaction", tcode)
    }
}
