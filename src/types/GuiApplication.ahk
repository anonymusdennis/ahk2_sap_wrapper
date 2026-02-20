#Requires AutoHotkey v2.0

class GuiApplication extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "/app") {
        super.__New(comObj, "GuiApplication", path, policy, strict)
    }

    Children {
        get {
            return this.InvokeGet("Children")
        }
    }

    Connections {
        get {
            return this.InvokeGet("Connections")
        }
    }

    OpenConnection(description, sync := unset) {
        if IsSet(sync) {
            return this.InvokeCall("OpenConnection", description, sync)
        }
        return this.InvokeCall("OpenConnection", description)
    }
}
