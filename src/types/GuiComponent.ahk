#Requires AutoHotkey v2.0

class GuiComponent extends SapComProxy {
    __New(comObj, policy := "", strict := false, path := "") {
        super.__New(comObj, "GuiComponent", path = "" ? "GuiComponent" : path, policy, strict)
    }
}
