#Requires AutoHotkey v2.0

class GuiCollection extends SapCollectionProxy {
    __New(comObj, policy := "", strict := false, path := "") {
        super.__New(comObj, "GuiCollection", path = "" ? "GuiCollection" : path, policy, strict)
    }
}
