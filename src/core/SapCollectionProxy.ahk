#Requires AutoHotkey v2.0

class SapCollectionProxy extends SapComProxy {
    __Item[index] {
        get {
            return this._GetItemAt(index)
        }
    }

    Count {
        get {
            return this.InvokeGet("Count")
        }
    }

    Length {
        get {
            return this.InvokeGet("Length")
        }
    }

    ; Enables `for item in collection` and `for index, item in collection`
    ; (indexes are 0-based, matching SAP collection semantics).
    __Enum(numberOfVars) {
        state := Map("index", 0, "count", this.Count, "collection", this, "vars", numberOfVars)
        return ObjBindMethod(SapCollectionProxy, "_EnumNext", state)
    }

    static _EnumNext(state, &var1, args*) {
        if (state["index"] >= state["count"]) {
            return false
        }
        currentIndex := state["index"]
        item := state["collection"]._GetItemAt(currentIndex)
        if (state["vars"] >= 2 && args.Length >= 1) {
            var1 := currentIndex
            %args[1]% := item
        } else {
            var1 := item
        }
        state["index"] += 1
        return true
    }

    ; Item() with ElementAt() fallback. The probe is done on the raw COM
    ; object so On_Error only fires when both accessors fail.
    _GetItemAt(index) {
        itemWorks := true
        try {
            this._com.Item(index)
        } catch {
            itemWorks := false
        }
        if (itemWorks) {
            return this.InvokeCall("Item", index)
        }
        return this.InvokeCall("ElementAt", index)
    }
}
