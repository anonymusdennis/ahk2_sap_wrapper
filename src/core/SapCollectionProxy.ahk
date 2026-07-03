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
        state := Map("index", 0, "count", this.Count, "collection", this)
        if (numberOfVars >= 2) {
            return ObjBindMethod(this, "_EnumNext2", state)
        }
        return ObjBindMethod(this, "_EnumNext1", state)
    }

    _EnumNext1(state, &item) {
        if (state["index"] >= state["count"]) {
            return false
        }
        item := state["collection"]._GetItemAt(state["index"])
        state["index"] += 1
        return true
    }

    _EnumNext2(state, &index, &item) {
        if (state["index"] >= state["count"]) {
            return false
        }
        index := state["index"]
        item := state["collection"]._GetItemAt(index)
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
