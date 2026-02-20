#Requires AutoHotkey v2.0

class SapCollectionProxy extends SapComProxy {
    __Item[index] {
        get {
            try {
                return this.InvokeCall("Item", index)
            } catch {
                return this.InvokeCall("ElementAt", index)
            }
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
}
