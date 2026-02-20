#Requires AutoHotkey v2.0

class SapTypeRegistry {
    static _allowlists := SapGeneratedAllowlists.GetAllowlists()
    static _inherits := SapGeneratedAllowlists.GetInheritance()
    static _resolved := Map()

    static DetectTypeName(comObj, fallback := "GuiUnknown") {
        try {
            typeName := comObj.Type
            if (typeName != "") {
                return typeName
            }
        } catch {
        }
        return fallback
    }

    static GetAllowlist(typeName) {
        if (this._resolved.Has(typeName)) {
            return this._resolved[typeName]
        }

        out := Map()
        this._MergeAllowlist(typeName, out)
        this._resolved[typeName] := out
        return out
    }

    static _MergeAllowlist(typeName, out) {
        if (this._allowlists.Has(typeName)) {
            allow := this._allowlists[typeName]
            for memberName, _ in allow {
                out[memberName] := true
            }
        }

        if (this._inherits.Has(typeName)) {
            baseType := this._inherits[typeName]
            if (baseType != "") {
                this._MergeAllowlist(baseType, out)
            }
        }
    }
}
