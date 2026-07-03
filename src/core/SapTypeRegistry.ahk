#Requires AutoHotkey v2.0

class SapTypeRegistry {
    static _allowlists := SapGeneratedAllowlists.GetAllowlists()
    static _bases := SapTypeRegistry._LoadBases()
    static _typeNumbers := SapTypeRegistry._LoadTypeNumbers()
    static _resolved := Map()

    static _LoadBases() {
        try {
            return SapGeneratedAllowlists.GetBaseTypes()
        } catch {
        }
        ; Legacy generated format: single parent per type.
        bases := Map()
        try {
            for typeName, baseName in SapGeneratedAllowlists.GetInheritance() {
                bases[typeName] := (baseName = "") ? [] : [baseName]
            }
        } catch {
        }
        return bases
    }

    static _LoadTypeNumbers() {
        try {
            return SapGeneratedTypeNumbers.GetTypeNumberMap()
        } catch {
        }
        return Map()
    }

    static DetectTypeName(comObj, fallback := "GuiUnknown") {
        try {
            typeName := comObj.Type
            if (typeName != "") {
                return typeName
            }
        } catch {
        }
        try {
            typeNumber := comObj.TypeAsNumber
            if (this._typeNumbers.Has(typeNumber)) {
                return this._typeNumbers[typeNumber]
            }
        } catch {
        }
        return fallback
    }

    static GetAllowlist(typeName) {
        if (this._resolved.Has(typeName)) {
            return this._resolved[typeName]
        }

        ; COM member access is case-insensitive, so strict-mode lookups are too.
        out := Map()
        out.CaseSense := "Off"
        this._MergeAllowlist(typeName, out, Map())
        this._resolved[typeName] := out
        return out
    }

    static _MergeAllowlist(typeName, out, seen) {
        if (typeName = "" || seen.Has(typeName)) {
            return
        }
        seen[typeName] := true

        if (this._allowlists.Has(typeName)) {
            allow := this._allowlists[typeName]
            for memberName, _ in allow {
                out[memberName] := true
            }
        }

        if (this._bases.Has(typeName)) {
            for baseType in this._bases[typeName] {
                this._MergeAllowlist(baseType, out, seen)
            }
        }
    }
}
