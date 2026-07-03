#Requires AutoHotkey v2.0

class SapComProxy {
    static _typeClassMap := ""

    __New(comObj, typeName := "GuiUnknown", path := "", policy := "", strict := false) {
        this.DefineProp("_com", {Value: comObj})
        this.DefineProp("_typeName", {Value: typeName})
        this.DefineProp("_path", {Value: path = "" ? typeName : path})
        this.DefineProp("_policy", {Value: IsObject(policy) ? policy : SapHookPolicy()})
        this.DefineProp("_strict", {Value: strict})
        this.DefineProp("_allow", {Value: SapTypeRegistry.GetAllowlist(typeName)})
        this.DefineProp("_staleRecovery", {Value: false})
        this.DefineProp("_findParentCom", {Value: ""})
        this.DefineProp("_findId", {Value: ""})
    }

    __Get(name, params) {
        if (SubStr(name, 1, 1) = "_") {
            if (this.HasOwnProp(name)) {
                return this.GetOwnPropDesc(name).Value
            }
            throw PropertyError("Unknown internal property.", -1, name)
        }
        value := this.InvokeGet(name)
        if (params.Length > 0) {
            return value[params*]
        }
        return value
    }

    __Set(name, params, value) {
        if (SubStr(name, 1, 1) = "_") {
            throw PropertyError("Internal property is read-only.", -1, name)
        }
        if (params.Length > 0) {
            target := this.InvokeGet(name)
            target[params*] := value
            return value
        }
        return this.InvokeSet(name, value)
    }

    __Call(name, params) {
        return this.InvokeCall(name, params*)
    }

    Raw() {
        return this._com
    }

    ; Opt-in: when a COM operation fails once, try to re-resolve this
    ; object's raw COM reference via its FindById origin and retry the
    ; operation a single time. Helps after server round trips invalidate
    ; references below window level.
    SetStaleRecovery(enabled := true) {
        this.DefineProp("_staleRecovery", {Value: enabled})
        return this
    }

    InvokeGet(member) {
        this._EnsureMemberAllowed(member)
        args := []
        this._CallPolicy("On_Call", "get", member, args)

        try {
            result := this._com.%member%
        } catch {
            if (this._TryStaleRetry("get", member)) {
                try {
                    result := this._com.%member%
                } catch {
                    this._RaiseComError("get", member, args)
                }
            } else {
                this._RaiseComError("get", member, args)
            }
        }

        wrapped := this._WrapResult(member, "get", result, args)
        this._CallPolicy("After_Call", "get", member, wrapped)
        return wrapped
    }

    InvokeSet(member, value) {
        this._EnsureMemberAllowed(member)
        args := [value]
        this._CallPolicy("On_Call", "set", member, args)

        try {
            this._com.%member% := value
        } catch {
            if (this._TryStaleRetry("set", member)) {
                try {
                    this._com.%member% := value
                } catch {
                    this._RaiseComError("set", member, args)
                }
            } else {
                this._RaiseComError("set", member, args)
            }
        }

        this._CallPolicy("After_Call", "set", member, value)
        return value
    }

    InvokeCall(member, args*) {
        this._EnsureMemberAllowed(member)
        this._CallPolicy("On_Call", "call", member, args)

        try {
            result := this._com.%member%(args*)
        } catch {
            if (this._TryStaleRetry("call", member)) {
                try {
                    result := this._com.%member%(args*)
                } catch {
                    this._RaiseComError("call", member, args)
                }
            } else {
                this._RaiseComError("call", member, args)
            }
        }

        wrapped := this._WrapResult(member, "call", result, args)
        this._CallPolicy("After_Call", "call", member, wrapped)
        return wrapped
    }

    _WrapResult(member, op, value, args) {
        if (!this._IsComObject(value)) {
            return value
        }

        typeName := SapTypeRegistry.DetectTypeName(value)
        childPath := this._BuildPath(member, op, args)
        typeClassMap := SapComProxy._GetTypeClassMap()
        if (typeClassMap.Has(typeName)) {
            proxyClass := typeClassMap[typeName]
            child := proxyClass(value, this._policy, this._strict, childPath)
        } else {
            child := SapComProxy(value, typeName, childPath, this._policy, this._strict)
        }

        if (this._staleRecovery) {
            child.DefineProp("_staleRecovery", {Value: true})
        }
        if (op = "call" && member = "FindById" && args.Length >= 1) {
            child.DefineProp("_findParentCom", {Value: this._com})
            child.DefineProp("_findId", {Value: args[1]})
        }
        return child
    }

    static _GetTypeClassMap() {
        if (IsObject(SapComProxy._typeClassMap)) {
            return SapComProxy._typeClassMap
        }
        typeClassMap := Map(
            "GuiCollection", GuiCollection,
            "GuiComponentCollection", GuiComponentCollection,
            "GuiApplication", GuiApplication,
            "GuiConnection", GuiConnection,
            "GuiSession", GuiSession,
            "GuiFrameWindow", GuiFrameWindow,
            "GuiVComponent", GuiVComponent
        )
        ; Merge generated typed wrappers when SapWrapper2.ahk is included.
        try {
            for typeName, proxyClass in SapGeneratedTypedWrappers.GetTypeClassMap() {
                if (!typeClassMap.Has(typeName)) {
                    typeClassMap[typeName] := proxyClass
                }
            }
        } catch {
        }
        SapComProxy._typeClassMap := typeClassMap
        return typeClassMap
    }

    _BuildPath(member, op, args) {
        if (op = "call" && member = "FindById" && args.Length >= 1) {
            childId := args[1]
            return this._JoinPath(this._path, childId)
        }
        if (op = "call" && member = "Item" && args.Length >= 1) {
            return this._JoinPath(this._path, "[" args[1] "]")
        }
        return this._JoinPath(this._path, member)
    }

    _JoinPath(basePath, child) {
        left := RTrim(basePath, "/")
        right := LTrim(child, "/")
        if (left = "") {
            return right
        }
        if (right = "") {
            return left
        }
        return left "/" right
    }

    _IsComObject(value) {
        try {
            t := ComObjType(value)
            return t != ""
        } catch {
            return false
        }
    }

    _EnsureMemberAllowed(member) {
        if (!this._strict) {
            return
        }
        if (this._allow.Has(member)) {
            return
        }
        throw Error("Member not allowlisted: " this._typeName "." member)
    }

    _TryStaleRetry(op, member) {
        if (!this._staleRecovery) {
            return false
        }
        if (!IsObject(this._findParentCom) || this._findId = "") {
            return false
        }
        refreshed := ""
        try {
            refreshed := this._findParentCom.FindById(this._findId)
        } catch {
            return false
        }
        if (!this._IsComObject(refreshed)) {
            return false
        }
        this.DefineProp("_com", {Value: refreshed})
        this._NotifyRetry(op, member)
        return true
    }

    _NotifyRetry(op, member) {
        try {
            if (HasMethod(this._policy, "On_Retry")) {
                this._policy.On_Retry(op, this._typeName, member, this._path, 1)
            }
        } catch {
        }
    }

    _ClassifyError(op, member, args) {
        ; When FindById fails, probe with raiseOnFailure=false to
        ; distinguish "control not found" from other COM failures.
        if (op = "call" && member = "FindById" && args.Length >= 1) {
            probe := ""
            try {
                probe := this._com.FindById(args[1], false)
            } catch {
                return "com-error"
            }
            if (!this._IsComObject(probe)) {
                return "not-found"
            }
            return "transient"
        }
        return "com-error"
    }

    _BuildError(op, member, reason := "") {
        text := "SAP COM " op " failed: " this._typeName "." member " @ " this._path
        if (reason = "not-found") {
            text .= " (control not found)"
        } else if (reason = "transient") {
            text .= " (control exists; call failed, possibly transient)"
        }
        return text " (LastError=" A_LastError ", may be unrelated for COM)"
    }

    _RaiseComError(op, member, args) {
        reason := this._ClassifyError(op, member, args)
        this._HandleError(op, member, args, reason)
        throw Error(this._BuildError(op, member, reason))
    }

    _HandleError(op, member, args, reason := "") {
        ; Optional structured hook (SapHookPolicy2 or any duck-typed policy).
        try {
            if (HasMethod(this._policy, "On_ErrorEx")) {
                info := Map(
                    "op", op,
                    "typeName", this._typeName,
                    "member", member,
                    "path", this._path,
                    "args", args,
                    "reason", reason
                )
                this._policy.On_ErrorEx(info)
            }
        } catch {
        }
        this._CallPolicy("On_Error", op, member, args)
    }

    _CallPolicy(methodName, op, member, data) {
        try {
            if (methodName = "After_Call") {
                this._policy.After_Call(op, this._typeName, member, this._path, data)
            } else if (methodName = "On_Error") {
                this._policy.On_Error(op, this._typeName, member, this._path, data)
            } else {
                this._policy.On_Call(op, this._typeName, member, this._path, data)
            }
        } catch {
        }
    }
}
