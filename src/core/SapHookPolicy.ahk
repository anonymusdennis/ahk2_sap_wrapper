#Requires AutoHotkey v2.0

class SapHookPolicy {
    On_Call(op, typeName, member, path, args) {
    }

    After_Call(op, typeName, member, path, result) {
    }

    On_Error(op, typeName, member, path, args) {
        MsgBox("SAP wrapper call failed.`n"
            . "Operation: " op "`n"
            . "Type: " typeName "`n"
            . "Member: " member "`n"
            . "Path: " path "`n"
            . "LastError (may be unrelated for COM): " A_LastError)
    }
}
