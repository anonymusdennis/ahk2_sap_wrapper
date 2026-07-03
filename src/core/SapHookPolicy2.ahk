#Requires AutoHotkey v2.0

; v2 hook policy. Extends the v1 SapHookPolicy with additional hooks.
; All policies are duck-typed: plain v1 SapHookPolicy instances keep
; working everywhere, and any object implementing a subset of these
; methods is accepted by the proxy layer.
class SapHookPolicy2 extends SapHookPolicy {
    ; Called before a stale-reference retry is attempted (see
    ; SapComProxy.SetStaleRecovery). attempt is 1-based.
    On_Retry(op, typeName, member, path, attempt) {
    }

    ; Structured error hook, called before On_Error. info is a Map with
    ; keys: op, typeName, member, path, args, reason.
    ; reason is one of:
    ;   "not-found"  - FindById probe confirmed the control does not exist
    ;   "transient"  - FindById probe found the control, so the failure
    ;                  may be transient (e.g. modal popup, busy session)
    ;   "com-error"  - any other COM failure
    On_ErrorEx(info) {
    }

    ; Hook point for popup detection/recovery. The proxy layer does not
    ; call this directly; tools (e.g. loaders) may call it when they
    ; detect an unexpected modal window so users can centralize their
    ; popup handling in one policy object.
    On_Popup(session, windowId) {
    }
}
