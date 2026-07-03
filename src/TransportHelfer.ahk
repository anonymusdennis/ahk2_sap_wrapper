#Requires AutoHotkey v2.0

#Include TransportHelferScript.ahk
#Include CodeReviewScriptParser.ahk
#Include SapWrapper.ahk
#Include Sm30SapSessions.ahk

class TransportHelfer {
    __New(session := "", policy := "") {
        this.session := session
        this.policy := IsObject(policy) ? policy : SapHookPolicy()
        this.transports := []
        this.transportIndex := 0
    }

    SetSession(session) {
        this.session := session
    }

    LoadTransportsFromText(text) {
        parsed := CodeReviewScriptParser.ParseText(text)
        cleaned := []
        for item in parsed.transports {
            if (item != "") {
                cleaned.Push(item)
            }
        }
        if (cleaned.Length = 0) {
            cleaned := TransportHelfer._ParsePlainList(text)
        }
        this.transports := cleaned
        this.transportIndex := cleaned.Length > 0 ? 1 : 0
        return this.GetState()
    }

    GetCurrentTransport() {
        if (this.transportIndex >= 1 && this.transportIndex <= this.transports.Length) {
            return this.transports[this.transportIndex]
        }
        return ""
    }

    SetTransportIndex(index) {
        if (index >= 1 && index <= this.transports.Length) {
            this.transportIndex := index
        }
        return this.GetState()
    }

    NextTransport() {
        if (this.transports.Length = 0) {
            return this.GetState()
        }
        this.transportIndex += 1
        if (this.transportIndex > this.transports.Length) {
            this.transportIndex := 1
        }
        return this.GetState()
    }

    OpenCurrentTransport() {
        this._EnsureSession()
        transportId := this.GetCurrentTransport()
        if (transportId = "") {
            throw Error("No transport selected. Load a transport list first.")
        }
        TransportHelferScript.OpenTransport(this.session, transportId)
        return { ok: true, message: "Opened " transportId " in SE01." }
    }

    ReleaseCurrentTransport() {
        this._EnsureSession()
        TransportHelferScript.ReleaseTransport(this.session)
        return { ok: true, message: "Release steps executed for " this.GetCurrentTransport() "." }
    }

    GetState() {
        return {
            transportIndex: this.transportIndex,
            transportCount: this.transports.Length,
            transportId: this.GetCurrentTransport()
        }
    }

    _EnsureSession() {
        if (!IsObject(this.session)) {
            throw Error("No SAP session selected.")
        }
    }

    static _ParsePlainList(text) {
        ids := []
        lines := StrSplit(text, "`n", "`r")
        for line in lines {
            trimmed := Trim(line)
            if (trimmed = "" || SubStr(trimmed, 1, 1) = "#") {
                continue
            }
            if (RegExMatch(trimmed, "^[A-Z0-9]+$")) {
                ids.Push(trimmed)
            }
        }
        return ids
    }
}
