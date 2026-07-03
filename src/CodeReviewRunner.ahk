#Requires AutoHotkey v2.0

#Include CodeReviewConfig.ahk
#Include CodeReviewAppPaths.ahk
#Include CodeReviewActions.ahk
#Include CodeReviewScriptParser.ahk
#Include SapWrapper.ahk
#Include Sm30SapSessions.ahk

; State-machine runner for the hard-coded SE01 review flow.
class CodeReviewRunner {
    static STATE_IDLE := "idle"
    static STATE_AT_CHECKPOINT_0 := "at_checkpoint_0"
    static STATE_AT_CHECKPOINT_1_WAIT := "at_checkpoint_1_wait"
    static STATE_STEP3_MENU_WAIT := "step3_menu_wait"
    static STATE_STEP3_COMPARE_WAIT := "step3_compare_wait"
    static STATE_STEP4_COMPARE := "step4_compare"

    __New(session := "", reviewDef := "", policy := "") {
        this.session := session
        this.policy := IsObject(policy) ? policy : SapHookPolicy()
        this.reviewDef := IsObject(reviewDef) ? reviewDef : CodeReviewConfig.LoadDefault()
        this.transports := []
        this.transportIndex := 1
        this.state := CodeReviewRunner.STATE_IDLE
        this.usedDefaultTreePath := false
        this.lastMessage := "Ready."
    }

    SetSession(session) {
        this.session := session
    }

    SetReviewDef(reviewDef) {
        this.reviewDef := reviewDef
    }

    LoadTransportsFromText(text) {
        parsed := CodeReviewScriptParser.ParseText(text)
        ids := parsed.transports
        cleaned := []
        for item in ids {
            if (item != "" && item != "(unknown)") {
                cleaned.Push(item)
            }
        }
        if (cleaned.Length = 0) {
            cleaned := CodeReviewRunner._ParsePlainTransportList(text)
        }
        this.transports := cleaned
        if (this.transportIndex > this.transports.Length) {
            this.transportIndex := this.transports.Length > 0 ? 1 : 0
        }
        if (this.transportIndex = 0 && this.transports.Length > 0) {
            this.transportIndex := 1
        }
        return this.GetState()
    }

    SetTransportIndex(index) {
        if (index >= 1 && index <= this.transports.Length) {
            this.transportIndex := index
        }
        return this.GetState()
    }

    NextTransport() {
        if (this.transportIndex < this.transports.Length) {
            this.transportIndex += 1
        }
        return this.GetState()
    }

    PrevTransport() {
        if (this.transportIndex > 1) {
            this.transportIndex -= 1
        }
        return this.GetState()
    }

    GetCurrentTransport() {
        if (this.transportIndex >= 1 && this.transportIndex <= this.transports.Length) {
            return this.transports[this.transportIndex]
        }
        return ""
    }

    GetState() {
        return {
            state: this.state,
            stateLabel: CodeReviewRunner._StateLabel(this.state),
            transportIndex: this.transportIndex,
            transportCount: this.transports.Length,
            transportId: this.GetCurrentTransport(),
            checkpointHint: CodeReviewRunner._CheckpointHint(this),
            actionHint: CodeReviewRunner._ActionHint(this),
            continueHint: CodeReviewRunner._ContinueHint(this),
            usedDefaultTreePath: this.usedDefaultTreePath,
            lastMessage: this.lastMessage,
            reviewLabel: this.reviewDef.label
        }
    }

    GoCheckpoint0() {
        this._EnsureSession()
        this._EnsureTransport()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        CodeReviewActions.RunBlock(this.session, this.reviewDef.checkpoint0, vars)
        this.state := CodeReviewRunner.STATE_AT_CHECKPOINT_0
        this.usedDefaultTreePath := false
        this.lastMessage := "Checkpoint 0 — transport opened."
        return this.GetState()
    }

    GoCheckpoint1() {
        this._EnsureSession()
        this._EnsureTransport()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        CodeReviewActions.RunBlock(this.session, this.reviewDef.checkpoint0, vars)
        CodeReviewActions.RunBlock(this.session, this.reviewDef.step1, vars)
        CodeReviewActions.RunBlock(this.session, this.reviewDef.step2, vars)
        this.state := CodeReviewRunner.STATE_AT_CHECKPOINT_1_WAIT
        this.usedDefaultTreePath := false
        this.lastMessage := "Checkpoint 1 — tree expanded. Numpad 5 = default coding, Numpad + = continue manually."
        return this.GetState()
    }

    ActionKey() {
        if (this.state = CodeReviewRunner.STATE_AT_CHECKPOINT_1_WAIT) {
            return this._RunDefaultTreePath()
        }
        if (this.state = CodeReviewRunner.STATE_STEP3_COMPARE_WAIT) {
            return this._SkipOpenCompare()
        }
        if (this.state = CodeReviewRunner.STATE_STEP4_COMPARE) {
            return this._ScrollDiff()
        }
        this.lastMessage := "Numpad 5 has no action in state " this.state "."
        return this.GetState()
    }

    ContinueKey() {
        if (this.state = CodeReviewRunner.STATE_AT_CHECKPOINT_0) {
            return this.GoCheckpoint1()
        }
        if (this.state = CodeReviewRunner.STATE_AT_CHECKPOINT_1_WAIT) {
            return this._ContinueFromTreeManual()
        }
        if (this.state = CodeReviewRunner.STATE_STEP3_MENU_WAIT) {
            return this._RunViewUser()
        }
        if (this.state = CodeReviewRunner.STATE_STEP3_COMPARE_WAIT) {
            return this._RunOpenCompare()
        }
        if (this.state = CodeReviewRunner.STATE_STEP4_COMPARE) {
            return this._FinishAndNextTransport()
        }
        this.lastMessage := "Numpad + has no action in state " this.state "."
        return this.GetState()
    }

    BeginComparisonFlow() {
        this._EnsureSession()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        CodeReviewActions.RunBlock(this.session, this.reviewDef.step3.menu, vars)
        this.state := CodeReviewRunner.STATE_STEP3_MENU_WAIT
        this.lastMessage := "Comparison menu opened. Press Numpad + to view user changes."
        return this.GetState()
    }

    _RunDefaultTreePath() {
        this._EnsureSession()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        CodeReviewActions.RunBlock(this.session, this.reviewDef.step2.defaultPath, vars)
        this.usedDefaultTreePath := true
        this.lastMessage := "Default coding opened."
        return this.BeginComparisonFlow()
    }

    _ContinueFromTreeManual() {
        this.usedDefaultTreePath := false
        this.lastMessage := "Continuing after manual tree navigation."
        return this.BeginComparisonFlow()
    }

    _RunViewUser() {
        this._EnsureSession()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        CodeReviewActions.RunBlock(this.session, this.reviewDef.step3.viewUser, vars)
        this.state := CodeReviewRunner.STATE_STEP3_COMPARE_WAIT
        this.lastMessage := "User column opened. Numpad + = open compare, Numpad 5 = you opened it manually."
        return this.GetState()
    }

    _RunOpenCompare() {
        this._EnsureSession()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        CodeReviewActions.RunBlock(this.session, this.reviewDef.step3.openCompare, vars)
        this.state := CodeReviewRunner.STATE_STEP4_COMPARE
        this.lastMessage := "Compare screen open. Numpad 5 = next diff. Numpad + = finish and next transport."
        return this.GetState()
    }

    _SkipOpenCompare() {
        this.state := CodeReviewRunner.STATE_STEP4_COMPARE
        this.lastMessage := "Skipped compare open (manual). Numpad 5 = next diff. Numpad + = finish."
        return this.GetState()
    }

    _ScrollDiff() {
        this._EnsureSession()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        CodeReviewActions.RunBlock(this.session, this.reviewDef.step4.diffScroll, vars)
        this.lastMessage := "Scrolled to next difference pair."
        return this.GetState()
    }

    _FinishAndNextTransport() {
        this._EnsureSession()
        vars := CodeReviewActions.BuildVars(this.reviewDef, this.GetCurrentTransport())
        if (Sm30JsonConfig._HasProp(this.reviewDef.step4, "finish")) {
            CodeReviewActions.RunBlock(this.session, this.reviewDef.step4.finish, vars)
        }
        hadNext := this.transportIndex < this.transports.Length
        if (hadNext) {
            this.transportIndex += 1
        }
        state := this.GoCheckpoint0()
        if (hadNext) {
            state.lastMessage := "Finished transport. Advanced to " this.GetCurrentTransport() "."
        } else {
            state.lastMessage := "Finished last transport. Back at checkpoint 0."
        }
        return state
    }

    _EnsureSession() {
        if (!IsObject(this.session)) {
            throw Error("No SAP session selected.")
        }
    }

    _EnsureTransport() {
        if (this.GetCurrentTransport() = "") {
            throw Error("No transport selected. Paste transport IDs and parse first.")
        }
    }

    static _ParsePlainTransportList(text) {
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

    static _StateLabel(state) {
        switch state {
            case CodeReviewRunner.STATE_IDLE:
                return "Idle"
            case CodeReviewRunner.STATE_AT_CHECKPOINT_0:
                return "Checkpoint 0 — transport screen"
            case CodeReviewRunner.STATE_AT_CHECKPOINT_1_WAIT:
                return "Checkpoint 1 — choose coding"
            case CodeReviewRunner.STATE_STEP3_MENU_WAIT:
                return "Step 3 — waiting to view user"
            case CodeReviewRunner.STATE_STEP3_COMPARE_WAIT:
                return "Step 3.2 — open compare or skip"
            case CodeReviewRunner.STATE_STEP4_COMPARE:
                return "Step 4 — compare differences"
            default:
                return state
        }
    }

    static _CheckpointHint(runner) {
        cp := runner.reviewDef.checkpoints
        if (runner.state = CodeReviewRunner.STATE_AT_CHECKPOINT_0 && Sm30JsonConfig._HasProp(cp, "0")) {
            hintObj := cp.%"0"%
            return hintObj.hint
        }
        if (runner.state = CodeReviewRunner.STATE_AT_CHECKPOINT_1_WAIT && Sm30JsonConfig._HasProp(cp, "1")) {
            hintObj := cp.%"1"%
            return hintObj.hint
        }
        return ""
    }

    static _ActionHint(runner) {
        if (runner.state = CodeReviewRunner.STATE_AT_CHECKPOINT_1_WAIT) {
            return "Numpad 5 — open first listed coding (default path)"
        }
        if (runner.state = CodeReviewRunner.STATE_STEP3_COMPARE_WAIT) {
            return "Numpad 5 — you opened compare manually"
        }
        if (runner.state = CodeReviewRunner.STATE_STEP4_COMPARE) {
            return "Numpad 5 — next difference (btn 8 + btn 7)"
        }
        return "Numpad 5 — no action here"
    }

    static _ContinueHint(runner) {
        switch runner.state {
            case CodeReviewRunner.STATE_IDLE:
                return "Numpad 0 — open transport / checkpoint 0"
            case CodeReviewRunner.STATE_AT_CHECKPOINT_0:
                return "Numpad + — run step 1+2 and go to checkpoint 1"
            case CodeReviewRunner.STATE_AT_CHECKPOINT_1_WAIT:
                return "Numpad + — continue after manual tree navigation"
            case CodeReviewRunner.STATE_STEP3_MENU_WAIT:
                return "Numpad + — run step 3.1 (view user)"
            case CodeReviewRunner.STATE_STEP3_COMPARE_WAIT:
                return "Numpad + — run step 3.2 (open compare)"
            case CodeReviewRunner.STATE_STEP4_COMPARE:
                return "Numpad + — close compare, next transport, checkpoint 0"
            default:
                return "Numpad + — continue"
        }
    }
}
