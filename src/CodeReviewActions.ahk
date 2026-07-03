#Requires AutoHotkey v2.0

#Include Sm30JsonConfig.ahk

; Execute JSON-defined SAP GUI actions for the code review player.
class CodeReviewActions {
    static RunList(session, actions, vars) {
        if (Type(actions) != "Array") {
            return
        }
        for action in actions {
            CodeReviewActions.RunOne(session, action, vars)
        }
    }

    static RunBlock(session, block, vars) {
        if (!IsObject(block)) {
            return
        }
        if (Sm30JsonConfig._HasProp(block, "steps") && Type(block.steps) = "Array") {
            for stepBlock in block.steps {
                if (Sm30JsonConfig._HasProp(stepBlock, "actions")) {
                    CodeReviewActions.RunList(session, stepBlock.actions, vars)
                }
            }
            return
        }
        if (Sm30JsonConfig._HasProp(block, "actions")) {
            CodeReviewActions.RunList(session, block.actions, vars)
        }
    }

    static RunOne(session, action, vars) {
        if (!IsObject(action)) {
            return
        }
        if (Sm30JsonConfig._HasProp(action, "action")) {
            CodeReviewActions._RunSpecial(session, action, vars)
            return
        }
        elementId := CodeReviewActions._Substitute(action.id, vars)
        if (Sm30JsonConfig._HasProp(action, "property")) {
            member := CodeReviewActions.NormalizeMember(action.property)
            value := CodeReviewActions._ResolveValue(action.value, vars)
            target := session.FindById(elementId)
            target.%member% := value
            return
        }
        if (Sm30JsonConfig._HasProp(action, "method")) {
            member := CodeReviewActions.NormalizeMember(action.method)
            target := session.FindById(elementId)
            args := CodeReviewActions._ResolveArgs(action, vars)
            if (args.Length = 0) {
                target.%member%()
            } else if (args.Length = 1) {
                target.%member%(args[1])
            } else if (args.Length = 2) {
                target.%member%(args[1], args[2])
            } else {
                target.%member%(args[1], args[2], args[3])
            }
        }
    }

    static _RunSpecial(session, action, vars) {
        actionName := action.action
        if (actionName = "expandTreeNodes") {
            treeId := CodeReviewActions._Substitute(action.treeId, vars)
            tree := session.FindById(treeId)
            topNode := Sm30JsonConfig._HasProp(action, "topNode") ? action.topNode : "          1"
            nodes := action.nodes
            for nodeKey in nodes {
                tree.ExpandNode(nodeKey)
                tree.TopNode := topNode
            }
        }
    }

    static BuildVars(reviewDef, transportId) {
        vars := Object()
        vars.transport := transportId
        if (Sm30JsonConfig._HasProp(reviewDef, "transportFieldId")) {
            vars.transportFieldId := reviewDef.transportFieldId
        }
        if (Sm30JsonConfig._HasProp(reviewDef, "transportCaretPosition")) {
            vars.transportCaretPosition := reviewDef.transportCaretPosition
        }
        return vars
    }

    static _Substitute(text, vars) {
        if (text = "") {
            return text
        }
        out := text
        for key, value in vars.OwnProps() {
            placeholder := "{" key "}"
            out := StrReplace(out, placeholder, String(value))
        }
        return out
    }

    static _ResolveValue(value, vars) {
        if (Type(value) = "Integer" || Type(value) = "Float") {
            return value
        }
        text := CodeReviewActions._Substitute(String(value), vars)
        if (RegExMatch(text, "^\d+$")) {
            return Integer(text)
        }
        return text
    }

    static _ResolveArgs(action, vars) {
        args := []
        if (!Sm30JsonConfig._HasProp(action, "args") || Type(action.args) != "Array") {
            return args
        }
        for argValue in action.args {
            args.Push(CodeReviewActions._ResolveValue(argValue, vars))
        }
        return args
    }

    static NormalizeMember(member) {
        known := Map(
            "findbyid", "FindById",
            "sendvkey", "SendVKey",
            "setfocus", "SetFocus",
            "caretposition", "CaretPosition",
            "text", "Text",
            "press", "Press",
            "expandnode", "ExpandNode",
            "topnode", "TopNode",
            "selectednode", "SelectedNode",
            "doubleclicknode", "DoubleClickNode",
            "select", "Select",
            "currentcellcolumn", "CurrentCellColumn",
            "clickcurrentcell", "ClickCurrentCell",
            "maximize", "Maximize"
        )
        key := StrLower(member)
        if (known.Has(key)) {
            return known[key]
        }
        return member
    }
}
