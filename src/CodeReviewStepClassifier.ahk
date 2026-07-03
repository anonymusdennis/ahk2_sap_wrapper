#Requires AutoHotkey v2.0

; Map SAP recorder lines to review categories and human-readable labels.
class CodeReviewStepClassifier {
    static Classify(elementId, member, op, value, args) {
        lowerId := StrLower(elementId)
        lowerMember := StrLower(member)

        if (InStr(lowerId, "/okcd") && op = "set") {
            return "transaction"
        }
        if (InStr(lowerId, "tr_trkorr") || InStr(lowerId, "trkorr")) {
            return "transport_input"
        }
        if (lowerMember = "sendvkey") {
            keyCode := Trim(args)
            if (keyCode = "0") {
                return "enter"
            }
            if (keyCode = "2") {
                return "grid_down"
            }
            if (keyCode = "3" || keyCode = "5") {
                return "back"
            }
            return "function_key"
        }
        if (InStr(lowerId, "/lbl[") && (lowerMember = "setfocus" || lowerMember = "caretposition")) {
            return "focus_label"
        }
        if (InStr(lowerId, "/btn[") && lowerMember = "press") {
            if (InStr(lowerId, "/tbar[1]/btn[8]")) {
                return "compare_next"
            }
            if (InStr(lowerId, "/tbar[1]/btn[7]")) {
                return "compare_prev"
            }
            if (InStr(lowerId, "/tbar[0]/btn[3]")) {
                return "nav_back"
            }
            return "button_press"
        }
        if (lowerMember = "expandnode") {
            return "tree_expand"
        }
        if (lowerMember = "topnode") {
            return "tree_scroll"
        }
        if (lowerMember = "selectednode" || lowerMember = "doubleclicknode") {
            return "tree_select"
        }
        if (InStr(lowerId, "/menu[") && lowerMember = "select") {
            return "menu_select"
        }
        if (lowerMember = "currentcellcolumn" || lowerMember = "clickcurrentcell") {
            if (InStr(StrUpper(value . args), "COMPARE")) {
                return "open_compare"
            }
            if (InStr(StrUpper(value . args), "USR")) {
                return "view_user"
            }
            return "grid_click"
        }
        if (lowerMember = "maximize") {
            return "window_setup"
        }
        if (InStr(lowerId, "editor") || InStr(lowerId, "/txt") || InStr(lowerId, "/cntl")) {
            if (lowerMember = "verticalscrollbar" || lowerMember = "horizontalscrollbar") {
                return "editor_scroll"
            }
            return "editor_view"
        }
        if (InStr(lowerId, "history") || InStr(lowerId, "versn") || InStr(lowerId, "version")) {
            return "code_history"
        }
        if (InStr(lowerId, "name") || InStr(lowerId, "/lbl[") && lowerMember = "setfocus") {
            return "view_name"
        }
        return "other"
    }

    static BuildLabel(kind, elementId, member, op, value, args) {
        switch kind {
            case "transaction":
                return "Start transaction " CodeReviewStepClassifier._Unquote(value)
            case "transport_input":
                return "Enter transport " CodeReviewStepClassifier._Unquote(value)
            case "enter":
                return "Press Enter"
            case "grid_down":
                return "Move down in transport list"
            case "focus_label":
                return "Focus list row (" CodeReviewStepClassifier._ShortId(elementId) ")"
            case "button_press":
                return "Press button (" CodeReviewStepClassifier._ShortId(elementId) ")"
            case "tree_expand":
                return "Expand tree node " CodeReviewStepClassifier._Unquote(args)
            case "tree_scroll":
                return "Scroll tree to node " CodeReviewStepClassifier._Unquote(args)
            case "tree_select":
                if (StrLower(member) = "doubleclicknode") {
                    return "Open tree node " CodeReviewStepClassifier._Unquote(args)
                }
                return "Select tree node " CodeReviewStepClassifier._Unquote(args)
            case "menu_select":
                return "Choose menu item (" CodeReviewStepClassifier._ShortId(elementId) ")"
            case "open_compare":
                return "Open version compare"
            case "view_user":
                return "View changed by user"
            case "grid_click":
                return "Click grid cell (" CodeReviewStepClassifier._ShortId(elementId) ")"
            case "compare_next":
                return "Next difference"
            case "compare_prev":
                return "Previous difference"
            case "nav_back":
                return "Navigate back"
            case "window_setup":
                return "Maximize SAP window"
            case "editor_scroll":
                return "Scroll in editor"
            case "editor_view":
                return "View in editor"
            case "code_history":
                return "View code history"
            case "view_name":
                return "View object name"
            case "back":
                return "Back / cancel key"
            default:
                action := op = "set" ? "Set " member : member
                return action " (" CodeReviewStepClassifier._ShortId(elementId) ")"
        }
    }

    static MacroGroupForKind(kind) {
        switch kind {
            case "window_setup", "transaction", "enter":
                return "setup"
            case "transport_input":
                return "transport"
            case "focus_label", "grid_down":
                return "locate_object"
            case "button_press", "tree_expand", "tree_scroll", "tree_select", "menu_select":
                return "open_object"
            case "view_user", "grid_click", "open_compare":
                return "inspect_change"
            case "compare_next", "compare_prev":
                return "review_diff"
            case "nav_back", "back":
                return "exit"
            case "editor_scroll", "editor_view":
                return "editor"
            case "code_history":
                return "history"
            case "view_name":
                return "name"
            default:
                return "other"
        }
    }

    static _ShortId(elementId) {
        if (InStr(elementId, "/")) {
            parts := StrSplit(elementId, "/")
            return parts[parts.Length]
        }
        return elementId
    }

    static _Unquote(text) {
        cleaned := Trim(text)
        if ((SubStr(cleaned, 1, 1) = '"' && SubStr(cleaned, -1) = '"')
            || (SubStr(cleaned, 1, 1) = "'" && SubStr(cleaned, -1) = "'")) {
            return SubStr(cleaned, 2, StrLen(cleaned) - 2)
        }
        return cleaned
    }
}
