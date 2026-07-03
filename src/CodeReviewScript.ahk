#Requires AutoHotkey v2.0

; Hard-coded SE01 transport code review SAP GUI steps.
class CodeReviewScript {
    static LABEL := "SE01 Transport Code Review"
    static TRANSPORT_FIELD := "wnd[0]/usr/tabsMAINTABSTRIP/tabpTSSN/ssubCOMMONSUBSCREEN:RDDM0001:0210/ctxtTRDYSE01SN-TR_TRKORR"
    static TRANSPORT_CARET := 10
    static TREE_SHELL := "wnd[0]/shellcont/shell"
    static GRID_SHELL := "wnd[0]/usr/cntlCONT/shellcont/shell/shellcont[0]/shell"

    static CHECKPOINT_0_HINT := "Numpad 0 jumps here. Numpad + runs step 1+2 to checkpoint 1."
    static CHECKPOINT_1_HINT := "Numpad 1 jumps here. Numpad 5 = default coding, Numpad + = continue manually."

    static RunCheckpoint0(session, transportId) {
        session.FindById("wnd[0]/tbar[0]/okcd").Text := "/nse01"
        session.FindById("wnd[0]").SendVKey(0)
        session.FindById("wnd[0]/usr/tabsMAINTABSTRIP/tabpTSSN").Select()
        transportField := session.FindById(CodeReviewScript.TRANSPORT_FIELD)
        transportField.Text := transportId
        transportField.CaretPosition := CodeReviewScript.TRANSPORT_CARET
        session.FindById("wnd[0]").SendVKey(0)
    }

    static RunStep1(session) {
        session.FindById("wnd[0]/usr/lbl[20,9]").SetFocus()
        session.FindById("wnd[0]/usr/lbl[20,9]").CaretPosition := 4
        session.FindById("wnd[0]/tbar[1]/btn[41]").Press()
    }

    static RunStep2ExpandTree(session) {
        tree := session.FindById(CodeReviewScript.TREE_SHELL)
        tree.ExpandNode("          5")
        tree.TopNode := "          1"
        tree.ExpandNode("          6")
        tree.TopNode := "          1"
    }

    static RunDefaultTreePath(session) {
        tree := session.FindById(CodeReviewScript.TREE_SHELL)
        tree.SelectedNode := "          7"
        tree.DoubleClickNode("          7")
    }

    static RunStep3Menu(session) {
        session.FindById("wnd[0]/mbar/menu[3]/menu[12]/menu[0]").Select()
    }

    static RunStep3ViewUser(session) {
        grid := session.FindById(CodeReviewScript.GRID_SHELL)
        grid.CurrentCellColumn := "USR"
        grid.ClickCurrentCell()
    }

    static RunStep3OpenCompare(session) {
        session.FindById("wnd[1]/tbar[0]/btn[12]").Press()
        grid := session.FindById(CodeReviewScript.GRID_SHELL)
        grid.CurrentCellColumn := "COMPARE"
        grid.ClickCurrentCell()
    }

    static RunStep4DiffScroll(session) {
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        session.FindById("wnd[0]/tbar[1]/btn[7]").Press()
    }

    static RunStep4Finish(session) {
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
        session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
        session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
        session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
        session.FindById("wnd[0]/usr/lbl[24,11]").SetFocus()
        session.FindById("wnd[0]/usr/lbl[24,11]").CaretPosition := 6
    }
}
