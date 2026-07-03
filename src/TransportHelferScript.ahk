#Requires AutoHotkey v2.0

; Hard-coded SAP steps for transport open / release.
class TransportHelferScript {
    static TRANSPORT_FIELD := "wnd[0]/usr/tabsMAINTABSTRIP/tabpTSSN/ssubCOMMONSUBSCREEN:RDDM0001:0210/ctxtTRDYSE01SN-TR_TRKORR"
    static TRANSPORT_CARET := 10
    static RELEASE_ORDER_ID := "50088725"
    static GRID_SHELL := "wnd[0]/usr/cntlCONT/shellcont/shell/shellcont[0]/shell"

    static OpenTransport(session, transportId) {
        session.FindById("wnd[0]/tbar[0]/okcd").Text := "/nse01"
        session.FindById("wnd[0]").SendVKey(0)
        session.FindById("wnd[0]/usr/tabsMAINTABSTRIP/tabpTSSN").Select()
        transportField := session.FindById(TransportHelferScript.TRANSPORT_FIELD)
        transportField.Text := transportId
        transportField.CaretPosition := TransportHelferScript.TRANSPORT_CARET
        session.FindById("wnd[0]").SendVKey(0)
    }

    static ReleaseTransport(session) {
        session.FindById("wnd[0]/usr/lbl[24,11]").CaretPosition := 6
        session.FindById("wnd[0]").SendVKey(5)
        session.FindById("wnd[1]").SendVKey(0)
        session.FindById("wnd[0]/tbar[1]/btn[9]").Press()
        session.FindById("wnd[1]/usr/btnBUTTON_1").Press()
        session.FindById("wnd[0]/usr/lbl[20,9]").SetFocus()
        session.FindById("wnd[0]/usr/lbl[20,9]").CaretPosition := 7
        session.FindById("wnd[0]/tbar[1]/btn[42]").Press()
        session.FindById("wnd[1]/usr/ctxtWA_SCR_CRMD_ORDERADM_H-OBJECT_ID").Text := TransportHelferScript.RELEASE_ORDER_ID
        session.FindById("wnd[1]").SendVKey(0)
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
        session.FindById(TransportHelferScript.TRANSPORT_FIELD).Text := ""
    }

    static RunCompareShortcut(session) {
        session.FindById("wnd[0]/mbar/menu[3]/menu[12]/menu[0]").Select()
        grid := session.FindById(TransportHelferScript.GRID_SHELL)
        grid.CurrentCellColumn := "USR"
        grid.ClickCurrentCell()
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        grid := session.FindById(TransportHelferScript.GRID_SHELL)
        grid.CurrentCellColumn := "COMPARE"
        grid.ClickCurrentCell()
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    }
}
