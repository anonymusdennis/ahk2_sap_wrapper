#Requires AutoHotkey v2.0

; Known SM30 customizing views with column metadata for bulk import.
class Sm30TableCatalog {
    static GetAll() {
        return [
            Sm30TableCatalog._PfepRunType()
        ]
    }

    static GetLabels() {
        labels := []
        for tableDef in Sm30TableCatalog.GetAll() {
            labels.Push(tableDef.label)
        }
        return labels
    }

    static GetByIndex(index) {
        tables := Sm30TableCatalog.GetAll()
        if (index < 1 || index > tables.Length) {
            return ""
        }
        return tables[index]
    }

    static FindByViewName(viewName) {
        for tableDef in Sm30TableCatalog.GetAll() {
            if (tableDef.viewName = viewName) {
                return tableDef
            }
        }
        return ""
    }

    static _PfepRunType() {
        return {
            label: "PFEP Run Type (/WUE/PFEPRUNTYPE)",
            viewName: "/WUE/PFEPRUNTYPE",
            tableId: "wnd[0]/usr/tbl/WUE/SAPLMMC_PFEPTCTRL_/WUE/PFEPRUNTYPE",
            description: "PFEP run type assignment per sales org and material numbers.",
            columns: [
                { index: 0, kind: "Text", prefix: "ctxt", field: "WUE/PFEPRUNTYPE-VKORG", name: "VKORG" },
                { index: 1, kind: "Text", prefix: "txt", field: "WUE/PFEPRUNTYPE-MATNR_V", name: "MATNR_V" },
                { index: 2, kind: "Text", prefix: "txt", field: "WUE/PFEPRUNTYPE-MATNR_B", name: "MATNR_B" },
                { index: 3, kind: "Key", prefix: "cmb", field: "WUE/PFEPRUNTYPE-/WUE/PFEP_RUN_TYPE", name: "PFEP_RUN_TYPE" }
            ],
            excelHeaders: ["VKORG", "MATNR_V", "MATNR_B", "PFEP_RUN_TYPE"]
        }
    }
}
