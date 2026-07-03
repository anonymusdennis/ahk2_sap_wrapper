#Requires AutoHotkey v2.0

; v2 entry point. Layers the generated typed wrappers, enum constants and
; convenience helpers on top of the v1 core. v1 scripts that include
; src/SapWrapper.ahk keep working unchanged; scripts that include this
; file additionally get:
;   - typed wrapper classes for the full SAP GUI Scripting object model
;     (src/generated/TypedWrappers.ahk), automatically used when wrapping
;     COM return values
;   - enum constant classes (GuiEventType, GuiMessageBoxType, ...) and the
;     TypeAsNumber fallback for type detection
;   - Sap.Attach() / Sap.App() convenience roots
;   - SapHookPolicy2 with On_Retry / On_ErrorEx / On_Popup hooks

#Include SapWrapper.ahk
#Include generated/TypedWrappers.ahk
#Include core/SapHookPolicy2.ahk
#Include core/Sap.ahk
