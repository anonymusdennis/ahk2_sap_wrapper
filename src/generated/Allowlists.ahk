#Requires AutoHotkey v2.0

class SapGeneratedAllowlists {
    static GetAllowlists() {
        allowlists := Map()
        allowlists["GuiAbapEditor"] := Map(
            "AutoBraceEnabled", true,
            "AutoComplete", true,
            "AutoCorrectEnabled", true,
            "AutoExpand", true,
            "AutoIndentEnabled", true,
            "Capitalize", true,
            "ClipboardCopy", true,
            "ClipboardCut", true,
            "ClipboardPaste", true,
            "ClipboardRingPaste", true,
            "CodeHintsEnabled", true,
            "CommentSelectedLines", true,
            "CorrectCapsEnabled", true,
            "Delete", true,
            "DeleteBack", true,
            "DeleteRange", true,
            "DeleteSelection", true,
            "DeleteWord", true,
            "DeleteWordBack", true,
            "DuplicateLine", true,
            "FormatSelectedLines", true,
            "GetAutoCompleteEntryCount", true,
            "GetAutoCompleteEntryText", true,
            "GetAutoCompleteIconType", true,
            "GetAutoCompleteSubIconType", true,
            "GetAutoCompleteToolTipDelay", true,
            "GetAutoCompleteToolbarButtonToolTip", true,
            "GetCurrentToolTipText", true,
            "GetCursorColumnPosition", true,
            "GetCursorLinePosition", true,
            "GetFirstVisibleLine", true,
            "GetHTMLClipboardContents", true,
            "GetLastVisibleLine", true,
            "GetLineCount", true,
            "GetLineText", true,
            "GetNumberedBookmarks", true,
            "GetRTFClipboardContents", true,
            "GetSelectedAutoComplete", true,
            "GetSelectedText", true,
            "GetStructureBlockEndLine", true,
            "GetStructureBlockStartLine", true,
            "GetUndoPosition", true,
            "GetWordWrapMode", true,
            "GetWordWrapPosition", true,
            "GoNextBookMark", true,
            "GoNumberedBookmark", true,
            "GoPreviousBookMark", true,
            "InsertTab", true,
            "InsertText", true,
            "IsAutoCompleteEntryBold", true,
            "IsAutoCompleteOpen", true,
            "IsAutoCompleteToolTipVisible", true,
            "IsAutoCompleteToolbarButtonPressed", true,
            "IsBookmark", true,
            "IsBreakpointSet", true,
            "IsLineCollapsed", true,
            "IsLineComment", true,
            "IsLineModified", true,
            "IsModified", true,
            "JoinSelectedLines", true,
            "LowerCase", true,
            "MoveCursorDocumentEnd", true,
            "MoveCursorLineDown", true,
            "MoveCursorLineEnd", true,
            "MoveCursorLineHome", true,
            "MoveCursorLineUp", true,
            "MoveLineDown", true,
            "MoveLineUp", true,
            "MoveWordLeft", true,
            "MoveWordRight", true,
            "OverwriteModeEnabled", true,
            "RemoveAllBookmarks", true,
            "RemoveAllBreakpoints", true,
            "RemoveBookmarks", true,
            "RemoveBreakpoint", true,
            "ReplaceSelection", true,
            "SaveToFile", true,
            "ScrollToLine", true,
            "SelectAll", true,
            "SelectBlockRange", true,
            "SelectRange", true,
            "SelectWordLeft", true,
            "SelectWordRight", true,
            "Sentencize", true,
            "SetAutoBrace", true,
            "SetAutoCorrect", true,
            "SetAutoIndent", true,
            "SetBookmarks", true,
            "SetBreakpoint", true,
            "SetCodeHints", true,
            "SetCorrectCaps", true,
            "SetLineFeedStyle", true,
            "SetOverwriteMode", true,
            "SetSelectionPosInLine", true,
            "SetSmartTab", true,
            "SetWordWrapMode", true,
            "SetWordWrapPosition", true,
            "SmartTabEnabled", true,
            "SortSelectedLines", true,
            "SwapCase", true,
            "ToggleCapsLock", true,
            "ToggleNumberedBookmark", true,
            "ToggleStructureBlock", true,
            "TransposeLine", true,
            "UnTab", true,
            "UncommentSelectedLines", true,
            "Undo", true,
            "UpperCase", true
        )
        allowlists["GuiApoGrid"] := Map(
            "CancelCut", true,
            "ClearSelection", true,
            "ColumnCount", true,
            "ContextMenu", true,
            "CurrentCellColumn", true,
            "CurrentCellRow", true,
            "Cut", true,
            "DeselectCell", true,
            "DeselectColumn", true,
            "DeselectRow", true,
            "DoubleClickCell", true,
            "FirstVisibleColumn", true,
            "FirstVisibleRow", true,
            "FixedColumnsLeft", true,
            "FixedColumnsRight", true,
            "FixedRowsBottom", true,
            "FixedRowsTop", true,
            "GetBgdColorInfo", true,
            "GetCellChangeable", true,
            "GetCellFormat", true,
            "GetCellTooltip", true,
            "GetCellValue", true,
            "GetFgdColorInfo", true,
            "GetIconInfo", true,
            "IsCellSelected", true,
            "IsColSelected", true,
            "IsRowSelected", true,
            "Paste", true,
            "PressEnter", true,
            "RowCount", true,
            "SelectAll", true,
            "SelectCell", true,
            "SelectColumn", true,
            "SelectRow", true,
            "SelectedCells", true,
            "SelectedColumns", true,
            "SelectedColumnsObject", true,
            "SelectedRows", true,
            "SelectedRowsObject", true,
            "SetCellValue", true,
            "VisibleColumnCount", true,
            "VisibleRowCount", true
        )
        allowlists["GuiApplication"] := Map(
            "ActiveSession", true,
            "AddHistoryEntry", true,
            "AllowSystemMessages", true,
            "ButtonbarVisible", true,
            "ConnectionErrorText", true,
            "Connections", true,
            "CreateGuiCollection", true,
            "DropHistory", true,
            "HistoryEnabled", true,
            "Ignore", true,
            "MajorVersion", true,
            "MinorVersion", true,
            "NewVisualDesign", true,
            "OpenConnection", true,
            "OpenConnectionByConnectionString", true,
            "Patchlevel", true,
            "RegisterROT", true,
            "Revision", true,
            "RevokeROT", true,
            "StatusbarVisible", true,
            "TitlebarVisible", true,
            "ToolbarVisible", true,
            "Utils", true,
            "onCreateSession", true,
            "onDestroySession", true,
            "onError", true,
            "onIgnoreSession", true
        )
        allowlists["GuiBarChart"] := Map(
            "BarCount", true,
            "ChartCount", true,
            "GetBarContent", true,
            "GetGridLineContent", true,
            "GridCount", true,
            "LinkCount", true,
            "SendData", true
        )
        allowlists["GuiBox"] := Map(
            "CharHeight", true,
            "CharLeft", true,
            "CharTop", true,
            "CharWidth", true
        )
        allowlists["GuiButton"] := Map(
            "Emphasized", true,
            "LeftLabel", true,
            "RightLabel", true
        )
        allowlists["GuiCTextField"] := Map()
        allowlists["GuiCalendar"] := Map(
            "ContextMenu", true,
            "CreateDate", true,
            "FirstVisibleDate", true,
            "FocusDate", true,
            "FocusedElement", true,
            "GetColor", true,
            "GetColorInfo", true,
            "GetDateTooltip", true,
            "GetDay", true,
            "GetMonth", true,
            "GetWeekNumber", true,
            "GetWeekday", true,
            "GetYear", true,
            "IsWeekend", true,
            "LastVisibleDate", true,
            "SelectMonth", true,
            "SelectRange", true,
            "SelectWeek", true,
            "SelectionInterval", true,
            "Today", true,
            "endSelection", true,
            "horizontal", true,
            "startSelection", true
        )
        allowlists["GuiChart"] := Map(
            "ValueChange", true
        )
        allowlists["GuiCheckBox"] := Map(
            "ColorIndex", true,
            "ColorIntensified", true,
            "ColorInverse", true,
            "Flushing", true,
            "GetListProperty", true,
            "GetListPropertyNonRec", true,
            "IsLeftLabel", true,
            "IsListElement", true,
            "IsRightLabel", true,
            "LeftLabel", true,
            "RightLabel", true,
            "RowText", true,
            "Selected", true
        )
        allowlists["GuiCollection"] := Map(
            "Add", true,
            "Count", true,
            "ElementAt", true,
            "Item", true,
            "Length", true,
            "NewEnum", true,
            "Type", true,
            "TypeAsNumber", true
        )
        allowlists["GuiColorSelector"] := Map(
            "ChangeSelection", true
        )
        allowlists["GuiComboBox"] := Map(
            "CharHeight", true,
            "CharLeft", true,
            "CharTop", true,
            "CharWidth", true,
            "CurListBoxEntry", true,
            "Flushing", true,
            "Highlighted", true,
            "IsLeftLabel", true,
            "IsListBoxActive", true,
            "IsRightLabel", true,
            "Key", true,
            "Modified", true,
            "Required", true,
            "RightLabel", true,
            "SetKeySpace", true,
            "ShowKey", true,
            "Value", true
        )
        allowlists["GuiComboBoxControl"] := Map(
            "Entries", true,
            "FireSelected", true,
            "LabelText", true,
            "Selected", true,
            "Text", true
        )
        allowlists["GuiComboBoxEntry"] := Map(
            "Key", true,
            "Pos", true,
            "Text", true,
            "Value", true
        )
        allowlists["GuiComponent"] := Map(
            "ContainerType", true,
            "Id", true,
            "Name", true,
            "Parent", true,
            "Type", true,
            "TypeAsNumber", true
        )
        allowlists["GuiComponentCollection"] := Map(
            "Count", true,
            "ElementAt", true,
            "Item", true,
            "Length", true,
            "NewEnum", true,
            "Type", true,
            "TypeAsNumber", true
        )
        allowlists["GuiConnection"] := Map(
            "Children", true,
            "CloseConnection", true,
            "CloseSession", true,
            "ConnectionString", true,
            "Description", true,
            "DisabledByServer", true,
            "Sessions", true
        )
        allowlists["GuiContainer"] := Map(
            "Children", true,
            "FindById", true
        )
        allowlists["GuiContainerShell"] := Map(
            "DockerIsVertical", true,
            "DockerPixelSize", true
        )
        allowlists["GuiContextMenu"] := Map(
            "Select", true
        )
        allowlists["GuiCustomControl"] := Map(
            "CharHeight", true,
            "CharLeft", true,
            "CharTop", true,
            "CharWidth", true
        )
        allowlists["GuiDialogShell"] := Map(
            "Close", true,
            "Title", true
        )
        allowlists["GuiEAIViewer2D"] := Map(
            "AnnotationEnabled", true,
            "AnnotationMode", true,
            "RedliningStream", true,
            "annotationTextRequest", true
        )
        allowlists["GuiEAIViewer3D"] := Map()
        allowlists["GuiEnum"] := Map(
            "Clone", true,
            "Next", true,
            "Reset", true,
            "Skip", true
        )
        allowlists["GuiFrameWindow"] := Map(
            "Close", true,
            "CompBitmap", true,
            "HardCopy", true,
            "HardCopyToMemory", true,
            "Iconify", true,
            "IsVKeyAllowed", true,
            "JumpBackward", true,
            "JumpForward", true,
            "Maximize", true,
            "Restore", true,
            "SendVKey", true,
            "ShowMessageBox", true,
            "TabBackward", true,
            "TabForward", true
        )
        allowlists["GuiGOSShell"] := Map()
        allowlists["GuiGraphAdapt"] := Map()
        allowlists["GuiGridView"] := Map(
            "ClearSelection", true,
            "Click", true,
            "ClickCurrentCell", true,
            "ColumnCount", true,
            "ColumnOrder", true,
            "ContextMenu", true,
            "CurrentCellColumn", true,
            "CurrentCellMoved", true,
            "CurrentCellRow", true,
            "DeleteRows", true,
            "DeselectColumn", true,
            "DoubleClick", true,
            "DoubleClickCurrentCell", true,
            "DuplicateRows", true,
            "FirstVisibleColumn", true,
            "FirstVisibleRow", true,
            "FrozenColumnCount", true,
            "GetCellChangeable", true,
            "GetCellCheckBoxChecked", true,
            "GetCellColor", true,
            "GetCellHeight", true,
            "GetCellIcon", true,
            "GetCellLeft", true,
            "GetCellMaxLength", true,
            "GetCellState", true,
            "GetCellTooltip", true,
            "GetCellTop", true,
            "GetCellType", true,
            "GetCellValue", true,
            "GetCellWidth", true,
            "GetColorInfo", true,
            "GetColumnDataType", true,
            "GetColumnPosition", true,
            "GetColumnSortType", true,
            "GetColumnTitles", true,
            "GetColumnTooltip", true,
            "GetColumnTotalType", true,
            "GetDisplayedColumnTitle", true,
            "GetRowTotalLevel", true,
            "GetSymbolInfo", true,
            "GetToolbarButtonChecked", true,
            "GetToolbarButtonEnabled", true,
            "GetToolbarButtonIcon", true,
            "GetToolbarButtonId", true,
            "GetToolbarButtonText", true,
            "GetToolbarButtonTooltip", true,
            "GetToolbarFocusButton", true,
            "HasCellF4Help", true,
            "HistoryCurEntry", true,
            "HistoryCurIndex", true,
            "HistoryIsActive", true,
            "HistoryList", true,
            "InsertRows", true,
            "IsCellHotspot", true,
            "IsCellSymbol", true,
            "IsCellTotalExpander", true,
            "IsColumnFiltered", true,
            "IsColumnKey", true,
            "IsTotalRowExpanded", true,
            "ModifyCell", true,
            "ModifyCheckBox", true,
            "MoveRows", true,
            "PressButton", true,
            "PressButtonCurrentCell", true,
            "PressColumnHeader", true,
            "PressEnter", true,
            "PressF1", true,
            "PressF4", true,
            "PressToolbarButton", true,
            "PressToolbarContextButton", true,
            "PressTotalRow", true,
            "PressTotalRowCurrentCell", true,
            "RowCount", true,
            "SelectAll", true,
            "SelectColumn", true,
            "SelectToolbarMenuItem", true,
            "SelectedCells", true,
            "SelectedColumns", true,
            "SelectedRows", true,
            "SelectionChanged", true,
            "SetColumnWidth", true,
            "SetCurrentCell", true,
            "Title", true,
            "ToolbarButtonCount", true,
            "TriggerModified", true,
            "VisibleRowCount", true
        )
        allowlists["GuiHTMLViewer"] := Map(
            "BrowserHandle", true,
            "ContextMenu", true,
            "DocumentComplete", true,
            "SapEvent", true
        )
        allowlists["GuiInputFieldControl"] := Map(
            "ButtonTooltip", true,
            "FindButtonActivated", true,
            "HistoryOpened", true,
            "LabelText", true,
            "Submit", true,
            "Text", true
        )
        allowlists["GuiLabel"] := Map(
            "CaretPosition", true,
            "ColorIndex", true,
            "ColorIntensified", true,
            "ColorInverse", true,
            "DisplayedText", true,
            "GetListProperty", true,
            "Highlighted", true,
            "IsHotspot", true,
            "IsLeftLabel", true,
            "IsListElement", true,
            "IsRightLabel", true,
            "MaxLength", true,
            "Numerical", true,
            "RowText", true
        )
        allowlists["GuiMainWindow"] := Map(
            "ButtonbarVisible", true,
            "ResizeWorkingPane", true,
            "ResizeWorkingPaneEx", true,
            "StatusbarVisible", true,
            "TitlebarVisible", true,
            "ToolbarVisible", true
        )
        allowlists["GuiMap"] := Map()
        allowlists["GuiMenu"] := Map(
            "Select", true
        )
        allowlists["GuiMenubar"] := Map()
        allowlists["GuiMessageWindow"] := Map(
            "FocusedButton", true,
            "HelpButtonHelpText", true,
            "HelpButtonText", true,
            "MessageText", true,
            "MessageType", true,
            "OKButtonHelpText", true,
            "OKButtonText", true,
            "Visible", true
        )
        allowlists["GuiModalWindow"] := Map(
            "PopupDialogText", true
        )
        allowlists["GuiNetChart"] := Map(
            "GetLinkContent", true,
            "GetNodeContent", true,
            "LinkCount", true,
            "NodeCount", true,
            "SendData", true
        )
        allowlists["GuiOfficeIntegration"] := Map(
            "AppendRow", true,
            "CloseDocument", true,
            "CustomEvent", true,
            "Document", true,
            "HostedApplication", true,
            "RemoveContent", true,
            "SaveDocument", true,
            "SetDocument", true
        )
        allowlists["GuiOkCodeField"] := Map(
            "Opened", true,
            "PressF1", true
        )
        allowlists["GuiPasswordField"] := Map()
        allowlists["GuiPicture"] := Map(
            "Click", true,
            "ClickControlArea", true,
            "ClickPictureArea", true,
            "ContextMenu", true,
            "DoubleClick", true,
            "DoubleClickControlArea", true,
            "DoubleClickPictureArea", true
        )
        allowlists["GuiRadioButton"] := Map(
            "CharHeight", true,
            "CharLeft", true,
            "CharTop", true,
            "CharWidth", true,
            "Flushing", true,
            "GroupCount", true,
            "GroupMembers", true,
            "GroupPos", true,
            "IsLeftLabel", true,
            "IsRightLabel", true,
            "LeftLabel", true,
            "RightLabel", true,
            "Select", true
        )
        allowlists["GuiSapChart"] := Map()
        allowlists["GuiScrollContainer"] := Map(
            "HorizontalScrollbar", true,
            "VerticalScrollbar", true
        )
        allowlists["GuiScrollbar"] := Map(
            "Maximum", true,
            "Minimum", true,
            "PageSize", true,
            "Position", true
        )
        allowlists["GuiSession"] := Map(
            "ActiveWindow", true,
            "AsStdNumberFormat", true,
            "Busy", true,
            "ClearErrorList", true,
            "CreateSession", true,
            "EnableJawsEvents", true,
            "EndTransaction", true,
            "ErrorList", true,
            "FindByPosition", true,
            "GetIconResourceName", true,
            "GetVKeyDescription", true,
            "Info", true,
            "IsActive", true,
            "IsListBoxActive", true,
            "ListBoxCurrEntry", true,
            "ListBoxCurrEntryHeight", true,
            "ListBoxCurrEntryLeft", true,
            "ListBoxCurrEntryTop", true,
            "ListBoxCurrEntryWidth", true,
            "ListBoxHeight", true,
            "ListBoxLeft", true,
            "ListBoxTop", true,
            "ListBoxWidth", true,
            "LockSessionUI", true,
            "PassportPreSystemId", true,
            "PassportSystemId", true,
            "PassportTransactionId", true,
            "ProgressPercent", true,
            "ProgressText", true,
            "Record", true,
            "RecordFile", true,
            "SaveAsUnicode", true,
            "SendCommand", true,
            "ShowDropdownKeys", true,
            "StartTransaction", true,
            "SuppressBackendPopups", true,
            "TestToolMode", true,
            "UnlockSessionUI", true,
            "onAbapScriptingEvent", true,
            "onActivated", true,
            "onAutomationFCode", true,
            "onChange", true,
            "onContextMenu", true,
            "onDestroy", true,
            "onEndRequest", true,
            "onError", true,
            "onFocusChanged", true,
            "onHistoryOpened", true,
            "onHit", true,
            "onProgressIndicator", true,
            "onStartRequest", true
        )
        allowlists["GuiSessionInfo"] := Map(
            "ApplicationServer", true,
            "Client", true,
            "Codepage", true,
            "Flushes", true,
            "Group", true,
            "GuiCodepage", true,
            "I18NMode", true,
            "InterpretationTime", true,
            "IsLowSpeedConnection", true,
            "Language", true,
            "MessageServer", true,
            "Program", true,
            "ResponseTime", true,
            "RoundTrips", true,
            "ScreenNumber", true,
            "ScriptingModeReadOnly", true,
            "ScriptingModeRecordingDisabled", true,
            "SessionNumber", true,
            "SystemName", true,
            "SystemNumber", true,
            "SystemSessionId", true,
            "Transaction", true,
            "UI_GUIDELINE", true,
            "User", true
        )
        allowlists["GuiShell"] := Map(
            "AccDescription", true,
            "CurrentContextMenu", true,
            "DragDropSupported", true,
            "Handle", true,
            "OcxEvents", true,
            "SelectContextMenuItem", true,
            "SelectContextMenuItemByPosition", true,
            "SelectContextMenuItemByText", true,
            "SubType", true
        )
        allowlists["GuiSimpleContainer"] := Map(
            "GetListProperty", true,
            "GetListPropertyNonRec", true,
            "IsListElement", true,
            "IsStepLoop", true,
            "LoopColCount", true,
            "LoopCurrentCol", true,
            "LoopCurrentColCount", true,
            "LoopCurrentRow", true,
            "LoopRowCount", true
        )
        allowlists["GuiSplit"] := Map(
            "GetColSize", true,
            "GetRowSize", true,
            "IsVertical", true,
            "SetColSize", true,
            "SetRowSize", true
        )
        allowlists["GuiSplitterContainer"] := Map(
            "IsVertical", true,
            "SashPosition", true
        )
        allowlists["GuiStage"] := Map(
            "ContextMenu", true,
            "DoubleClick", true,
            "SelectItems", true
        )
        allowlists["GuiStatusPane"] := Map()
        allowlists["GuiStatusbar"] := Map(
            "DoubleClick", true,
            "Handle", true,
            "MessageAsPopup", true,
            "MessageId", true,
            "MessageNumber", true,
            "MessageParameter", true,
            "MessageType", true
        )
        allowlists["GuiTab"] := Map(
            "ScrollToLeft", true,
            "Select", true
        )
        allowlists["GuiTabStrip"] := Map(
            "CharHeight", true,
            "CharLeft", true,
            "CharTop", true,
            "CharWidth", true,
            "LeftTab", true,
            "SelectedTab", true
        )
        allowlists["GuiTableColumn"] := Map(
            "Count", true,
            "DefaultTooltip", true,
            "ElementAt", true,
            "Fixed", true,
            "IconName", true,
            "Item", true,
            "Length", true,
            "NewEnum", true,
            "Selected", true,
            "Title", true,
            "Tooltip", true,
            "Type", true,
            "TypeAsNumber", true
        )
        allowlists["GuiTableControl"] := Map(
            "CharHeight", true,
            "CharLeft", true,
            "CharTop", true,
            "CharWidth", true,
            "ColSelectMode", true,
            "Columns", true,
            "ConfigureLayout", true,
            "CurrentCol", true,
            "CurrentRow", true,
            "DeselectAllColumns", true,
            "GetAbsoluteRow", true,
            "GetCell", true,
            "HorizontalScrollbar", true,
            "ReorderTable", true,
            "RowCount", true,
            "RowSelectMode", true,
            "Rows", true,
            "SelectAllColumns", true,
            "TableFieldName", true,
            "VerticalScrollbar", true,
            "VisibleRowCount", true
        )
        allowlists["GuiTableRow"] := Map(
            "Count", true,
            "ElementAt", true,
            "Item", true,
            "Length", true,
            "NewEnum", true,
            "Selectable", true,
            "Selected", true,
            "Type", true,
            "TypeAsNumber", true
        )
        allowlists["GuiTextedit"] := Map(
            "CaretPosition", true,
            "ContextMenu", true,
            "CurrentColumn", true,
            "CurrentLine", true,
            "DisplayedText", true,
            "DoubleClick", true,
            "FirstVisibleLine", true,
            "GetLineText", true,
            "GetListProperty", true,
            "GetListPropertyNonRec", true,
            "GetUnprotectedTextPart", true,
            "Highlighted", true,
            "HistoryCurEntry", true,
            "HistoryCurIndex", true,
            "HistoryIsActive", true,
            "HistoryList", true,
            "IsBreakpointLine", true,
            "IsCommentLine", true,
            "IsHighlightedLine", true,
            "IsHotspot", true,
            "IsLeftLabel", true,
            "IsListElement", true,
            "IsOField", true,
            "IsProtectedLine", true,
            "IsRightLabel", true,
            "IsSelectedLine", true,
            "LastVisibleLine", true,
            "LeftLabel", true,
            "LineCount", true,
            "MaxLength", true,
            "ModifiedStatusChanged", true,
            "MultipleFilesDropped", true,
            "NumberOfUnprotectedTextParts", true,
            "Numerical", true,
            "PressF1", true,
            "PressF4", true,
            "Required", true,
            "RightLabel", true,
            "SelectedText", true,
            "SelectionEndColumn", true,
            "SelectionEndLine", true,
            "SelectionIndexEnd", true,
            "SelectionIndexStart", true,
            "SelectionStartColumn", true,
            "SelectionStartLine", true,
            "SetSelectionIndexes", true,
            "SetUnprotectedTextPart", true,
            "SingleFileDropped", true
        )
        allowlists["GuiTitlebar"] := Map()
        allowlists["GuiToolbar"] := Map(
            "ButtonCount", true,
            "FocusedButton", true,
            "GetButtonChecked", true,
            "GetButtonEnabled", true,
            "GetButtonIcon", true,
            "GetButtonId", true,
            "GetButtonText", true,
            "GetButtonTooltip", true,
            "GetButtonType", true,
            "GetMenuItemIdFromPosition", true,
            "PressButton", true,
            "PressContextButton", true,
            "SelectMenuItem", true,
            "SelectMenuItemByText", true
        )
        allowlists["GuiTree"] := Map(
            "ChangeCheckbox", true,
            "ClickLink", true,
            "CollapseNode", true,
            "ColumnOrder", true,
            "DefaultContextMenu", true,
            "DoubleClickItem", true,
            "DoubleClickNode", true,
            "EnsureVisibleHorizontalItem", true,
            "ExpandNode", true,
            "FindNodeKeyByPath", true,
            "GetAbapImage", true,
            "GetAllNodeKeys", true,
            "GetCheckBoxState", true,
            "GetColumnCol", true,
            "GetColumnHeaders", true,
            "GetColumnIndexFromName", true,
            "GetColumnNames", true,
            "GetColumnTitleFromName", true,
            "GetColumnTitles", true,
            "GetFocusedNodeKey", true,
            "GetHierarchyLevel", true,
            "GetHierarchyTitle", true,
            "GetIsDisabled", true,
            "GetIsHighLighted", true,
            "GetItemHeight", true,
            "GetItemLeft", true,
            "GetItemStyle", true,
            "GetItemText", true,
            "GetItemTextColor", true,
            "GetItemToolTip", true,
            "GetItemTop", true,
            "GetItemType", true,
            "GetItemWidth", true,
            "GetListTreeNodeItemCount", true,
            "GetNextNodeKey", true,
            "GetNodeAbapImage", true,
            "GetNodeChildrenCount", true,
            "GetNodeChildrenCountByPath", true,
            "GetNodeHeight", true,
            "GetNodeIndex", true,
            "GetNodeItemHeaders", true,
            "GetNodeKeyByPath", true,
            "GetNodeLeft", true,
            "GetNodePathByKey", true,
            "GetNodeStyle", true,
            "GetNodeTextByKey", true,
            "GetNodeTextByPath", true,
            "GetNodeTextColor", true,
            "GetNodeToolTip", true,
            "GetNodeTop", true,
            "GetNodeWidth", true,
            "GetNodesCol", true,
            "GetParent", true,
            "GetPreviousNodeKey", true,
            "GetSelectedNodes", true,
            "GetSelectionMode", true,
            "GetStyleDescription", true,
            "GetSubNodesCol", true,
            "GetTreeType", true,
            "HeaderContextMenu", true,
            "IsFolder", true,
            "IsFolderExpandable", true,
            "IsFolderExpanded", true,
            "ItemContextMenu", true,
            "NodeContextMenu", true,
            "PressButton", true,
            "PressHeader", true,
            "PressKey", true,
            "SelectColumn", true,
            "SelectItem", true,
            "SelectNode", true,
            "SelectedItemColumn", true,
            "SelectedItemNode", true,
            "SelectedNode", true,
            "SetCheckBoxState", true,
            "SetColumnWidth", true,
            "TopNode", true,
            "UnselectAll", true,
            "UnselectColumn", true,
            "UnselectNode", true
        )
        allowlists["GuiUserArea"] := Map(
            "CurrentContextMenu", true,
            "FindByLabel", true,
            "HorizontalScrollbar", true,
            "IsOTFPreview", true,
            "VerticalScrollbar", true
        )
        allowlists["GuiUtils"] := Map(
            "CloseFile", true,
            "MESSAGE_OPTION_OK", true,
            "MESSAGE_OPTION_OKCANCEL", true,
            "MESSAGE_OPTION_YESNO", true,
            "MESSAGE_RESULT_CANCEL", true,
            "MESSAGE_RESULT_NO", true,
            "MESSAGE_RESULT_OK", true,
            "MESSAGE_RESULT_YES", true,
            "MESSAGE_TYPE_ERROR", true,
            "MESSAGE_TYPE_INFORMATION", true,
            "MESSAGE_TYPE_PLAIN", true,
            "MESSAGE_TYPE_QUESTION", true,
            "MESSAGE_TYPE_WARNING", true,
            "OpenFile", true,
            "ShowMessageBox", true,
            "Write", true,
            "WriteLine", true
        )
        allowlists["GuiVComponent"] := Map(
            "AccLabelCollection", true,
            "AccText", true,
            "AccTextOnRequest", true,
            "AccTooltip", true,
            "Changeable", true,
            "DefaultTooltip", true,
            "DumpState", true,
            "Height", true,
            "IconName", true,
            "IsSymbolFont", true,
            "Left", true,
            "Modified", true,
            "ParentFrame", true,
            "ScreenLeft", true,
            "ScreenTop", true,
            "SetFocus", true,
            "Text", true,
            "Tooltip", true,
            "Top", true,
            "Visualize", true,
            "Width", true
        )
        allowlists["GuiVContainer"] := Map(
            "FindAllByName", true,
            "FindAllByNameEx", true,
            "FindByName", true
        )
        allowlists["GuiVHViewSwitch"] := Map(
            "GetScriptingEngine", true,
            "onAbapScriptingEvent", true,
            "onActivated", true,
            "onAutomationFCode", true,
            "onChange", true,
            "onContextMenu", true,
            "onCreateSession", true,
            "onDestroy", true,
            "onDestroySession", true,
            "onEndRequest", true,
            "onError", true,
            "onFocusChanged", true,
            "onHistoryOpened", true,
            "onHit", true,
            "onIgnoreSession", true,
            "onProgressIndicator", true,
            "onStartRequest", true
        )
        return allowlists
    }

    static GetInheritance() {
        inheritance := Map()
        inheritance["GuiAbapEditor"] := "GuiShell"
        inheritance["GuiApoGrid"] := "GuiShell"
        inheritance["GuiApplication"] := "GuiContainer"
        inheritance["GuiBarChart"] := "GuiShell"
        inheritance["GuiBox"] := "GuiVComponent"
        inheritance["GuiButton"] := "GuiVComponent"
        inheritance["GuiCTextField"] := ""
        inheritance["GuiCalendar"] := "GuiShell"
        inheritance["GuiChart"] := ""
        inheritance["GuiCheckBox"] := "GuiVComponent"
        inheritance["GuiCollection"] := ""
        inheritance["GuiColorSelector"] := "GuiShell"
        inheritance["GuiComboBox"] := "GuiVComponent"
        inheritance["GuiComboBoxControl"] := ""
        inheritance["GuiComboBoxEntry"] := ""
        inheritance["GuiComponent"] := ""
        inheritance["GuiComponentCollection"] := ""
        inheritance["GuiConnection"] := "GuiContainer"
        inheritance["GuiContainer"] := "GuiComponent"
        inheritance["GuiContainerShell"] := "GuiVContainer"
        inheritance["GuiContextMenu"] := "GuiMenu"
        inheritance["GuiCustomControl"] := "GuiVContainer"
        inheritance["GuiDialogShell"] := "GuiVContainer"
        inheritance["GuiEAIViewer2D"] := "GuiShell"
        inheritance["GuiEAIViewer3D"] := "GuiShell"
        inheritance["GuiEnum"] := "GuiVContainer"
        inheritance["GuiFrameWindow"] := "GuiVContainer"
        inheritance["GuiGOSShell"] := "GuiVContainer"
        inheritance["GuiGraphAdapt"] := ""
        inheritance["GuiGridView"] := "GuiShell"
        inheritance["GuiHTMLViewer"] := "GuiShell"
        inheritance["GuiInputFieldControl"] := ""
        inheritance["GuiLabel"] := "GuiVComponent"
        inheritance["GuiMainWindow"] := "GuiFrameWindow"
        inheritance["GuiMap"] := ""
        inheritance["GuiMenu"] := "GuiVContainer"
        inheritance["GuiMenubar"] := "GuiVContainer"
        inheritance["GuiMessageWindow"] := ""
        inheritance["GuiModalWindow"] := "GuiFrameWindow"
        inheritance["GuiNetChart"] := ""
        inheritance["GuiOfficeIntegration"] := ""
        inheritance["GuiOkCodeField"] := "GuiVComponent"
        inheritance["GuiPasswordField"] := ""
        inheritance["GuiPicture"] := "GuiShell"
        inheritance["GuiRadioButton"] := "GuiVComponent"
        inheritance["GuiSapChart"] := ""
        inheritance["GuiScrollContainer"] := ""
        inheritance["GuiScrollbar"] := ""
        inheritance["GuiSession"] := "GuiContainer"
        inheritance["GuiSessionInfo"] := ""
        inheritance["GuiShell"] := "GuiVContainer"
        inheritance["GuiSimpleContainer"] := "GuiVContainer"
        inheritance["GuiSplit"] := "GuiShell"
        inheritance["GuiSplitterContainer"] := ""
        inheritance["GuiStage"] := ""
        inheritance["GuiStatusPane"] := ""
        inheritance["GuiStatusbar"] := "GuiVComponent"
        inheritance["GuiTab"] := "GuiVContainer"
        inheritance["GuiTabStrip"] := "GuiVContainer"
        inheritance["GuiTableColumn"] := "GuiComponentCollection"
        inheritance["GuiTableControl"] := ""
        inheritance["GuiTableRow"] := "GuiComponentCollection"
        inheritance["GuiTextedit"] := "GuiShell"
        inheritance["GuiTitlebar"] := ""
        inheritance["GuiToolbar"] := "GuiVContainer"
        inheritance["GuiTree"] := ""
        inheritance["GuiUserArea"] := "GuiVContainer"
        inheritance["GuiUtils"] := ""
        inheritance["GuiVComponent"] := "GuiComponent"
        inheritance["GuiVContainer"] := "GuiContainer"
        inheritance["GuiVHViewSwitch"] := "GuiVComponent"
        return inheritance
    }
}
