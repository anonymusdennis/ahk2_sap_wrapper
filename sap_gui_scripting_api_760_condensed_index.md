# SAP GUI Scripting API (SAP GUI for Windows 7.60 PL1) — condensed complete member list

This is a **compressed, JS/TS-like signature index** of the SAP GUI Scripting COM object model, extracted from the attached SAP GUI Scripting API developer guide.

Type mapping used here: `Byte → boolean`, `Long/Integer → number`, `String → string`, `Variant/Object → any`.

## GuiAbapEditor (extends GuiShell)
_Doc section 1.2.1, pages 16–29._

```ts
interface GuiAbapEditor extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  AutoBraceEnabled(): boolean;
  AutoComplete(): void;
  AutoCorrectEnabled(): boolean;
  AutoExpand(): void;
  AutoIndentEnabled(): boolean;
  Capitalize(): void;
  ClipboardCopy(): void;
  ClipboardCut(): void;
  ClipboardPaste(): void;
  ClipboardRingPaste(Index: number): void;
  CodeHintsEnabled(): boolean;
  CommentSelectedLines(): void;
  CorrectCapsEnabled(): boolean;
  Delete(): void;
  DeleteBack(): void;
  DeleteRange(LineStart: number, ColumnStart: number, LineEnd: number, ● ColumnStart (p2): void;
  DeleteSelection(): void;
  DeleteWord(): void;
  DeleteWordBack(): void;
  DuplicateLine(): void;
  FormatSelectedLines(): void;
  GetAutoCompleteEntryCount(): number;
  GetAutoCompleteEntryText(Index: number): string;
  GetAutoCompleteIconType(Index: number): string;
  GetAutoCompleteSubIconType(Index: number): string;
  GetAutoCompleteToolbarButtonToolTip(Index: number): string;
  GetAutoCompleteToolTipDelay(): number;
  GetCurrentToolTipText(): string;
  GetCursorColumnPosition(): number;
  GetCursorLinePosition(): number;
  GetFirstVisibleLine(): number;
  GetHTMLClipboardContents(): string;
  GetLastVisibleLine(): number;
  GetLineCount(): number;
  GetLineText(Line: number): string;
  GetNumberedBookmarks(Line: number): any;
  GetRTFClipboardContents(): string;
  GetSelectedAutoComplete(): number;
  GetSelectedText(): string;
  GetStructureBlockEndLine(Line: number): number;
  GetStructureBlockStartLine(Line: number): number;
  GetUndoPosition(): number;
  GetWordWrapMode(): number;
  GetWordWrapPosition(): number;
  GoNextBookMark(): void;
  GoNumberedBookmark(Mark: number): void;
  GoPreviousBookMark(): void;
  InsertTab(): void;
  InsertText(Text: string, Line: number, Column: number): void;
  IsAutoCompleteEntryBold(Index: number): boolean;
  IsAutoCompleteOpen(): boolean;
  IsAutoCompleteToolbarButtonPressed(Index: number): boolean;
  IsAutoCompleteToolTipVisible(): boolean;
  IsBookmark(Line: number): boolean;
  IsBreakpointSet(Line: number): boolean;
  IsLineCollapsed(Line: number): boolean;
  IsLineComment(Line: number): boolean;
  IsLineModified(Line: number): boolean;
  IsModified(): boolean;
  JoinSelectedLines(): void;
  LowerCase(): void;
  MoveCursorDocumentEnd(): void;
  MoveCursorLineDown(): void;
  MoveCursorLineEnd(): void;
  MoveCursorLineHome(): void;
  MoveCursorLineUp(): void;
  MoveLineDown(): void;
  MoveLineUp(): void;
  MoveWordLeft(): void;
  MoveWordRight(): void;
  OverwriteModeEnabled(): boolean;
  RemoveAllBookmarks(): void;
  RemoveAllBreakpoints(): void;
  RemoveBookmarks(Bookmarks: string): void;
  RemoveBreakpoint(Line: number): void;
  ReplaceSelection(Text: string): void;
  SaveToFile(p1: string): void;
  ScrollToLine(Line: number): void;
  SelectAll(): void;
  SelectBlockRange(LineStart: number, ColumnStart: number, LineEnd: number, ColumnEnd: tion): void;
  SelectRange(LineStart: number, ColumnStart: number, LineEnd: number, ColumnEnd: where): void;
  SelectWordLeft(): void;
  SelectWordRight(): void;
  Sentencize(): void;
  SetAutoBrace(Status: boolean): void;
  SetAutoCorrect(Status: boolean): void;
  SetAutoIndent(Status: boolean): void;
  SetBookmarks(Bookmarks: string): void;
  SetBreakpoint(Line: number): void;
  SetCodeHints(Status: boolean): void;
  SetCorrectCaps(Status: boolean): void;
  SetLineFeedStyle(p1: number): void;
  SetOverwriteMode(Status: boolean): void;
  SetSelectionPosInLine(Line: number, Column: number): void;
  SetSmartTab(Status: boolean): void;
  SetWordWrapMode(Mode: number): void;
  SetWordWrapPosition(Pos: number): void;
  SmartTabEnabled(): boolean;
  SortSelectedLines(): void;
  SwapCase(): void;
  ToggleCapsLock(): void;
  ToggleNumberedBookmark(Mark: number, Line: number): void;
  ToggleStructureBlock(Line: number): void;
  TransposeLine(): void;
  UncommentSelectedLines(): void;
  Undo(UndoPosition: number): void;
  UnTab(): void;
  UpperCase(): void;
}
```

## GuiApoGrid (extends GuiShell)
_Doc section 1.2.2, pages 29–37._

```ts
interface GuiApoGrid extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  readonly ColumnCount: number;
  CurrentCellColumn: number;
  readonly CurrentCellRow: number;
  FirstVisibleColumn: number;
  FirstVisibleRow: number;
  FixedColumnsLeft: number;
  FixedColumnsRight: number;
  FixedRowsBottom: number;
  readonly FixedRowsTop: number;
  readonly RowCount: number;
  SelectedCells: any;
  SelectedColumns: string;
  SelectedColumnsObject: any;
  readonly SelectedRows: string;
  SelectedRowsObject: any;
  VisibleColumnCount: number;
  VisibleRowCount: number;
  CancelCut(): void;
  ClearSelection(): void;
  ContextMenu(Column: number, Row: number): void;
  Cut(): void;
  DeselectCell(Column: number, Row: number): void;
  DeselectColumn(Column: number): void;
  DeselectRow(Row: number): void;
  DoubleClickCell(Column: number, Row: number): void;
  GetBgdColorInfo(Row: number, Column: number): string;
  GetCellChangeable(Column: number, Row: number): boolean;
  GetCellFormat(Column: number, Row: number): string;
  GetCellTooltip(Column: number, Row: number): string;
  GetCellValue(Column: number, Row: number): string;
  GetFgdColorInfo(Row: number, Column: number): string;
  GetIconInfo(Row: number, Column: number): string;
  IsCellSelected(Column: number, Row: number): boolean;
  IsColSelected(col: number): boolean;
  IsRowSelected(Row: number): boolean;
  Paste(CellValues: any, ColumnCount: number): number;
  PressEnter(): void;
  SelectAll(): void;
  SelectCell(Column: number, Row: number): void;
  SelectColumn(Column: number): void;
  SelectRow(Row: number): void;
  SetCellValue(Column: number, Row: number, Value: string): string;
}
```

## GuiApplication (extends GuiContainer)
_Doc section 1.2.3, pages 37–47._

```ts
interface GuiApplication extends GuiContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  ActiveSession: any;
  AllowSystemMessages: boolean;
  ButtonbarVisible: boolean;
  ConnectionErrorText: string;
  Connections: GuiComponentCollection;
  HistoryEnabled: boolean;
  readonly MajorVersion: number;
  readonly MinorVersion: number;
  NewVisualDesign: boolean;
  readonly Patchlevel: number;
  readonly Revision: number;
  StatusbarVisible: boolean;
  TitlebarVisible: boolean;
  ToolbarVisible: boolean;
  readonly Utils: GuiUtils;
  AddHistoryEntry(Fieldname: string, a suggestion. With this function, Value: string): boolean;
  CreateGuiCollection(): any;
  DropHistory(): boolean;
  Ignore(WindowHandle: number): void;
  OpenConnection(Description: string, Sync?: any, Raise?: any): GuiConnection;
  OpenConnectionByConnectionString(ConnectString: string, Sync?: any, Raise?: any): GuiConnection;
  RegisterROT(): boolean;
  RevokeROT(): void;
  // Events
  onCreateSession?(handler: (Session: GuiSession) => void): void;
  onDestroySession?(handler: (Session: GuiSession) => void): void;
  onError?(handler: (ErrorId: number, Desc1: string, Desc2: string, Desc3: string, Desc4: string) => void): void;
  onIgnoreSession?(handler: (SessionMainWindowHandle: SAP) => void): void;
}
```

## GuiBarChart (extends GuiShell)
_Doc section 1.2.4, pages 47–51._

```ts
interface GuiBarChart extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  readonly ChartCount: number;
  BarCount(chartId: number): number;
  GetBarContent(chartId: number, barId: number, textId: number): string;
  GetGridLineContent(chartId: number, gridlineId: number, textId: number): string;
  GridCount(chartId: number): number;
  LinkCount(chartId: number): number;
  SendData(Data: string): void;
}
```

## GuiBox (extends GuiVComponent)
_Doc section 1.2.5, pages 51–53._

```ts
interface GuiBox extends GuiVComponent {
  // Inherits members from:
  // - GuiComponent
  // - GuiVComponent
  readonly CharHeight: number;
  readonly CharLeft: number;
  readonly CharTop: number;
  readonly CharWidth: number;
}
```

## GuiButton (extends GuiVComponent)
_Doc section 1.2.6, pages 53–55._

```ts
interface GuiButton extends GuiVComponent {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiVComponent
  readonly Emphasized: boolean;
  LeftLabel: GuiVComponent;
  RightLabel: GuiVComponent;
}
```

## GuiCalendar (extends GuiShell)
_Doc section 1.2.7, pages 55–61._

```ts
interface GuiCalendar extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  readonly endSelection: string;
  FirstVisibleDate: string;
  FocusDate: string;
  readonly FocusedElement: number;
  readonly horizontal: number;
  LastVisibleDate: string;
  SelectionInterval: string;
  startSelection: string;
  readonly Today: string;
  ContextMenu(CtxMenuId: number, CtxMenuCellRow: number, CtxMenuCellCol: number, DateBegin: string, DateEnd: string): void;
  CreateDate(day: number, month: number, year: number): string;
  GetColor(from: string): number;
  GetColorInfo(Color: number): string;
  GetDateTooltip(date: string): string;
  GetDay(date: string): number;
  GetMonth(date: string): number;
  GetWeekday(date: string): string;
  GetWeekNumber(date: string): number;
  GetYear(date: string): number;
  IsWeekend(date: string): number;
  SelectMonth(month: number, year: number): void;
  SelectRange(from: string, to: string): void;
  SelectWeek(week: number, year: number): void;
}
```

## GuiChart
_Doc section 1.2.8, pages 61–65._

```ts
interface GuiChart {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  ValueChange(Series: number, Point: number, XValue: string, YValue: string, DataChange: boolean, Id: string, ZValue: string, ChangeFlag: number): void;
}
```

## GuiCheckBox (extends GuiVComponent)
_Doc section 1.2.9, pages 65–68._

```ts
interface GuiCheckBox extends GuiVComponent {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiVComponent
  readonly ColorIndex: number;
  ColorIntensified: boolean;
  readonly ColorInverse: boolean;
  readonly Flushing: boolean;
  readonly IsLeftLabel: boolean;
  readonly IsListElement: boolean;
  readonly IsRightLabel: boolean;
  LeftLabel: GuiVComponent;
  RightLabel: GuiVComponent;
  readonly RowText: string;
  Selected: boolean;
  GetListProperty(Property: string): string;
  GetListPropertyNonRec(Property: string): string;
}
```

## GuiCollection
_Doc section 1.2.10, pages 68–69._

```ts
interface GuiCollection {
  readonly Count: number;
  readonly Length: number;
  readonly NewEnum: Unknown;
  readonly Type: string;
  readonly TypeAsNumber: number;
  Add(Item: any): void;
  ElementAt(Index: number, an exception is): any;
  Item(Index: any): any;
}
```

## GuiColorSelector (extends GuiShell)
_Doc section 1.2.11, pages 69–72._

```ts
interface GuiColorSelector extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  ChangeSelection(i: number): void;
}
```

## GuiComboBox (extends GuiVComponent)
_Doc section 1.2.12, pages 72–75._

```ts
interface GuiComboBox extends GuiVComponent {
  // Inherits members from:
  // - GuiComponent
  // - GuiVComponent
  readonly CharHeight: number;
  readonly CharLeft: number;
  readonly CharTop: number;
  readonly CharWidth: number;
  CurListBoxEntry: GuiComboBoxEntry;
  readonly Flushing: boolean;
  readonly Highlighted: boolean;
  readonly IsLeftLabel: boolean;
  IsListBoxActive: boolean;
  readonly IsRightLabel: boolean;
  Key: string;
  Modified: boolean;
  readonly Required: boolean;
  RightLabel: GuiVComponent;
  readonly ShowKey: boolean;
  Value: string;
  SetKeySpace(): void;
}
```

## GuiComboBoxControl
_Doc section 1.2.13, pages 75–78._

```ts
interface GuiComboBoxControl {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  Entries: GuiCollection;
  readonly LabelText: string;
  Selected: string;
  readonly Text: string;
  FireSelected(): void;
}
```

## GuiComboBoxEntry
_Doc section 1.2.14, pages 78–78._

```ts
interface GuiComboBoxEntry {
  readonly Key: string;
  readonly Pos: number;
  readonly Text: string;
  readonly Value: string;
}
```

## GuiComponent
_Doc section 1.2.15, pages 78–79._

```ts
interface GuiComponent {
  readonly ContainerType: boolean;
  readonly Id: string;
  readonly Name: string;
  readonly Parent: any;
  readonly Type: string;
  readonly TypeAsNumber: number;
}
```

## GuiComponentCollection
_Doc section 1.2.16, pages 79–81._

```ts
interface GuiComponentCollection {
  readonly Count: number;
  readonly Length: number;
  readonly NewEnum: Unknown;
  readonly Type: string;
  readonly TypeAsNumber: number;
  ElementAt(Index: number, the exception): GuiComponent;
  Item(Index: any): GuiComponent;
}
```

## GuiConnection (extends GuiContainer)
_Doc section 1.2.17, pages 81–83._

```ts
interface GuiConnection extends GuiContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  Children: GuiComponentCollection;
  ConnectionString: string;
  readonly Description: string;
  DisabledByServer: boolean;
  Sessions: GuiComponentCollection;
  CloseConnection(): void;
  CloseSession(Id: string, too.): void;
}
```

## GuiContainer (extends GuiComponent)
_Doc section 1.2.18, pages 83–84._

```ts
interface GuiContainer extends GuiComponent {
  // Inherits members from:
  // - GuiComponent
  Children: GuiComponentCollection;
  FindById(Id: string, Raise?: any, this prefix is truncated. If no descendant with the): GuiComponent;
}
```

## GuiContainerShell (extends GuiVContainer)
_Doc section 1.2.19, pages 84–87._

```ts
interface GuiContainerShell extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  DockerIsVertical: boolean;
  DockerPixelSize: number;
}
```

## GuiContextMenu (extends GuiMenu)
_Doc section 1.2.20, pages 87–88._

```ts
interface GuiContextMenu extends GuiMenu {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  Select(): void;
}
```

## GuiCTextField (extends GuiTextField)
_Doc section 1.2.21, pages 88–91._

```ts
interface GuiCTextField {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiVComponent
}
```

## GuiCustomControl (extends GuiVContainer)
_Doc section 1.2.22, pages 91–93._

```ts
interface GuiCustomControl extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  readonly CharHeight: number;
  readonly CharLeft: number;
  readonly CharTop: number;
  readonly CharWidth: number;
}
```

## GuiDialogShell (extends GuiVContainer)
_Doc section 1.2.23, pages 93–96._

```ts
interface GuiDialogShell extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  readonly Title: string;
  Close(): void;
}
```

## GuiEAIViewer2D (extends GuiShell)
_Doc section 1.2.24, pages 96–98._

```ts
interface GuiEAIViewer2D extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  AnnotationEnabled: number;
  AnnotationMode: number;
  RedliningStream: string;
  annotationTextRequest(strText: string): void;
}
```

## GuiEAIViewer3D (extends GuiShell)
_Doc section 1.2.25, pages 98–101._

```ts
interface GuiEAIViewer3D extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
}
```

## GuiEnum (extends GuiVContainer)
_Doc section 1.2.26, pages 101–101._

```ts
interface GuiEnum extends GuiVContainer {
  Clone(ppenum: GuiEnum): HResult;
  Next(celt: ULong, rgvar: any, pceltFetched: ULong): HResult;
  Reset(): HResult;
  Skip(celt: ULong): HResult;
}
```

## GuiFrameWindow (extends GuiVContainer)
_Doc section 1.2.27, pages 101–108._

```ts
interface GuiFrameWindow extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  Close(): void;
  CompBitmap(Filename1: string, Filename2: string): number;
  HardCopy(Filename: string, tion succeeds, ImageType?: path, then the fi le Variant, xPos?: any, yPos?: any, nWidth?: any, nHeight?: any): string;
  HardCopyToMemory(ImageType?: The): any;
  Iconify(): void;
  IsVKeyAllowed(VKey: number): boolean;
  JumpBackward(): void;
  JumpForward(): void;
  Maximize(): void;
  Restore(): void;
  SendVKey(VKey: number): void;
  ShowMessageBox(Title: string, Text: string, MsgIcon: number, MsgType: number): number;
  TabBackward(): void;
  TabForward(): void;
}
```

## GuiGOSShell (extends GuiVContainer)
_Doc section 1.2.28, pages 108–110._

```ts
interface GuiGOSShell extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
}
```

## GuiGraphAdapt
_Doc section 1.2.29, pages 110–113._

```ts
interface GuiGraphAdapt {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
}
```

## GuiGridView (extends GuiShell)
_Doc section 1.2.30, pages 113–128._

```ts
interface GuiGridView extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  readonly ColumnCount: number;
  ColumnOrder: any;
  CurrentCellColumn: string;
  CurrentCellRow: number;
  FirstVisibleColumn: string;
  FirstVisibleRow: number;
  FrozenColumnCount: number;
  readonly RowCount: number;
  SelectedCells: any;
  SelectedColumns: any;
  SelectedRows: string;
  readonly Title: string;
  ToolbarButtonCount: number;
  VisibleRowCount: number;
  ClearSelection(): void;
  Click(Row: number, Column: string): void;
  ClickCurrentCell(): void;
  ContextMenu(): void;
  CurrentCellMoved(): void;
  DeleteRows(Rows: string, otherwise an exception is raised.): void;
  DeselectColumn(Column: string): void;
  DoubleClick(Row: number, Column: string): void;
  DoubleClickCurrentCell(): void;
  DuplicateRows(Rows: string): void;
  GetCellChangeable(Row: number, Column: string): boolean;
  GetCellCheckBoxChecked(Row: number, Column: string): boolean;
  GetCellColor(Row: number, Column: string): number;
  GetCellHeight(Row: number, Column: string): number;
  GetCellIcon(Row: number, Column: string): string;
  GetCellLeft(Row: number, Column: string): number;
  GetCellMaxLength(Row: number, Column: string): number;
  GetCellState(Row: number, Column: string): string;
  GetCellTooltip(Row: number, Column: string): string;
  GetCellTop(Row: number, Column: string): number;
  GetCellType(Row: number, Column: string): string;
  GetCellValue(Row: number, Column: string): string;
  GetCellWidth(Row: number, Column: string): number;
  GetColorInfo(Color: number): string;
  GetColumnDataType(Column: string): string;
  GetColumnPosition(Column: string): number;
  GetColumnSortType(Column: string): string;
  GetColumnTitles(Column: string): any;
  GetColumnTooltip(Column: string): string;
  GetColumnTotalType(Column: string): string;
  GetDisplayedColumnTitle(Column: string): string;
  GetRowTotalLevel(Row: number): number;
  GetSymbolInfo(Symbol: string): string;
  GetToolbarButtonChecked(ButtonPos: number): boolean;
  GetToolbarButtonEnabled(ButtonPos: number): boolean;
  GetToolbarButtonIcon(ButtonPos: number): string;
  GetToolbarButtonId(ButtonPos: number): string;
  GetToolbarButtonText(ButtonPos: number): string;
  GetToolbarButtonTooltip(ButtonPos: number): string;
  GetToolbarFocusButton(): number;
  HasCellF4Help(Row: number, Column: string): boolean;
  HistoryCurEntry(Row: number, Column: string): string;
  HistoryCurIndex(Row: number, Column: string): number;
  HistoryIsActive(Row: number, Column: string): boolean;
  HistoryList(Row: number, Column: string): GuiCollection;
  InsertRows(Rows: string, moving the old): void;
  IsCellHotspot(Row: number, Column: string): boolean;
  IsCellSymbol(Row: number, Column: string): boolean;
  IsCellTotalExpander(Row: number, Column: string): boolean;
  IsColumnFiltered(Column: string): boolean;
  IsColumnKey(Column: string): boolean;
  IsTotalRowExpanded(Row: number): boolean;
  ModifyCell(Row: number, Otherwise, Column: string, Value: string): void;
  ModifyCheckBox(Row: number, Column: string, Checked: Boolean): void;
  MoveRows(FromRow: number, ToRow: number, DestRow: number): void;
  PressButton(Row: number, button, Column: string): void;
  PressButtonCurrentCell(): void;
  PressColumnHeader(Column: string): void;
  PressEnter(): void;
  PressF1(): void;
  PressF4(): void;
  PressToolbarButton(Id: string): void;
  PressToolbarContextButton(Id: string): void;
  PressTotalRow(Row: number, Column: string): void;
  PressTotalRowCurrentCell(): void;
  SelectAll(): void;
  SelectColumn(Column: string): void;
  SelectionChanged(): void;
  SelectToolbarMenuItem(Id: string): void;
  SetColumnWidth(Column: string, Width: number): void;
  SetCurrentCell(Row: number, Column: string): void;
  TriggerModified(): void;
}
```

## GuiHTMLViewer (extends GuiShell)
_Doc section 1.2.31, pages 128–132._

```ts
interface GuiHTMLViewer extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  BrowserHandle: any;
  DocumentComplete: number;
  ContextMenu(): void;
  SapEvent(FrameName: string, PostData: string, If the form is to be submitted using the GET method, Url: string): void;
}
```

## GuiInputFieldControl
_Doc section 1.2.32, pages 132–134._

```ts
interface GuiInputFieldControl {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  ButtonTooltip: string;
  FindButtonActivated: Boolean;
  readonly HistoryOpened: boolean;
  readonly LabelText: string;
  Text: string;
  Submit(): void;
}
```

## GuiLabel (extends GuiVComponent)
_Doc section 1.2.33, pages 134–141._

```ts
interface GuiLabel extends GuiVComponent {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiVComponent
  CaretPosition: number;
  readonly ColorIndex: number;
  ColorIntensified: boolean;
  readonly ColorInverse: boolean;
  DisplayedText: string;
  readonly Highlighted: boolean;
  readonly IsHotspot: boolean;
  readonly IsLeftLabel: boolean;
  readonly IsListElement: boolean;
  readonly IsRightLabel: boolean;
  readonly MaxLength: number;
  readonly Numerical: boolean;
  readonly RowText: string;
  GetListProperty(Property: string): string;
}
```

## GuiMainWindow (extends GuiFrameWindow)
_Doc section 1.2.34, pages 141–147._

```ts
interface GuiMainWindow extends GuiFrameWindow {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiFrameWindow
  // - GuiVComponent
  ButtonbarVisible: boolean;
  StatusbarVisible: hide;
  TitlebarVisible: boolean;
  ToolbarVisible: boolean;
  ResizeWorkingPane(Width: number, Height: number, ThrowOnFail: Boolean): void;
  ResizeWorkingPaneEx(Width: number, Height: number, ThrowOnFail: Boolean): void;
}
```

## GuiMap
_Doc section 1.2.35, pages 147–150._

```ts
interface GuiMap {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
}
```

## GuiMenu (extends GuiVContainer)
_Doc section 1.2.36, pages 150–151._

```ts
interface GuiMenu extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  Select(): void;
}
```

## GuiMenubar (extends GuiVContainer)
_Doc section 1.2.37, pages 151–153._

```ts
interface GuiMenubar extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
}
```

## GuiMessageWindow
_Doc section 1.2.38, pages 153–155._

```ts
interface GuiMessageWindow {
  // Inherits members from:
  // - GuiComponent
  // - GuiVComponent
  readonly FocusedButton: number;
  HelpButtonHelpText: string;
  HelpButtonText: string;
  MessageText: string;
  MessageType: number;
  OKButtonHelpText: string;
  OKButtonText: string;
  Visible: boolean;
}
```

## GuiModalWindow (extends GuiFrameWindow)
_Doc section 1.2.39, pages 155–158._

```ts
interface GuiModalWindow extends GuiFrameWindow {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiFrameWindow
  // - GuiVComponent
  // - GuiVContainer
  PopupDialogText: string;
}
```

## GuiNetChart
_Doc section 1.2.40, pages 158–161._

```ts
interface GuiNetChart {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  readonly LinkCount: number;
  readonly NodeCount: number;
  GetLinkContent(linkId: number, textId: number, do be determined during recording.): string;
  GetNodeContent(nodeId: number, textId: number, do be determined during recording.): string;
  SendData(Data: string): void;
}
```

## GuiOfficeIntegration
_Doc section 1.2.41, pages 161–164._

```ts
interface GuiOfficeIntegration {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  readonly Document: any;
  readonly HostedApplication: are;
  AppendRow(Name: string, Row: string): void;
  CloseDocument(Cookie: number, EverChanged: boolean, ChangedAfterSave: boolean): void;
  CustomEvent(Cookie: number, EventName: string, ParamCount: number, Par1?: any, Par2?: any, Par3?: any, Par4?: any, Par5?: any, Par6?: any, Par7?: any, Par8?: any, Par9?: any, Par10?: any, Par11?: any, Par12?: any): void;
  RemoveContent(Name: string): void;
  SaveDocument(Cookie: number, Changed: boolean): void;
  SetDocument(Index: number, Document: string): void;
}
```

## GuiOkCodeField (extends GuiVComponent)
_Doc section 1.2.42, pages 164–166._

```ts
interface GuiOkCodeField extends GuiVComponent {
  // Inherits members from:
  // - GuiComponent
  // - GuiVComponent
  readonly Opened: boolean;
  PressF1(): void;
}
```

## GuiPasswordField (extends GuiTextField)
_Doc section 1.2.43, pages 166–169._

```ts
interface GuiPasswordField {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiVComponent
}
```

## GuiPicture (extends GuiShell)
_Doc section 1.2.44, pages 169–173._

```ts
interface GuiPicture extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  Click(): void;
  ClickControlArea(x: number, y: number): void;
  ClickPictureArea(x: number, y: displayed): void;
  ContextMenu(x: number, y: number): void;
  DoubleClick(): void;
  DoubleClickControlArea(x: number, y: number): void;
  DoubleClickPictureArea(x: number, y: nates): void;
}
```

## GuiRadioButton (extends GuiVComponent)
_Doc section 1.2.45, pages 173–177._

```ts
interface GuiRadioButton extends GuiVComponent {
  // Inherits members from:
  // - GuiButton
  // - GuiComponent
  // - GuiVComponent
  readonly CharHeight: number;
  readonly CharLeft: number;
  readonly CharTop: number;
  readonly CharWidth: number;
  readonly Flushing: boolean;
  readonly GroupCount: number;
  GroupMembers: The;
  readonly GroupPos: number;
  readonly IsLeftLabel: boolean;
  readonly IsRightLabel: boolean;
  LeftLabel: GuiVComponent;
  RightLabel: uiVComponent;
  Select(): void;
}
```

## GuiSapChart
_Doc section 1.2.46, pages 177–179._

```ts
interface GuiSapChart {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
}
```

## GuiScrollbar
_Doc section 1.2.47, pages 179–179._

```ts
interface GuiScrollbar {
  readonly Maximum: number;
  readonly Minimum: number;
  readonly PageSize: number;
  Position: number;
}
```

## GuiScrollContainer
_Doc section 1.2.48, pages 179–181._

```ts
interface GuiScrollContainer {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  HorizontalScrollbar: GuiScrollbar;
  VerticalScrollbar: GuiScrollbar;
}
```

## GuiSession (extends GuiContainer)
_Doc section 1.2.49, pages 181–192._

```ts
interface GuiSession extends GuiContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  ActiveWindow: GuiFrameWindow;
  Busy: boolean;
  ErrorList: GuiCollection;
  readonly Info: GuiSessionInfo;
  IsActive: boolean;
  IsListBoxActive: boolean;
  ListBoxCurrEntry: number;
  ListBoxCurrEntryHeight: number;
  ListBoxCurrEntryLeft: number;
  ListBoxCurrEntryTop: number;
  ListBoxCurrEntryWidth: number;
  readonly ListBoxHeight: number;
  readonly ListBoxLeft: number;
  readonly ListBoxTop: number;
  readonly ListBoxWidth: number;
  PassportPreSystemId: string;
  PassportSystemId: string;
  PassportTransactionId: string;
  ProgressPercent: LongPublic;
  readonly ProgressText: string;
  Record: boolean;
  RecordFile: string;
  SaveAsUnicode: boolean;
  ShowDropdownKeys: boolean;
  SuppressBackendPopups: boolean;
  TestToolMode: number;
  AsStdNumberFormat(Number: string): string;
  ClearErrorList(): void;
  CreateSession(): void;
  EnableJawsEvents(): void;
  EndTransaction(): void;
  FindByPosition(x: number, y: number, Raise?: any): GuiCollection;
  GetIconResourceName(Text: string): string;
  GetVKeyDescription(VKey: number): string;
  LockSessionUI(): void;
  SendCommand(Command: string): void;
  StartTransaction(Transaction: string): void;
  UnlockSessionUI(): void;
  // Events
  onAbapScriptingEvent?(handler: (param: string) => void): void;
  onActivated?(handler: (Session: GuiSession) => void): void;
  onAutomationFCode?(handler: (Session: GuiSession, FunctionCode: string) => void): void;
  onChange?(handler: (Session: GuiSession, Component: GuiComponent, CommandArray: any) => void): void;
  onContextMenu?(handler: (Session: GuiSession, Component: GuiVComponent) => void): void;
  onDestroy?(handler: (Session: GuiSession) => void): void;
  onEndRequest?(handler: (Session: GuiSession) => void): void;
  onError?(handler: (Session: GuiSession, ErrorId: number, Desc1: string, Desc2: string, Desc3: string, Desc4: string) => void): void;
  onFocusChanged?(handler: (Session: GuiSession, NewFocusedControl: GuiVComponent) => void): void;
  onHistoryOpened?(handler: (Session: GuiSession, NewFocusedControl: GuiVComponent) => void): void;
  onHit?(handler: (Session: GuiSession, in this mode a SAP GUI component is identified, Component: GuiComponent, InnerObject: string) => void): void;
  onProgressIndicator?(handler: (percentage: number, Text: string) => void): void;
  onStartRequest?(handler: (Session: GuiSession) => void): void;
}
```

## GuiSessionInfo
_Doc section 1.2.50, pages 192–195._

```ts
interface GuiSessionInfo {
  ApplicationServer: string;
  readonly Client: string;
  readonly Codepage: number;
  readonly Flushes: number;
  readonly Group: string;
  readonly GuiCodepage: number;
  readonly I18NMode: boolean;
  InterpretationTime: number;
  IsLowSpeedConnection: boolean;
  readonly Language: string;
  MessageServer: string;
  readonly Program: string;
  readonly ResponseTime: number;
  readonly RoundTrips: number;
  readonly ScreenNumber: number;
  ScriptingModeReadOnly: boolean;
  ScriptingModeRecordingDisabled: boolean;
  readonly SessionNumber: number;
  readonly SystemName: string;
  readonly SystemNumber: number;
  SystemSessionId: string;
  readonly Transaction: string;
  readonly UI_GUIDELINE: string;
  readonly User: string;
}
```

## GuiShell (extends GuiVContainer)
_Doc section 1.2.51, pages 195–198._

```ts
interface GuiShell extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  AccDescription: string;
  CurrentContextMenu: GuiContextMenu;
  DragDropSupported: boolean;
  readonly Handle: number;
  OcxEvents: GuiCollection;
  readonly SubType: string;
  SelectContextMenuItem(FunctionCode: string): void;
  SelectContextMenuItemByPosition(PositionDesc: string): void;
  SelectContextMenuItemByText(Text: string): void;
}
```

## GuiSimpleContainer (extends GuiVContainer)
_Doc section 1.2.52, pages 198–201._

```ts
interface GuiSimpleContainer extends GuiVContainer {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  readonly IsListElement: boolean;
  readonly IsStepLoop: boolean;
  readonly LoopColCount: number;
  readonly LoopCurrentCol: number;
  LoopCurrentColCount: number;
  readonly LoopCurrentRow: number;
  readonly LoopRowCount: number;
  GetListProperty(Property: string): string;
  GetListPropertyNonRec(Property: string): string;
}
```

## GuiSplit (extends GuiShell)
_Doc section 1.2.53, pages 201–204._

```ts
interface GuiSplit extends GuiShell {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  readonly IsVertical: number;
  GetColSize(Id: number): number;
  GetRowSize(Id: number): number;
  SetColSize(Id: number, Size: number): void;
  SetRowSize(Id: number, Size: number): void;
}
```

## GuiSplitterContainer
_Doc section 1.2.54, pages 204–206._

```ts
interface GuiSplitterContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  readonly IsVertical: boolean;
  SashPosition: number;
}
```

## GuiStage
_Doc section 1.2.55, pages 206–209._

```ts
interface GuiStage {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  ContextMenu(strId: string): void;
  DoubleClick(strId: string): void;
  SelectItems(strItems: string): void;
}
```

## GuiStatusbar (extends GuiVComponent)
_Doc section 1.2.56, pages 209–212._

```ts
interface GuiStatusbar extends GuiVComponent {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  readonly Handle: number;
  readonly MessageAsPopup: boolean;
  readonly MessageId: string;
  MessageNumber: string;
  MessageParameter: string;
  readonly MessageType: string;
  DoubleClick(): void;
}
```

## GuiStatusPane
_Doc section 1.2.57, pages 212–213._

```ts
interface GuiStatusPane {
  // Inherits members from:
  // - GuiComponent
  // - GuiVComponent
}
```

## GuiTab (extends GuiVContainer)
_Doc section 1.2.58, pages 213–215._

```ts
interface GuiTab extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  ScrollToLeft(): void;
  Select(): void;
}
```

## GuiTableColumn (extends GuiComponentCollection)
_Doc section 1.2.59, pages 215–217._

```ts
interface GuiTableColumn extends GuiComponentCollection {
  readonly Count: number;
  DefaultTooltip: string;
  readonly Fixed: boolean;
  readonly IconName: string;
  readonly Length: number;
  readonly NewEnum: Unknown;
  Selected: boolean;
  readonly Title: string;
  readonly Tooltip: string;
  readonly Type: string;
  readonly TypeAsNumber: number;
  ElementAt(member can be found for the given index, Index: number): GuiComponent;
  Item(Index: any, an): GuiComponent;
}
```

## GuiTableControl (extends theGuiVContainer)
_Doc section 1.2.60, pages 217–221._

```ts
interface GuiTableControl {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  readonly CharHeight: number;
  readonly CharLeft: number;
  readonly CharTop: number;
  readonly CharWidth: number;
  ColSelectMode: GuiTableSelectionType;
  Columns: GuiCollection;
  readonly CurrentCol: number;
  readonly CurrentRow: number;
  HorizontalScrollbar: GuiScrollbar;
  readonly RowCount: number;
  readonly Rows: GuiCollection;
  RowSelectMode: GuiTableSelectionType;
  TableFieldName: string;
  VerticalScrollbar: GuiScrollbar;
  VisibleRowCount: number;
  ConfigureLayout(): void;
  DeselectAllColumns(): void;
  GetAbsoluteRow(Index: number): GuiTableRow;
  GetCell(Row: number, Column: number): GuiVComponent;
  ReorderTable(Permutation: string): void;
  SelectAllColumns(): void;
}
```

## GuiTableRow (extends GuiComponentCollection)
_Doc section 1.2.61, pages 221–223._

```ts
interface GuiTableRow extends GuiComponentCollection {
  readonly Count: number;
  readonly Length: number;
  readonly NewEnum: Unknown;
  Selectable: boolean;
  Selected: boolean;
  readonly Type: string;
  readonly TypeAsNumber: number;
  ElementAt(member can be found for the given index, Index: number): GuiComponent;
  Item(Index: any, an): GuiComponent;
}
```

## GuiTabStrip (extends GuiVContainer)
_Doc section 1.2.62, pages 223–226._

```ts
interface GuiTabStrip extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  // - GuiVContainer
  readonly CharHeight: number;
  readonly CharLeft: number;
  readonly CharTop: number;
  readonly CharWidth: number;
  readonly LeftTab: GuiTab;
  readonly SelectedTab: GuiTab;
}
```

## GuiTextedit (extends GuiShell)
_Doc section 1.2.63, pages 226–235._

```ts
interface GuiTextedit extends GuiShell {
  // Inherits members from:
  // - GuiBox
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  CaretPosition: number;
  readonly CurrentColumn: number;
  readonly CurrentLine: number;
  DisplayedText: string;
  FirstVisibleLine: number;
  readonly Highlighted: boolean;
  HistoryCurEntry: string;
  HistoryCurIndex: number;
  HistoryIsActive: boolean;
  HistoryList: GuiCollection;
  readonly IsHotspot: boolean;
  readonly IsLeftLabel: boolean;
  readonly IsListElement: boolean;
  readonly IsOField: boolean;
  readonly IsRightLabel: boolean;
  LastVisibleLine: number;
  LeftLabel: GuiVComponent;
  readonly LineCount: number;
  readonly MaxLength: number;
  NumberOfUnprotectedTextParts: number;
  readonly Numerical: boolean;
  readonly Required: boolean;
  RightLabel: GuiVComponent;
  readonly SelectedText: string;
  SelectionEndColumn: number;
  SelectionEndLine: number;
  SelectionIndexEnd: number;
  SelectionIndexStart: number;
  SelectionStartColumn: number;
  SelectionStartLine: number;
  ContextMenu(): void;
  DoubleClick(): void;
  GetLineText(nLine: number): string;
  GetListProperty(Property: string): string;
  GetListPropertyNonRec(Property: string): string;
  GetUnprotectedTextPart(Part: number): string;
  IsBreakpointLine(nLine: number): boolean;
  IsCommentLine(nLine: number): boolean;
  IsHighlightedLine(nLine: number): boolean;
  IsProtectedLine(nLine: number): boolean;
  IsSelectedLine(nLine: number): boolean;
  ModifiedStatusChanged(Status: Boolean): void;
  MultipleFilesDropped(): void;
  PressF1(): void;
  PressF4(): void;
  SetSelectionIndexes(Start: number, End: the): void;
  SetUnprotectedTextPart(True if it was possible to perform the assignment. Otherwise, Part: number, Text: string): boolean;
  SingleFileDropped(Filename: string): void;
}
```

## GuiTitlebar (extends theGuiVContainer)
_Doc section 1.2.65, pages 235–237._

```ts
interface GuiTitlebar {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
}
```

## GuiToolbar (extends GuiVContainer)
_Doc section 1.2.66, pages 237–243._

```ts
interface GuiToolbar extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  ButtonCount: number;
  FocusedButton: number;
  GetButtonChecked(ButtonPos: number): boolean;
  GetButtonEnabled(ButtonPos: number): boolean;
  GetButtonIcon(ButtonPos: number): string;
  GetButtonId(ButtonPos: number): string;
  GetButtonText(ButtonPos: number): string;
  GetButtonTooltip(ButtonPos: number): string;
  GetButtonType(ButtonPos: number, "Group", "CheckBox"): string;
  GetMenuItemIdFromPosition(Pos: number): string;
  PressButton(Id: string): void;
  PressContextButton(Id: string): void;
  SelectMenuItem(Id: string): void;
  SelectMenuItemByText(strText: string): void;
}
```

## GuiTree
_Doc section 1.2.68, pages 243–256._

```ts
interface GuiTree {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  ColumnOrder: any;
  SelectedNode: string;
  TopNode: string;
  ChangeCheckbox(NodeKey: string, ItemName: string, Checked: Boolean): void;
  ClickLink(NodeKey: string, ItemName: string): void;
  CollapseNode(NodeKey: string): void;
  DefaultContextMenu(): void;
  DoubleClickItem(NodeKey: string, ItemName: string): void;
  DoubleClickNode(NodeKey: string): void;
  EnsureVisibleHorizontalItem(NodeKey: string, ItemName: string): void;
  ExpandNode(NodeKey: string): void;
  FindNodeKeyByPath(Path: string): string;
  GetAbapImage(Key: string, Name: string): string;
  GetAllNodeKeys(): any;
  GetCheckBoxState(NodeKey: string, ItemName: string): boolean;
  GetColumnCol(colName: string): any;
  GetColumnHeaders(): any;
  GetColumnIndexFromName(Key: string): number;
  GetColumnNames(): any;
  GetColumnTitleFromName(Key: string): string;
  GetColumnTitles(): any;
  GetFocusedNodeKey(): string;
  GetHierarchyLevel(Key: string): number;
  GetHierarchyTitle(): string;
  GetIsDisabled(NodeKey: string, ItemName: string): boolean;
  GetIsHighLighted(NodeKey: string, ItemName: string): boolean;
  GetItemHeight(NodeKey: string, ItemName: string): number;
  GetItemLeft(NodeKey: string, ItemName: string): number;
  GetItemStyle(NodeKey: string, ItemName: string): number;
  GetItemText(Key: string, Name: string): string;
  GetItemTextColor(Key: string, Name: string): ULong;
  GetItemToolTip(Key: string, Name: string): string;
  GetItemTop(NodeKey: string, ItemName: string): number;
  GetItemType(Key: string, Name: string): number;
  GetItemWidth(NodeKey: string, ItemName: string): number;
  GetListTreeNodeItemCount(NodeKey: string): number;
  GetNextNodeKey(NodeKey: string): string;
  GetNodeAbapImage(Key: string): string;
  GetNodeChildrenCount(Key: string): number;
  GetNodeChildrenCountByPath(Path: string): number;
  GetNodeHeight(Key: string): number;
  GetNodeIndex(Key: string): number;
  GetNodeItemHeaders(NodeKey: string): any;
  GetNodeKeyByPath(Path: string): string;
  GetNodeLeft(Key: string): number;
  GetNodePathByKey(Key: string): string;
  GetNodesCol(): any;
  GetNodeStyle(NodeKey: string): number;
  GetNodeTextByKey(Path: string): string;
  GetNodeTextByPath(Path: string): string;
  GetNodeTextColor(Key: string): ULong;
  GetNodeToolTip(NodeKey: string): string;
  GetNodeTop(Key: string): number;
  GetNodeWidth(Key: string): number;
  GetParent(CKey: string): string;
  GetPreviousNodeKey(NodeKey: string): string;
  GetSelectedNodes(): any;
  GetSelectionMode(): number;
  GetStyleDescription(nStyle: number): string;
  GetSubNodesCol(Path: string): any;
  GetTreeType(): number;
  HeaderContextMenu(HeaderName: string): void;
  IsFolder(NodeKey: string): boolean;
  IsFolderExpandable(NodeKey: string): boolean;
  IsFolderExpanded(NodeKey: string): boolean;
  ItemContextMenu(NodeKey: string, ItemName: string): void;
  NodeContextMenu(NodeKey: string): void;
  PressButton(NodeKey: string, ItemName: string): void;
  PressHeader(HeaderName: string): void;
  PressKey(Key: string): void;
  SelectColumn(ColumnName: string): void;
  SelectedItemColumn(): string;
  SelectedItemNode(): string;
  SelectItem(NodeKey: string, ItemName: string): void;
  SelectNode(NodeKey: string): void;
  SetCheckBoxState(checkbox is unchecked, NodeKey: string, ItemName: string, state: number): void;
  SetColumnWidth(ColumnName: string, Width: number): void;
  UnselectAll(): void;
  UnselectColumn(ColumnName: string): void;
  UnselectNode(NodeKey: string): void;
}
```

## GuiUserArea (extends GuiVContainer)
_Doc section 1.2.69, pages 256–259._

```ts
interface GuiUserArea extends GuiVContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiShell
  // - GuiVComponent
  // - GuiVContainer
  CurrentContextMenu: GuiContextMenu;
  HorizontalScrollbar: GuiScrollbar;
  readonly IsOTFPreview: boolean;
  VerticalScrollbar: GuiScrollbar;
  FindByLabel(Text: string, Type: string): GuiComponent;
}
```

## GuiUtils
_Doc section 1.2.70, pages 259–262._

```ts
interface GuiUtils {
  MESSAGE_OPTION_OK: number;
  MESSAGE_OPTION_OKCANCEL: number;
  MESSAGE_OPTION_YESNO: number;
  MESSAGE_RESULT_CANCEL: number;
  MESSAGE_RESULT_NO: number;
  MESSAGE_RESULT_OK: number;
  MESSAGE_RESULT_YES: number;
  MESSAGE_TYPE_ERROR: number;
  MESSAGE_TYPE_INFORMATION: number;
  MESSAGE_TYPE_PLAIN: number;
  MESSAGE_TYPE_QUESTION: number;
  MESSAGE_TYPE_WARNING: number;
  CloseFile(File: number): void;
  OpenFile(Name: string): number;
  ShowMessageBox(Title: string, Text: string, MsgIcon: number, MsgType: number): number;
  Write(File: number, Text: string): void;
  WriteLine(File: number, Text: string): void;
}
```

## GuiVComponent (extends GuiComponent)
_Doc section 1.2.71, pages 262–267._

```ts
interface GuiVComponent extends GuiComponent {
  AccLabelCollection: GuiComponentCollection;
  readonly AccText: string;
  AccTextOnRequest: string;
  readonly AccTooltip: string;
  readonly Changeable: boolean;
  DefaultTooltip: string;
  readonly Height: number;
  readonly IconName: string;
  readonly IsSymbolFont: boolean;
  readonly Left: number;
  Modified: boolean;
  ParentFrame: GuiComponent;
  readonly ScreenLeft: number;
  readonly ScreenTop: number;
  Text: string;
  readonly Tooltip: string;
  Top: number;
  readonly Width: number;
  DumpState(InnerObject: string): GuiCollection;
  SetFocus(): void;
  Visualize(On: Boolean, InnerObject?: such, such as cells. The format of the inner): boolean;
}
```

## GuiVContainer (extends GuiContainer)
_Doc section 1.2.72, pages 267–269._

```ts
interface GuiVContainer extends GuiContainer {
  // Inherits members from:
  // - GuiComponent
  // - GuiContainer
  // - GuiVComponent
  FindAllByName(Name: string, may however be several matching objects, Type: string): GuiComponentCollection;
  FindAllByNameEx(Name: string, may however be several matching objects, Type: number): GuiComponentCollection;
  FindByName(Name: string, Type: string, but it): GuiComponent;
}
```

## GuiVHViewSwitch (extends GuiVComponent)
_Doc section 1.2.73, pages 269–302._

```ts
interface GuiVHViewSwitch extends GuiVComponent {
  // Inherits members from:
  // - GuiComponent
  // - GuiVComponent
  GetScriptingEngine() This method returns the GuiApplication COM interface for As Object the currently running SAP GUI process. Return Type: GuiApplication  Example Visual Basic Script (Visual Basic) Set SapGuiAuto = GetObject( "SAPGUI") Set application = SapGuiAuto.GetScriptingEngine WinBatch by Wilson WindowWare (Visual Basic) SapGuiAuto = ObjectGet("SAPGUI"); Application = SapGuiAuto.GetScriptingEngine; SAP GUI Scripting API 296 PUBLIC SAP GUI Scripting ROT Entry Helper 3 SAP GUI Scripting ROT Access Helper Description This library contains just one object. It is required in cases where the scripting engine does not allow accessing entries in the Running Object Table. This is for example the case for the following engines ● JAWS script, the scripting language of the JAWS Screenreader, by Freedom Scientific ● AutoIt, a freeware Windows automation language For these languages you need to create a CSapROTWrapper object first, before you can access the ROT entry and get the GuiApplication interface of the running SAP GUI process. 3.1 CSapROTWrapper Object SAP GUI Scripting API SAP GUI Scripting ROT Access Helper PUBLIC 297 Methods Syntax Method Visual Basic Parameter GetROTEntry strDisplayName Public Function This method returns the ROT entry of SAP GUI. G etROTEntry( ByVal strDisplayName As  Example String): any;
  // Events
  onAbapScriptingEvent?(handler: (param: string) => void): void;
  onAbapScriptingEvent?(handler: (param: string) => void): void;
  onActivated?(handler: (Session: GuiSession) => void): void;
  onActivated?(handler: (Session: GuiSession) => void): void;
  onAutomationFCode?(handler: (Session: GuiSession, FunctionCode: string) => void): void;
  onAutomationFCode?(handler: (Session: GuiSession, FunctionCode: string) => void): void;
  onChange?(handler: (Session: GuiSession, Component: GuiComponent, CommandArray: any) => void): void;
  onChange?(handler: (Session: GuiSession, Component: GuiComponent, CommandArray: any) => void): void;
  onContextMenu?(handler: (Session: GuiSession, Component: GuiVComponent) => void): void;
  onContextMenu?(handler: (Session: GuiSession, Component: GuiVComponent) => void): void;
  onCreateSession?(handler: (Session: GuiSession) => void): void;
  onCreateSession?(handler: (Session: GuiSession) => void): void;
  onDestroy?(handler: (Session: GuiSession) => void): void;
  onDestroy?(handler: (Session: GuiSession) => void): void;
  onDestroySession?(handler: (Session: GuiSession) => void): void;
  onDestroySession?(handler: (Session: GuiSession) => void): void;
  onEndRequest?(handler: (SessionSession: GuiSession) => void): void;
  onEndRequest?(handler: (SessionSession: GuiSession) => void): void;
  onError?(handler: (Session: GuiSession, ErrorId: number, Desc1: string, Desc2: string, Desc3: string, Desc4: string) => void): void;
  onError?(handler: (ErrorId: number, Desc1: string, Desc2: string, Desc3: string, Desc4: string) => void): void;
  onError?(handler: (Session: GuiSession, ErrorId: number, Desc1: string, Desc2: string, Desc3: string, Desc4: string) => void): void;
  onError?(handler: (ErrorId: number, Desc1: string, Desc2: string, Desc3: string, Desc4: string) => void): void;
  onFocusChanged?(handler: (Session: GuiSession, NewFocusedControl: GuiVComponent) => void): void;
  onFocusChanged?(handler: (Session: GuiSession, NewFocusedControl: GuiVComponent) => void): void;
  onHistoryOpened?(handler: (Session: GuiSession, NewFocusedControl: GuiVComponent) => void): void;
  onHistoryOpened?(handler: (Session: GuiSession, NewFocusedControl: GuiVComponent) => void): void;
  onHit?(handler: (SessionSession: GuiSession, GUI. If in this mode a SAP GUI component is identified, Component: GuiComponent, InnerObject: string) => void): void;
  onHit?(handler: (SessionSession: GuiSession, GUI. If in this mode a SAP GUI component is identified, Component: GuiComponent, InnerObject: string) => void): void;
  onIgnoreSession?(handler: (SessionMainWindowHandle: SAP) => void): void;
  onIgnoreSession?(handler: (SessionMainWindowHandle: SAP) => void): void;
  onProgressIndicator?(handler: (percentage: number, Text: string) => void): void;
  onProgressIndicator?(handler: (percentage: number, Text: string) => void): void;
  onStartRequest?(handler: (Session: GuiSession) => void): void;
  onStartRequest?(handler: (Session: GuiSession) => void): void;
}
```

# Enumerations

## GuiComponentType
```ts
enum GuiComponentType {
  GuiApplication = 10,
  GuiBox = 62,
  GuiButton = 40,
  GuiCheckBox = 42,
  GuiCollection = 120,
  GuiComboBox = 34,
  GuiComponent = 0,
  GuiComponentCollection = 128,
  GuiConnection = 11,
  GuiContainer = 70,
  GuiContainerShell = 51,
  GuiContextMenu = 127,
  GuiCTextField = 32,
  GuiCustomControl = 50,
  GuiDialogShell = 125,
  GuiDockShell = 126,
  GuiFrameWindow = 20,
  GuiGOSShell = 123,
  GuiLabel = 30,
  GuiListContainer = 73,
  GuiMainWindow = 21,
  GuiMenu = 110,
  GuiMenubar = 111,
  GuiMessageWindow = 23,
  GuiModalWindow = 22,
  GuiOkCodeField = 35,
  GuiPasswordField = 33,
  GuiRadioButton = 41,
  GuiScrollbar = 100,
  GuiScrollContainer = 72,
  GuiSession = 12,
  GuiSessionInfo = 121,
  GuiShell = 122,
  GuiSimpleContainer = 71,
  GuiSplitterContainer = 75,
  GuiSplitterShell = 124,
  GuiStatusbar = 103,
  GuiStatusPane = 43,
  GuiTab = 91,
  GuiTableColumn = 81,
  GuiTableControl = 80,
  GuiTableRow = 82,
  GuiTabStrip = 90,
  GuiTextField = 31,
  GuiTitlebar = 102,
  GuiToolbar = 101,
  GuiUnknown = -1,
  GuiUserArea = 74,
  GuiVComponent = 1,
  GuiVContainer = 2,
  GuiVHViewSwitch = 129,
}
```

## GuiErrorType
```ts
enum GuiErrorType {
  Gui_Err_AccessDenied = 633, // Access denied.
  Gui_Err_Bad_Focus = 634, // Can not set focus to this object.
  Gui_Err_Bad_Index_Type = 618, // Bad index type for collection access.
  Gui_Err_Control_Label = 615, // The control could not be found by label.
  Gui_Err_Control_Name = 608, // The control could not be found by name.
  Gui_Err_Control_Position = 616, // The control could not be found by position.
  Gui_Err_Disconnected = 621, // The object invoked has disconnected from its clients.
  Gui_Err_Enumerator_Index = 614, // The enumerator of the collection cannot find an element with the s pecified
  Gui_Err_Enumerator_Reset = 612, // The enumerator of the collection cannot be reset.
  Gui_Err_FindById = 619, // The control could not be found by id.
  Gui_Err_FindByName = 620, // The control could not be found by name.
  Gui_Err_FindByPos = 632, // The control could not be found by position.
  Gui_Err_Front_Module = 602, // The path of the 'sapfront.dll' could not be determined.
  Gui_Err_Init = 601, // Sapgui engine cannot be initialized.
  Gui_Err_Int_Get_Ses = 629, // Can not get session from TLS
  Gui_Err_Int_GetCtrl_Failed = 625, // Could not get ctrl (Internal Error)
  Gui_Err_Int_GetFocusMan = 627, // Could not get focus manager from session (Internal Error)
  Gui_Err_Int_GetSes = 626, // Could not get session from ctrl (Internal Error)
  Gui_Err_Int_Invalid_Test = 628, // Invalid test tool mode
  Gui_Err_Int_View_Not_Set = 630, // View not set (Internal Error)
  Gui_Err_Invalid_Argument = 613, // The method got an invalid argument.
  Gui_Err_Invalid_Context = 603, // Function called in invalid thread context
  Gui_Err_Invalid_Window = 611, // The required window is invalid.
  Gui_Err_Logon_Module = 604, // The 'Sapgui Logon Component' could not be instantiated.
  Gui_Err_Menu_Disabled = 623, // The menu item is disabled.
  Gui_Err_No_Memory = 607, // The system is out of memory.
  Gui_Err_No_Wrapper = 622, // No wrapper available for this control.
  Gui_Err_Not_Implemented = 610, // The method or property is currently not implemented.
  Gui_Err_Permission_Denied = 637, // Permission denied.
  Gui_Err_Property_Readonly = 609, // The property is readonly.
  Gui_Err_Resize_Failed = 631, // Resize failed.
  Gui_Err_Sapgui_Module = 605, // The 'Sapgui Component' could not be instantiated.
  Gui_Err_Save_Image = 635, // Error saving image.
  Gui_Err_Scripting_Disa = 624, // Scripting is disabled by the server.
  Gui_Err_Session_Index = 606, // The session index is out of range.
  Gui_Err_Shortcut_Evalua = 636, // Shortcut evaluation failed.
  Gui_Err_SL_No_Entry = 1000, // Not a valid SAPLogon entry
  Gui_Err_VKey_Disabled = 617, // The virtual key is not enabled.
}
```

## GuiEventType
```ts
enum GuiEventType {
  SapApplicationCreateSessionEvent = 2002, // ApplicationCreateSession
  SapApplicationDestroySessionEvent = 2003, // ApplicationDestroySession
  SapApplicationErrorEvent = 2004, // ApplicationErrorEvent
  SapApplicationIgnoreSessionEvent = 2005, // ApplicationIgnoreSession
  SapContextMenuEvent = 1282, // ContextMenu
  SapCustomDataChangedEvent = 1280, // CustomDataChanged
  SapDefaultEvent = 0, // Default
  SapHitSelectEvent = 1281, // Hit
  SapSessionAbapScriptingEvent = 1289, // AbapScriptingEvent
  SapSessionActivatedEvent = 1285, // SessionActivated
  SapSessionAutoFCodeEvent = 1284, // SessionAutoFCode
  SapSessionDestroyEvent = 1283, // SessionDestroy
  SapSessionEndRequestEvent = 515, // SessionEndRequest
  SapSessionErrorEvent = 516, // SessionError
  SapSessionFocusChangedEvent = 1286, // SessionFocusChanged
  SapSessionHistoryOpenedEvent = 1287, // SessionHistoryOpened
  SapSessionProgressIndicatorEvent = 1288, // ProgressIndicatorOpened
  SapSessionStartRequestEvent = 514, // SessionStartRequest
}
```

## GuiImageType
```ts
enum GuiImageType {
  BMP = 0,
  GIF = 2,
  JPEG = 1,
  PNG = 2,
}
```

## GuiMagicDispIDs
```ts
enum GuiMagicDispIDs {
  GuiDispIDBTPress = 32200,
  GuiDispIDCBChecked = 32011,
  GuiDispIDCBCurListBoxEntry = 32305,
  GuiDispIDCBEntries = 32302,
  GuiDispIDCBEntryKey = 33800,
  GuiDispIDCBEntryPos = 33802,
  GuiDispIDCBEntryValue = 33801,
  GuiDispIDCBIsListBoxActive = 32304,
  GuiDispIDCBKey = 32300,
  GuiDispIDCBKeySpace = 32303,
  GuiDispIDCBShowKey = 32306,
  GuiDispIDCBValue = 32301,
  GuiDispIDCollAdd = 33103,
  GuiDispIDCollCount = 33100,
  GuiDispIDCollElAt = 33102,
  GuiDispIDCollLength = 33101,
  GuiDispIDConConnString = 33003,
  GuiDispIDConDescription = 33002,
  GuiDispIDConDisabled = 33001,
  GuiDispIDConnClose = 32831,
  GuiDispIDConSessions = 33000,
  GuiDispIDCTFindAllByName = 32035,
  GuiDispIDCTFindAllByNameEx = 32036,
  GuiDispIDCTFindById = 32029,
  GuiDispIDCTFindByLabel = 32027,
  GuiDispIDCTFindByName = 32026,
  GuiDispIDCTFindByNameEx = 32034,
  GuiDispIDCTFindByPosition = 32028,
  GuiDispIDDockerIsVertical = 34301,
  GuiDispIDDockerPixelSize = 34300,
  GuiDispIDEngAddHist = 32913,
  GuiDispIDEngButtonB = 32903,
  GuiDispIDEngCon = 32900,
  GuiDispIDEngConnErr = 32924,
  GuiDispIDEngCrColl = 32911,
  GuiDispIDEngDropHist = 32914,
  GuiDispIDEngGetEng = 1,
  GuiDispIDEngHistEnabled = 32916,
  GuiDispIDEngIgnore = 32908,
  GuiDispIDEngInplace = 32907,
  GuiDispIDEngMajor = 32909,
  GuiDispIDEngMinor = 32910,
  GuiDispIDEngNoSysMsg = 32925,
  GuiDispIDEngOpenCon = 32905,
  GuiDispIDEngOpenConEx = 32918,
  GuiDispIDEngOpenWDCon = 32926,
  GuiDispIDEngPatchlevel = 32919,
  GuiDispIDEngQuit = 32906,
  GuiDispIDEngRegister = 32921,
  GuiDispIDEngRevision = 32920,
  GuiDispIDEngRevoke = 32923,
  GuiDispIDEngStatusB = 32902,
  GuiDispIDEngTheme = 32912,
  GuiDispIDEngTitleB = 32904,
  GuiDispIDEngToolB = 32901,
  GuiDispIDEngUtils = 32917,
  GuiDispIDEngWDSessions = 32927,
  GuiDispIDErrDesc1 = 33601,
  GuiDispIDErrDesc2 = 33602,
  GuiDispIDErrDesc3 = 33603,
  GuiDispIDErrDesc4 = 33604,
  GuiDispIDErrNo = 33600,
  GuiDispIDGActiveSession = 32049,
  GuiDispIDGActiveSession2 = 32075,
  GuiDispIDGCAccDescription = 33703,
  GuiDispIDGCAccLabelCol = 32043,
  GuiDispIDGCAccText = 32044,
  GuiDispIDGCAccTextOnReq = 32045,
  GuiDispIDGCAccTooltip = 32042,
  GuiDispIDGCChangeable = 32009,
  GuiDispIDGCCharHeight = 32073,
  GuiDispIDGCCharLeft = 32070,
  GuiDispIDGCCharTop = 32071,
  GuiDispIDGCCharWidth = 32072,
  GuiDispIDGCChildren = 32019,
  GuiDispIDGCClass = 32017,
  GuiDispIDGCColorIndex = 32058,
  GuiDispIDGCColorIntensified = 32059,
  GuiDispIDGCColorInverse = 32060,
  GuiDispIDGCCtxMnu = 33701,
  GuiDispIDGCDefaultTooltip = 32069,
  GuiDispIDGCDisplayedText = 32074,
  GuiDispIDGCDragDrop = 33706,
  GuiDispIDGCDumpState = 31194,
  GuiDispIDGCFlushing = 33704,
  GuiDispIDGCHeight = 32006,
  GuiDispIDGCHwnd = 33702,
  GuiDispIDGCIcon = 32037,
  GuiDispIDGCId = 32025,
  GuiDispIDGCIsContainer = 32033,
  GuiDispIDGCIsHotspot = 32051,
  GuiDispIDGCIsList = 32052,
  GuiDispIDGCIsStepLoop = 32062,
  GuiDispIDGCIsSymbolFont = 32061,
  GuiDispIDGCLeft = 32003,
  GuiDispIDGCLeftLabel = 32040,
  GuiDispIDGCLoopCurrentCol = 32065,
  GuiDispIDGCLoopCurrentRow = 32066,
  GuiDispIDGCLoopHeight = 32064,
  GuiDispIDGCLoopWidth = 32063,
  GuiDispIDGCModified = 32030,
  GuiDispIDGCName = 32001,
  GuiDispIDGCOcxEvents = 33705,
  GuiDispIDGCParent = 32038,
  GuiDispIDGCParentFrame = 32050,
  GuiDispIDGCRightLabel = 32041,
  GuiDispIDGCRowText = 32053,
  GuiDispIDGCScreenLeft = 32046,
  GuiDispIDGCScreenTop = 32047,
  GuiDispIDGCSession = 32018,
  GuiDispIDGCSetFocus = 32024,
  GuiDispIDGCShortId = 32031,
  GuiDispIDGCShowContextMenu = 32068,
  GuiDispIDGCSubType = 33700,
  GuiDispIDGCText = 32000,
  GuiDispIDGCTitle = 32048,
  GuiDispIDGCTooltip = 32008,
  GuiDispIDGCTop = 32004,
  GuiDispIDGCType = 32015,
  GuiDispIDGCTypeAsNum = 32032,
  GuiDispIDGCVisualize = 32039,
  GuiDispIDGCWidth = 32005,
  GuiDispIDGECATTReplay = 32076,
  GuiDispIDGetAbsoluteRow = 33407,
  GuiDispIDGMSelect = 33300,
  GuiDispIDGMWFocusedButton = 32433,
  GuiDispIDGMWHelpButtonHelpText = 32440,
  GuiDispIDGMWHelpButtonText = 32435,
  GuiDispIDGMWMessageText = 32437,
  GuiDispIDGMWMessageType = 32436,
  GuiDispIDGMWOKButtonHelpText = 32439,
  GuiDispIDGMWOKButtonText = 32434,
  GuiDispIDGMWVisible = 32438,
  GuiDispIDGUCharHeight = 32603,
  GuiDispIDGUCharWidth = 32602,
  GuiDispIDGUHorizontalScrollbar = 32600,
  GuiDispIDGUListNav = 32605,
  GuiDispIDGUOTFPreview = 32606,
  GuiDispIDGUResize = 32604,
  GuiDispIDGUVerticalScrollbar = 32601,
  GuiDispIDGWButtonB = 32425,
  GuiDispIDGWClose = 32414,
  GuiDispIDGWCompBitmap = 32443,
  GuiDispIDGWGuiFocus = 32422,
  GuiDispIDGWHandle = 32420,
  GuiDispIDGWHardCopy = 32415,
  GuiDispIDGWHardCopyMem = 32441,
  GuiDispIDGWIconic = 32400,
  GuiDispIDGWIconify = 32408,
  GuiDispIDGWIsPopupDialog = 32427,
  GuiDispIDGWJumpBackward = 32432,
  GuiDispIDGWJumpForward = 32431,
  GuiDispIDGWMaximize = 32410,
  GuiDispIDGWMoveWindow = 32407,
  GuiDispIDGWPopupDialogText = 32428,
  GuiDispIDGWRestore = 32409,
  GuiDispIDGWSpyMode = 32413,
  GuiDispIDGWStatusB = 32424,
  GuiDispIDGWSysFocus = 32421,
  GuiDispIDGWTabBackward = 32430,
  GuiDispIDGWTabForward = 32429,
  GuiDispIDGWTitleB = 32426,
  GuiDispIDGWToolB = 32423,
  GuiDispIDGWVKAllowed = 32412,
  GuiDispIDGWWPHeight = 32417,
  GuiDispIDGWWPMsgBox = 32419,
  GuiDispIDGWWPResize = 32418,
  GuiDispIDGWWPResizeEx = 32442,
  GuiDispIDGWWPWidth = 32416,
  GuiDispIDIsListBoxActive = 32840,
  GuiDispIDLCursor = 32022,
  GuiDispIDLHighlighted = 32100,
  GuiDispIDLIsLeftLabel = 32101,
  GuiDispIDLIsRightLabel = 32102,
  GuiDispIDListBoxCurrEntry = 32849,
  GuiDispIDListBoxCurrEntryHeight = 32848,
  GuiDispIDListBoxCurrEntryLeft = 32846,
  GuiDispIDListBoxCurrEntryTop = 32845,
  GuiDispIDListBoxCurrEntryWidth = 32847,
  GuiDispIDListBoxHeight = 32844,
  GuiDispIDListBoxLeft = 32842,
  GuiDispIDListBoxTop = 32841,
  GuiDispIDListBoxWidth = 32843,
  GuiDispIDLListProperty = 32103,
  GuiDispIDLMaxLength = 32012,
  GuiDispIDLNumerical = 32013,
  GuiDispIDLPassword = 32016,
  GuiDispIDLSimpleListProperty = 32104,
  GuiDispIDMsgAsPopup = 34004,
  GuiDispIDMsgId = 34001,
  GuiDispIDMsgNumber = 34002,
  GuiDispIDMsgPar = 34003,
  GuiDispIDMsgType = 34000,
  GuiDispIDOcxCallbackChange = 200889,
  GuiDispIDOcxCallbackHighlight = 200890,
  GuiDispIDOcxCallbackHit = 200891,
  GuiDispIDOcxControl = 271062,
  GuiDispIDOcxGetRect = 31192,
  GuiDispIDOcxHit = 31195,
  GuiDispIDOcxHitTest = 31193,
  GuiDispIDOcxHover = 31196,
  GuiDispIDOcxIsReadOnlyCall = 31191,
  GuiDispIDOcxNotify = 31199,
  GuiDispIDOcxNotifyContEvSink = 31197,
  GuiDispIDOcxNotifyCtrlEvent = 31198,
  GuiDispIDOKF1 = 32351,
  GuiDispIDOKOpened = 32350,
  GuiDispIDRBGroupColl = 32504,
  GuiDispIDRBGroupCount = 32502,
  GuiDispIDRBGroupPos = 32503,
  GuiDispIDRBSelect = 32501,
  GuiDispIDRBSelected = 32500,
  GuiDispIDSBDblClick = 32750,
  GuiDispIDScrollMax = 33904,
  GuiDispIDScrollMin = 33905,
  GuiDispIDScrollPage = 33903,
  GuiDispIDScrollPos = 33902,
  GuiDispIDScrollRange = 33900,
  GuiDispIDSesActivWin = 32800,
  GuiDispIDSesBusy = 32803,
  GuiDispIDSesClearErrorList = 32825,
  GuiDispIDSesClose = 32811,
  GuiDispIDSesCmd = 32805,
  GuiDispIDSesCmdAsync = 32806,
  GuiDispIDSesCreate = 32812,
  GuiDispIDSesEnableAccSymbols = 32830,
  GuiDispIDSesEnableAccTabChain = 32829,
  GuiDispIDSesEnableJaws = 32828,
  GuiDispIDSesEndT = 32810,
  GuiDispIDSesErrorList = 32824,
  GuiDispIDSesFindByPos = 32818,
  GuiDispIDSesIconDesc = 33525,
  GuiDispIDSesInfo = 32802,
  GuiDispIDSesInfoAppSr = 33508,
  GuiDispIDSesInfoCl = 33509,
  GuiDispIDSesInfoCP = 33512,
  GuiDispIDSesInfoDisRec = 33521,
  GuiDispIDSesInfoDynp = 33506,
  GuiDispIDSesInfoFlush = 33503,
  GuiDispIDSesInfoForceNot = 33522,
  GuiDispIDSesInfoGrpN = 33515,
  GuiDispIDSesInfoGuiCP = 33523,
  GuiDispIDSesInfoI18N = 33524,
  GuiDispIDSesInfoITime = 33501,
  GuiDispIDSesInfoLang = 33511,
  GuiDispIDSesInfoModeNo = 33517,
  GuiDispIDSesInfoMsgSrc = 33514,
  GuiDispIDSesInfoMsgSrv = 33513,
  GuiDispIDSesInfoProg = 33505,
  GuiDispIDSesInfoReadOnly = 33520,
  GuiDispIDSesInfoRound = 33504,
  GuiDispIDSesInfoRTime = 33500,
  GuiDispIDSesInfoSesCtx = 33518,
  GuiDispIDSesInfoSysN = 33507,
  GuiDispIDSesInfoSysNo = 33516,
  GuiDispIDSesInfoTrans = 33502,
  GuiDispIDSesInfoUser = 33510,
  GuiDispIDSesInfoWAN = 33519,
  GuiDispIDSesIsActive = 32819,
  GuiDispIDSesLockSessionUI = 32826,
  GuiDispIDSesMenu = 32807,
  GuiDispIDSesPPPSyId = 32821,
  GuiDispIDSesPPSyId = 32822,
  GuiDispIDSesPPTaId = 32820,
  GuiDispIDSesProgressPercent = 32832,
  GuiDispIDSesProgressText = 32833,
  GuiDispIDSesRecFile = 32814,
  GuiDispIDSesRecord = 32804,
  GuiDispIDSesRunScrCtrl = 32816,
  GuiDispIDSesSaveAsUnicode = 32823,
  GuiDispIDSesShowKeys = 33527,
  GuiDispIDSesStartT = 32809,
  GuiDispIDSesStdNumFmt = 33526,
  GuiDispIDSesSuppressBackendPopups = 32834,
  GuiDispIDSesTestTool = 32813,
  GuiDispIDSesUnlockSessionUI = 32827,
  GuiDispIDSesVKey = 32808,
  GuiDispIDSesVKeyDesc = 32817,
  GuiDispIDSHSelCtxtMenIt = 34100,
  GuiDispIDSHSelCtxtMenItPos = 34102,
  GuiDispIDSHSelCtxtMenItTxt = 34101,
  GuiDispIDSplitterIsVertical = 34400,
  GuiDispIDSplitterSashPosition = 34401,
  GuiDispIDTableBase = 33400,
  GuiDispIDTableColBase = 33420,
  GuiDispIDTableColFixed = 33421,
  GuiDispIDTableColSelected = 33422,
  GuiDispIDTableColSelectMode = 33401,
  GuiDispIDTableColTitle = 33420,
  GuiDispIDTableColumns = 33402,
  GuiDispIDTableConfigureLayout = 33406,
  GuiDispIDTableCurrentCol = 33410,
  GuiDispIDTableCurrentRow = 33411,
  GuiDispIDTableDeselAllCols = 33414,
  GuiDispIDTableFieldName = 33409,
  GuiDispIDTabLeftTab = 33200,
  GuiDispIDTableGetCell = 33415,
  GuiDispIDTableReorderTable = 33405,
  GuiDispIDTableRowBase = 33430,
  GuiDispIDTableRowCount = 33412,
  GuiDispIDTableRows = 33404,
  GuiDispIDTableRowSelectable = 33431,
  GuiDispIDTableRowSelected = 33430,
  GuiDispIDTableRowSelectMode = 33403,
  GuiDispIDTableSelAllCols = 33408,
  GuiDispIDTableVisRowCount = 33413,
  GuiDispIDTabSelTab = 33201,
  GuiDispIDTBSelect = 32700,
  GuiDispIDTBToLeft = 32701,
  GuiDispIDTHistoryCurEntry = 32057,
  GuiDispIDTHistoryCurIndex = 32056,
  GuiDispIDTHistoryIsActive = 32054,
  GuiDispIDTHistoryList = 32055,
  GuiDispIDTIsOField = 32067,
  GuiDispIDTRequired = 32014,
  GuiDispIDUtilCloseFile = 34202,
  GuiDispIDUtilMsgBox = 34200,
  GuiDispIDUtilMsgOptOK = 34220,
  GuiDispIDUtilMsgOptOKCan = 34222,
  GuiDispIDUtilMsgOptYesNo = 34221,
  GuiDispIDUtilMsgResCancel = 34230,
  GuiDispIDUtilMsgResNo = 34233,
  GuiDispIDUtilMsgResOK = 34231,
  GuiDispIDUtilMsgResYes = 34232,
  GuiDispIDUtilMsgTypeE = 34208,
  GuiDispIDUtilMsgTypeI = 34205,
  GuiDispIDUtilMsgTypeP = 34209,
  GuiDispIDUtilMsgTypeQ = 34206,
  GuiDispIDUtilMsgTypeW = 34207,
  GuiDispIDUtilOpenFile = 34201,
  GuiDispIDUtilWriteFile = 34203,
  GuiDispIDUtilWriteLnFile = 34204,
}
```

## GuiMessageBoxOption
```ts
enum GuiMessageBoxOption {
  MSG_OPTION_OK = 0, // Constant value to be used when calling
  MSG_OPTION_OKCANCEL = 2, // Constant value to be used when calling
  MSG_OPTION_YESNO = 1, // Constant value to be used when calling
}
```

## GuiMessageBoxResult
```ts
enum GuiMessageBoxResult {
  MSG_RESULT_CANCEL = 0, // Constant value to be used as a return
  MSG_RESULT_NO = 3, // Constant value to be used as a return
  MSG_RESULT_OK = 1, // Constant value to be used as a return
  MSG_RESULT_YES = 2, // Constant value to be used as a return
}
```

## GuiMessageBoxType
```ts
enum GuiMessageBoxType {
  MSG_TYPE_ERROR = 3, // Constant value to be used when calling
  MSG_TYPE_INFORMATION = 0, // Constant value to be used when calling
  MSG_TYPE_PLAIN = 4, // Constant value to be used when calling
  MSG_TYPE_QUESTION = 1, // Constant value to be used when calling
  MSG_TYPE_WARNING = 2, // Constant value to be used when calling
}
```

## GuiScrollbarType
```ts
enum GuiScrollbarType {
  GuiScrollbarTypeHorizontal = 2,
  GuiScrollbarTypeUnknown = 0,
  GuiScrollbarTypeVertical = 1,
}
```

## GuiTableSelectionType
```ts
enum GuiTableSelectionType {
  MULTIPLE_INTERVAL_SELECTION = 2, // Several columns/rows can be selected.
  NO_SELECTION = 0, // No selection possible. (0)
  SINGLE_SELECTION = 1, // One column/row can be selected. (1)
}
```
