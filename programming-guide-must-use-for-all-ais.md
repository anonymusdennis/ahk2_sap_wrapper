# Programming Guide for AI Assistants

This repository is a **long-lived home for AutoHotkey v2 tools that automate SAP GUI Scripting**. It is not limited to one feature. Today it contains the SAP COM wrapper library plus the SM30 bulk-import tool; **future tools should follow the same patterns documented here.**

**Read this file before changing any code in the repo.**

---

## Repository purpose

| Layer | Location | Role |
|-------|----------|------|
| SAP wrapper library | `src/SapWrapper.ahk`, `src/core/`, `src/types/` | Typed, hookable COM proxy for SAP GUI Scripting |
| Shared utilities | `src/Sm30AppPaths.ahk`, `src/Sm30JsonConfig.ahk`, `src/core/SapFileLogger.ahk`, … | Reuse across tools where applicable |
| Runnable tools | `examples/<tool-name>/` | One folder per tool: launcher `.ahk2`, `config/`, `data/`, runtime `logs/` |
| One-off / demo scripts | `examples/*.ahk2` | Small scripts not tied to a full tool folder |
| Reference docs | `task.md`, `sap_gui_scripting_api_*.md` | Wrapper design and SAP API allowlists |

When adding a **new tool**, prefer:

```text
examples/my-new-tool/
  my_new_tool.ahk2          # launcher (#Requires v2.0, #Include ../../src/...)
  config/                   # JSON, INI, or other user-editable settings
  data/                     # sample input files
  logs/                     # created at runtime (gitignored)
src/
  MyNewTool.ahk             # optional shared library class(es)
```

Keep **libraries in `src/`**, **entry points in `examples/`**.

---

## Hard rules (every file in this repo)

These come from `task.md` and real runtime constraints (including **v2.1-alpha** builds):

| Rule | Detail |
|------|--------|
| Runnable scripts | `.ahk2` extension |
| Libraries / includes | `.ahk` under `src/` |
| Header | `#Requires AutoHotkey v2.0` unless there is a documented reason otherwise |
| No lambdas | No `() => expr` or arrow functions |
| No exception variables | `try { } catch { }` only — not `catch err`, `catch as e`, etc. |
| Error messages | Build from context + `A_LastError`; never `e.Message` |
| Minimal diffs | Match existing style; no drive-by refactors |
| Tests | Add only when requested or they cover real failure modes |

**Do not** require AutoHotkey v2.1 stable or `JsonLoad` — many users run **v2.0 or v2.1-alpha**. Use repo utilities (`Sm30JsonConfig`, `FileOpen`) instead.

---

## AHK v2 pitfalls (all tools)

### 1. Class methods share local variable names

**Symptom:** `Error: This local variable has not been assigned a value` on e.g. `gui := Gui(...)`.

**Cause:** Local names are shared across all methods of a class. A parameter `gui` in `_OnResize(gui, ...)` breaks another method that uses local `gui`.

**Fix:** Use distinct names: `excelWin`, `runWin`, `senderGui`, `mainWin`, etc.

```ahk
; BAD
_BuildWindow() { gui := Gui() }
_OnResize(gui, minMax, width, height, *) { }

; GOOD
_BuildWindow() { excelWin := Gui(); this.excelGui := excelWin }
_OnResize(senderGui, minMax, width, height, *) { }
```

---

### 2. Plain objects vs Map

On some builds (including **v2.1-alpha**):

- **`obj[key] := value`** on plain **`Object()`** can fail with *no property named `__Item`* — use **`obj.%key% := value`** for identifier keys (`id`, `index`, `kind`, …).
- **`for key, value in obj`** on plain **`Object()`** can throw *Value not enumerable* — return parser-built objects as-is; only iterate **`Map`** when converting to plain objects.
- **`HasOwnProp()`** on **`Map`** fails — convert to plain **`Object()`** first or use **`.Has(key)`**.
- **`<` / `>` on strings** (e.g. `ch < "0"`) can throw *Expected a Number but got a String* — use **`InStr("0123456789", ch)`** for digit checks in parsers.
- Do not wrap JSON load in a bare `catch { }` that hides the real error.

---

### 3. Other common mistakes

| Mistake | Use instead |
|---------|-------------|
| `obj.Has("x")` on plain objects | `obj.HasOwnProp("x")` |
| `array.Join()` | Manual join or a small helper |
| `FileRead(path, "UTF-8")` on alpha | `FileOpen(path, "r", "UTF-8")` with fallback |
| `#Requires AutoHotkey v2.1` + `JsonLoad` | Built-in `Sm30JsonParser` in `Sm30JsonConfig.ahk` |
| Inline arrow callbacks | `ObjBindMethod(this, "_Handler")` |
| Global `HasOwnProp("x")` | `obj.HasOwnProp("x")` on a specific object |

---

## GUI tools (any future app)

### Layout and paths

Use **`Sm30AppPaths`** (or copy the pattern into a tool-specific `AppPaths` class):

- **`A_ScriptDir`** = folder of the script **or** compiled `.exe`
- Ship **`config/`**, **`data/`**, **`logs/`** next to the launcher/exe
- **`src/`** is compiled into the exe; runtime config is **not**

```text
MyTool.exe          (or my_tool.ahk2)
config/
data/
logs/
```

### Long-running COM work (Excel, SAP, file I/O)

COM calls **block the GUI thread**. Never open Excel or SAP from a button handler without feedback.

1. Show loading state immediately (disable buttons, status text).
2. Defer work: `SetTimer(ObjBindMethod(this, "_ProcessDeferred"), -1)`.
3. Re-enable UI in `finally` or after the deferred method completes.

For Excel: one `Excel.Application` session per operation; set `ScreenUpdating := false`, `EnableEvents := false`.

### Session selection

Do **not** rely on `ActiveSession` for tools that must work while another window is focused.

- List sessions: `Sm30SapSessions.List(policy)` (or same hierarchy walk).
- Attach with `Sm30BulkLoader.FromSession(session, policy)` or raw `GuiSession` from the wrapper.

---

## JSON configuration (any tool)

Pattern used by SM30; reuse for other tools:

| Piece | File |
|-------|------|
| Path helpers | `src/Sm30AppPaths.ahk` — generalize or duplicate per tool |
| Parser | `src/Sm30JsonConfig.ahk` — `LoadFile`, `LoadAllFromDir` |
| Tool catalog | e.g. `Sm30TableCatalog.ahk` loading `config/tables/*.json` |

Rules:

- One JSON file per entity (table, view, job type, …).
- Validate required fields after parse; throw **specific** errors (no silent catch).
- Property names like `"index"` and `"kind"` must use **`obj[key] := value`**, not fragile dynamic `%key%` where avoidable.

Example table config: `examples/sm30/config/tables/pfepruntype.json`.

---

## SAP GUI wrapper (all SAP automation)

### Attach chain

```ahk
#Include src/SapWrapper.ahk

policy := SapHookPolicy()   ; or LoggingSapHookPolicy(logger)
app := GuiApplication(ComObjGet("SAPGUI").GetScriptingEngine, policy)
session := app.Children[0].Children[0]   ; first connection, first session only
```

Hierarchy:

```text
GuiApplication
  └── Children[i]  → GuiConnection
        └── Children[j]  → GuiSession
```

Session label: `session.Info` → `SystemName`, `Client`, `User`, `Transaction`, `SessionNumber`.

### Hooks and logging

- Extend **`SapHookPolicy`**: `On_Call`, `After_Call`, `On_Error`.
- File logging: **`SapFileLogger`** + **`LoggingSapHookPolicy`** in `src/core/SapFileLogger.ahk`.
- Logs go under the tool’s **`logs/`** directory (see `Sm30BulkImportGui` for pattern).

### Stale COM references

After **`SendVKey`**, toolbar presses, or dialog recovery, **re-fetch** controls:

```ahk
control := session.FindById(storedRelativeId)
```

Do not keep table/cell references across SAP round trips.

### Allowlists

Wrapper types enforce SAP member allowlists from `src/generated/Allowlists.ahk`. Use documented SAP names (`FindById`, `SendVKey`, …). See `task.md` and `sap_gui_scripting_api_760_condensed_index.md`.

---

## Tool catalog (current)

| Tool | Launcher | Library | Notes |
|------|----------|---------|-------|
| SAP attach demo | `examples/demo_rot_attach.ahk2` | `SapWrapper.ahk` | Minimal ROT attach |
| SM30 bulk fill (script) | `examples/sm30_bulk_fill.ahk2` | `Sm30BulkLoader.ahk` | Inline rows / CSV |
| SM30 dummy test | `examples/sm30_dummy_fill_test.ahk2` | `Sm30BulkLoader.ahk` | Stress + duplicates |
| **SM30 import GUI** | `examples/sm30/sm30_bulk_import_gui.ahk2` | `Sm30BulkImportGui.ahk` | Excel + JSON config + GUI |

When you add a row to this table, add a short section under **Tool-specific notes** below.

---

## Tool-specific notes: SM30 bulk loader

The most complex tool today; many lessons apply to other table-based SAP automation.

### Cell paths (from Script Recorder only)

```ahk
{ index: 0, kind: "Text", prefix: "ctxt", field: "VIEW-FIELDNAME" }
```

Path format:

```text
{tablePath}/{prefix}/{field}[{columnIndex},{visibleRowIndex}]
```

Kinds: `Text` → `.Text`, `Key` → `.Key`, `Selected` → checkbox.

### Scroll model

| Term | Meaning |
|------|---------|
| `absoluteRow` | Logical row cursor in the loader |
| `scrollPos` | `VerticalScrollbar.Position` (top visible absolute index) |
| `visibleRow` | `absoluteRow - scrollPos` |
| `VisibleRowCount` | On-screen rows (changes on window resize) |
| `RowCount` | **Not** filled row count — includes empty preallocated rows |

Re-read scroll metrics after scroll, Enter, skip recovery, or resize.

### Error recovery (duplicate key SV009)

**Works:**

```text
while status bar shows error (MessageType E or A, non-empty text):
    press skip button (default wnd[0]/tbar[1]/btn[20])
    wait until session not Busy
until no error
press Enter once to continue
```

After page commit with `skipped` errors:

```ahk
dataIndex += rowsOnPage
absoluteRow += rowsOnPage - skipped
```

**Never:**

- Stop skipping because status bar text unchanged (SAP repeats same SV009 text).
- Set `absoluteRow := table.RowCount` after skips.

### Writable-cell fallback

If bottom of planned page is not `Changeable`: count trailing blocked rows upward, scroll up by that many, retry. Do not use `RowCount` as the next write target.

### Throughput

- Default **`SetFillMode("page")`** — column-major, one Enter per page.
- **`SetFillMode("row")`** — debug only.
- **`ResizeWorkingPaneEx`** — default height **2000** in `OpenView()`.

### Key SM30 files

| File | Purpose |
|------|---------|
| `src/Sm30BulkLoader.ahk` | Bulk fill, scroll, recovery |
| `src/Sm30BulkImportGui.ahk` | Import GUI |
| `src/Sm30TableCatalog.ahk` | Loads `config/tables/*.json` |
| `src/Sm30ExcelImport.ahk` | Excel COM reader |
| `src/Sm30SapSessions.ahk` | Session list |
| `examples/sm30/config/tables/*.json` | Table definitions |

---

## Checklist before proposing changes

1. **Which tool?** Identify launcher under `examples/` and libraries in `src/`.
2. **Read logs** in the tool’s `logs/` folder — include SAP status bar, scroll, row counts when debugging table fill.
3. **Preserve repo rules** — no lambdas, no `catch err`, v2.0-compatible APIs.
4. **Paths** — use `A_ScriptDir` / `Sm30AppPaths` pattern; config beside exe when compiled.
5. **GUI** — defer blocking COM; unique local variable names in classes.
6. **JSON** — `Object()` + bracket assign + plain-object conversion; clear error messages.
7. **SAP** — refresh COM refs; don’t trust `RowCount` as “next empty row”.
8. **Scope** — smallest correct fix; update this guide if you discover a new cross-tool pitfall.

---

## Document maintenance

When a fix or pattern applies **beyond one tool**, add it under:

- **Hard rules**, **AHK v2 pitfalls**, **GUI tools**, **JSON configuration**, or **SAP GUI wrapper**

When a fix applies **only to one tool**, add it under **Tool-specific notes** and one line in the **Tool catalog** table.

Remove or shorten **session-specific chronologies** over time; keep durable rules, not narrative history.
