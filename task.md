# Agent Handover — AHK2 SAP GUI Scripting Wrapper/Proxy

Repo: `anonymusdennis/ahk2_sap_wrapper`  
Primary sources in repo:
- `sap_gui_scripting_api_760_condensed_index.md` (condensed complete member index)
- `sap_gui_scripting_api_761.txt` (detailed SAP GUI Scripting API extract)

## Goal

Build an AutoHotkey v2 wrapper library for SAP GUI Scripting COM objects that:

1. Wraps *all* COM method calls and property gets/sets with hooks:
   - `On_Call()` runs before the COM operation
   - `After_Call()` runs after success
   - `On_Error()` runs on failure (popup recovery will be implemented later by repo owner)

2. Supports **chained / stacked calls** as used in recorded SAP scripts, e.g.:
   - `session.FindById("wnd[2]/usr").Children[i].Text`
   - `session.FindById("wnd[0]").SendVKey(0)`
   The key requirement is that **intermediate returned COM objects** are automatically wrapped, so the hook pipeline continues across the chain.

3. Provides **autocomplete-friendly typed wrapper classes**:
   - Wrapper methods/properties must use the **same names as SAP GUI members** (case-sensitive as used in scripting, e.g. `FindById`, `SendVKey`, `ActiveWindow`, `Info`, `Children`, `Text`).
   - Use allowlists generated from the documentation member list (the condensed index is the best input for this).

4. Uses allowlists to block invalid members per type (optional strict mode):
   - Example: `GuiSession` allowlist includes `Busy`, `Info`, `StartTransaction`, etc.
   - Accessing a non-allowlisted member throws a wrapper error (custom error message).

---

## HARD CONSTRAINTS (AHK2 engine limitations)

The owner’s AHK2 build does NOT support:

- Arrow functions / lambdas such as: `() => expr`
- Exception variables in catch:
  - Allowed: `try { ... } catch { ... }`
  - Not allowed: `catch err`, `catch as e`, `catch e`, etc.

So:
- Your code must not rely on `e.Message`.
- If you need error info, capture it using contextual data you already know (type/member/path/op), and optionally use `A_LastError` where meaningful (COM errors may not map cleanly).

---

## Architecture Overview

### Core design principle
AHK can intercept member access on *AHK objects* via meta-methods, but not directly on raw COM objects. Therefore:

- All SAP COM objects must be wrapped in an AHK proxy object.
- Every proxy operation (get/set/call) must route through a single central `Invoke()` pipeline.
- Any operation that returns a COM object must wrap that return into another proxy so chaining still works.

### Required building blocks

#### 1) Hook policy class (user extensible)
File: `src/core/SapHookPolicy.ahk`

Methods:
- `On_Call(op, typeName, member, path, args)`
- `After_Call(op, typeName, member, path, result)`
- `On_Error(op, typeName, member, path, args)`

Default `On_Error` should show a simple `MsgBox` placeholder (owner will replace with popup handling).

#### 2) Central invoker + allowlist enforcement
File: `src/core/SapInvoke.ahk` (or inside proxy base)

Responsibilities:
- Validate `member` is allowlisted for `typeName` (unless strict mode disabled)
- Perform the operation inside `try/catch`
- Call hook policy methods accordingly
- Wrap returned COM objects or collections into proxies

NOTE: No exception variable available. `catch { ... }` must still call `On_Error()` and then rethrow via `throw Error("...context...")` or `throw`.

#### 3) Generic COM proxy base (dynamic chaining support)
File: `src/core/SapComProxy.ahk`

State:
- `_com` raw COM reference
- `_path` string (e.g. `"/app/con[0]/ses[0]/wnd[0]/usr/txtFOO"`)
- `_typeName` (e.g. `"GuiSession"`)
- `_allow` (Map/Set of member names)
- `_policy` (SapHookPolicy instance)
- optional `_strict` boolean

Must implement:
- `__Get(name, params := unset)` => property get via invoker
- `__Set(name, value, params := unset)` => property set via invoker
- `__Call(name, params*)` => method call via invoker
- `Raw()` => returns `_com` escape hatch

Critical: Wrap return values:
- If return is COM object -> return new proxy for that object (type determined via `.Type` property if possible)
- If return is a collection (GuiCollection/GuiComponentCollection) -> return collection proxy that supports `[i]` indexing

Type detection:
- Many SAP objects provide `.Type` (string) and `.TypeAsNumber` (long) per docs.
- Best practice: attempt to read `.Type` safely; on failure default to `"GuiUnknown"`.

#### 4) Collection proxy with indexing
File: `src/core/SapCollectionProxy.ahk`

Needed because chaining often includes:
- `.Children[i]`
- `.Sessions[0]`, `.Connections[0]`

SAP docs show collections:
- `GuiComponentCollection` has `.Item(Index)` and `.ElementAt(Index)`
- `GuiCollection` has `.Item(Index)` and `.ElementAt(Index)`
Properties: `Count`, `Length`, `Type`, `TypeAsNumber`

Implement:
- `__Item[index]` getter mapping to COM `Item(index)` (fallback `ElementAt(index)` if needed)
- `Count`, `Length` properties pass-through (with hooks)

Return values from `.Item()` must be wrapped.

#### 5) Typed wrappers for autocomplete (generated + thin)
Directory:
- `src/types/GuiSession.ahk`
- `src/types/GuiApplication.ahk`
- `src/types/GuiConnection.ahk`
- `src/types/GuiFrameWindow.ahk`
- `src/types/GuiVComponent.ahk`
- etc.

These classes can either:
- (A) extend `SapComProxy` and just set allowlist + type name
- or (B) contain an internal `SapComProxy` and forward

Preferred: extend `SapComProxy` for simpler chaining.

Each typed wrapper should:
- Provide explicit methods/properties for key members to improve autocomplete.
- Still allow dynamic access via inherited `__Get/__Call` for completeness.

---

## Allowlists / Member registry

Create generated files:
- `src/generated/Allowlists.ahk`
  - Map typeName => Set/Map of allowed member names

Input: `sap_gui_scripting_api_760_condensed_index.md`
- Parse sections `interface GuiSession ...`
- Collect method names and property names
- Also include inherited members if you want strict allowlists to be accurate:
  - ex: `GuiSession extends GuiContainer` => add `GuiContainer` allowlist
  - For first iteration, it’s acceptable to:
    - include only declared members + the known common base (`GuiComponent`, `GuiContainer`, `GuiVComponent`, `GuiVContainer`)
    - or run in non-strict mode by default

Strict mode suggestion:
- default strict = false (log blocked member rather than fail)
- ability to enable strict per proxy instance

---

## Path handling

Every proxy should carry a human-readable `_path` used in diagnostics and error messages:
- Root objects:
  - `GuiApplication` path `"GuiApplication"` or `"/app"`
  - `GuiConnection` path `"/app/con[0]"` etc.
  - `GuiSession` path `"/app/con[0]/ses[0]"`

`FindById` should produce a child proxy whose path becomes:
- `${parent._path}/${id}` (normalize slashes as desired)

This is vital for debugging failures like `Gui_Err_FindById (619)`.

---

## Important SAP behavior to account for (from docs)

1) **Round trips invalidate references below window level**
Docs explicitly warn for `GuiButton.Press`: after server communication, stored references to elements below window become invalid.
Implication:
- Wrapper should optionally support a “stale object” strategy later (not required now), but at least the error messages should point to this.

2) `GuiSession.Busy`
Docs: while busy, scripting calls block. Wrapper can optionally `WaitNotBusy` in `On_Call` for stability (feature flag).

---

## Public API / usage examples to provide

Create `examples/demo.ahk`:

- Attach to ROT:
  - `SapGuiAuto := ComObjGet("SAPGUI")`
  - `appCom := SapGuiAuto.GetScriptingEngine`
- Wrap:
  - `app := GuiApplicationProxy(appCom, policy)`
  - `con := app.Children[0]` (or `app.Connections[0]`)
  - `ses := con.Children[0]`
  - `ses.FindById("wnd[0]/tbar[0]/okcd").Text := "/nSE16"`
  - `ses.FindById("wnd[0]").SendVKey(0)`

Also include an example of `session.Info.SystemName` access.

---

## Minimal file layout proposal

- `src/`
  - `core/`
    - `SapHookPolicy.ahk`
    - `SapComProxy.ahk`
    - `SapCollectionProxy.ahk`
    - `SapTypeRegistry.ahk` (type detection + allowlist lookup)
  - `generated/`
    - `Allowlists.ahk`
    - `TypeNumbers.ahk` (optional map from TypeAsNumber to TypeName using enum GuiComponentType)
  - `types/`
    - `GuiApplication.ahk`
    - `GuiConnection.ahk`
    - `GuiSession.ahk`
    - `GuiFrameWindow.ahk`
    - `GuiComponent.ahk`
    - `GuiContainer.ahk`
    - `GuiVComponent.ahk`
    - `GuiVContainer.ahk`
    - `GuiCollection.ahk`
    - `GuiComponentCollection.ahk`
- `examples/`
  - `demo_rot_attach.ahk`
- `README.md`

---

## Implementation notes for AHK2 meta methods

- Implement `__Get/__Set/__Call` on proxy classes.
- Avoid lambdas; use named helper methods or `ObjBindMethod(this, "MethodName", fixedArgs...)` to defer execution if needed.
- Since we can’t capture exception objects, rely on:
  - contextual strings (typeName, member, args, path, operation)
  - optionally `A_LastError` (may be 0 for COM)
- For deferred execution, structure Invoke() like:
  - `InvokeCall(member, args*)` does the `try { return this._com.%member%(args*) } catch { ... }`
  - `InvokeGet(member)` uses `try { v := this._com.%member% } catch { ... }`
  - `InvokeSet(member, value)` uses `try { this._com.%member% := value } catch { ... }`

---

## Acceptance criteria

- Chained access works with hooks applied at each step:
  - `ses.FindById("wnd[0]").SendVKey(0)` triggers `On_Call` + `After_Call`
  - `ses.FindById("wnd[0]/usr").Children[0].Text` triggers hooks across FindById, Children, Item/index, Text get
- Collections support `[i]` indexing.
- Typed wrapper classes exist at least for: GuiApplication, GuiConnection, GuiSession, GuiFrameWindow, GuiVComponent, GuiCollection, GuiComponentCollection.
- No usage of arrow functions/lambdas; no `catch e`.
- A basic demo runs without syntax errors.

---

## Owner TODO hooks

- Replace `On_Error` MsgBox placeholder with:
  - random popup detection/closing
  - optional retry logic for transient failures


for the first step please wait for the files i want to give to you as a help 
th first is the official spa dcoumentation and the second a pseudocode implemantation in js-like code please note that that is in no way ahk2 code 
please wait
