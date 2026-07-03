#!/usr/bin/env python3
"""Generate AHK v2 source files from the SAP GUI Scripting API condensed index.

Parses ``sap_gui_scripting_api_760_condensed_index.md`` and emits:

- ``src/generated/Allowlists.ahk``    -- per-type member allowlists + inheritance
- ``src/generated/TypeNumbers.ahk``   -- enum constants + TypeAsNumber -> type name map
- ``src/generated/TypedWrappers.ahk`` -- typed proxy classes for autocomplete

Run from the repository root (or anywhere):

    python scripts/generate_allowlists.py

Notes on the source document:
- Interface sections look like ``## GuiButton (extends GuiVComponent)``.
  Some headings carry a doc typo ``(extends theGuiVContainer)`` -- the
  leading ``the`` is stripped.
- Some sections have no ``(extends ...)`` in the heading but list bases in
  ``// Inherits members from:`` comments inside the ``ts`` block. Both
  sources are merged.
- ``GuiTextField`` has no section of its own in the condensed index even
  though ``GuiCTextField``/``GuiPasswordField`` extend it. Its members are
  synthesized below from section 1.2.64 of the full API guide
  (sap_gui_scripting_api_761).
"""

from __future__ import annotations

import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_MD = REPO_ROOT / "sap_gui_scripting_api_760_condensed_index.md"
OUT_DIR = REPO_ROOT / "src" / "generated"

IDENT_RE = re.compile(r"^[A-Za-z_]\w*$")
HEADING_RE = re.compile(r"^##\s+(Gui\w+)(?:\s+\(extends\s+(?:the)?(Gui\w+)\))?\s*$")
INTERFACE_RE = re.compile(r"^\s*interface\s+(Gui\w+)(?:\s+extends\s+(?:the)?(Gui\w+))?\s*\{")
ENUM_RE = re.compile(r"^\s*enum\s+(Gui\w+)\s*\{")
INHERIT_COMMENT_RE = re.compile(r"^\s*//\s*-\s*(?:the)?(Gui\w+)\s*$")
METHOD_RE = re.compile(r"^\s*([A-Za-z_]\w*)\s*\(")
PROPERTY_RE = re.compile(r"^\s*(readonly\s+)?([A-Za-z_]\w*)\s*:\s*[^;]*;\s*$")
ENUM_ENTRY_RE = re.compile(r"^\s*([A-Za-z_]\w*)\s*=\s*(-?\d+)\s*,?(?:\s*//.*)?$")

# GuiTextField is referenced as a base type but has no own section in the
# condensed index. Members transcribed from the full API guide, section
# 1.2.64 (GuiTextField, pages 231-234).
SYNTHESIZED_INTERFACES = {
    "GuiTextField": {
        "bases": ["GuiVComponent"],
        "methods": ["GetListProperty", "GetListPropertyNonRec"],
        "properties": {
            "CaretPosition": False,
            "DisplayedText": True,
            "Highlighted": True,
            "HistoryCurEntry": True,
            "HistoryCurIndex": True,
            "HistoryIsActive": True,
            "HistoryList": True,
            "IsHotspot": True,
            "IsLeftLabel": True,
            "IsListElement": True,
            "IsOField": True,
            "IsRightLabel": True,
            "LeftLabel": True,
            "MaxLength": True,
            "Numerical": True,
            "Required": True,
            "RightLabel": True,
        },
    },
}

# Types that already have hand-written wrapper classes under src/types/.
# The generator must not emit classes for these to avoid duplicate class
# definitions when both files are included.
HAND_WRITTEN_TYPES = {
    "GuiApplication",
    "GuiCollection",
    "GuiComponent",
    "GuiComponentCollection",
    "GuiConnection",
    "GuiContainer",
    "GuiFrameWindow",
    "GuiSession",
    "GuiVComponent",
    "GuiVContainer",
}

# Member names that collide with AHK Object/Class built-ins or the proxy
# infrastructure itself. They stay in the allowlists (dynamic access still
# works through __Get/__Call) but no explicit class member is emitted.
MEMBER_NAME_BLOCKLIST = {
    "Base",
    "Clone",
    "DefineProp",
    "DeleteProp",
    "GetOwnPropDesc",
    "HasOwnProp",
    "HasProp",
    "HasMethod",
    "OwnProps",
    "Raw",
    "__Init",
    "__New",
    "__Get",
    "__Set",
    "__Call",
    "__Item",
    "__Enum",
    "__Class",
}


class Interface:
    def __init__(self, name):
        self.name = name
        self.bases = []          # ordered, de-duplicated base type names
        self.methods = []        # ordered method names
        self.properties = {}     # name -> readonly (bool)

    def add_base(self, base):
        if base and base != self.name and base not in self.bases:
            self.bases.append(base)

    def add_method(self, name):
        if name not in self.methods:
            self.methods.append(name)

    def add_property(self, name, readonly):
        if name not in self.properties:
            self.properties[name] = readonly


def parse_document(text):
    interfaces = {}
    enums = {}

    current_heading = None
    heading_base = None
    current_iface = None
    current_enum = None
    in_ts_block = False

    for raw_line in text.splitlines():
        line = raw_line.rstrip()

        heading = HEADING_RE.match(line)
        if heading:
            current_heading = heading.group(1)
            heading_base = heading.group(2)
            current_iface = None
            current_enum = None
            continue

        if line.strip().startswith("```"):
            in_ts_block = not in_ts_block
            if not in_ts_block:
                current_iface = None
                current_enum = None
            continue

        if not in_ts_block:
            continue

        iface_match = INTERFACE_RE.match(line)
        if iface_match:
            name = iface_match.group(1)
            iface = interfaces.setdefault(name, Interface(name))
            iface.add_base(iface_match.group(2))
            if current_heading == name:
                iface.add_base(heading_base)
            current_iface = iface
            current_enum = None
            continue

        enum_match = ENUM_RE.match(line)
        if enum_match:
            current_enum = enums.setdefault(enum_match.group(1), {})
            current_iface = None
            continue

        if current_enum is not None:
            entry = ENUM_ENTRY_RE.match(line)
            if entry:
                current_enum[entry.group(1)] = int(entry.group(2))
            continue

        if current_iface is None:
            continue

        inherit = INHERIT_COMMENT_RE.match(line)
        if inherit:
            current_iface.add_base(inherit.group(1))
            continue

        if line.strip().startswith("//") or line.strip() in ("{", "}"):
            continue

        prop = PROPERTY_RE.match(line)
        if prop and "(" not in line.split(":", 1)[0]:
            current_iface.add_property(prop.group(2), bool(prop.group(1)))
            continue

        method = METHOD_RE.match(line)
        if method:
            current_iface.add_method(method.group(1))
            continue

    return interfaces, enums


def add_synthesized(interfaces):
    for name, spec in SYNTHESIZED_INTERFACES.items():
        iface = interfaces.setdefault(name, Interface(name))
        for base in spec["bases"]:
            iface.add_base(base)
        for method in spec["methods"]:
            iface.add_method(method)
        for prop, readonly in spec["properties"].items():
            iface.add_property(prop, readonly)


def collection_types(interfaces):
    """Types that (transitively) extend GuiComponentCollection or GuiCollection."""
    result = set()

    def is_collection(name, seen=None):
        if name in ("GuiCollection", "GuiComponentCollection"):
            return True
        if seen is None:
            seen = set()
        if name in seen or name not in interfaces:
            return False
        seen.add(name)
        return any(is_collection(base, seen) for base in interfaces[name].bases)

    for name in interfaces:
        if is_collection(name):
            result.add(name)
    return result


def flatten_members(interfaces, name, seen=None):
    """All members of a type including inherited ones: (methods, properties)."""
    if seen is None:
        seen = set()
    if name in seen or name not in interfaces:
        return [], {}
    seen.add(name)
    iface = interfaces[name]
    methods = list(iface.methods)
    properties = dict(iface.properties)
    for base in iface.bases:
        base_methods, base_properties = flatten_members(interfaces, base, seen)
        for m in base_methods:
            if m not in methods:
                methods.append(m)
        for p, readonly in base_properties.items():
            if p not in properties:
                properties[p] = readonly
    return methods, properties


HEADER = "; AUTO-GENERATED by scripts/generate_allowlists.py -- do not edit by hand.\n; Source: sap_gui_scripting_api_760_condensed_index.md\n"


def emit_allowlists(interfaces):
    lines = ["#Requires AutoHotkey v2.0", "", HEADER.rstrip(), "",
             "class SapGeneratedAllowlists {",
             "    static GetAllowlists() {",
             "        allowlists := Map()"]
    for name in sorted(interfaces):
        iface = interfaces[name]
        members = sorted(set(iface.methods) | set(iface.properties), key=str.lower)
        if not members:
            lines.append(f'        allowlists["{name}"] := Map()')
            continue
        lines.append(f'        allowlists["{name}"] := Map(')
        for i, member in enumerate(members):
            comma = "," if i < len(members) - 1 else ""
            lines.append(f'            "{member}", true{comma}')
        lines.append("        )")
    lines.append("        return allowlists")
    lines.append("    }")
    lines.append("")
    lines.append("    ; Every type may inherit from multiple documented bases.")
    lines.append("    static GetBaseTypes() {")
    lines.append("        bases := Map()")
    for name in sorted(interfaces):
        joined = ", ".join(f'"{b}"' for b in interfaces[name].bases)
        lines.append(f'        bases["{name}"] := [{joined}]')
    lines.append("        return bases")
    lines.append("    }")
    lines.append("")
    lines.append("    ; Kept for backward compatibility: primary (first) base per type.")
    lines.append("    static GetInheritance() {")
    lines.append("        inheritance := Map()")
    for name in sorted(interfaces):
        primary = interfaces[name].bases[0] if interfaces[name].bases else ""
        lines.append(f'        inheritance["{name}"] := "{primary}"')
    lines.append("        return inheritance")
    lines.append("    }")
    lines.append("}")
    lines.append("")
    return "\n".join(lines)


def emit_type_numbers(enums):
    lines = ["#Requires AutoHotkey v2.0", "", HEADER.rstrip(), "",
             "class SapGeneratedTypeNumbers {",
             "    ; Map of GuiComponentType number => type name (TypeAsNumber fallback).",
             "    static GetTypeNumberMap() {",
             "        typeNumbers := Map()"]
    component_type = enums.get("GuiComponentType", {})
    for name, value in sorted(component_type.items(), key=lambda kv: kv[1]):
        lines.append(f'        typeNumbers[{value}] := "{name}"')
    lines.append("        return typeNumbers")
    lines.append("    }")
    lines.append("")
    lines.append("    static GetEnums() {")
    lines.append("        enums := Map()")
    for enum_name in sorted(enums):
        lines.append(f'        enums["{enum_name}"] := Map(')
        entries = sorted(enums[enum_name].items())
        for i, (name, value) in enumerate(entries):
            comma = "," if i < len(entries) - 1 else ""
            lines.append(f'            "{name}", {value}{comma}')
        lines.append("        )")
    lines.append("        return enums")
    lines.append("    }")
    lines.append("}")
    lines.append("")

    # Constant classes for convenient literal access, e.g. GuiEventType.Click.
    for enum_name in sorted(enums):
        if enum_name == "GuiComponentType":
            # Names collide with the wrapper classes (GuiButton etc.), so the
            # component type enum is only exposed through GetEnums()/GetTypeNumberMap().
            continue
        lines.append(f"class {enum_name} {{")
        for name, value in sorted(enums[enum_name].items()):
            lines.append(f"    static {name} := {value}")
        lines.append("}")
        lines.append("")
    return "\n".join(lines)


def emit_typed_wrappers(interfaces, collections):
    lines = ["#Requires AutoHotkey v2.0", "", HEADER.rstrip(), "",
             "; Typed proxy classes for autocomplete over the full SAP GUI",
             "; Scripting object model. Types that already have hand-written",
             "; classes under src/types/ are intentionally not generated here.",
             ""]

    generated = [n for n in sorted(interfaces) if n not in HAND_WRITTEN_TYPES]

    lines.append("class SapGeneratedTypedWrappers {")
    lines.append("    static GetTypeClassMap() {")
    lines.append("        typeClasses := Map()")
    for name in generated:
        lines.append(f'        typeClasses["{name}"] := {name}')
    lines.append("        return typeClasses")
    lines.append("    }")
    lines.append("}")
    lines.append("")

    for name in generated:
        base_class = "SapCollectionProxy" if name in collections else "SapComProxy"
        methods, properties = flatten_members(interfaces, name)
        lines.append(f"class {name} extends {base_class} {{")
        lines.append(f'    __New(comObj, policy := "", strict := false, path := "") {{')
        lines.append(f'        super.__New(comObj, "{name}", path = "" ? "{name}" : path, policy, strict)')
        lines.append("    }")
        emitted = set()
        for prop in sorted(properties, key=str.lower):
            if prop in MEMBER_NAME_BLOCKLIST or prop in emitted:
                continue
            emitted.add(prop)
            readonly = properties[prop]
            lines.append("")
            lines.append(f"    {prop} {{")
            lines.append("        get {")
            lines.append(f'            return this.InvokeGet("{prop}")')
            lines.append("        }")
            if not readonly:
                lines.append("        set {")
                lines.append(f'            return this.InvokeSet("{prop}", value)')
                lines.append("        }")
            lines.append("    }")
        for method in sorted(methods, key=str.lower):
            if method in MEMBER_NAME_BLOCKLIST or method in emitted:
                continue
            emitted.add(method)
            lines.append("")
            lines.append(f"    {method}(args*) {{")
            lines.append(f'        return this.InvokeCall("{method}", args*)')
            lines.append("    }")
        lines.append("}")
        lines.append("")
    return "\n".join(lines)


def main():
    text = SOURCE_MD.read_text(encoding="utf-8")
    interfaces, enums = parse_document(text)
    add_synthesized(interfaces)

    collections = collection_types(interfaces)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "Allowlists.ahk").write_text(emit_allowlists(interfaces), encoding="utf-8", newline="\n")
    (OUT_DIR / "TypeNumbers.ahk").write_text(emit_type_numbers(enums), encoding="utf-8", newline="\n")
    (OUT_DIR / "TypedWrappers.ahk").write_text(emit_typed_wrappers(interfaces, collections), encoding="utf-8", newline="\n")

    print(f"Interfaces: {len(interfaces)} (collections: {len(collections)})")
    print(f"Enums: {len(enums)}")
    missing_bases = [n for n, i in interfaces.items() if not i.bases and n not in ("GuiComponent",)]
    print(f"Types without bases: {sorted(missing_bases)}")


if __name__ == "__main__":
    main()
