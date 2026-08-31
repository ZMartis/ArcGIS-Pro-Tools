"""Exercise the field-resolution ladder without ArcGIS installed.

Loads the definition half of Wood_Pole_Replacement.py against a stub arcpy so
the pure logic (label parsing, alias matching, heuristics) can be tested with
realistic county layer schemas.
"""
import pathlib, re, sys, types


class Field:
    def __init__(self, name, alias=None, type="String"):
        self.name, self.aliasName, self.type = name, alias or name, type


class LabelClass:
    def __init__(self, expression, visible=True):
        self.expression, self.visible = expression, visible


class FakeLayer:
    """Stands in for an arcpy Layer: a name, a field list, optional labels."""
    def __init__(self, name, fields, labels=None):
        self.name = name
        self._fields = [f if isinstance(f, Field) else Field(*f) for f in fields]
        self._labels = labels

    def listLabelClasses(self):
        if self._labels is None:
            raise RuntimeError("streaming layer refuses label classes")
        return self._labels


arcpy = types.ModuleType("arcpy")
arcpy.env = types.SimpleNamespace(overwriteOutput=True, scratchGDB="memory")
arcpy.GetParameter = lambda i: None
arcpy.GetParameterAsText = lambda i: ""
arcpy.AddMessage = lambda m: None
arcpy.ListFields = lambda src: src._fields
sys.modules["arcpy"] = arcpy

TOOL = pathlib.Path(__file__).resolve().parent.parent / "Wood_Pole_Replacement.py"
source = TOOL.read_text()
head = source.split("# --- Buffer poles ---")[0]
head = "\n".join(
    line for line in head.splitlines()
    if not re.match(r"(import pandas|from openpyxl)", line)
)
ns = {}
exec(compile(head, str(TOOL), "exec"), ns)

resolve_fields = ns["resolve_fields"]
format_result = ns["format_result"]
_resolve_sj_field = ns["_resolve_sj_field"]
_parse_ids = ns["_parse_ids"]
find_field = ns["find_field"]
infer_role = ns["infer_role"]

failures = []


def check(label, got, want):
    ok = got == want
    print(f"  {'PASS' if ok else 'FAIL'}  {label}")
    if not ok:
        print(f"        got  {got!r}\n        want {want!r}")
        failures.append(label)


print("\n== The bug that made everything report 'Yes' ==")
# Field names with spaces in a label expression: the old \w+ regex could not match these.
mdot = FakeLayer(
    "MDOT LegalSystem",
    [("OBJECTID", "OBJECTID", "OID"), ("LEGALSYST", "Legal System"), ("PR", "PR Number", "Integer")],
    [LabelClass('$feature["Legal System"]')],
)
check("$feature[\"Legal System\"] resolves",
      resolve_fields(mdot, "MDOT LegalSystem", mdot, role="road")[0], ["LEGALSYST"])

vb = FakeLayer("County Drains", [("DRAINNAME", "County Drain Name")],
               [LabelClass("[County Drain Name]")])
check("[County Drain Name] resolves",
      resolve_fields(vb, "County Drains", vb)[0], ["DRAINNAME"])

print("\n== Same thing named differently in each county ==")
counties = {
    "Kent County Drains":    [("Label", "Label"), ("OBJECTID", "OBJECTID", "OID")],
    "Ottawa County Drains":  [("COUNTY_DRAIN_NAME", "County Drain Name")],
    "Allegan Drain Network": [("DRAIN_NO", "Drain Number", "Integer"), ("DRAIN_NM", "Drain Name")],
    "Barry Co Drains":       [("NAME", "NAME"), ("SHAPE_Length", "SHAPE_Length", "Double")],
}
expected = ["Label", "COUNTY_DRAIN_NAME", "DRAIN_NM", "NAME"]
for (name, fields), want in zip(counties.items(), expected):
    lyr = FakeLayer(name, fields)  # no labels — heuristics only
    check(f"{name}", resolve_fields(lyr, name, lyr, role=infer_role(name))[0], [want])

print("\n== Known aliases for the named layers ==")
soil = FakeLayer("Soils", [("musym", "Mapunit Symbol"), ("muname", "Mapunit Name"),
                           ("mukey", "Mapunit Key", "Integer")])
check("soil takes both fields",
      resolve_fields(soil, "Soils", soil, role="soil",
                     preferred=[["Mapunit Symbol", "MUSYM"], ["Mapunit Name", "MUNAME"]],
                     want=2)[0],
      ["musym", "muname"])

plss = FakeLayer("PLSS", [("STR_Label", "STR_Label"), ("TOWN", "TOWN", "Integer")])
check("STR_Label", resolve_fields(plss, "PLSS", plss, role="plss",
                                  preferred=[["STR_Label", "STR"]])[0], ["STR_Label"])

print("\n== Overrides and the give-up path ==")
weird = FakeLayer("Mystery Parcels", [("F17", "F17"), ("F18", "F18")])
check("override wins",
      resolve_fields(weird, "Mystery Parcels", weird,
                     overrides={ns["_norm"]("Mystery Parcels"): ["F18"]})[0], ["F18"])
fields, provenance = resolve_fields(weird, "Mystery Parcels", weird)
check("no match falls back to Yes/No", fields, [])
check("and names the available fields", "F17, F18" in provenance, True)

opaque = FakeLayer("Wetlands", [("OBJECTID", "OBJECTID", "OID"),
                                ("Shape_Area", "Shape_Area", "Double")])
check("geometry/OID-only layer gives up cleanly",
      resolve_fields(opaque, "Wetlands", opaque)[0], [])

print("\n== SpatialJoin field collisions ==")
sj = ["OBJECTID", "StructureN", "Label", "StructureN_1", "Label_1"]
check("target id keeps its name", _resolve_sj_field(sj, "StructureN", False), "StructureN")
check("join value takes the suffix", _resolve_sj_field(sj, "Label", True), "Label_1")
check("no collision, no suffix", _resolve_sj_field(["Label", "Pole"], "Label", True), "Label")
check("missing field is None", _resolve_sj_field(["A"], "Label", True), None)

print("\n== Result formatting ==")
hits = {"P1": {("MdB", "Miami loam"), ("MdC", "Miami sandy loam")}, "P2": set()}
check("multi-field joined", format_result(hits, "P1"), "MdB - Miami loam, MdC - Miami sandy loam")
check("symbol column", format_result(hits, "P1", index=0), "MdB, MdC")
check("name column", format_result(hits, "P1", index=1), "Miami loam, Miami sandy loam")
check("no hit", format_result(hits, "P2"), "No")
check("absent pole", format_result(hits, "P9"), "No")

print("\n== Blank answers say why they are blank ==")
# "No" is a real answer for a constraint layer and a meaningless one for a
# county lookup, so the two must not print the same thing.
blank_answer = ns["blank_answer"]
ns["pole_ids"] = ["P1", "P2", "P3"]

check("statewide lookup that matched nothing flags itself",
      blank_answer("County", {}, count=42, expect_every_structure=True),
      "No match - see Run Log")
check("empty layer is named as empty",
      blank_answer("County", {}, count=0, expect_every_structure=True),
      "Layer empty")
check("empty beats no-match when both are true",
      blank_answer("County", {}, count=0, expect_every_structure=False),
      "Layer empty")
check("constraint layer with no hits is a real No",
      blank_answer("Cemeteries", {}, count=17, expect_every_structure=False),
      "No")
check("partial coverage is still a real No",
      blank_answer("County", {"P1": {("Kent",)}}, count=42, expect_every_structure=True),
      "No")

check("default reaches the cell",
      format_result({}, "P1", default="Layer empty"), "Layer empty")
check("default does not override a real hit",
      format_result({"P1": {("Kent",)}}, "P1", default="Layer empty"), "Kent")
check("a hit whose name is blank still reads as a hit",
      format_result({"P1": {("Yes",)}}, "P1", default="Layer empty"), "Yes")

print("\n== Problem answers are styled, not mistaken for results ==")
problems = ["Layer empty", "Layer failed - see Run Log", "No match - see Run Log",
            "Incomplete - NWI Wetlands unreadable"]
prefixes = ("Layer empty", "Layer failed", "No match", "Incomplete")
check("every diagnostic string is recognised",
      all(v.startswith(prefixes) for v in problems), True)
check("a real answer is not", any(v.startswith(prefixes) for v in ["No", "Kent", "Yes"]), False)

print("\n== Pole ID parsing ==")
check("comma list", _parse_ids(" 101, 102 ,103 ", True), ["101", "102", "103"])
check("single is stringified", _parse_ids(1042, False), ["1042"])
check("null", _parse_ids(None, False), [])
check("blank", _parse_ids("   ", False), [])

print("\n== 2026 structure fields ==")
structures = FakeLayer("Structures_2026", [
    ("StructureN", "Structure Number"), ("Latitude", "Latitude", "Double"),
    ("Longitude", "Longitude", "Double"), ("Height1", "Height", "Double"),
    ("Installati", "Installation Date", "Date"), ("Structure_T", "Structure Type"),
])
idx = ns["field_index"](structures)
for column, candidates in ns["STRUCTURE_ATTRS"]:
    check(f"{column} found", find_field(idx, candidates) is not None, True)

legacy = FakeLayer("Structures_2024", [
    ("Structures_2024_StructureN", "StructureN"),
    ("Structures_Latitude", "Latitude", "Double"),
    ("Structures_Longitude", "Longitude", "Double"),
])
lidx = ns["field_index"](legacy)
check("2024 Latitude still resolves", find_field(lidx, ["Latitude", "Structures_Latitude"]),
      "Structures_Latitude")
check("2024 Height absent -> None", find_field(lidx, ["Height1", "Height"]), None)

print(f"\n{'ALL PASS' if not failures else str(len(failures)) + ' FAILURE(S): ' + ', '.join(failures)}")
sys.exit(1 if failures else 0)
