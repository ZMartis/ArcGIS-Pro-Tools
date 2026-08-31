# Wood Pole Replacement Tool — 2026 Update Plan

Scope: field renames for the 2026 streaming structure layer, replacing "Yes" answers with
actual feature names, six new named-layer analyses, a redesigned SESC check, and an opt-in
2-mile T&E review.

> **Status:** implemented on branch `2026-update`. The resolution ladder is covered by
> `tests/test_field_resolution.py`, which stubs `arcpy` and runs without ArcGIS installed.
>
> **Note on the screenshots:** the attached images (`image001`–`image009`) did not come
> through as readable files, so every field name below is taken from the written feedback.
> Two need a quick confirm before coding — flagged as **CONFIRM** where they appear.

---

## 1. Root cause: why constraints report "Yes" instead of a name

The tool already tries to be smart about this. `get_label_field()` reads the layer's label
class and pulls the field out of the Arcade expression. It fails in three ways:

| # | Problem | Location |
|---|---|---|
| 1 | `re.search(r'\$feature\.(\w+)', expr)` — `\w+` excludes spaces, so `$feature["Legal System"]` and `$feature['County Drain Name']` never match | `Wood_Pole_Replacement.py:114` |
| 2 | `re.search(r'\[(\w+)\]', expr)` — same problem for `[County Drain Name]` | `Wood_Pole_Replacement.py:118` |
| 3 | Streaming/web layers often raise or return `[]` from `listLabelClasses()`; the bare `except Exception: pass` swallows it silently | `Wood_Pole_Replacement.py:121` |

Any of these returns `None`, and `get_intersect_hits()` then takes the no-label branch that
hardcodes `{"Yes"}` (`Wood_Pole_Replacement.py:174`). There is no second attempt.

A related latent bug: when a label field *is* found, SpatialJoin field-name collision is
handled by guessing `f"{label_field}_1"` (`Wood_Pole_Replacement.py:154`). ArcGIS can produce
`_12`, or sanitize spaces to underscores, so this guess can also miss.

**This is the single highest-value fix** — it is what makes county drains, parcels, MDOT
jurisdiction, and soil all work, and it is why they currently don't.

---

## 2. Field resolution overhaul

Replace `get_label_field()` with a `resolve_fields(layer, role, overrides)` ladder. It returns
a list of field names (soil needs two) plus a provenance string for logging.

1. **Explicit override** — new *Field Overrides* value-table parameter (`Layer Name` →
   `Field(s)`, comma-separated). Always wins. This is the escape hatch for the county layer
   that defeats every heuristic.
2. **Known name/alias for this analysis** — compare against field *names and aliases*,
   normalized case-insensitively with spaces and underscores collapsed. This resolves
   `"Legal System"`, `"Mapunit Symbol"` and `"Mapunit Name"`, which are aliases rather than
   field names, and it runs ahead of the label expression because for a named layer we already
   know what we're looking for. Generic constraint layers have no known alias, so for them the
   ladder collapses to override → label → heuristics.
3. **Label expression parse** — rewritten to handle all real-world forms:
   `$feature.NAME`, `$feature["Name With Spaces"]`, `$feature['Name']`, `[Name With Spaces]`,
   and concatenations (captures every field referenced, in order). Iterates all label classes,
   preferring visible ones, not just `[0]`.
4. **Heuristic scoring** over name + alias:
   - Strong: exact `LABEL`, `STR_LABEL`, `NAME`, `MUSYM`, `MUNAME`
   - Boost on role keywords — drain: `DRAIN`, `DRAINNAME`; parcel: `PARCEL`, `PIN`, `OWNER`,
     `TAXID`; township: `TWP`, `TOWNSHIP`, `MUNI`; county: `COUNTY`, `CNTY`
   - Penalize: `OBJECTID`, `FID`, `SHAPE*`, `GLOBALID`, `*_ID`, pure-numeric and date fields
5. **Fail loudly, not silently** — if nothing scores, fall back to Yes/No *and* log the
   layer's full field list so you know exactly what to type into the override table next run.

Every resolution is written to `arcpy.AddMessage` **and** to a new `Run_Log` worksheet, so the
tool's guess is auditable rather than invisible.

Also fix the SpatialJoin collision handling: resolve the output field by scanning the actual
`arcpy.ListFields(sj_out)` result for a normalized match instead of guessing a `_1` suffix.

**Multi-field support:** `get_intersect_hits()` gains a list-of-fields signature so soil can
return `"MdB — Miami loam, 2 to 6 percent slopes"` as one value.

---

## 3. 2026 structure layer — field renames and new attributes

Renames (`Wood_Pole_Replacement.py:201` is the hardcoded one that will crash today):

| Old | New |
|---|---|
| `Structures_2024_StructureN` | `StructureN` |
| `Structures_Latitude` | `Latitude` |
| `Structures_Longitude` | `Longitude` |

New attributes to carry into the matrix:

| Output column | Source field | |
|---|---|---|
| `Height` | `Height1` | |
| `Install_Date` | `Installati` | |
| `Structure_Type` | `Structure T` | **CONFIRM** — written with a space; likely `Structure_T`. The resolver will try `Structure T`, `Structure_T`, and alias match, so either works. |

Implementation notes:

- Resolve every structure field through the same tolerant resolver — try new name, then old
  name, then alias. The layer is a stream that changed once and will change again; a rename
  should produce a warning and a blank column, never a crash.
- `StructureN` is passed via parameter 5, so that rename is really just the parameter default.
  Latitude/Longitude are hardcoded and must change in code.

---

## 4. Buffer change: 100 ft → 500 ft, for SESC only

The 500 ft buffer replaces the 100 ft buffer as geometry, but not as an analysis. It exists to
answer one question — the SESC check in §6 — and nothing else is measured against it.

- `pole_buffer_100` → `pole_buffer_500`, `"100 Feet"` → `"500 Feet"`.
- **The `100ft_Pole_*` columns go away rather than becoming `500ft_Pole_*`.** Constraint layers
  are a work-zone question: each one produces a `60ft_Pole_*` and an `AR_*` column, and that is
  all. A county road ROW 300 ft from the structure is not a constraint on the job.
- This drops the second spatial join per constraint layer, so the constraint loop gets
  meaningfully faster, and the `"Same as 60ft"` collapsing logic is no longer needed.
- The wide-buffer highlighting in the Excel formatter (yellow fill, bolding features the wider
  buffer found that the work zone did not) goes with those columns. Hit columns are orange,
  informational columns are plain.
- Derived output parameter still returns the 500 ft buffer, so it lands in the Contents pane
  where the 100 ft buffer used to.
- The AOI used to download streaming layers is built from the widest buffer in play, since SESC
  still needs water and wetland data out to 500 ft.

---

## 5. New named-layer analyses

These get their own optional parameters rather than being sniffed out of the map by name —
streaming layer names vary too much to pattern-match reliably.

| Analysis | Geometry tested | Field(s) | Output column(s) |
|---|---|---|---|
| County | pole point | `LABEL` | `County` |
| Township / City | pole point | `LABEL` | `Township_City` |
| Section-Township-Range | pole point | `STR_Label` | `Sec_Twp_Rng` |
| MDOT Jurisdiction | 60 ft work zone + access routes | `Legal System` (alias) | `MDOT_Jurisdiction`, `AR_MDOT_Jurisdiction` |
| Soil | **pole point only** — per your note, not the 60 ft buffer | `Mapunit Symbol` + `Mapunit Name` | `Soil_Symbol`, `Soil_Name` |
| Parcel | pole point | auto-resolved | `Parcel` |

MDOT is deliberately tested against the work zone rather than the pole: the ROW permit is
triggered by disturbance in the ROW, which is what the 60 ft buffer represents. Access routes
are checked too, since they also enter ROW.

Point-in-polygon lookups (county, township, STR, soil, parcel) join against the pole points
rather than a buffer, which is much cheaper. They stay one-to-many so a structure sitting on a
parcel or map-unit boundary reports both, instead of silently picking one.

---

## 6. SESC redesign

Current workflow: a pre-buffered 500 ft open-water layer, tested against the pole and the
60 ft buffer, answer is Yes/No.

New workflow, as requested:

- Buffer the **pole** by 500 ft — this is now the standard large buffer from §4, so no extra
  geometry is created.
- Test that buffer against a **multi-value SESC Layers parameter** so you choose per county
  what counts: open water only, water + PEM, or water + all wetlands.
- County drains are included here and run through the full resolver from §2, so you get
  *"Miller Drain"*, not *"Yes"*.
- For wetland layers, also pull the type/class field (`ATTRIBUTE`, `WETLAND_TYPE`) so PEM vs
  PFO is visible in the output rather than something you have to go look up.

Output: `SESC_Trigger` (Yes/No) and `SESC_Features` (e.g. `"Miller Drain; PEM wetland"`).

**One decision to make.** Buffering the pole by 500 ft is mathematically identical to your old
"water buffered 500 ft, test the pole" check — same answer. But your old check *also* triggered
off the 60 ft buffer, which is effectively a 560 ft reach. Since earth disturbance actually
happens across the whole work zone, the technically correct test is `60 ft work zone + 500 ft`
= 560 ft from the pole. I'll build it as you asked (500 ft from the pole) with a checkbox to
use the work zone instead, so you don't silently lose the 500–560 ft band you were catching
before.

---

## 7. Two-mile T&E review

Opt-in checkbox, off by default, project-wide rather than per-pole:

1. Dissolve all poles + access routes into one project footprint.
2. Buffer that footprint by 2 miles — one buffer, not one per pole. This is strictly better
   than picking a central structure by hand, since it covers linear projects that a single
   2-mile circle would miss.
3. Download MNFI through the existing `ensure_feature_layer()` path, using the 2-mile buffer as
   the AOI.
4. One spatial join; write results to a separate `T&E_2mi` worksheet listing each element
   occurrence with its name and status fields (auto-resolved).

Cost is one buffer + one join + one web download. It will not meaningfully slow the per-pole
matrix, and it's skipped entirely when the box is unchecked.

---

## 8. Robustness pass

Real defects found while reading, independent of the new features:

| Issue | Location | Fix |
|---|---|---|
| Hardcoded `min_col=4` in the Excel formatter — silently mis-colors once Height/Install/Type/County/etc. columns are added | `Wood_Pole_Replacement.py:306` | Classify each column from its header instead of assuming a start position |
| `constraint_100ft_sel` feature layer leaks if `get_intersect_hits` raises | `Wood_Pole_Replacement.py:254` | `try/finally` cleanup |
| Inconsistent pole-ID key types — buffers key on raw `row[0]`, access routes on `str(pid).strip()` | `Wood_Pole_Replacement.py:263-277` | Normalize every ID to a stripped string once, at read time |
| One bad constraint layer kills the whole run | constraint loop, `Wood_Pole_Replacement.py:236` | Per-layer `try/except`; record failures and keep going |
| Duplicate pole IDs silently collapse in `pole_coords` | `Wood_Pole_Replacement.py:199` | Warn on duplicates |
| `pole_coords[pid]` raises `KeyError` on any mismatch | `Wood_Pole_Replacement.py:283` | `.get()` with blank fallback |
| Empty pole selection produces a confusing downstream failure | — | Guard and exit with a clear message |
| `iter_rows()` takes `max_col`, but `ws.cell()` takes `column` — mixing them up raises only at the very end of a run, after every spatial join is already paid for | Excel formatter | Fixed, plus a static AST check in the test suite that validates openpyxl keyword names without needing openpyxl installed |

**Blank answers explain themselves.** `"No"` currently means three different things: a genuine
non-overlap, a layer that downloaded zero features, and a layer that failed. For a county or
township lookup it is never a real answer — every structure is in some county — so a silent
`"No"` there hides a broken layer.

Root cause found in the field: `ensure_feature_layer()` calls `SelectLayerByLocation` then
`CopyFeatures`, and **`CopyFeatures` succeeds on an empty selection**. A streaming layer whose
spatial selection silently fails returns zero features, every join misses, and every cell reads
`"No"`. An empty download is now treated as a failed strategy so it falls through to the next
one, and the answers are separated:

| Cell reads | Means |
|---|---|
| `No` | genuinely no overlap |
| `Layer empty` | the layer returned zero features in the project area |
| `No match - see Run Log` | a statewide lookup matched no structure at all — almost certainly wrong |
| `Layer failed - see Run Log` | the layer raised, and was skipped |
| `Incomplete - <layer> unreadable` | SESC, where one of the chosen water layers could not be read |

These are grey-filled and italic red, so they never read as a clean result at a glance. The Run
Log gains a feature count and an `N of M structures matched` line per layer, plus a note when a
layer is reprojected on the fly.

**Workbook structure** becomes three sheets:

- `Pole Matrix` — Pole_ID, Latitude, Longitude, Height, Install_Date, Structure_Type, County,
  Township_City, Sec_Twp_Rng, Soil_Symbol, Soil_Name, Parcel, MDOT_Jurisdiction, SESC_Trigger,
  SESC_Features, then a `60ft_Pole_*` and `AR_*` pair per constraint layer.
- `T&E_2mi` — only when enabled.
- `Run_Log` — which field was chosen for each layer and how it was chosen, plus any layer
  failures. This is what turns a wrong guess into a two-second override instead of a mystery.

---

## 9. Toolbox parameters to wire manually

Existing indices 0–6 are unchanged. The two derived outputs move from 7–8 to the end.

| Idx | Name | Data type | Notes |
|---|---|---|---|
| 0 | Poles | Feature Layer | unchanged |
| 1 | Access Routes | Feature Layer | unchanged |
| 2 | Constraint Layers | Feature Layer, multi-value | unchanged |
| 3 | Output Folder | Folder | unchanged |
| 4 | Project Name | String | unchanged |
| 5 | Pole ID Field | Field, obtained from 0 | default `StructureN` |
| 6 | AR Pole ID Field | Field, obtained from 1 | unchanged |
| 7 | County Layer | Feature Layer, optional | new |
| 8 | Township / City Layer | Feature Layer, optional | new |
| 9 | PLSS (Sec-Twp-Rng) Layer | Feature Layer, optional | new |
| 10 | MDOT LegalSystem Layer | Feature Layer, optional | new |
| 11 | Soil Layer | Feature Layer, optional | new |
| 12 | Parcel Layer | Feature Layer, optional | new |
| 13 | SESC Layers | Feature Layer, multi-value, optional | new |
| 14 | SESC from work zone (560 ft) | Boolean, optional, default `false` | new — see §6 |
| 15 | Run 2-mile T&E Review | Boolean, optional, default `false` | new |
| 16 | MNFI Layer | Feature Layer, optional | new — enable when 15 is checked |
| 17 | Field Overrides | Value Table: `Layer Name` (String), `Field(s)` (String), optional | new — see §2 |
| 18 | 60 ft Buffer | Feature Class, **Derived**, Output | moved from 7 |
| 19 | 500 ft Buffer | Feature Class, **Derived**, Output | moved from 8 |

Every new layer parameter is optional — leave one blank and that analysis is skipped and its
columns omitted, so the tool still runs on projects where you don't have that data.

---

## 10. Build order

Phased so the blocking problem is fixed first and each phase is independently useful.

**Phase 1 — unblock the 2026 layer.** Field renames, the three new structure attributes,
100 ft → 500 ft. Small, self-contained, gets the tool running again on current data.

**Phase 2 — the "which drain?" fix.** Field resolution ladder, override parameter, `Run_Log`
sheet, SpatialJoin field-collision fix. This is the highest-value change and it makes every
subsequent phase work.

**Phase 3 — new analyses.** County, township, STR, MDOT, soil, parcel, and the SESC redesign.
Depends on Phase 2.

**Phase 4 — 2-mile T&E.** Fully separable; can be deferred without holding anything up.

The robustness fixes in §8 fold into whichever phase touches the relevant code.
