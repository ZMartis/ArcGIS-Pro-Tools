import arcpy
import os
import re
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl import load_workbook
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import CellRichText, TextBlock

arcpy.env.overwriteOutput = True

# --- Inputs ---

poles = arcpy.GetParameter(0)
access_routes = arcpy.GetParameter(1)
constraint_layers = arcpy.GetParameter(2)
out_folder = arcpy.GetParameterAsText(3)
project_name = arcpy.GetParameterAsText(4)
pole_id_field = arcpy.GetParameterAsText(5)
ar_pole_id_field = arcpy.GetParameterAsText(
    6
)  # Field in access routes with comma-separated pole IDs

gdb = arcpy.env.scratchGDB


def ensure_feature_layer(layer, scratch_gdb, aoi_features=None, aoi_extent_fc=None):
    desc = arcpy.Describe(layer)

    # Detect web/streaming layers — these have shapeType but don't
    # work reliably with geoprocessing tools like SpatialJoin
    is_web = getattr(layer, "isWebLayer", False)
    if not is_web and hasattr(desc, "catalogPath"):
        is_web = str(desc.catalogPath).startswith("http")
    if not is_web and hasattr(desc, "path"):
        is_web = str(desc.path).startswith("http")

    # Local feature layer/class — safe to use directly
    if not is_web and hasattr(desc, "shapeType"):
        return layer

    # Web/service/streaming layer — three strategies, each falling back to the next:
    #   1. Select by actual project features (most precise)
    #   2. Select by simple AOI rectangle (simpler query)
    #   3. CopyFeatures with extent environment (last resort)
    layer_name = getattr(layer, "name", layer)
    arcpy.AddMessage(f"Copying web layer '{layer_name}' locally...")
    output_fc = arcpy.CreateUniqueName("web_layer", scratch_gdb)

    def _try_select_and_copy(select_features):
        arcpy.management.MakeFeatureLayer(layer, "web_download_sel")
        for i, aoi_fc in enumerate(select_features):
            sel_type = "NEW_SELECTION" if i == 0 else "ADD_TO_SELECTION"
            arcpy.management.SelectLayerByLocation(
                "web_download_sel", "INTERSECT", aoi_fc,
                selection_type=sel_type,
            )
        result = arcpy.management.CopyFeatures("web_download_sel", output_fc)[0]
        arcpy.management.Delete("web_download_sel")
        return result

    # Strategy 1: Select by actual project features
    if aoi_features:
        try:
            temp_fc = _try_select_and_copy(aoi_features)
            arcpy.AddMessage(f"  Downloaded via project feature selection.")
        except Exception as e:
            arcpy.AddMessage(f"  Selection by project features failed ({e}), trying next method...")
            if arcpy.Exists("web_download_sel"):
                arcpy.management.Delete("web_download_sel")
            temp_fc = None
        if temp_fc:
            return arcpy.management.MakeFeatureLayer(
                temp_fc, arcpy.CreateUniqueName("temp_lyr")
            )[0]

    # Strategy 2: Select by simple AOI rectangle
    if aoi_extent_fc:
        try:
            temp_fc = _try_select_and_copy([aoi_extent_fc])
            arcpy.AddMessage(f"  Downloaded via AOI rectangle selection.")
        except Exception as e:
            arcpy.AddMessage(f"  Selection by AOI rectangle failed ({e}), trying next method...")
            if arcpy.Exists("web_download_sel"):
                arcpy.management.Delete("web_download_sel")
            temp_fc = None
        if temp_fc:
            return arcpy.management.MakeFeatureLayer(
                temp_fc, arcpy.CreateUniqueName("temp_lyr")
            )[0]

    # Strategy 3: CopyFeatures with extent environment (last resort)
    arcpy.AddMessage(f"  Falling back to extent-based download for '{layer_name}'...")
    original_extent = arcpy.env.extent
    if aoi_extent_fc:
        arcpy.env.extent = arcpy.Describe(aoi_extent_fc).extent
    try:
        temp_fc = arcpy.management.CopyFeatures(layer, output_fc)[0]
    finally:
        arcpy.env.extent = original_extent

    return arcpy.management.MakeFeatureLayer(
        temp_fc, arcpy.CreateUniqueName("temp_lyr")
    )[0]


def get_label_field(layer):
    """Extract the primary label field name from a layer's label classes."""
    try:
        label_classes = layer.listLabelClasses()
        if label_classes:
            expr = label_classes[0].expression
            # Arcade format: $feature.FIELDNAME
            match = re.search(r'\$feature\.(\w+)', expr)
            if match:
                return match.group(1)
            # Python/VBScript format: [FIELDNAME]
            match = re.search(r'\[(\w+)\]', expr)
            if match:
                return match.group(1)
    except Exception:
        pass
    return None


def _parse_ids(raw_id, multi_id):
    """Parse an ID value, optionally splitting on commas."""
    if raw_id is None:
        return []
    if multi_id:
        return [p.strip() for p in str(raw_id).split(",") if p.strip()]
    return [raw_id]


def get_intersect_hits(target, join_layer, id_field, label_field, scratch_gdb, multi_id=False):
    """
    Spatial intersection returning {id: set of values}.
    With label_field: values are the label field values from intersecting features.
    Without label_field: values are {"Yes"} for any intersecting feature.
    If multi_id is True, id_field values are split on commas.
    """
    hits = {}

    if label_field:
        sj_out = arcpy.CreateUniqueName("sj_temp", scratch_gdb)
        arcpy.analysis.SpatialJoin(
            target, join_layer, sj_out,
            join_operation="JOIN_ONE_TO_MANY",
            join_type="KEEP_COMMON",
            match_option="INTERSECT",
        )

        sj_fields = [f.name for f in arcpy.ListFields(sj_out)]
        actual_lf = label_field if label_field in sj_fields else f"{label_field}_1"

        with arcpy.da.SearchCursor(sj_out, [id_field, actual_lf]) as cursor:
            for row in cursor:
                val = row[1]
                for pid in _parse_ids(row[0], multi_id):
                    if pid not in hits:
                        hits[pid] = set()
                    if val is not None and str(val).strip():
                        hits[pid].add(str(val).strip())

        arcpy.management.Delete(sj_out)
    else:
        lyr_name = "temp_intersect_sel"
        arcpy.management.MakeFeatureLayer(target, lyr_name)
        arcpy.management.SelectLayerByLocation(lyr_name, "INTERSECT", join_layer)

        with arcpy.da.SearchCursor(lyr_name, [id_field]) as cursor:
            for row in cursor:
                for pid in _parse_ids(row[0], multi_id):
                    hits[pid] = {"Yes"}

        arcpy.management.Delete(lyr_name)

    return hits


def format_result(hits, pid):
    """Format intersection hits into a result string."""
    if pid in hits and hits[pid]:
        return ", ".join(sorted(hits[pid]))
    return "No"


# --- Buffer poles ---

pole_buffer_60 = os.path.join(gdb, "PoleBuffer_60ft")
arcpy.analysis.Buffer(poles, pole_buffer_60, "60 Feet", dissolve_option="NONE")

pole_buffer_100 = os.path.join(gdb, "PoleBuffer_100ft")
arcpy.analysis.Buffer(poles, pole_buffer_100, "100 Feet", dissolve_option="NONE")

# --- Build pole list with coordinates ---

pole_ids = []
pole_coords = {}

with arcpy.da.SearchCursor(pole_buffer_60, [pole_id_field, "Structures_Latitude", "Structures_Longitude"]) as cursor:
    for row in cursor:
        pole_ids.append(row[0])
        pole_coords[row[0]] = {"Latitude": row[1], "Longitude": row[2]}

# --- Build AOI for web layer downloads ---

aoi_buffer_dissolved = os.path.join(gdb, "aoi_buffer_dissolved")
arcpy.management.Dissolve(pole_buffer_100, aoi_buffer_dissolved)

aoi_ar_dissolved = os.path.join(gdb, "aoi_ar_dissolved")
arcpy.management.Dissolve(access_routes, aoi_ar_dissolved)

aoi_features = [aoi_buffer_dissolved, aoi_ar_dissolved]

pole_ext = arcpy.Describe(pole_buffer_100).extent
ar_ext = arcpy.Describe(access_routes).extent
sr = arcpy.Describe(pole_buffer_100).spatialReference

aoi_polygon = arcpy.Polygon(arcpy.Array([
    arcpy.Point(min(pole_ext.XMin, ar_ext.XMin), min(pole_ext.YMin, ar_ext.YMin)),
    arcpy.Point(min(pole_ext.XMin, ar_ext.XMin), max(pole_ext.YMax, ar_ext.YMax)),
    arcpy.Point(max(pole_ext.XMax, ar_ext.XMax), max(pole_ext.YMax, ar_ext.YMax)),
    arcpy.Point(max(pole_ext.XMax, ar_ext.XMax), min(pole_ext.YMin, ar_ext.YMin)),
]), sr)

aoi_extent_fc = os.path.join(gdb, "aoi_extent_poly")
arcpy.management.CopyFeatures([aoi_polygon], aoi_extent_fc)

# --- Initialize results dictionary ---

results = {pid: {} for pid in pole_ids}

# --- Loop through constraints ---

for constraint in constraint_layers:

    constraint_name = constraint.name
    label_field = get_label_field(constraint)
    constraint_lyr = ensure_feature_layer(
        constraint, gdb, aoi_features=aoi_features, aoi_extent_fc=aoi_extent_fc
    )

    if label_field:
        arcpy.AddMessage(f"Using label field '{label_field}' for {constraint_name}")
    else:
        arcpy.AddMessage(f"No label field found for {constraint_name}, using Yes/No")

    # Run 100ft first — only features that intersect the larger buffer can intersect the smaller one
    hits_100 = get_intersect_hits(pole_buffer_100, constraint_lyr, pole_id_field, label_field, gdb)

    if hits_100:
        # Filter constraint to only features that intersected the 100ft buffer
        arcpy.management.MakeFeatureLayer(constraint_lyr, "constraint_100ft_sel")
        arcpy.management.SelectLayerByLocation("constraint_100ft_sel", "INTERSECT", pole_buffer_100)
        hits_60 = get_intersect_hits(pole_buffer_60, "constraint_100ft_sel", pole_id_field, label_field, gdb)
        arcpy.management.Delete("constraint_100ft_sel")
    else:
        hits_60 = {}

    hits_ar = get_intersect_hits(access_routes, constraint_lyr, ar_pole_id_field, label_field, gdb, multi_id=True)

    for pid in pole_ids:
        val_60 = format_result(hits_60, pid)
        results[pid][f"60ft_Pole_{constraint_name}"] = val_60

        set_60 = hits_60.get(pid, set())
        set_100 = hits_100.get(pid, set())

        if set_100 and set_100 <= set_60:
            # 100ft hits are the same as 60ft — no new info
            results[pid][f"100ft_Pole_{constraint_name}"] = "Same as 60ft"
        else:
            results[pid][f"100ft_Pole_{constraint_name}"] = format_result(hits_100, pid)

        pid_str = str(pid).strip()
        results[pid][f"AR_{constraint_name}"] = format_result(hits_ar, pid_str)

# --- Create DataFrame ---

df = pd.DataFrame.from_dict(results, orient="index")
df.insert(0, "Pole_ID", df.index)
df.insert(1, "Latitude", df["Pole_ID"].map(lambda pid: pole_coords[pid]["Latitude"]))
df.insert(2, "Longitude", df["Pole_ID"].map(lambda pid: pole_coords[pid]["Longitude"]))
df.reset_index(drop=True, inplace=True)

# --- Export with color coding ---

excel_path = os.path.join(out_folder, f"{project_name}_Pole_Constraint_Matrix.xlsx")
df.to_excel(excel_path, index=False)

# Apply color coding
orange_fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
black_font = Font(color="000000", bold=True)
bold_inline = InlineFont(b=True)
regular_inline = InlineFont()

wb = load_workbook(excel_path)
ws = wb.active

# Map headers for column lookups
headers = {cell.column: cell.value for cell in ws[1]}
header_to_col = {cell.value: cell.column for cell in ws[1]}

for row in ws.iter_rows(min_row=2, min_col=4, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
        val = str(cell.value) if cell.value else ""

        if not val or val == "No" or val == "Same as 60ft":
            continue

        header = headers.get(cell.column, "")

        if header and header.startswith("100ft_Pole_"):
            # 100ft column with new features — yellow fill, bold the new ones
            constraint_suffix = header[len("100ft_Pole_"):]
            sixty_col = header_to_col.get(f"60ft_Pole_{constraint_suffix}")

            sixty_features = set()
            if sixty_col:
                sixty_val = ws.cell(row=cell.row, column=sixty_col).value
                if sixty_val and str(sixty_val) != "No":
                    sixty_features = {f.strip() for f in str(sixty_val).split(", ")}

            hundred_features = [f.strip() for f in val.split(", ")]

            # Build rich text: bold features that are new (not in 60ft)
            parts = []
            for j, feat in enumerate(hundred_features):
                if j > 0:
                    parts.append(TextBlock(regular_inline, ", "))
                if feat not in sixty_features:
                    parts.append(TextBlock(bold_inline, feat))
                else:
                    parts.append(TextBlock(regular_inline, feat))

            cell.value = CellRichText(*parts)
            cell.fill = yellow_fill
        else:
            # 60ft or AR column — orange fill
            cell.fill = orange_fill
            cell.font = black_font

wb.save(excel_path)
arcpy.AddMessage(f"Matrix created: {excel_path}")

# --- Add buffer layers to Contents pane ---

arcpy.SetParameter(7, pole_buffer_60)
arcpy.SetParameter(8, pole_buffer_100)
