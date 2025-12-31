# -*- coding: utf-8 -*-
"""
Set Comment Values for Elements (natural/numeric sort by Mark)
- Select category, view, filter by type, and assign numbered Comments based on Mark order.
- Natural sort: numeric parts inside Mark are treated as numbers (IPP-2 < IPP-10).
"""
__title__ = 'Set Mark\nValues (Natural Sort)'
__doc__ = 'Assign mark values to filtered elements by type in a selected view (natural numeric ordering)'

import clr
import re
from pyrevit import revit, forms, script

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    View,
    StorageType,
    FamilyInstance
)

doc = revit.doc
output = script.get_output()

# -------------------------
# Step 0: Category options map
# -------------------------
category_options = [
    'Walls',
    'Doors',
    'Windows',
    'Structural Framing',
    'Structural Columns',
    'Floors',
    'Furniture',
    'Generic Models'
]
category_map = {
    'Walls': BuiltInCategory.OST_Walls,
    'Doors': BuiltInCategory.OST_Doors,
    'Windows': BuiltInCategory.OST_Windows,
    'Structural Framing': BuiltInCategory.OST_StructuralFraming,
    'Structural Columns': BuiltInCategory.OST_StructuralColumns,
    'Floors': BuiltInCategory.OST_Floors,
    'Furniture': BuiltInCategory.OST_Furniture,
    'Generic Models': BuiltInCategory.OST_GenericModel
}

# -------------------------
# Step 1: Select Category
# -------------------------
selected_category_name = forms.SelectFromList.show(
    category_options,
    title='Step 1: Select Category',
    button_name='Select'
)
if not selected_category_name:
    script.exit()

selected_category = category_map[selected_category_name]

# -------------------------
# Step 2: Select View
# -------------------------
all_views = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType().ToElements()
valid_views = [v for v in all_views if (v is not None) and (not v.IsTemplate)]
if not valid_views:
    forms.alert('No valid views found', exitscript=True)
view_dict = {v.Name: v for v in valid_views}

selected_view_name = forms.SelectFromList.show(
    sorted(view_dict.keys()),
    title='Step 2: Select View',
    button_name='Select'
)
if not selected_view_name:
    script.exit()

selected_view = view_dict[selected_view_name]

# -------------------------
# Step 3: Collect all elements of category in view
# -------------------------
collector = FilteredElementCollector(doc, selected_view.Id).WhereElementIsNotElementType().OfCategory(selected_category)
try:
    all_elements = list(collector.ToElements())
except Exception:
    all_elements = [e for e in collector]

if not all_elements:
    forms.alert('No {} found in view "{}"'.format(selected_category_name, selected_view.Name), exitscript=True)

output.print_md("Found **{}** {} in view **{}**".format(len(all_elements), selected_category_name, selected_view.Name))

# -------------------------
# Step 4: Group elements by Type (friendly type name)
# -------------------------
def friendly_type_name(el):
    """Return a friendly type name for an element with safe fallbacks."""
    try:
        tid = el.GetTypeId()
        if tid and tid.IntegerValue != 0:
            t = doc.GetElement(tid)
            if t is not None:
                # most type elements have a Name property
                name = getattr(t, "Name", None)
                if name:
                    return name
                # fallback to looking up common params
                try:
                    p = t.LookupParameter("Type Name")
                    if p:
                        try:
                            return p.AsString() or p.AsValueString() or "<unknown type>"
                        except:
                            return getattr(t, "Name", "<unknown type>")
                except:
                    pass
    except Exception:
        pass

    # If family instance, include family : symbol
    try:
        if isinstance(el, FamilyInstance):
            try:
                sym = getattr(el, "Symbol", None)
                if sym:
                    fam = getattr(sym, "Family", None)
                    fam_name = getattr(fam, "Name", None) if fam else None
                    sym_name = getattr(sym, "Name", None)
                    if fam_name and sym_name:
                        return "{} : {}".format(fam_name, sym_name)
                    elif sym_name:
                        return sym_name
            except:
                pass
    except:
        pass

    # final fallback: class name or category name
    try:
        return el.GetType().Name
    except:
        try:
            cat = getattr(el, "Category", None)
            if cat and getattr(cat, "Name", None):
                return cat.Name
        except:
            pass
    return "<unknown>"

type_dict = {}
for elem in all_elements:
    try:
        tname = friendly_type_name(elem)
        type_dict.setdefault(tname, []).append(elem)
    except Exception:
        # ignore problematic element
        continue

if not type_dict:
    forms.alert("No element types discovered for the selected category in view.", exitscript=True)

# Build choices for user including counts
type_options = ["{} ({} elements)".format(tname, len(elems)) for tname, elems in sorted(type_dict.items())]

selected_type_option = forms.SelectFromList.show(
    type_options,
    title='Step 3: Select Element Type',
    button_name='Select',
    multi_select=False
)
if not selected_type_option:
    script.exit()

selected_type_name = selected_type_option.split(' (')[0]
filtered_elements = type_dict.get(selected_type_name, [])

output.print_md("Filtered to **{}** elements of type **{}**".format(len(filtered_elements), selected_type_name))


# -------------------------
# Step 5: Prefix Input
# -------------------------
prefix = forms.ask_for_string(
    title="Comments Prefix",
    prompt="Enter prefix (example: BS1-GF-IPP15-)"
)
if prefix is None:
    script.exit()

# -------------------------
# Helper: robust Mark reading
# -------------------------
def read_mark(el):
    p = el.LookupParameter("Mark")
    if p is None:
        return None
    try:
        if p.StorageType == StorageType.String:
            v = p.AsString()
            return v if v is not None else None
        if p.StorageType == StorageType.Integer:
            return str(p.AsInteger())
        if p.StorageType == StorageType.Double:
            try:
                return p.AsValueString()
            except:
                return str(p.AsDouble())
    except Exception:
        try:
            return p.AsValueString()
        except:
            return None
    return None

# -------------------------
# Step 6: Collect marked elements from filtered_elements (NOT all_elements)
# -------------------------
marked_elements = []
skipped_no_mark = []
for e in filtered_elements:
    mv = read_mark(e)
    if mv is None:
        skipped_no_mark.append(e)
    else:
        mv_str = str(mv).strip()
        if mv_str == "":
            skipped_no_mark.append(e)
        else:
            marked_elements.append((mv_str, e))

# If nothing to do, inform user
if not marked_elements:
    forms.alert("No elements with a Mark value found in the selected type.", exitscript=True)

# -------------------------
# Step 7: Natural / alphanumeric sort helper
# -------------------------
_digits_re = re.compile(r'(\d+)')

def alphanum_key(s):
    """
    Turn a string into a list of string and numeric chunks:
    "IPP-12-A3" -> ["IPP-", 12, "-A", 3]
    This sorts numerically on numbers and lexicographically on text,
    producing natural sort order.
    """
    parts = _digits_re.split(s)
    key = []
    for p in parts:
        if p.isdigit():
            key.append(int(p))
        else:
            key.append(p.lower())
    return tuple(key)

# Sort the marked_elements by the natural key of the mark text
marked_elements.sort(key=lambda pair: alphanum_key(pair[0]))

# -------------------------
# Step 8: Write Comments with incremental numbering
# -------------------------
count = 1
updated = 0
skipped_no_comments = []
failed = []

with revit.Transaction("Set Comments from Mark Order (Natural Sort)"):
    for mark_text, elem in marked_elements:
        try:
            c = elem.LookupParameter("Comments")
            if c is None or c.IsReadOnly:
                skipped_no_comments.append(elem)
                continue
            # Set as text (Comments typically string)
            try:
                c.Set("{}{:03d}".format(prefix, count))
            except Exception as ex:
                failed.append((elem, str(ex)))
                continue
            count += 1
            updated += 1
        except Exception as ex:
            failed.append((elem, str(ex)))

# -------------------------
# Result summary
# -------------------------
output.print_md("## ✅ Completed")
output.print_md("* Selected Category: `{}`".format(selected_category_name))
output.print_md("* Selected View: `{}`".format(selected_view.Name))
output.print_md("* Selected Type: `{}`".format(selected_type_name))
output.print_md("* Elements in selected type: **{}**".format(len(filtered_elements)))
output.print_md("* Elements with Mark (to be numbered): **{}**".format(len(marked_elements)))
output.print_md("* Elements updated (Comments set): **{}**".format(updated))
if skipped_no_mark:
    output.print_md("* Skipped (no Mark): **{}**".format(len(skipped_no_mark)))
if skipped_no_comments:
    output.print_md("* Skipped (no Comments param or read-only): **{}**".format(len(skipped_no_comments)))
if failed:
    output.print_md("* Failed to update: **{}**".format(len(failed)))
    for f in failed[:20]:
        elem, msg = f
        output.print_md("  * Id {} : {}".format(getattr(elem, "Id", "<unknown>"), msg))
