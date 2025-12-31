# -*- coding: utf-8 -*-
import clr
from pyrevit import revit, script

# Revit API
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Workset, FilteredWorksetCollector

output = script.get_output()
doc = revit.doc   # ✅ CORRECT for pyRevit

# -------------------------------------------------
# WORKSET NAMES
# -------------------------------------------------
workset_names = [
    "AR-Civil Works-GF",
    "AR-Civil Works-FF",
    "LINK-CAD-ST",
    "LINK-RVT-AR",
    "LINK-RVT-EL",
    "LINK-RVT-GN",
    "LINK-RVT-ID",
    "LINK-RVT-LA",
    "LINK-RVT-ME",
    "LINK-RVT-PL",
    "LINK-RVT-ST",
    "ST-Beam-FF",
    "ST-Beam-RF",
    "ST-BLOCK WALL-FF",
    "ST-BLOCK WALL-GF",
    "ST-BLOCK WALL-PA",
    "ST-CIP-FF",
    "ST-CIP-RF",
    "ST-Column-FF",
    "ST-Column-GF",
    "ST-Erection Mark-FF",
    "ST-Erection Mark-GF",
    "ST-Erection Mark-RF",
    "ST-Foundation",
    "ST-Hollowcore-FF",
    "ST-Hollowcore-RF",
    "ST-Reveals-FF",
    "ST-Reveals-GF",
    "ST-Reveals-RF",
    "ST-SLAB-FF",
    "ST-SLAB-GF",
    "ST-SLAB-RF",
    "ST-STAIR-FF",
    "ST-STAIR-GF",
    "ST-Stru Opngs-FF",
    "ST-Stru Opngs-RF",
    "ST-WALL-FF",
    "ST-WALL-GF",
    "ST-WALL-RF"
]

# -------------------------------------------------
# CHECK WORKSHARING
# -------------------------------------------------
if not doc.IsWorkshared:
    output.print_md("### ❌ Project is not workshared\nEnable Worksharing before running this tool.")
    script.exit()

# Existing worksets
existing = [ws.Name for ws in FilteredWorksetCollector(doc)]

created = []
skipped = []

# -------------------------------------------------
# CREATE WORKSETS
# -------------------------------------------------
with revit.Transaction("Create Standard Worksets"):
    for name in workset_names:
        if name in existing:
            skipped.append(name)
        else:
            Workset.Create(doc, name)
            created.append(name)

# -------------------------------------------------
# RESULT
# -------------------------------------------------
output.print_md("## ✅ Workset Creation Complete")

if created:
    output.print_md("### 🆕 Created")
    for ws in created:
        output.print_md("- " + ws)

if skipped:
    output.print_md("### ⚠️ Already Existing")
    for ws in skipped:
        output.print_md("- " + ws)
