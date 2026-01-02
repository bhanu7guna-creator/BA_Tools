# -*- coding: utf-8 -*-
import clr
import System
from pyrevit import revit, script, forms

# Revit API
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Workset, FilteredWorksetCollector

from System.Windows.Forms import (
    Form, Panel, Label, Button, CheckedListBox,
    FormBorderStyle, DockStyle
)
from System.Drawing import Color, Font, FontStyle, Size, Point

doc = revit.doc
output = script.get_output()

# -------------------------------------------------
# WORKSET MASTER LIST
# -------------------------------------------------
WORKSETS = [
    "AR-Civil Works-GF", "AR-Civil Works-FF",
    "LINK-CAD-ST", "LINK-RVT-AR", "LINK-RVT-EL",
    "LINK-RVT-GN", "LINK-RVT-ID", "LINK-RVT-LA",
    "LINK-RVT-ME", "LINK-RVT-PL", "LINK-RVT-ST",
    "ST-Beam-FF", "ST-Beam-RF",
    "ST-BLOCK WALL-FF", "ST-BLOCK WALL-GF", "ST-BLOCK WALL-PA",
    "ST-CIP-FF", "ST-CIP-RF",
    "ST-Column-FF", "ST-Column-GF",
    "ST-Erection Mark-FF", "ST-Erection Mark-GF", "ST-Erection Mark-RF",
    "ST-Foundation",
    "ST-Hollowcore-FF", "ST-Hollowcore-RF",
    "ST-Reveals-FF", "ST-Reveals-GF", "ST-Reveals-RF",
    "ST-SLAB-FF", "ST-SLAB-GF", "ST-SLAB-RF",
    "ST-STAIR-FF", "ST-STAIR-GF",
    "ST-Stru Opngs-FF", "ST-Stru Opngs-RF",
    "ST-WALL-FF", "ST-WALL-GF", "ST-WALL-RF"
]

# -------------------------------------------------
# CHECK WORKSHARING
# -------------------------------------------------
if not doc.IsWorkshared:
    forms.alert("Project is not workshared.", exitscript=True)

existing = [ws.Name for ws in FilteredWorksetCollector(doc)]
available = [w for w in WORKSETS if w not in existing]

if not available:
    forms.alert("All worksets already exist.", exitscript=True)

# -------------------------------------------------
# MODERN UI FORM
# -------------------------------------------------
class WorksetForm(Form):
    def __init__(self):
        self.selected = []
        self.InitUI()

    def InitUI(self):
        self.Text = "Create Worksets"
        self.Size = Size(520, 640)
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.BackColor = Color.FromArgb(18, 18, 24)
        self.MaximizeBox = False

        accent = Color.FromArgb(0, 120, 215)
        text = Color.FromArgb(230, 230, 235)
        panel_bg = Color.FromArgb(28, 28, 36)

        # Header
        header = Panel(
            Dock=DockStyle.Top,
            Height=80,
            BackColor=panel_bg
        )
        self.Controls.Add(header)

        Label(
            Text="WORKSET CREATOR",
            Font=Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor=text,
            Location=Point(20, 18),
            AutoSize=True,
            Parent=header
        )

        Label(
            Text="Select worksets to create",
            Font=Font("Segoe UI", 9),
            ForeColor=Color.FromArgb(160, 160, 170),
            Location=Point(20, 45),
            AutoSize=True,
            Parent=header
        )

        # List
        self.listbox = CheckedListBox()
        self.listbox.Location = Point(20, 100)
        self.listbox.Size = Size(460, 420)
        self.listbox.Font = Font("Segoe UI", 9)
        self.listbox.BackColor = panel_bg
        self.listbox.ForeColor = text
        self.listbox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle

        for w in available:
            self.listbox.Items.Add(w)

        self.Controls.Add(self.listbox)

        # Buttons
        def add_btn(text, x, handler, color):
            b = Button(
                Text=text,
                Location=Point(x, 540),
                Size=Size(140, 38),
                BackColor=color,
                ForeColor=Color.White,
                FlatStyle=System.Windows.Forms.FlatStyle.Flat
            )
            b.FlatAppearance.BorderSize = 0
            b.Click += handler
            self.Controls.Add(b)

        add_btn("Select All", 20, self.SelectAll, Color.FromArgb(60, 60, 70))
        add_btn("Clear", 180, self.ClearAll, Color.FromArgb(60, 60, 70))
        add_btn("Create", 340, self.Create, accent)

    def SelectAll(self, s, e):
        for i in range(self.listbox.Items.Count):
            self.listbox.SetItemChecked(i, True)

    def ClearAll(self, s, e):
        for i in range(self.listbox.Items.Count):
            self.listbox.SetItemChecked(i, False)

    def Create(self, s, e):
        self.selected = [self.listbox.Items[i] for i in range(self.listbox.Items.Count) if self.listbox.GetItemChecked(i)]
        self.DialogResult = System.Windows.Forms.DialogResult.OK
        self.Close()

# -------------------------------------------------
# RUN
# -------------------------------------------------
form = WorksetForm()
if form.ShowDialog() == System.Windows.Forms.DialogResult.OK:
    with revit.Transaction("Create Selected Worksets"):
        for w in form.selected:
            Workset.Create(doc, w)

    output.print_md("## ✅ Created Worksets")
    for w in form.selected:
        output.print_md("- {}".format(w))
