# -*- coding: utf-8 -*-
__title__ = "Export\nSheets"
__doc__ = "Export sheets to PDF/DWG with modern UI"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from System.Collections.Generic import List
from pyrevit import revit, script, forms

import os
import System
from System.Windows.Forms import (
    Form, Label, CheckedListBox, Button, RadioButton, TextBox,
    FolderBrowserDialog, DialogResult, ComboBox, FormBorderStyle,
    FlatStyle, BorderStyle
)
from System.Drawing import Color, Font, FontStyle, Size, Point

doc = revit.doc
output = script.get_output()

# ==============================================================================
# COLLECT SHEETS
# ==============================================================================

sheets = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Sheets) \
    .WhereElementIsNotElementType() \
    .ToElements()

if not sheets:
    forms.alert("No sheets found in the project.", exitscript=True)

sheets = sorted(sheets, key=lambda s: s.SheetNumber)

# ==============================================================================
# UI CLASS
# ==============================================================================

class ModernSheetExporter(Form):
    def __init__(self, sheets, doc):
        self.sheets = sheets
        self.doc = doc
        self.export_path = r"C:\Temp\Sheets"
        self.selected_sheets = []
        self.dwg_setups = {}
        self.InitializeUI()

    def InitializeUI(self):
        self.Text = "SHEET EXPORTER"
        self.Width = 1000
        self.Height = 750
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.BackColor = Color.FromArgb(20, 20, 28)
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

        # ---------------- SHEET LIST ----------------
        list_label = Label()
        list_label.Text = "SELECT SHEETS:"
        list_label.Location = Point(20, 0)
        list_label.Size = Size(200, 20)
        list_label.ForeColor = Color.White
        list_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(list_label)

        self.sheet_list = CheckedListBox()
        self.sheet_list.Location = Point(20, 25)
        self.sheet_list.Size = Size(450, 600)
        self.sheet_list.Font = Font("Consolas", 9)
        self.sheet_list.BackColor = Color.FromArgb(25, 25, 35)
        self.sheet_list.ForeColor = Color.White
        self.sheet_list.CheckOnClick = True
        self.sheet_list.BorderStyle = BorderStyle.FixedSingle

        for s in self.sheets:
            self.sheet_list.Items.Add("{} - {}".format(s.SheetNumber, s.Name))

        self.Controls.Add(self.sheet_list)

        # ---------------- SELECT/DESELECT BUTTONS ----------------
        select_all = Button(Text="Select All", Location=Point(20, 635))
        select_all.Size = Size(100, 25)
        select_all.BackColor = Color.FromArgb(60, 60, 70)
        select_all.ForeColor = Color.White
        select_all.Click += self.SelectAll
        self.Controls.Add(select_all)

        deselect_all = Button(Text="Deselect All", Location=Point(130, 635))
        deselect_all.Size = Size(100, 25)
        deselect_all.BackColor = Color.FromArgb(60, 60, 70)
        deselect_all.ForeColor = Color.White
        deselect_all.Click += self.DeselectAll
        self.Controls.Add(deselect_all)

        # ---------------- EXPORT OPTIONS ----------------
        options_label = Label()
        options_label.Text = "EXPORT FORMAT:"
        options_label.Location = Point(500, 0)
        options_label.Size = Size(200, 20)
        options_label.ForeColor = Color.White
        options_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(options_label)

        self.pdf_radio = RadioButton()
        self.pdf_radio.Text = "PDF Only"
        self.pdf_radio.Location = Point(500, 30)
        self.pdf_radio.Size = Size(150, 25)
        self.pdf_radio.Checked = True
        self.pdf_radio.ForeColor = Color.White
        self.pdf_radio.Font = Font("Segoe UI", 9)
        self.Controls.Add(self.pdf_radio)

        self.dwg_radio = RadioButton()
        self.dwg_radio.Text = "DWG Only"
        self.dwg_radio.Location = Point(500, 60)
        self.dwg_radio.Size = Size(150, 25)
        self.dwg_radio.ForeColor = Color.White
        self.dwg_radio.Font = Font("Segoe UI", 9)
        self.dwg_radio.CheckedChanged += self.OnFormatChanged
        self.Controls.Add(self.dwg_radio)

        self.both_radio = RadioButton()
        self.both_radio.Text = "PDF + DWG"
        self.both_radio.Location = Point(500, 90)
        self.both_radio.Size = Size(150, 25)
        self.both_radio.ForeColor = Color.White
        self.both_radio.Font = Font("Segoe UI", 9)
        self.both_radio.CheckedChanged += self.OnFormatChanged
        self.Controls.Add(self.both_radio)

        # ---------------- DWG SETTINGS DROPDOWN ----------------
        dwg_label = Label()
        dwg_label.Text = "DWG EXPORT SETUP:"
        dwg_label.Location = Point(500, 130)
        dwg_label.Size = Size(200, 20)
        dwg_label.ForeColor = Color.White
        dwg_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(dwg_label)

        self.dwg_combo = ComboBox()
        self.dwg_combo.Location = Point(500, 155)
        self.dwg_combo.Size = Size(350, 25)
        self.dwg_combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.dwg_combo.BackColor = Color.FromArgb(25, 25, 35)
        self.dwg_combo.ForeColor = Color.White
        self.dwg_combo.Enabled = False
        self.Controls.Add(self.dwg_combo)

        # ---------------- PATH ----------------
        path_label = Label()
        path_label.Text = "EXPORT LOCATION:"
        path_label.Location = Point(500, 200)
        path_label.Size = Size(200, 20)
        path_label.ForeColor = Color.White
        path_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(path_label)

        self.path_box = TextBox()
        self.path_box.Text = self.export_path
        self.path_box.Location = Point(500, 225)
        self.path_box.Size = Size(350, 25)
        self.path_box.ReadOnly = True
        self.path_box.BackColor = Color.FromArgb(25, 25, 35)
        self.path_box.ForeColor = Color.White
        self.Controls.Add(self.path_box)

        browse = Button()
        browse.Text = "Browse"
        browse.Location = Point(860, 223)
        browse.Size = Size(100, 27)
        browse.Click += self.Browse
        browse.BackColor = Color.FromArgb(74, 144, 226)
        browse.ForeColor = Color.White
        browse.FlatStyle = FlatStyle.Flat
        self.Controls.Add(browse)

        # ---------------- EXPORT BUTTON ----------------
        export_btn = Button()
        export_btn.Text = "EXPORT SHEETS"
        export_btn.Location = Point(500, 275)
        export_btn.Size = Size(460, 45)
        export_btn.BackColor = Color.FromArgb(46, 204, 113)
        export_btn.ForeColor = Color.White
        export_btn.FlatStyle = FlatStyle.Flat
        export_btn.Font = Font("Segoe UI", 11, FontStyle.Bold)
        export_btn.Click += self.ExportSheets
        self.Controls.Add(export_btn)

        # ---------------- LOG ----------------
        log_label = Label()
        log_label.Text = "EXPORT LOG:"
        log_label.Location = Point(500, 335)
        log_label.Size = Size(200, 20)
        log_label.ForeColor = Color.White
        log_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(log_label)

        self.log = TextBox()
        self.log.Location = Point(500, 360)
        self.log.Size = Size(460, 300)
        self.log.Multiline = True
        self.log.ReadOnly = True
        self.log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        self.log.BackColor = Color.FromArgb(15, 15, 20)
        self.log.ForeColor = Color.FromArgb(200, 200, 200)
        self.log.Font = Font("Consolas", 8)
        self.log.BorderStyle = BorderStyle.FixedSingle
        self.Controls.Add(self.log)
        
        # Load DWG setups AFTER log is created
        self.LoadDWGSetups()

    def LoadDWGSetups(self):
        """Load available DWG export setups"""
        try:
            predefined_setups = DWGExportOptions.GetPredefinedOptions(self.doc)
            
            if predefined_setups:
                for setup_name in predefined_setups.Keys:
                    self.dwg_setups[setup_name] = predefined_setups[setup_name]
                    self.dwg_combo.Items.Add(setup_name)
                
                if self.dwg_combo.Items.Count > 0:
                    self.dwg_combo.SelectedIndex = 0
                    self.Log("✓ Loaded {} DWG export setup(s)".format(self.dwg_combo.Items.Count))
            else:
                self.dwg_combo.Items.Add("No DWG setups available")
                self.dwg_combo.SelectedIndex = 0
                self.Log("⚠ No DWG export setups found in project")
        except Exception as e:
            self.Log("✗ Error loading DWG setups: {}".format(str(e)))

    def OnFormatChanged(self, sender, args):
        """Enable/disable DWG combo based on format selection"""
        needs_dwg = self.dwg_radio.Checked or self.both_radio.Checked
        self.dwg_combo.Enabled = needs_dwg

    def SelectAll(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, True)

    def DeselectAll(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, False)

    def Browse(self, sender, args):
        dlg = FolderBrowserDialog()
        dlg.SelectedPath = self.export_path
        if dlg.ShowDialog() == DialogResult.OK:
            self.export_path = dlg.SelectedPath
            self.path_box.Text = self.export_path

    def Log(self, msg):
        """Add message to log - safely handles case where log doesn't exist yet"""
        if hasattr(self, 'log') and self.log:
            self.log.AppendText(msg + "\r\n")
            self.log.SelectionStart = self.log.Text.Length
            self.log.ScrollToCaret()

    # ==============================================================================
    # EXPORT LOGIC
    # ==============================================================================

    def ExportSheets(self, sender, args):
        self.selected_sheets = [
            self.sheets[i] for i in self.sheet_list.CheckedIndices
        ]

        if not self.selected_sheets:
            forms.alert("Please select at least one sheet to export.")
            return

        if not os.path.exists(self.export_path):
            try:
                os.makedirs(self.export_path)
                self.Log("✓ Created export directory: {}".format(self.export_path))
            except Exception as e:
                self.Log("✗ Failed to create directory: {}".format(str(e)))
                return

        export_pdf = self.pdf_radio.Checked or self.both_radio.Checked
        export_dwg = self.dwg_radio.Checked or self.both_radio.Checked

        pdf_count = 0
        dwg_count = 0

        self.log.Clear()
        self.Log("=" * 50)
        self.Log("EXPORT STARTED")
        self.Log("=" * 50)
        self.Log("Sheets to export: {}".format(len(self.selected_sheets)))
        self.Log("Export path: {}".format(self.export_path))
        self.Log("")

        # ---------------- PDF EXPORT ----------------
        try:
            if export_pdf:
                self.Log("=" * 50)
                self.Log("PDF EXPORT")
                self.Log("=" * 50)
                
                pdf_opts = PDFExportOptions()
                pdf_opts.CombineMultipleSelectedViewsSheets = False

                for sheet in self.selected_sheets:
                    try:
                        view_set = ViewSet()
                        view_set.Insert(sheet)

                        name = "{} - {}".format(sheet.SheetNumber, sheet.Name)
                        # Clean filename
                        for c in '<>:"|?*/\\':
                            name = name.replace(c, "_")
                        
                        # Remove any leading/trailing spaces
                        name = name.strip()

                        # Use the correct export method signature
                        result = self.doc.Export(self.export_path, name, view_set, pdf_opts)
                        
                        if result:
                            pdf_count += 1
                            self.Log("✓ PDF: {}".format(name))
                        else:
                            self.Log("⚠ PDF export returned False: {}".format(name))

                    except Exception as e:
                        self.Log("✗ PDF FAILED {}: {}".format(sheet.SheetNumber, str(e)))

                self.Log("")

            # ---------------- DWG EXPORT ----------------
            if export_dwg:
                self.Log("=" * 50)
                self.Log("DWG EXPORT")
                self.Log("=" * 50)

                if not self.dwg_setups:
                    self.Log("✗ No DWG export setups available")
                else:
                    selected_setup = str(self.dwg_combo.SelectedItem)
                    
                    if selected_setup in self.dwg_setups:
                        dwg_opts = self.dwg_setups[selected_setup]
                        self.Log("Using DWG setup: {}".format(selected_setup))
                        self.Log("")

                        for sheet in self.selected_sheets:
                            try:
                                view_set = ViewSet()
                                view_set.Insert(sheet)

                                file_name = "{} - {}".format(sheet.SheetNumber, sheet.Name)
                                # Clean filename
                                for c in '<>:"|?*/\\':
                                    file_name = file_name.replace(c, "_")
                                
                                # Remove any leading/trailing spaces
                                file_name = file_name.strip()

                                result = self.doc.Export(self.export_path, file_name, view_set, dwg_opts)
                                
                                if result:
                                    dwg_count += 1
                                    self.Log("✓ DWG: {}".format(file_name))
                                else:
                                    self.Log("⚠ DWG export returned False: {}".format(file_name))

                            except Exception as e:
                                self.Log("✗ DWG FAILED {}: {}".format(sheet.SheetNumber, str(e)))
                    else:
                        self.Log("✗ Selected DWG setup not found")

                self.Log("")

        except Exception as e:
            self.Log("✗ EXPORT ERROR: {}".format(str(e)))

        # Final Summary
        self.Log("=" * 50)
        self.Log("EXPORT COMPLETE")
        self.Log("=" * 50)
        if export_pdf:
            self.Log("PDFs exported: {} / {}".format(pdf_count, len(self.selected_sheets)))
        if export_dwg:
            self.Log("DWGs exported: {} / {}".format(dwg_count, len(self.selected_sheets)))
        self.Log("Location: {}".format(self.export_path))


# ==============================================================================
# RUN
# ==============================================================================

form = ModernSheetExporter(sheets, doc)
form.ShowDialog()