# -*- coding: utf-8 -*-
"""
Revision Manager - Modern UI
Set selected revisions on selected sheets with advanced filtering
"""
__title__ = 'Revision\nManager'
__doc__ = 'Assign revisions to sheets with modern interface'

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
import System
from System.Windows.Forms import (
    Form, Label, CheckedListBox, Button, ComboBox, GroupBox,
    RadioButton, Panel, TextBox,
    FormBorderStyle, ComboBoxStyle, FlatStyle, BorderStyle
)
from System.Drawing import (
    Color, Font, FontStyle, Size, Point
)
from System.Collections.Generic import List as DotNetList

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ==============================================================================
# COLLECT REVISIONS AND SHEETS
# ==============================================================================

def get_all_revisions():
    """Get all revisions in the project"""
    all_revisions = FilteredElementCollector(doc)\
        .OfClass(Revision)\
        .ToElements()
    
    # Sort by sequence number
    revisions_sorted = sorted(all_revisions, key=lambda r: r.SequenceNumber)
    
    return revisions_sorted


def get_all_sheets():
    """Get all sheets in the project"""
    sheets = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Sheets)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Sort by sheet number
    sheets_sorted = sorted(sheets, key=lambda s: s.SheetNumber)
    
    return sheets_sorted


def get_sheet_revisions(sheet):
    """Get revisions assigned to a sheet"""
    try:
        rev_ids = list(sheet.GetAdditionalRevisionIds())
        revisions = [doc.GetElement(rev_id) for rev_id in rev_ids]
        return [r for r in revisions if r is not None]
    except:
        return []


# Collect data
all_revisions = get_all_revisions()
all_sheets = get_all_sheets()

if not all_revisions:
    forms.alert('No revisions found in the project', exitscript=True)

if not all_sheets:
    forms.alert('No sheets found in the project', exitscript=True)

output.print_md("**Found {} revisions**".format(len(all_revisions)))
output.print_md("**Found {} sheets**".format(len(all_sheets)))


# ==============================================================================
# MODERN REVISION MANAGER UI
# ==============================================================================

class RevisionManager(Form):
    def __init__(self, revisions, sheets, document):
        self.revisions = revisions
        self.sheets = sheets
        self.doc = document
        self.selected_sheets = []
        self.selected_revisions = []
        self.result = None
        
        self.InitializeUI()
    
    def InitializeUI(self):
        # Form properties
        self.Text = "REVISION MANAGER v2.5.0"
        self.Width = 1200
        self.Height = 750
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.BackColor = Color.FromArgb(20, 20, 28)
        self.ForeColor = Color.FromArgb(220, 220, 220)
        
        # Fonts
        self.title_font = Font("Segoe UI", 14, FontStyle.Bold)
        self.label_font = Font("Consolas", 9, FontStyle.Regular)
        self.input_font = Font("Consolas", 10, FontStyle.Regular)
        
        # Colors
        self.bg_panel = Color.FromArgb(30, 30, 40)
        self.accent_color = Color.FromArgb(74, 144, 226)
        self.remove_color = Color.FromArgb(200, 60, 60)
        self.text_color = Color.FromArgb(220, 220, 220)
        self.text_muted = Color.FromArgb(140, 140, 150)
        
        y_offset = 20
        
        # ===== HEADER =====
        header_panel = Panel()
        header_panel.Location = Point(20, y_offset)
        header_panel.Size = Size(1140, 80)
        header_panel.BackColor = self.bg_panel
        self.Controls.Add(header_panel)
        
        title = Label()
        title.Text = "REVISION MANAGER"
        title.Font = self.title_font
        title.ForeColor = Color.White
        title.Location = Point(20, 15)
        title.Size = Size(500, 30)
        header_panel.Controls.Add(title)
        
        subtitle = Label()
        subtitle.Text = "ASSIGN OR REMOVE REVISIONS FROM SHEETS"
        subtitle.Font = Font("Consolas", 8, FontStyle.Regular)
        subtitle.ForeColor = self.text_muted
        subtitle.Location = Point(20, 45)
        subtitle.Size = Size(500, 20)
        header_panel.Controls.Add(subtitle)
        
        y_offset += 100
        
        # ===== LEFT PANEL - REVISIONS =====
        left_x = 20
        
        rev_label = Label()
        rev_label.Text = "STEP 1: SELECT REVISIONS"
        rev_label.Font = self.label_font
        rev_label.ForeColor = self.text_muted
        rev_label.Location = Point(left_x, y_offset)
        rev_label.Size = Size(300, 20)
        self.Controls.Add(rev_label)
        
        # Filter options
        filter_group = GroupBox()
        filter_group.Text = "FILTER"
        filter_group.Location = Point(left_x + 300, y_offset - 5)
        filter_group.Size = Size(260, 50)
        filter_group.Font = Font("Consolas", 8, FontStyle.Regular)
        filter_group.ForeColor = self.text_muted
        filter_group.BackColor = self.bg_panel
        self.Controls.Add(filter_group)
        
        self.all_rev_radio = RadioButton()
        self.all_rev_radio.Text = "All"
        self.all_rev_radio.Location = Point(15, 20)
        self.all_rev_radio.Size = Size(60, 20)
        self.all_rev_radio.Font = Font("Consolas", 9, FontStyle.Regular)
        self.all_rev_radio.ForeColor = self.text_color
        self.all_rev_radio.Checked = True
        self.all_rev_radio.CheckedChanged += self.OnRevisionFilterChanged
        filter_group.Controls.Add(self.all_rev_radio)
        
        self.unissued_radio = RadioButton()
        self.unissued_radio.Text = "Unissued Only"
        self.unissued_radio.Location = Point(80, 20)
        self.unissued_radio.Size = Size(120, 20)
        self.unissued_radio.Font = Font("Consolas", 9, FontStyle.Regular)
        self.unissued_radio.ForeColor = self.text_color
        self.unissued_radio.CheckedChanged += self.OnRevisionFilterChanged
        filter_group.Controls.Add(self.unissued_radio)
        
        # Select buttons
        select_all_rev_btn = Button()
        select_all_rev_btn.Text = "All"
        select_all_rev_btn.Location = Point(left_x + 370, y_offset + 28)
        select_all_rev_btn.Size = Size(60, 24)
        select_all_rev_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        select_all_rev_btn.BackColor = self.accent_color
        select_all_rev_btn.ForeColor = Color.White
        select_all_rev_btn.FlatStyle = FlatStyle.Flat
        select_all_rev_btn.FlatAppearance.BorderSize = 0
        select_all_rev_btn.Click += self.SelectAllRevisions
        self.Controls.Add(select_all_rev_btn)
        
        clear_all_rev_btn = Button()
        clear_all_rev_btn.Text = "Clear"
        clear_all_rev_btn.Location = Point(left_x + 435, y_offset + 28)
        clear_all_rev_btn.Size = Size(60, 24)
        clear_all_rev_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        clear_all_rev_btn.BackColor = Color.FromArgb(80, 80, 100)
        clear_all_rev_btn.ForeColor = Color.White
        clear_all_rev_btn.FlatStyle = FlatStyle.Flat
        clear_all_rev_btn.FlatAppearance.BorderSize = 0
        clear_all_rev_btn.Click += self.ClearAllRevisions
        self.Controls.Add(clear_all_rev_btn)
        
        y_offset += 60
        
        # Revision CheckedListBox
        self.revision_list = CheckedListBox()
        self.revision_list.Location = Point(left_x, y_offset)
        self.revision_list.Size = Size(540, 280)
        self.revision_list.Font = Font("Consolas", 9, FontStyle.Regular)
        self.revision_list.BackColor = Color.FromArgb(15, 15, 20)
        self.revision_list.ForeColor = self.text_color
        self.revision_list.BorderStyle = BorderStyle.FixedSingle
        self.revision_list.CheckOnClick = True
        
        self.PopulateRevisionList()
        
        self.Controls.Add(self.revision_list)
        
        y_offset += 290
        
        # Revision count
        self.revision_count_label = Label()
        self.revision_count_label.Text = "Selected: 0 revisions"
        self.revision_count_label.Font = Font("Consolas", 9, FontStyle.Regular)
        self.revision_count_label.ForeColor = self.text_muted
        self.revision_count_label.Location = Point(left_x, y_offset)
        self.revision_count_label.Size = Size(540, 20)
        self.Controls.Add(self.revision_count_label)
        
        self.revision_list.ItemCheck += self.UpdateRevisionCount
        
        # ===== RIGHT PANEL - SHEETS =====
        right_x = 580
        y_offset = 120
        
        sheet_label = Label()
        sheet_label.Text = "STEP 2: SELECT SHEETS"
        sheet_label.Font = self.label_font
        sheet_label.ForeColor = self.text_muted
        sheet_label.Location = Point(right_x, y_offset)
        sheet_label.Size = Size(300, 20)
        self.Controls.Add(sheet_label)
        
        # Sheet filter
        filter_label = Label()
        filter_label.Text = "Filter:"
        filter_label.Font = Font("Consolas", 8, FontStyle.Regular)
        filter_label.ForeColor = self.text_muted
        filter_label.Location = Point(right_x + 300, y_offset + 2)
        filter_label.Size = Size(50, 20)
        self.Controls.Add(filter_label)
        
        self.sheet_filter_textbox = TextBox()
        self.sheet_filter_textbox.Location = Point(right_x + 350, y_offset)
        self.sheet_filter_textbox.Size = Size(210, 20)
        self.sheet_filter_textbox.Font = Font("Consolas", 9, FontStyle.Regular)
        self.sheet_filter_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.sheet_filter_textbox.ForeColor = self.text_color
        self.sheet_filter_textbox.TextChanged += self.OnSheetFilterChanged
        self.Controls.Add(self.sheet_filter_textbox)
        
        # Sheet select buttons
        select_all_sheet_btn = Button()
        select_all_sheet_btn.Text = "All"
        select_all_sheet_btn.Location = Point(right_x + 300, y_offset + 28)
        select_all_sheet_btn.Size = Size(60, 24)
        select_all_sheet_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        select_all_sheet_btn.BackColor = self.accent_color
        select_all_sheet_btn.ForeColor = Color.White
        select_all_sheet_btn.FlatStyle = FlatStyle.Flat
        select_all_sheet_btn.FlatAppearance.BorderSize = 0
        select_all_sheet_btn.Click += self.SelectAllSheets
        self.Controls.Add(select_all_sheet_btn)
        
        clear_all_sheet_btn = Button()
        clear_all_sheet_btn.Text = "Clear"
        clear_all_sheet_btn.Location = Point(right_x + 365, y_offset + 28)
        clear_all_sheet_btn.Size = Size(60, 24)
        clear_all_sheet_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        clear_all_sheet_btn.BackColor = Color.FromArgb(80, 80, 100)
        clear_all_sheet_btn.ForeColor = Color.White
        clear_all_sheet_btn.FlatStyle = FlatStyle.Flat
        clear_all_sheet_btn.FlatAppearance.BorderSize = 0
        clear_all_sheet_btn.Click += self.ClearAllSheets
        self.Controls.Add(clear_all_sheet_btn)
        
        filtered_btn = Button()
        filtered_btn.Text = "Filtered"
        filtered_btn.Location = Point(right_x + 430, y_offset + 28)
        filtered_btn.Size = Size(90, 24)
        filtered_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        filtered_btn.BackColor = Color.FromArgb(60, 100, 140)
        filtered_btn.ForeColor = Color.White
        filtered_btn.FlatStyle = FlatStyle.Flat
        filtered_btn.FlatAppearance.BorderSize = 0
        filtered_btn.Click += self.SelectFiltered
        self.Controls.Add(filtered_btn)
        
        y_offset += 60
        
        # Sheet CheckedListBox
        self.sheet_list = CheckedListBox()
        self.sheet_list.Location = Point(right_x, y_offset)
        self.sheet_list.Size = Size(560, 280)
        self.sheet_list.Font = Font("Consolas", 9, FontStyle.Regular)
        self.sheet_list.BackColor = Color.FromArgb(15, 15, 20)
        self.sheet_list.ForeColor = self.text_color
        self.sheet_list.BorderStyle = BorderStyle.FixedSingle
        self.sheet_list.CheckOnClick = True
        
        # Add sheets
        for sheet in self.sheets:
            display_text = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            self.sheet_list.Items.Add(display_text)
        
        self.Controls.Add(self.sheet_list)
        
        y_offset += 290
        
        # Sheet count
        self.sheet_count_label = Label()
        self.sheet_count_label.Text = "Selected: 0 / {}".format(len(self.sheets))
        self.sheet_count_label.Font = Font("Consolas", 9, FontStyle.Regular)
        self.sheet_count_label.ForeColor = self.text_muted
        self.sheet_count_label.Location = Point(right_x, y_offset)
        self.sheet_count_label.Size = Size(560, 20)
        self.Controls.Add(self.sheet_count_label)
        
        self.sheet_list.ItemCheck += self.UpdateSheetCount
        self.sheet_list.ItemCheck += self.UpdatePreview
        
        # ===== BOTTOM - STATUS AND APPLY =====
        y_offset = 590
        
        # Status Panel
        status_group = GroupBox()
        status_group.Text = "PREVIEW & STATUS"
        status_group.Location = Point(20, y_offset)
        status_group.Size = Size(920, 100)
        status_group.Font = self.label_font
        status_group.ForeColor = self.text_muted
        status_group.BackColor = self.bg_panel
        self.Controls.Add(status_group)
        
        self.status_text = TextBox()
        self.status_text.Location = Point(15, 25)
        self.status_text.Size = Size(890, 65)
        self.status_text.Font = Font("Consolas", 9, FontStyle.Regular)
        self.status_text.BackColor = Color.FromArgb(10, 10, 15)
        self.status_text.ForeColor = self.text_muted
        self.status_text.BorderStyle = BorderStyle.None
        self.status_text.ReadOnly = True
        self.status_text.Multiline = True
        self.status_text.WordWrap = True
        self.status_text.Text = "Select revisions and sheets to see preview..."
        status_group.Controls.Add(self.status_text)
        
        # Apply Button
        self.apply_btn = Button()
        self.apply_btn.Text = "APPLY REVISIONS"
        self.apply_btn.Location = Point(960, y_offset + 10)
        self.apply_btn.Size = Size(200, 38)
        self.apply_btn.Font = Font("Consolas", 9, FontStyle.Bold)
        self.apply_btn.BackColor = self.accent_color
        self.apply_btn.ForeColor = Color.White
        self.apply_btn.FlatStyle = FlatStyle.Flat
        self.apply_btn.FlatAppearance.BorderSize = 0
        self.apply_btn.Click += self.ApplyRevisions
        self.Controls.Add(self.apply_btn)
        
        # Remove Button
        self.remove_btn = Button()
        self.remove_btn.Text = "REMOVE REVISIONS"
        self.remove_btn.Location = Point(960, y_offset + 52)
        self.remove_btn.Size = Size(200, 38)
        self.remove_btn.Font = Font("Consolas", 9, FontStyle.Bold)
        self.remove_btn.BackColor = self.remove_color
        self.remove_btn.ForeColor = Color.White
        self.remove_btn.FlatStyle = FlatStyle.Flat
        self.remove_btn.FlatAppearance.BorderSize = 0
        self.remove_btn.Click += self.RemoveRevisions
        self.Controls.Add(self.remove_btn)
    
    def PopulateRevisionList(self):
        """Populate revision list based on filter"""
        self.revision_list.Items.Clear()
        
        for rev in self.revisions:
            # Filter logic
            if self.unissued_radio.Checked and rev.Issued:
                continue
            
            issued_text = "" if not rev.Issued else " [ISSUED]"
            display_text = "Rev {} - {}{}".format(
                rev.SequenceNumber,
                rev.Description or "(No Description)",
                issued_text
            )
            self.revision_list.Items.Add(display_text)
    
    def OnRevisionFilterChanged(self, sender, args):
        """Refresh revision list when filter changes"""
        self.PopulateRevisionList()
        self.UpdateRevisionCount(None, None)
    
    def OnSheetFilterChanged(self, sender, args):
        """Filter sheets based on text input"""
        filter_text = self.sheet_filter_textbox.Text.lower()
        
        self.sheet_list.Items.Clear()
        
        for sheet in self.sheets:
            display_text = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            if not filter_text or filter_text in display_text.lower():
                self.sheet_list.Items.Add(display_text)
    
    def SelectAllRevisions(self, sender, args):
        for i in range(self.revision_list.Items.Count):
            self.revision_list.SetItemChecked(i, True)
    
    def ClearAllRevisions(self, sender, args):
        for i in range(self.revision_list.Items.Count):
            self.revision_list.SetItemChecked(i, False)
    
    def SelectAllSheets(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, True)
    
    def ClearAllSheets(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, False)
    
    def SelectFiltered(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, True)
    
    def UpdateRevisionCount(self, sender, args):
        if sender:
            timer = System.Windows.Forms.Timer()
            timer.Interval = 10
            timer.Tick += lambda s, e: self.DoUpdateRevCount(s, e)
            timer.Start()
        else:
            self.DoUpdateRevCount(None, None)
    
    def DoUpdateRevCount(self, sender, args):
        if sender:
            sender.Stop()
        count = self.revision_list.CheckedItems.Count
        self.revision_count_label.Text = "Selected: {} revisions".format(count)
    
    def UpdateSheetCount(self, sender, args):
        timer = System.Windows.Forms.Timer()
        timer.Interval = 10
        timer.Tick += lambda s, e: self.DoUpdateSheetCount(s, e)
        timer.Start()
    
    def DoUpdateSheetCount(self, sender, args):
        sender.Stop()
        count = self.sheet_list.CheckedItems.Count
        self.sheet_count_label.Text = "Selected: {} / {}".format(count, len(self.sheets))
    
    def UpdatePreview(self, sender, args):
        """Update preview when selection changes"""
        timer = System.Windows.Forms.Timer()
        timer.Interval = 10
        timer.Tick += lambda s, e: self.DoUpdatePreview(s, e)
        timer.Start()
    
    def DoUpdatePreview(self, sender, args):
        sender.Stop()
        
        rev_count = self.revision_list.CheckedItems.Count
        sheet_count = self.sheet_list.CheckedItems.Count
        
        if rev_count == 0 or sheet_count == 0:
            self.status_text.Text = "Select revisions and sheets to see preview..."
            return
        
        preview_text = "Ready to apply {} revision(s) to {} sheet(s)\n\n".format(
            rev_count,
            sheet_count
        )
        
        preview_text += "Revisions:\n"
        for i in range(min(rev_count, 3)):
            preview_text += "  • {}\n".format(self.revision_list.CheckedItems[i])
        if rev_count > 3:
            preview_text += "  ... and {} more\n".format(rev_count - 3)
        
        preview_text += "\nSheets:\n"
        for i in range(min(sheet_count, 3)):
            preview_text += "  • {}\n".format(self.sheet_list.CheckedItems[i])
        if sheet_count > 3:
            preview_text += "  ... and {} more".format(sheet_count - 3)
        
        self.status_text.Text = preview_text
    
    def AddLog(self, message):
        log_line = "{}\r\n".format(message)
        self.status_text.AppendText(log_line)
        self.status_text.SelectionStart = len(self.status_text.Text)
        self.status_text.ScrollToCaret()
        self.Refresh()
    
    def ApplyRevisions(self, sender, args):
        # Get selected revisions
        selected_rev_indices = []
        for i in range(self.revision_list.CheckedItems.Count):
            selected_rev_indices.append(self.revision_list.CheckedIndices[i])
        
        if len(selected_rev_indices) == 0:
            forms.alert("Please select at least one revision", title="Validation Error")
            return
        
        # Get selected sheets
        selected_sheet_indices = []
        for i in range(self.sheet_list.CheckedItems.Count):
            selected_sheet_indices.append(self.sheet_list.CheckedIndices[i])
        
        if len(selected_sheet_indices) == 0:
            forms.alert("Please select at least one sheet", title="Validation Error")
            return
        
        # Map indices to actual objects
        display_revisions = [r for r in self.revisions if not self.unissued_radio.Checked or not r.Issued]
        self.selected_revisions = [display_revisions[i] for i in selected_rev_indices]
        
        # For sheets, need to map from filtered list back to original
        filter_text = self.sheet_filter_textbox.Text.lower()
        filtered_sheets = []
        for sheet in self.sheets:
            display_text = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            if not filter_text or filter_text in display_text.lower():
                filtered_sheets.append(sheet)
        
        self.selected_sheets = [filtered_sheets[i] for i in selected_sheet_indices]
        
        # Confirm
        confirm_msg = "Apply {} revision(s) to {} sheet(s)?\n\nContinue?".format(
            len(self.selected_revisions),
            len(self.selected_sheets)
        )
        
        if not forms.alert(confirm_msg, yes=True, no=True):
            return
        
        # Update UI
        self.apply_btn.Enabled = False
        self.apply_btn.Text = "APPLYING..."
        self.status_text.Clear()
        
        self.AddLog("Starting revision assignment...")
        self.AddLog("Revisions: {}".format(len(self.selected_revisions)))
        self.AddLog("Sheets: {}".format(len(self.selected_sheets)))
        self.AddLog("")
        
        success_count = 0
        error_count = 0
        
        with revit.Transaction("Apply Revisions to Sheets"):
            for sheet in self.selected_sheets:
                try:
                    # Get existing revisions
                    rev_ids = list(sheet.GetAdditionalRevisionIds())
                    
                    # Add new revisions if not already present
                    for rev in self.selected_revisions:
                        if rev.Id not in rev_ids:
                            rev_ids.append(rev.Id)
                    
                    # Set revisions
                    sheet.SetAdditionalRevisionIds(DotNetList[ElementId](rev_ids))
                    
                    success_count += 1
                    self.AddLog("✓ {} - {}".format(sheet.SheetNumber, sheet.Name))
                    
                except Exception as e:
                    error_count += 1
                    self.AddLog("✗ {} - {}: {}".format(
                        sheet.SheetNumber,
                        sheet.Name,
                        str(e)
                    ))
        
        # Summary
        self.AddLog("")
        self.AddLog("=== COMPLETE ===")
        self.AddLog("Success: {}".format(success_count))
        self.AddLog("Errors: {}".format(error_count))
        
        msg = "Revision assignment complete!\n\n"
        msg += "Success: {}\n".format(success_count)
        msg += "Errors: {}".format(error_count)
        
        forms.alert(msg, title="Complete")
        
        self.result = True
        self.apply_btn.Enabled = True
        self.apply_btn.Text = "APPLY REVISIONS"
    
    def RemoveRevisions(self, sender, args):
        """Remove selected revisions from selected sheets"""
        # Get selected revisions
        selected_rev_indices = []
        for i in range(self.revision_list.CheckedItems.Count):
            selected_rev_indices.append(self.revision_list.CheckedIndices[i])
        
        if len(selected_rev_indices) == 0:
            forms.alert("Please select at least one revision to remove", title="Validation Error")
            return
        
        # Get selected sheets
        selected_sheet_indices = []
        for i in range(self.sheet_list.CheckedItems.Count):
            selected_sheet_indices.append(self.sheet_list.CheckedIndices[i])
        
        if len(selected_sheet_indices) == 0:
            forms.alert("Please select at least one sheet", title="Validation Error")
            return
        
        # Map indices to actual objects
        display_revisions = [r for r in self.revisions if not self.unissued_radio.Checked or not r.Issued]
        self.selected_revisions = [display_revisions[i] for i in selected_rev_indices]
        
        # For sheets, need to map from filtered list back to original
        filter_text = self.sheet_filter_textbox.Text.lower()
        filtered_sheets = []
        for sheet in self.sheets:
            display_text = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            if not filter_text or filter_text in display_text.lower():
                filtered_sheets.append(sheet)
        
        self.selected_sheets = [filtered_sheets[i] for i in selected_sheet_indices]
        
        # Confirm
        confirm_msg = "Remove {} revision(s) from {} sheet(s)?\n\nThis will only remove the selected revisions.\nContinue?".format(
            len(self.selected_revisions),
            len(self.selected_sheets)
        )
        
        if not forms.alert(confirm_msg, yes=True, no=True):
            return
        
        # Update UI
        self.remove_btn.Enabled = False
        self.remove_btn.Text = "REMOVING..."
        self.status_text.Clear()
        
        self.AddLog("Starting revision removal...")
        self.AddLog("Revisions: {}".format(len(self.selected_revisions)))
        self.AddLog("Sheets: {}".format(len(self.selected_sheets)))
        self.AddLog("")
        
        success_count = 0
        error_count = 0
        skipped_count = 0
        
        # Get IDs of revisions to remove
        revisions_to_remove = [rev.Id for rev in self.selected_revisions]
        
        with revit.Transaction("Remove Revisions from Sheets"):
            for sheet in self.selected_sheets:
                try:
                    # Get existing revisions
                    rev_ids = list(sheet.GetAdditionalRevisionIds())
                    original_count = len(rev_ids)
                    
                    # Remove selected revisions
                    rev_ids = [rid for rid in rev_ids if rid not in revisions_to_remove]
                    
                    removed_count = original_count - len(rev_ids)
                    
                    if removed_count == 0:
                        skipped_count += 1
                        self.AddLog("○ {} - {} (no revisions to remove)".format(
                            sheet.SheetNumber, 
                            sheet.Name
                        ))
                    else:
                        # Set updated revisions
                        sheet.SetAdditionalRevisionIds(DotNetList[ElementId](rev_ids))
                        
                        success_count += 1
                        self.AddLog("✓ {} - {} ({} removed)".format(
                            sheet.SheetNumber,
                            sheet.Name,
                            removed_count
                        ))
                    
                except Exception as e:
                    error_count += 1
                    self.AddLog("✗ {} - {}: {}".format(
                        sheet.SheetNumber,
                        sheet.Name,
                        str(e)
                    ))
        
        # Summary
        self.AddLog("")
        self.AddLog("=== COMPLETE ===")
        self.AddLog("Removed: {}".format(success_count))
        self.AddLog("Skipped: {}".format(skipped_count))
        self.AddLog("Errors: {}".format(error_count))
        
        msg = "Revision removal complete!\n\n"
        msg += "Removed: {}\n".format(success_count)
        msg += "Skipped: {}\n".format(skipped_count)
        msg += "Errors: {}".format(error_count)
        
        forms.alert(msg, title="Complete")
        
        self.result = True
        self.remove_btn.Enabled = True
        self.remove_btn.Text = "REMOVE REVISIONS"


# ==============================================================================
# SHOW FORM
# ==============================================================================

form = RevisionManager(all_revisions, all_sheets, doc)
form.ShowDialog()

if form.result:
    output.print_md("### ✅ Operation Completed Successfully")
else:
    output.print_md("### Operation Cancelled")