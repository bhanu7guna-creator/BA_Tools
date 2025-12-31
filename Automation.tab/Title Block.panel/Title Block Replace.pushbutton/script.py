# -*- coding: utf-8 -*-
"""
Title Block Replacer - Modern UI
Replace title blocks on selected sheets with advanced filtering
"""
__title__ = 'Replace\nTitle Blocks'
__doc__ = 'Replace title blocks on selected sheets'

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
import System
from System.Windows.Forms import (
    Form, Label, CheckedListBox, Button, ComboBox, GroupBox,
    FolderBrowserDialog, DialogResult, Panel, TextBox,
    FormBorderStyle, ComboBoxStyle, FlatStyle, BorderStyle
)
from System.Drawing import (
    Color, Font, FontStyle, Size, Point
)

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ==============================================================================
# COLLECT TITLE BLOCKS AND SHEETS
# ==============================================================================

def get_all_titleblock_types():
    """Get all title block family types in the project"""
    titleblock_types = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_TitleBlocks)\
        .WhereElementIsElementType()\
        .ToElements()
    
    # Group by family name
    titleblock_dict = {}
    for tb_type in titleblock_types:
        family_name = tb_type.FamilyName
        type_name = Element.Name.GetValue(tb_type)
        
        if family_name not in titleblock_dict:
            titleblock_dict[family_name] = []
        
        titleblock_dict[family_name].append({
            'type': tb_type,
            'name': type_name,
            'family': family_name
        })
    
    return titleblock_dict


def get_all_sheets():
    """Get all sheets in the project"""
    sheets = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Sheets)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Sort by sheet number
    sheets_sorted = sorted(sheets, key=lambda s: s.SheetNumber)
    
    return sheets_sorted


def get_sheet_titleblock(sheet):
    """Get the title block instance on a sheet"""
    titleblocks = FilteredElementCollector(doc, sheet.Id)\
        .OfCategory(BuiltInCategory.OST_TitleBlocks)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    if titleblocks:
        return titleblocks[0]
    return None


# Collect data
titleblock_families = get_all_titleblock_types()
all_sheets = get_all_sheets()

if not titleblock_families:
    forms.alert('No title block families found in the project', exitscript=True)

if not all_sheets:
    forms.alert('No sheets found in the project', exitscript=True)

output.print_md("**Found {} title block families**".format(len(titleblock_families)))
output.print_md("**Found {} sheets**".format(len(all_sheets)))


# ==============================================================================
# MODERN TITLE BLOCK REPLACER UI
# ==============================================================================

class TitleBlockReplacer(Form):
    def __init__(self, titleblocks, sheets, document):
        self.titleblocks = titleblocks
        self.sheets = sheets
        self.doc = document
        self.selected_sheets = []
        self.selected_titleblock_type = None
        self.result = None
        
        self.InitializeUI()
    
    def InitializeUI(self):
        # Form properties
        self.Text = "TITLE BLOCK REPLACER v2.4.0"
        self.Width = 1150
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
        self.border_color = Color.FromArgb(100, 150, 200)
        self.accent_color = Color.FromArgb(74, 144, 226)
        self.text_color = Color.FromArgb(220, 220, 220)
        self.text_muted = Color.FromArgb(140, 140, 150)
        
        y_offset = 20
        
        # ===== HEADER =====
        header_panel = Panel()
        header_panel.Location = Point(20, y_offset)
        header_panel.Size = Size(1090, 80)
        header_panel.BackColor = self.bg_panel
        self.Controls.Add(header_panel)
        
        title = Label()
        title.Text = "TITLE BLOCK REPLACER"
        title.Font = self.title_font
        title.ForeColor = Color.White
        title.Location = Point(20, 15)
        title.Size = Size(500, 30)
        header_panel.Controls.Add(title)
        
        subtitle = Label()
        subtitle.Text = "REPLACE TITLE BLOCKS ON SELECTED SHEETS"
        subtitle.Font = Font("Consolas", 8, FontStyle.Regular)
        subtitle.ForeColor = self.text_muted
        subtitle.Location = Point(20, 45)
        subtitle.Size = Size(500, 20)
        header_panel.Controls.Add(subtitle)
        
        y_offset += 100
        
        # ===== LEFT PANEL - TITLE BLOCK SELECTION =====
        left_x = 20
        
        tb_select_label = Label()
        tb_select_label.Text = "STEP 1: SELECT NEW TITLE BLOCK"
        tb_select_label.Font = self.label_font
        tb_select_label.ForeColor = self.text_muted
        tb_select_label.Location = Point(left_x, y_offset)
        tb_select_label.Size = Size(500, 20)
        self.Controls.Add(tb_select_label)
        
        y_offset += 30
        
        # Title Block Family Dropdown
        family_label = Label()
        family_label.Text = "Title Block Family:"
        family_label.Font = self.label_font
        family_label.ForeColor = self.text_color
        family_label.Location = Point(left_x, y_offset)
        family_label.Size = Size(150, 20)
        self.Controls.Add(family_label)
        
        self.family_combo = ComboBox()
        self.family_combo.Location = Point(left_x + 160, y_offset)
        self.family_combo.Size = Size(360, 30)
        self.family_combo.Font = self.input_font
        self.family_combo.BackColor = Color.FromArgb(15, 15, 20)
        self.family_combo.ForeColor = self.text_color
        self.family_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.family_combo.FlatStyle = FlatStyle.Flat
        
        # Add families
        self.family_combo.Items.Add("-- Select Title Block Family --")
        for family_name in sorted(self.titleblocks.keys()):
            self.family_combo.Items.Add(family_name)
        self.family_combo.SelectedIndex = 0
        self.family_combo.SelectedIndexChanged += self.OnFamilyChanged
        
        self.Controls.Add(self.family_combo)
        
        y_offset += 40
        
        # Title Block Type Dropdown
        type_label = Label()
        type_label.Text = "Title Block Type:"
        type_label.Font = self.label_font
        type_label.ForeColor = self.text_color
        type_label.Location = Point(left_x, y_offset)
        type_label.Size = Size(150, 20)
        self.Controls.Add(type_label)
        
        self.type_combo = ComboBox()
        self.type_combo.Location = Point(left_x + 160, y_offset)
        self.type_combo.Size = Size(360, 30)
        self.type_combo.Font = self.input_font
        self.type_combo.BackColor = Color.FromArgb(15, 15, 20)
        self.type_combo.ForeColor = self.text_color
        self.type_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.type_combo.FlatStyle = FlatStyle.Flat
        self.type_combo.Enabled = False
        
        self.type_combo.Items.Add("-- Select Type --")
        self.type_combo.SelectedIndex = 0
        
        self.Controls.Add(self.type_combo)
        
        y_offset += 60
        
        # ===== SHEET SELECTION =====
        sheet_label = Label()
        sheet_label.Text = "STEP 2: SELECT SHEETS TO UPDATE"
        sheet_label.Font = self.label_font
        sheet_label.ForeColor = self.text_muted
        sheet_label.Location = Point(left_x, y_offset)
        sheet_label.Size = Size(300, 20)
        self.Controls.Add(sheet_label)
        
        # Filter input
        filter_label = Label()
        filter_label.Text = "Search:"
        filter_label.Font = Font("Consolas", 9, FontStyle.Regular)
        filter_label.ForeColor = self.text_muted
        filter_label.Location = Point(left_x + 300, y_offset + 2)
        filter_label.Size = Size(60, 20)
        self.Controls.Add(filter_label)
        
        self.filter_textbox = TextBox()
        self.filter_textbox.Location = Point(left_x + 365, y_offset)
        self.filter_textbox.Size = Size(155, 20)
        self.filter_textbox.Font = Font("Consolas", 9, FontStyle.Regular)
        self.filter_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.filter_textbox.ForeColor = self.text_color
        self.filter_textbox.TextChanged += self.OnFilterChanged
        self.Controls.Add(self.filter_textbox)
        
        y_offset += 30
        
        # Select/Clear/Filtered buttons
        select_all_btn = Button()
        select_all_btn.Text = "Select All"
        select_all_btn.Location = Point(left_x, y_offset)
        select_all_btn.Size = Size(85, 26)
        select_all_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        select_all_btn.BackColor = self.accent_color
        select_all_btn.ForeColor = Color.White
        select_all_btn.FlatStyle = FlatStyle.Flat
        select_all_btn.FlatAppearance.BorderSize = 0
        select_all_btn.Click += self.SelectAll
        self.Controls.Add(select_all_btn)
        
        clear_all_btn = Button()
        clear_all_btn.Text = "Clear All"
        clear_all_btn.Location = Point(left_x + 90, y_offset)
        clear_all_btn.Size = Size(85, 26)
        clear_all_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        clear_all_btn.BackColor = Color.FromArgb(80, 80, 100)
        clear_all_btn.ForeColor = Color.White
        clear_all_btn.FlatStyle = FlatStyle.Flat
        clear_all_btn.FlatAppearance.BorderSize = 0
        clear_all_btn.Click += self.ClearAll
        self.Controls.Add(clear_all_btn)
        
        filter_btn = Button()
        filter_btn.Text = "Select Filtered"
        filter_btn.Location = Point(left_x + 180, y_offset)
        filter_btn.Size = Size(110, 26)
        filter_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        filter_btn.BackColor = Color.FromArgb(60, 100, 140)
        filter_btn.ForeColor = Color.White
        filter_btn.FlatStyle = FlatStyle.Flat
        filter_btn.FlatAppearance.BorderSize = 0
        filter_btn.Click += self.SelectFiltered
        self.Controls.Add(filter_btn)
        
        # Sheet count
        self.sheet_count_label = Label()
        self.sheet_count_label.Text = "0 / {} selected".format(len(self.sheets))
        self.sheet_count_label.Font = Font("Consolas", 8, FontStyle.Regular)
        self.sheet_count_label.ForeColor = self.text_muted
        self.sheet_count_label.Location = Point(left_x + 300, y_offset + 4)
        self.sheet_count_label.Size = Size(220, 20)
        self.Controls.Add(self.sheet_count_label)
        
        y_offset += 35
        
        # Sheet CheckedListBox
        self.sheet_list = CheckedListBox()
        self.sheet_list.Location = Point(left_x, y_offset)
        self.sheet_list.Size = Size(520, 320)
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
        
        # Update count on check
        self.sheet_list.ItemCheck += self.UpdateSheetCount
        
        # ===== RIGHT PANEL - PREVIEW AND STATUS =====
        right_x = 560
        y_offset = 120
        
        # Current Title Blocks Info
        info_group = GroupBox()
        info_group.Text = "CURRENT TITLE BLOCKS"
        info_group.Location = Point(right_x, y_offset)
        info_group.Size = Size(550, 220)
        info_group.Font = self.label_font
        info_group.ForeColor = self.text_muted
        info_group.BackColor = self.bg_panel
        self.Controls.Add(info_group)
        
        self.current_tb_text = TextBox()
        self.current_tb_text.Location = Point(15, 30)
        self.current_tb_text.Size = Size(520, 180)
        self.current_tb_text.Font = Font("Consolas", 9, FontStyle.Regular)
        self.current_tb_text.BackColor = Color.FromArgb(10, 10, 15)
        self.current_tb_text.ForeColor = self.text_muted
        self.current_tb_text.BorderStyle = BorderStyle.None
        self.current_tb_text.ReadOnly = True
        self.current_tb_text.Multiline = True
        self.current_tb_text.WordWrap = True
        self.current_tb_text.Text = "Select sheets to see current title blocks..."
        info_group.Controls.Add(self.current_tb_text)
        
        y_offset += 240
        
        # Status Panel
        status_group = GroupBox()
        status_group.Text = "REPLACEMENT LOG"
        status_group.Location = Point(right_x, y_offset)
        status_group.Size = Size(550, 280)
        status_group.Font = self.label_font
        status_group.ForeColor = self.text_muted
        status_group.BackColor = self.bg_panel
        self.Controls.Add(status_group)
        
        self.status_text = TextBox()
        self.status_text.Location = Point(15, 30)
        self.status_text.Size = Size(520, 240)
        self.status_text.Font = Font("Consolas", 9, FontStyle.Regular)
        self.status_text.BackColor = Color.FromArgb(10, 10, 15)
        self.status_text.ForeColor = self.text_muted
        self.status_text.BorderStyle = BorderStyle.None
        self.status_text.ReadOnly = True
        self.status_text.Multiline = True
        self.status_text.WordWrap = True
        self.status_text.Text = (
            "• Select new title block family and type\n"
            "• Select sheets to update\n"
            "• Click REPLACE TITLE BLOCKS\n"
            "• Existing title blocks will be replaced\n"
            "• Position will be maintained"
        )
        status_group.Controls.Add(self.status_text)
        
        # Update current title blocks on selection
        self.sheet_list.ItemCheck += self.UpdateCurrentTitleBlocks
        
        # ===== REPLACE BUTTON =====
        y_offset = 640
        
        self.replace_btn = Button()
        self.replace_btn.Text = "REPLACE TITLE BLOCKS"
        self.replace_btn.Location = Point(20, y_offset)
        self.replace_btn.Size = Size(1090, 55)
        self.replace_btn.Font = Font("Consolas", 12, FontStyle.Bold)
        self.replace_btn.BackColor = self.accent_color
        self.replace_btn.ForeColor = Color.White
        self.replace_btn.FlatStyle = FlatStyle.Flat
        self.replace_btn.FlatAppearance.BorderSize = 0
        self.replace_btn.Click += self.ReplaceTitleBlocks
        self.Controls.Add(self.replace_btn)
    
    def OnFamilyChanged(self, sender, args):
        """When family is selected, populate type dropdown"""
        self.type_combo.Items.Clear()
        self.type_combo.Items.Add("-- Select Type --")
        
        if self.family_combo.SelectedIndex > 0:
            family_name = self.family_combo.SelectedItem.ToString()
            types = self.titleblocks[family_name]
            
            for tb_type in types:
                self.type_combo.Items.Add(tb_type['name'])
            
            self.type_combo.Enabled = True
            if len(types) == 1:
                self.type_combo.SelectedIndex = 1
            else:
                self.type_combo.SelectedIndex = 0
        else:
            self.type_combo.Enabled = False
            self.type_combo.SelectedIndex = 0
    
    def OnFilterChanged(self, sender, args):
        """Filter sheets based on text input"""
        filter_text = self.filter_textbox.Text.lower()
        
        self.sheet_list.Items.Clear()
        
        for sheet in self.sheets:
            display_text = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            if not filter_text or filter_text in display_text.lower():
                self.sheet_list.Items.Add(display_text)
        
        self.UpdateSheetCount(None, None)
    
    def SelectAll(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, True)
    
    def ClearAll(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, False)
    
    def SelectFiltered(self, sender, args):
        """Select all currently visible (filtered) sheets"""
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, True)
    
    def UpdateSheetCount(self, sender, args):
        if sender:
            System.Windows.Forms.Timer().Tick += lambda s, e: self.DoUpdateCount(s, e, sender)
        else:
            self.DoUpdateCount(None, None, self.sheet_list)
    
    def DoUpdateCount(self, sender, args, listbox):
        if sender:
            sender.Stop()
        count = listbox.CheckedItems.Count
        total = len([s for s in self.sheets if not self.filter_textbox.Text or 
                    self.filter_textbox.Text.lower() in "{} - {}".format(s.SheetNumber, s.Name).lower()])
        self.sheet_count_label.Text = "{} / {} selected".format(count, total)
    
    def UpdateCurrentTitleBlocks(self, sender, args):
        """Show current title blocks for selected sheets"""
        System.Windows.Forms.Timer().Tick += lambda s, e: self.DoUpdateCurrentTB(s, e)
    
    def DoUpdateCurrentTB(self, sender, args):
        sender.Stop()
        
        checked_count = self.sheet_list.CheckedItems.Count
        if checked_count == 0:
            self.current_tb_text.Text = "Select sheets to see current title blocks..."
            return
        
        # Get current title blocks
        tb_info = {}
        filter_text = self.filter_textbox.Text.lower()
        
        for i in range(self.sheet_list.CheckedItems.Count):
            idx = self.sheet_list.CheckedIndices[i]
            
            # Map back to original sheet considering filter
            filtered_sheets = [s for s in self.sheets if not filter_text or 
                             filter_text in "{} - {}".format(s.SheetNumber, s.Name).lower()]
            
            if idx < len(filtered_sheets):
                sheet = filtered_sheets[idx]
                
                current_tb = get_sheet_titleblock(sheet)
                if current_tb:
                    tb_type = self.doc.GetElement(current_tb.GetTypeId())
                    tb_name = "{} : {}".format(
                        tb_type.FamilyName,
                        Element.Name.GetValue(tb_type)
                    )
                    
                    if tb_name not in tb_info:
                        tb_info[tb_name] = []
                    tb_info[tb_name].append(sheet.SheetNumber)
        
        # Display info
        info_text = "Selected {} sheet(s):\n\n".format(checked_count)
        for tb_name, sheet_numbers in tb_info.items():
            info_text += "• {}\n".format(tb_name)
            info_text += "  Sheets: {}\n\n".format(", ".join(sheet_numbers[:5]))
            if len(sheet_numbers) > 5:
                info_text += "  ... and {} more\n\n".format(len(sheet_numbers) - 5)
        
        self.current_tb_text.Text = info_text
    
    def AddLog(self, message):
        timestamp = System.DateTime.Now.ToString("HH:mm:ss")
        log_line = "[{}] {}\r\n".format(timestamp, message)
        self.status_text.AppendText(log_line)
        self.status_text.SelectionStart = len(self.status_text.Text)
        self.status_text.ScrollToCaret()
        self.Refresh()
    
    def ReplaceTitleBlocks(self, sender, args):
        # Validate selections
        if self.family_combo.SelectedIndex == 0:
            forms.alert("Please select a title block family", title="Validation Error")
            return
        
        if self.type_combo.SelectedIndex == 0:
            forms.alert("Please select a title block type", title="Validation Error")
            return
        
        checked_indices = []
        for i in range(self.sheet_list.CheckedItems.Count):
            checked_indices.append(self.sheet_list.CheckedIndices[i])
        
        if len(checked_indices) == 0:
            forms.alert("Please select at least one sheet", title="Validation Error")
            return
        
        # Get selected title block type
        family_name = self.family_combo.SelectedItem.ToString()
        type_name = self.type_combo.SelectedItem.ToString()
        
        selected_tb = None
        for tb_data in self.titleblocks[family_name]:
            if tb_data['name'] == type_name:
                selected_tb = tb_data['type']
                break
        
        if not selected_tb:
            forms.alert("Could not find selected title block type", title="Error")
            return
        
        # Map checked indices to actual sheets (considering filter)
        filter_text = self.filter_textbox.Text.lower()
        filtered_sheets = [s for s in self.sheets if not filter_text or 
                         filter_text in "{} - {}".format(s.SheetNumber, s.Name).lower()]
        
        self.selected_sheets = [filtered_sheets[i] for i in checked_indices if i < len(filtered_sheets)]
        
        # Confirm
        confirm_msg = "Replace title blocks on {} sheet(s) with:\n\n{} : {}\n\nContinue?".format(
            len(self.selected_sheets),
            family_name,
            type_name
        )
        
        if not forms.alert(confirm_msg, yes=True, no=True):
            return
        
        # Update UI
        self.replace_btn.Enabled = False
        self.replace_btn.Text = "REPLACING..."
        self.status_text.Clear()
        
        self.AddLog("Starting title block replacement...")
        self.AddLog("Target: {} : {}".format(family_name, type_name))
        self.AddLog("Sheets: {}".format(len(self.selected_sheets)))
        self.AddLog("")
        
        success_count = 0
        error_count = 0
        
        with revit.Transaction("Replace Title Blocks"):
            # Activate symbol if needed
            if not selected_tb.IsActive:
                selected_tb.Activate()
                self.doc.Regenerate()
                self.AddLog("Activated title block symbol")
            
            for sheet in self.selected_sheets:
                try:
                    # Get current title block
                    current_tb = get_sheet_titleblock(sheet)
                    
                    if current_tb:
                        # Get current location
                        current_location = current_tb.Location.Point
                        
                        # Delete current title block
                        self.doc.Delete(current_tb.Id)
                        
                        # Create new title block
                        new_tb = self.doc.Create.NewFamilyInstance(
                            current_location,
                            selected_tb,
                            sheet
                        )
                        
                        success_count += 1
                        self.AddLog("✓ {} - {}".format(sheet.SheetNumber, sheet.Name))
                    else:
                        # No existing title block, create new one
                        new_tb = self.doc.Create.NewFamilyInstance(
                            XYZ.Zero,
                            selected_tb,
                            sheet
                        )
                        success_count += 1
                        self.AddLog("✓ {} - {} (new)".format(sheet.SheetNumber, sheet.Name))
                    
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
        
        msg = "Title block replacement complete!\n\n"
        msg += "Success: {}\n".format(success_count)
        msg += "Errors: {}".format(error_count)
        
        forms.alert(msg, title="Complete")
        
        self.result = True
        self.replace_btn.Enabled = True
        self.replace_btn.Text = "REPLACE TITLE BLOCKS"


# ==============================================================================
# SHOW FORM
# ==============================================================================

form = TitleBlockReplacer(titleblock_families, all_sheets, doc)
form.ShowDialog()

if form.result:
    output.print_md("### ✅ Title Block Replacement Completed")
else:
    output.print_md("### Operation Cancelled")