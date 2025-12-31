# -*- coding: utf-8 -*-
"""
Title Block Batch Update
Select multiple sheets and update title block parameters at once
"""
__title__ = 'Title Block\nUpdate'
__doc__ = 'Batch update title block parameters on multiple sheets'

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
import System
from System.Windows.Forms import (
    Form, Label, CheckedListBox, Button, TextBox, RadioButton,
    Panel, GroupBox, DateTimePicker,
    FormBorderStyle, FlatStyle, BorderStyle
)
from System.Drawing import (
    Color, Font, FontStyle, Size, Point
)

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ==============================================================================
# COLLECT SHEETS AND TITLE BLOCKS
# ==============================================================================

def get_all_sheets_with_titleblocks():
    """Get all sheets with their title blocks"""
    sheets = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Sheets)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    sheet_data = []
    
    for sheet in sheets:
        # Get title block on sheet
        titleblocks = FilteredElementCollector(doc, sheet.Id)\
            .OfCategory(BuiltInCategory.OST_TitleBlocks)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        if titleblocks:
            titleblock = titleblocks[0]
            sheet_data.append({
                'sheet': sheet,
                'titleblock': titleblock,
                'sheet_number': sheet.SheetNumber,
                'sheet_name': sheet.Name
            })
    
    # Sort by sheet number
    sheet_data_sorted = sorted(sheet_data, key=lambda x: x['sheet_number'])
    
    return sheet_data_sorted


# Get title block parameter names (common ones)
TITLEBLOCK_PARAMS = {
    'Drawn By': 'Drawn By',
    'Checked By': 'Checked By',
    'Designed By': 'Designed By',
    'Approved By': 'Approved By',
    'Sheet Issue Date': 'Sheet Issue Date',
    'Project Number': 'Project Number',
    'Project Name': 'Project Name',
    'Client Name': 'Client Name',
    'Project Address': 'Project Address'
}

# Collect data
all_sheets = get_all_sheets_with_titleblocks()

if not all_sheets:
    forms.alert('No sheets with title blocks found', exitscript=True)

output.print_md("**Found {} sheets with title blocks**".format(len(all_sheets)))


# ==============================================================================
# MODERN TITLE BLOCK UPDATE UI
# ==============================================================================

class TitleBlockBatchUpdate(Form):
    def __init__(self, sheet_data, document):
        self.sheet_data = sheet_data
        self.doc = document
        self.selected_sheets = []
        self.result = None
        
        self.InitializeUI()
    
    def InitializeUI(self):
        # Form properties
        self.Text = "TITLE BLOCK BATCH UPDATE"
        self.Width = 900
        self.Height = 800
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
        self.text_color = Color.FromArgb(220, 220, 220)
        self.text_muted = Color.FromArgb(140, 140, 150)
        
        y_offset = 20
        
        # ===== HEADER =====
        header_panel = Panel()
        header_panel.Location = Point(20, y_offset)
        header_panel.Size = Size(840, 70)
        header_panel.BackColor = self.bg_panel
        self.Controls.Add(header_panel)
        
        title = Label()
        title.Text = "TITLE BLOCK UPDATE"
        title.Font = self.title_font
        title.ForeColor = Color.White
        title.Location = Point(20, 10)
        title.Size = Size(400, 30)
        header_panel.Controls.Add(title)
        
        subtitle = Label()
        subtitle.Text = "BATCH UPDATE TITLE BLOCK PARAMETERS"
        subtitle.Font = Font("Consolas", 8, FontStyle.Regular)
        subtitle.ForeColor = self.text_muted
        subtitle.Location = Point(20, 40)
        subtitle.Size = Size(400, 20)
        header_panel.Controls.Add(subtitle)
        
        y_offset += 90
        
        # ===== SHEET SELECTION =====
        sheet_label = Label()
        sheet_label.Text = "SELECT SHEETS"
        sheet_label.Font = self.label_font
        sheet_label.ForeColor = self.text_muted
        sheet_label.Location = Point(20, y_offset)
        sheet_label.Size = Size(400, 20)
        self.Controls.Add(sheet_label)
        
        y_offset += 25
        
        # Selection buttons
        self.select_all_radio = RadioButton()
        self.select_all_radio.Text = "Select all"
        self.select_all_radio.Location = Point(20, y_offset)
        self.select_all_radio.Size = Size(100, 20)
        self.select_all_radio.Font = Font("Consolas", 9, FontStyle.Regular)
        self.select_all_radio.ForeColor = self.text_color
        self.select_all_radio.CheckedChanged += self.OnSelectAllChanged
        self.Controls.Add(self.select_all_radio)
        
        self.select_none_radio = RadioButton()
        self.select_none_radio.Text = "Select none"
        self.select_none_radio.Location = Point(150, y_offset)
        self.select_none_radio.Size = Size(120, 20)
        self.select_none_radio.Font = Font("Consolas", 9, FontStyle.Regular)
        self.select_none_radio.ForeColor = self.text_color
        self.select_none_radio.CheckedChanged += self.OnSelectNoneChanged
        self.Controls.Add(self.select_none_radio)
        
        # Sheet count
        self.sheet_count_label = Label()
        self.sheet_count_label.Text = "0 / {} selected".format(len(self.sheet_data))
        self.sheet_count_label.Font = Font("Consolas", 8, FontStyle.Regular)
        self.sheet_count_label.ForeColor = self.text_muted
        self.sheet_count_label.Location = Point(300, y_offset + 2)
        self.sheet_count_label.Size = Size(200, 20)
        self.Controls.Add(self.sheet_count_label)
        
        y_offset += 30
        
        # Sheet CheckedListBox
        self.sheet_list = CheckedListBox()
        self.sheet_list.Location = Point(20, y_offset)
        self.sheet_list.Size = Size(400, 380)
        self.sheet_list.Font = Font("Consolas", 9, FontStyle.Regular)
        self.sheet_list.BackColor = Color.FromArgb(15, 15, 20)
        self.sheet_list.ForeColor = self.text_color
        self.sheet_list.BorderStyle = BorderStyle.FixedSingle
        self.sheet_list.CheckOnClick = True
        
        # Add sheets
        for sheet_data in self.sheet_data:
            display_text = "{} - {}".format(
                sheet_data['sheet_number'],
                sheet_data['sheet_name']
            )
            self.sheet_list.Items.Add(display_text)
        
        self.Controls.Add(self.sheet_list)
        
        self.sheet_list.ItemCheck += self.UpdateSheetCount
        
        # ===== PARAMETER INPUTS =====
        right_x = 440
        y_offset = 110
        
        param_label = Label()
        param_label.Text = "TITLE BLOCK PARAMETERS"
        param_label.Font = self.label_font
        param_label.ForeColor = self.text_muted
        param_label.Location = Point(right_x, y_offset)
        param_label.Size = Size(420, 20)
        self.Controls.Add(param_label)
        
        y_offset += 30
        
        # Parameter input panel
        param_panel = Panel()
        param_panel.Location = Point(right_x, y_offset)
        param_panel.Size = Size(420, 380)
        param_panel.BackColor = self.bg_panel
        self.Controls.Add(param_panel)
        
        panel_y = 20
        
        # Drawn By
        self.CreateParameterInput(
            param_panel,
            "Drawn By:",
            panel_y,
            "drawn_by"
        )
        panel_y += 50
        
        # Checked By
        self.CreateParameterInput(
            param_panel,
            "Checked By:",
            panel_y,
            "checked_by"
        )
        panel_y += 50
        
        # Designed By
        self.CreateParameterInput(
            param_panel,
            "Designed By:",
            panel_y,
            "designed_by"
        )
        panel_y += 50
        
        # Approved By
        self.CreateParameterInput(
            param_panel,
            "Approved By:",
            panel_y,
            "approved_by"
        )
        panel_y += 50
        
        # Sheet Issue Date
        date_label = Label()
        date_label.Text = "Sheet Issue Date:"
        date_label.Location = Point(15, panel_y)
        date_label.Size = Size(120, 20)
        date_label.Font = self.label_font
        date_label.ForeColor = self.text_color
        param_panel.Controls.Add(date_label)
        
        self.issue_date_picker = DateTimePicker()
        self.issue_date_picker.Location = Point(140, panel_y - 2)
        self.issue_date_picker.Size = Size(260, 25)
        self.issue_date_picker.Font = Font("Consolas", 9, FontStyle.Regular)
        self.issue_date_picker.Format = System.Windows.Forms.DateTimePickerFormat.Short
        param_panel.Controls.Add(self.issue_date_picker)
        
        panel_y += 50
        
        # Project Number
        self.CreateParameterInput(
            param_panel,
            "Project Number:",
            panel_y,
            "project_number"
        )
        panel_y += 50
        
        # Clear button
        clear_btn = Button()
        clear_btn.Text = "Clear All Fields"
        clear_btn.Location = Point(140, panel_y)
        clear_btn.Size = Size(260, 30)
        clear_btn.Font = Font("Consolas", 9, FontStyle.Bold)
        clear_btn.BackColor = Color.FromArgb(80, 80, 100)
        clear_btn.ForeColor = Color.White
        clear_btn.FlatStyle = FlatStyle.Flat
        clear_btn.FlatAppearance.BorderSize = 0
        clear_btn.Click += self.ClearFields
        param_panel.Controls.Add(clear_btn)
        
        # ===== BUTTONS =====
        y_offset = 710
        
        cancel_btn = Button()
        cancel_btn.Text = "CANCEL"
        cancel_btn.Location = Point(20, y_offset)
        cancel_btn.Size = Size(410, 45)
        cancel_btn.Font = Font("Consolas", 11, FontStyle.Bold)
        cancel_btn.BackColor = Color.FromArgb(80, 80, 100)
        cancel_btn.ForeColor = Color.White
        cancel_btn.FlatStyle = FlatStyle.Flat
        cancel_btn.FlatAppearance.BorderSize = 0
        cancel_btn.Click += lambda s, e: self.Close()
        self.Controls.Add(cancel_btn)
        
        self.update_btn = Button()
        self.update_btn.Text = "UPDATE"
        self.update_btn.Location = Point(450, y_offset)
        self.update_btn.Size = Size(410, 45)
        self.update_btn.Font = Font("Consolas", 11, FontStyle.Bold)
        self.update_btn.BackColor = self.accent_color
        self.update_btn.ForeColor = Color.White
        self.update_btn.FlatStyle = FlatStyle.Flat
        self.update_btn.FlatAppearance.BorderSize = 0
        self.update_btn.Click += self.UpdateTitleBlocks
        self.Controls.Add(self.update_btn)
    
    def CreateParameterInput(self, parent, label_text, y_pos, attr_name):
        """Helper to create parameter input field"""
        label = Label()
        label.Text = label_text
        label.Location = Point(15, y_pos)
        label.Size = Size(120, 20)
        label.Font = self.label_font
        label.ForeColor = self.text_color
        parent.Controls.Add(label)
        
        textbox = TextBox()
        textbox.Location = Point(140, y_pos - 2)
        textbox.Size = Size(260, 25)
        textbox.Font = self.input_font
        textbox.BackColor = Color.FromArgb(15, 15, 20)
        textbox.ForeColor = self.text_color
        parent.Controls.Add(textbox)
        
        # Store reference
        setattr(self, attr_name + "_textbox", textbox)
    
    def OnSelectAllChanged(self, sender, args):
        if self.select_all_radio.Checked:
            for i in range(self.sheet_list.Items.Count):
                self.sheet_list.SetItemChecked(i, True)
    
    def OnSelectNoneChanged(self, sender, args):
        if self.select_none_radio.Checked:
            for i in range(self.sheet_list.Items.Count):
                self.sheet_list.SetItemChecked(i, False)
    
    def UpdateSheetCount(self, sender, args):
        System.Windows.Forms.Timer().Tick += lambda s, e: self.DoUpdateCount(s, e)
    
    def DoUpdateCount(self, sender, args):
        sender.Stop()
        count = self.sheet_list.CheckedItems.Count
        self.sheet_count_label.Text = "{} / {} selected".format(count, len(self.sheet_data))
    
    def ClearFields(self, sender, args):
        """Clear all input fields"""
        self.drawn_by_textbox.Text = ""
        self.checked_by_textbox.Text = ""
        self.designed_by_textbox.Text = ""
        self.approved_by_textbox.Text = ""
        self.project_number_textbox.Text = ""
        self.issue_date_picker.Value = System.DateTime.Now
    
    def GetParameterValue(self, element, param_name):
        """Get parameter value from element"""
        param = element.LookupParameter(param_name)
        if param and not param.IsReadOnly:
            return param
        return None
    
    def UpdateTitleBlocks(self, sender, args):
        # Get selected sheets
        checked_indices = []
        for i in range(self.sheet_list.CheckedItems.Count):
            checked_indices.append(self.sheet_list.CheckedIndices[i])
        
        if len(checked_indices) == 0:
            forms.alert("Please select at least one sheet", title="Validation Error")
            return
        
        self.selected_sheets = [self.sheet_data[i] for i in checked_indices]
        
        # Get parameter values
        params_to_update = {}
        
        drawn_by = self.drawn_by_textbox.Text.strip()
        if drawn_by:
            params_to_update['Drawn By'] = drawn_by
        
        checked_by = self.checked_by_textbox.Text.strip()
        if checked_by:
            params_to_update['Checked By'] = checked_by
        
        designed_by = self.designed_by_textbox.Text.strip()
        if designed_by:
            params_to_update['Designed By'] = designed_by
        
        approved_by = self.approved_by_textbox.Text.strip()
        if approved_by:
            params_to_update['Approved By'] = approved_by
        
        project_number = self.project_number_textbox.Text.strip()
        if project_number:
            params_to_update['Project Number'] = project_number
        
        # Issue date - always include if changed from default
        issue_date = self.issue_date_picker.Value.ToString("MM/dd/yyyy")
        params_to_update['Sheet Issue Date'] = issue_date
        
        if len(params_to_update) == 0:
            forms.alert("Please fill in at least one parameter to update", title="Validation Error")
            return
        
        # Confirm
        param_list = "\n".join(["- {}: {}".format(k, v) for k, v in params_to_update.items()])
        confirm_msg = "Update {} sheet(s) with:\n\n{}\n\nContinue?".format(
            len(self.selected_sheets),
            param_list
        )
        
        if not forms.alert(confirm_msg, yes=True, no=True):
            return
        
        # Update UI
        self.update_btn.Enabled = False
        self.update_btn.Text = "UPDATING..."
        
        success_count = 0
        error_count = 0
        
        with revit.Transaction("Update Title Blocks"):
            for sheet_data in self.selected_sheets:
                try:
                    titleblock = sheet_data['titleblock']
                    sheet_number = sheet_data['sheet_number']
                    
                    updated_params = []
                    failed_params = []
                    
                    for param_name, param_value in params_to_update.items():
                        param = self.GetParameterValue(titleblock, param_name)
                        
                        if param:
                            try:
                                param.Set(param_value)
                                updated_params.append(param_name)
                            except:
                                failed_params.append(param_name)
                        else:
                            # Try on sheet if not on titleblock
                            sheet_param = self.GetParameterValue(sheet_data['sheet'], param_name)
                            if sheet_param:
                                try:
                                    sheet_param.Set(param_value)
                                    updated_params.append(param_name)
                                except:
                                    failed_params.append(param_name)
                            else:
                                failed_params.append(param_name)
                    
                    if updated_params:
                        success_count += 1
                        output.print_md("✓ {} - Updated: {}".format(
                            sheet_number,
                            ", ".join(updated_params)
                        ))
                        if failed_params:
                            output.print_md("  ⚠ Not found/read-only: {}".format(
                                ", ".join(failed_params)
                            ))
                    else:
                        error_count += 1
                        output.print_md("✗ {} - No parameters updated".format(sheet_number))
                
                except Exception as e:
                    error_count += 1
                    output.print_md("✗ {} - Error: {}".format(
                        sheet_data['sheet_number'],
                        str(e)
                    ))
        
        # Summary
        msg = "Title block update complete!\n\n"
        msg += "Sheets Updated: {}\n".format(success_count)
        msg += "Errors: {}".format(error_count)
        
        forms.alert(msg, title="Complete")
        
        self.result = True
        self.Close()


# ==============================================================================
# SHOW FORM
# ==============================================================================

form = TitleBlockBatchUpdate(all_sheets, doc)
form.ShowDialog()

if form.result:
    output.print_md("### ✅ Title Blocks Updated Successfully")
else:
    output.print_md("### Operation Cancelled")