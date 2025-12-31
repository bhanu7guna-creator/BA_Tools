# -*- coding: utf-8 -*-
"""
View Title on Sheet Renamer - Modern UI
Rename view titles on sheets with find/replace, prefix & suffix
"""
__title__ = 'Rename View\nTitles'
__doc__ = 'Rename viewport titles on sheets with modern interface'

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
import System
import re

from System.Windows.Forms import (
    Form, Label, CheckedListBox, Button, ComboBox, GroupBox,
    RadioButton, Panel, TextBox, CheckBox,
    FormBorderStyle, ComboBoxStyle, FlatStyle, BorderStyle
)
from System.Drawing import (
    Color, Font, FontStyle, Size, Point
)

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ==============================================================================
# COLLECT VIEWPORT DATA
# ==============================================================================

def get_all_viewports():
    """Get all viewports in the project"""
    sheets = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Sheets)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    viewport_data = []
    
    for sheet in sheets:
        viewport_ids = sheet.GetAllViewports()
        
        for vp_id in viewport_ids:
            vp = doc.GetElement(vp_id)
            view = doc.GetElement(vp.ViewId)
            
            if view:
                title_param = vp.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                title = title_param.AsString() if title_param else ""
                
                viewport_data.append({
                    'viewport': vp,
                    'view': view,
                    'sheet': sheet,
                    'sheet_number': sheet.SheetNumber,
                    'sheet_name': sheet.Name,
                    'view_name': view.Name,
                    'title': title or ""
                })
    
    # Sort by sheet number
    viewport_data_sorted = sorted(viewport_data, key=lambda x: x['sheet_number'])
    
    return viewport_data_sorted


# Collect viewport data
all_viewports = get_all_viewports()

if not all_viewports:
    forms.alert('No views found on sheets', exitscript=True)

output.print_md("**Found {} views on sheets**".format(len(all_viewports)))


# ==============================================================================
# MODERN VIEW TITLE RENAMER UI
# ==============================================================================

class ViewTitleRenamer(Form):
    def __init__(self, viewport_data, document):
        self.viewport_data = viewport_data
        self.filtered_data = list(viewport_data)
        self.doc = document
        self.selected_viewports = []
        self.result = None
        
        self.InitializeUI()
    
    def InitializeUI(self):
        # Form properties
        self.Text = "VIEW TITLE RENAMER v2.4.0"
        self.Width = 1200
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
        header_panel.Size = Size(1140, 80)
        header_panel.BackColor = self.bg_panel
        self.Controls.Add(header_panel)
        
        title = Label()
        title.Text = "VIEW TITLE RENAMER"
        title.Font = self.title_font
        title.ForeColor = Color.White
        title.Location = Point(20, 15)
        title.Size = Size(500, 30)
        header_panel.Controls.Add(title)
        
        subtitle = Label()
        subtitle.Text = "RENAME VIEW TITLES ON SHEETS"
        subtitle.Font = Font("Consolas", 8, FontStyle.Regular)
        subtitle.ForeColor = self.text_muted
        subtitle.Location = Point(20, 45)
        subtitle.Size = Size(500, 20)
        header_panel.Controls.Add(subtitle)
        
        y_offset += 100
        
        # ===== LEFT PANEL - VIEW SELECTION =====
        left_x = 20
        
        view_label = Label()
        view_label.Text = "STEP 1: SELECT VIEWS ON SHEETS"
        view_label.Font = self.label_font
        view_label.ForeColor = self.text_muted
        view_label.Location = Point(left_x, y_offset)
        view_label.Size = Size(300, 20)
        self.Controls.Add(view_label)
        
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
        self.filter_textbox.Size = Size(235, 20)
        self.filter_textbox.Font = Font("Consolas", 9, FontStyle.Regular)
        self.filter_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.filter_textbox.ForeColor = self.text_color
        self.filter_textbox.TextChanged += self.OnFilterChanged
        self.Controls.Add(self.filter_textbox)
        
        y_offset += 30
        
        # Selection buttons
        select_all_btn = Button()
        select_all_btn.Text = "Select All"
        select_all_btn.Location = Point(left_x, y_offset)
        select_all_btn.Size = Size(90, 26)
        select_all_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        select_all_btn.BackColor = self.accent_color
        select_all_btn.ForeColor = Color.White
        select_all_btn.FlatStyle = FlatStyle.Flat
        select_all_btn.FlatAppearance.BorderSize = 0
        select_all_btn.Click += self.SelectAll
        self.Controls.Add(select_all_btn)
        
        clear_all_btn = Button()
        clear_all_btn.Text = "Clear All"
        clear_all_btn.Location = Point(left_x + 95, y_offset)
        clear_all_btn.Size = Size(90, 26)
        clear_all_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        clear_all_btn.BackColor = Color.FromArgb(80, 80, 100)
        clear_all_btn.ForeColor = Color.White
        clear_all_btn.FlatStyle = FlatStyle.Flat
        clear_all_btn.FlatAppearance.BorderSize = 0
        clear_all_btn.Click += self.ClearAll
        self.Controls.Add(clear_all_btn)
        
        filter_btn = Button()
        filter_btn.Text = "Select Filtered"
        filter_btn.Location = Point(left_x + 190, y_offset)
        filter_btn.Size = Size(110, 26)
        filter_btn.Font = Font("Consolas", 8, FontStyle.Bold)
        filter_btn.BackColor = Color.FromArgb(60, 100, 140)
        filter_btn.ForeColor = Color.White
        filter_btn.FlatStyle = FlatStyle.Flat
        filter_btn.FlatAppearance.BorderSize = 0
        filter_btn.Click += self.SelectFiltered
        self.Controls.Add(filter_btn)
        
        # Count label
        self.view_count_label = Label()
        self.view_count_label.Text = "0 / {} selected".format(len(self.viewport_data))
        self.view_count_label.Font = Font("Consolas", 8, FontStyle.Regular)
        self.view_count_label.ForeColor = self.text_muted
        self.view_count_label.Location = Point(left_x + 310, y_offset + 4)
        self.view_count_label.Size = Size(290, 20)
        self.Controls.Add(self.view_count_label)
        
        y_offset += 35
        
        # View CheckedListBox
        self.view_list = CheckedListBox()
        self.view_list.Location = Point(left_x, y_offset)
        self.view_list.Size = Size(600, 450)
        self.view_list.Font = Font("Consolas", 9, FontStyle.Regular)
        self.view_list.BackColor = Color.FromArgb(15, 15, 20)
        self.view_list.ForeColor = self.text_color
        self.view_list.BorderStyle = BorderStyle.FixedSingle
        self.view_list.CheckOnClick = True
        
        # Add viewports
        self.PopulateViewList()
        
        self.Controls.Add(self.view_list)
        
        self.view_list.ItemCheck += self.UpdateViewCount
        self.view_list.ItemCheck += lambda s, e: self.DelayedPreview()
        
        # ===== RIGHT PANEL - RENAME OPTIONS =====
        right_x = 640
        y_offset = 120
        
        options_label = Label()
        options_label.Text = "STEP 2: CONFIGURE RENAME OPTIONS"
        options_label.Font = self.label_font
        options_label.ForeColor = self.text_muted
        options_label.Location = Point(right_x, y_offset)
        options_label.Size = Size(500, 20)
        self.Controls.Add(options_label)
        
        y_offset += 30
        
        # Options Panel
        options_panel = Panel()
        options_panel.Location = Point(right_x, y_offset)
        options_panel.Size = Size(520, 280)
        options_panel.BackColor = self.bg_panel
        self.Controls.Add(options_panel)
        
        # Find
        find_label = Label()
        find_label.Text = "Find Text:"
        find_label.Location = Point(20, 20)
        find_label.Size = Size(100, 20)
        find_label.Font = self.label_font
        find_label.ForeColor = self.text_color
        options_panel.Controls.Add(find_label)
        
        self.find_textbox = TextBox()
        self.find_textbox.Location = Point(130, 18)
        self.find_textbox.Size = Size(370, 25)
        self.find_textbox.Font = self.input_font
        self.find_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.find_textbox.ForeColor = self.text_color
        self.find_textbox.TextChanged += lambda s, e: self.UpdatePreview()
        options_panel.Controls.Add(self.find_textbox)
        
        # Replace
        replace_label = Label()
        replace_label.Text = "Replace With:"
        replace_label.Location = Point(20, 58)
        replace_label.Size = Size(100, 20)
        replace_label.Font = self.label_font
        replace_label.ForeColor = self.text_color
        options_panel.Controls.Add(replace_label)
        
        self.replace_textbox = TextBox()
        self.replace_textbox.Location = Point(130, 56)
        self.replace_textbox.Size = Size(370, 25)
        self.replace_textbox.Font = self.input_font
        self.replace_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.replace_textbox.ForeColor = self.text_color
        self.replace_textbox.TextChanged += lambda s, e: self.UpdatePreview()
        options_panel.Controls.Add(self.replace_textbox)
        
        # Use View Name checkbox
        self.use_view_name_check = CheckBox()
        self.use_view_name_check.Text = "Use View Name instead of Current Title"
        self.use_view_name_check.Location = Point(130, 95)
        self.use_view_name_check.Size = Size(370, 20)
        self.use_view_name_check.Font = Font("Consolas", 9, FontStyle.Regular)
        self.use_view_name_check.ForeColor = self.text_color
        self.use_view_name_check.CheckedChanged += lambda s, e: self.UpdatePreview()
        options_panel.Controls.Add(self.use_view_name_check)
        
        # Add Sheet Number checkbox
        self.add_sheet_number_check = CheckBox()
        self.add_sheet_number_check.Text = "Add Sheet Number as Prefix"
        self.add_sheet_number_check.Location = Point(130, 120)
        self.add_sheet_number_check.Size = Size(370, 20)
        self.add_sheet_number_check.Font = Font("Consolas", 9, FontStyle.Regular)
        self.add_sheet_number_check.ForeColor = self.text_color
        self.add_sheet_number_check.CheckedChanged += lambda s, e: self.UpdatePreview()
        options_panel.Controls.Add(self.add_sheet_number_check)
        
        # Prefix
        prefix_label = Label()
        prefix_label.Text = "Custom Prefix:"
        prefix_label.Location = Point(20, 158)
        prefix_label.Size = Size(100, 20)
        prefix_label.Font = self.label_font
        prefix_label.ForeColor = self.text_color
        options_panel.Controls.Add(prefix_label)
        
        self.prefix_textbox = TextBox()
        self.prefix_textbox.Location = Point(130, 156)
        self.prefix_textbox.Size = Size(160, 25)
        self.prefix_textbox.Font = self.input_font
        self.prefix_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.prefix_textbox.ForeColor = self.text_color
        self.prefix_textbox.TextChanged += lambda s, e: self.UpdatePreview()
        options_panel.Controls.Add(self.prefix_textbox)
        
        # Suffix
        suffix_label = Label()
        suffix_label.Text = "Custom Suffix:"
        suffix_label.Location = Point(20, 196)
        suffix_label.Size = Size(100, 20)
        suffix_label.Font = self.label_font
        suffix_label.ForeColor = self.text_color
        options_panel.Controls.Add(suffix_label)
        
        self.suffix_textbox = TextBox()
        self.suffix_textbox.Location = Point(130, 194)
        self.suffix_textbox.Size = Size(160, 25)
        self.suffix_textbox.Font = self.input_font
        self.suffix_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.suffix_textbox.ForeColor = self.text_color
        self.suffix_textbox.TextChanged += lambda s, e: self.UpdatePreview()
        options_panel.Controls.Add(self.suffix_textbox)
        
        # Case sensitive
        self.case_sensitive_check = CheckBox()
        self.case_sensitive_check.Text = "Case Sensitive"
        self.case_sensitive_check.Location = Point(130, 235)
        self.case_sensitive_check.Size = Size(150, 20)
        self.case_sensitive_check.Font = Font("Consolas", 9, FontStyle.Regular)
        self.case_sensitive_check.ForeColor = self.text_color
        self.case_sensitive_check.CheckedChanged += lambda s, e: self.UpdatePreview()
        options_panel.Controls.Add(self.case_sensitive_check)
        
        y_offset += 300
        
        # Preview Panel
        preview_group = GroupBox()
        preview_group.Text = "PREVIEW"
        preview_group.Location = Point(right_x, y_offset)
        preview_group.Size = Size(520, 170)
        preview_group.Font = self.label_font
        preview_group.ForeColor = self.text_muted
        preview_group.BackColor = self.bg_panel
        self.Controls.Add(preview_group)
        
        self.preview_text = TextBox()
        self.preview_text.Location = Point(15, 25)
        self.preview_text.Size = Size(490, 135)
        self.preview_text.Font = Font("Consolas", 9, FontStyle.Regular)
        self.preview_text.BackColor = Color.FromArgb(10, 10, 15)
        self.preview_text.ForeColor = Color.FromArgb(100, 200, 100)
        self.preview_text.BorderStyle = BorderStyle.None
        self.preview_text.ReadOnly = True
        self.preview_text.Multiline = True
        self.preview_text.WordWrap = False
        self.preview_text.Text = "Select views and configure options to see preview..."
        preview_group.Controls.Add(self.preview_text)
        
        # ===== APPLY BUTTON =====
        y_offset = 690
        
        self.apply_btn = Button()
        self.apply_btn.Text = "APPLY RENAME"
        self.apply_btn.Location = Point(20, y_offset)
        self.apply_btn.Size = Size(1140, 55)
        self.apply_btn.Font = Font("Consolas", 12, FontStyle.Bold)
        self.apply_btn.BackColor = self.accent_color
        self.apply_btn.ForeColor = Color.White
        self.apply_btn.FlatStyle = FlatStyle.Flat
        self.apply_btn.FlatAppearance.BorderSize = 0
        self.apply_btn.Click += self.ApplyRename
        self.Controls.Add(self.apply_btn)
    
    def PopulateViewList(self):
        """Populate view list with current filtered data"""
        self.view_list.Items.Clear()
        
        for vp_data in self.filtered_data:
            display_text = "{} | {} | {}".format(
                vp_data['sheet_number'],
                vp_data['view_name'],
                vp_data['title'] or "(No Title)"
            )
            self.view_list.Items.Add(display_text)
    
    def OnFilterChanged(self, sender, args):
        """Filter views based on text input"""
        filter_text = self.filter_textbox.Text.lower()
        
        self.filtered_data = []
        for vp_data in self.viewport_data:
            search_text = "{} {} {}".format(
                vp_data['sheet_number'],
                vp_data['view_name'],
                vp_data['title']
            ).lower()
            
            if not filter_text or filter_text in search_text:
                self.filtered_data.append(vp_data)
        
        self.PopulateViewList()
        self.UpdateViewCount(None, None)
    
    def SelectAll(self, sender, args):
        for i in range(self.view_list.Items.Count):
            self.view_list.SetItemChecked(i, True)
    
    def ClearAll(self, sender, args):
        for i in range(self.view_list.Items.Count):
            self.view_list.SetItemChecked(i, False)
    
    def SelectFiltered(self, sender, args):
        for i in range(self.view_list.Items.Count):
            self.view_list.SetItemChecked(i, True)
    
    def UpdateViewCount(self, sender, args):
        if sender:
            System.Windows.Forms.Timer().Tick += lambda s, e: self.DoUpdateCount(s, e)
        else:
            self.DoUpdateCount(None, None)
    
    def DoUpdateCount(self, sender, args):
        if sender:
            sender.Stop()
        count = self.view_list.CheckedItems.Count
        total = len(self.filtered_data)
        self.view_count_label.Text = "{} / {} selected".format(count, total)
    
    def DelayedPreview(self):
        """Delay preview update to avoid frequent updates during checking"""
        timer = System.Windows.Forms.Timer()
        timer.Interval = 100
        timer.Tick += lambda s, e: self.DoDelayedPreview(s)
        timer.Start()
    
    def DoDelayedPreview(self, sender):
        sender.Stop()
        self.UpdatePreview()
    
    def GenerateNewTitle(self, vp_data):
        """Generate new title based on current settings"""
        # Base title
        if self.use_view_name_check.Checked:
            base_title = vp_data['view_name']
        else:
            base_title = vp_data['title']
        
        # Find and replace
        find_text = self.find_textbox.Text
        if find_text:
            replace_text = self.replace_textbox.Text
            if self.case_sensitive_check.Checked:
                base_title = base_title.replace(find_text, replace_text)
            else:
                base_title = re.sub(re.escape(find_text), replace_text, base_title, flags=re.IGNORECASE)
        
        # Add sheet number prefix
        if self.add_sheet_number_check.Checked:
            base_title = vp_data['sheet_number'] + " - " + base_title
        
        # Add custom prefix
        custom_prefix = self.prefix_textbox.Text
        if custom_prefix:
            base_title = custom_prefix + base_title
        
        # Add custom suffix
        custom_suffix = self.suffix_textbox.Text
        if custom_suffix:
            base_title = base_title + custom_suffix
        
        return base_title
    
    def UpdatePreview(self):
        """Update preview with sample renames"""
        checked_count = self.view_list.CheckedItems.Count
        
        if checked_count == 0:
            self.preview_text.Text = "Select views to see preview..."
            return
        
        # Get sample (first 3 checked items)
        samples = []
        for i in range(min(checked_count, 3)):
            idx = self.view_list.CheckedIndices[i]
            
            if idx < len(self.filtered_data):
                vp_data = self.filtered_data[idx]
                old_title = vp_data['title'] or "(No Title)"
                new_title = self.GenerateNewTitle(vp_data)
                
                samples.append("{}\n  → {}".format(old_title, new_title))
        
        preview_text = "Preview ({} views):\n\n".format(checked_count)
        preview_text += "\n\n".join(samples)
        
        if checked_count > 3:
            preview_text += "\n\n... and {} more".format(checked_count - 3)
        
        self.preview_text.Text = preview_text
    
    def ApplyRename(self, sender, args):
        # Get selected viewports
        checked_indices = []
        for i in range(self.view_list.CheckedItems.Count):
            checked_indices.append(self.view_list.CheckedIndices[i])
        
        if len(checked_indices) == 0:
            forms.alert("Please select at least one view", title="Validation Error")
            return
        
        self.selected_viewports = [self.filtered_data[i] for i in checked_indices if i < len(self.filtered_data)]
        
        # Confirm
        confirm_msg = "Rename {} view title(s)?\n\nContinue?".format(len(self.selected_viewports))
        
        if not forms.alert(confirm_msg, yes=True, no=True):
            return
        
        # Update UI
        self.apply_btn.Enabled = False
        self.apply_btn.Text = "RENAMING..."
        
        success_count = 0
        error_count = 0
        
        with revit.Transaction("Rename View Titles on Sheets"):
            for vp_data in self.selected_viewports:
                try:
                    viewport = vp_data['viewport']
                    old_title = vp_data['title']
                    new_title = self.GenerateNewTitle(vp_data)
                    
                    # Set new title
                    title_param = viewport.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                    if title_param and not title_param.IsReadOnly:
                        title_param.Set(new_title)
                        success_count += 1
                        output.print_md("✓ {} | {} → {}".format(
                            vp_data['sheet_number'],
                            old_title or "(No Title)",
                            new_title
                        ))
                    else:
                        error_count += 1
                        output.print_md("✗ {} | {} - Parameter read-only".format(
                            vp_data['sheet_number'],
                            vp_data['view_name']
                        ))
                
                except Exception as e:
                    error_count += 1
                    output.print_md("✗ {} | {} - {}".format(
                        vp_data['sheet_number'],
                        vp_data['view_name'],
                        str(e)
                    ))
        
        # Summary
        msg = "View title rename complete!\n\n"
        msg += "Success: {}\n".format(success_count)
        msg += "Errors: {}".format(error_count)
        
        forms.alert(msg, title="Complete")
        
        self.result = True
        self.apply_btn.Enabled = True
        self.apply_btn.Text = "APPLY RENAME"


# ==============================================================================
# SHOW FORM
# ==============================================================================

form = ViewTitleRenamer(all_viewports, doc)
form.ShowDialog()

if form.result:
    output.print_md("### ✅ View Titles Renamed Successfully")
else:
    output.print_md("### Operation Cancelled")