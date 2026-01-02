# -*- coding: utf-8 -*-
"""
Set Mark Values - Modern UI Enhanced
"""
__title__ = 'Set Mark\nValues'
__doc__ = 'Assign mark values to filtered elements by type and level'

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
import System
from System.Windows.Forms import Form, Label, Button, ComboBox, Panel, TextBox, FormBorderStyle, ComboBoxStyle, FlatStyle, BorderStyle, ScrollBars
from System.Drawing import Color, Font, FontStyle, Size, Point

doc = revit.doc
output = script.get_output()

def get_all_levels():
    levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    return sorted(levels, key=lambda l: l.Elevation)

def get_elements_by_category_and_level(category, level):
    all_elements = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
    filtered = []
    for elem in all_elements:
        elem_level_id = None
        
        # Try multiple level parameters in order of priority
        try:
            level_param = elem.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
            if level_param and level_param.HasValue:
                elem_level_id = level_param.AsElementId()
        except:
            pass
        
        if not elem_level_id or elem_level_id == ElementId.InvalidElementId:
            try:
                level_param = elem.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
                if level_param and level_param.HasValue:
                    elem_level_id = level_param.AsElementId()
            except:
                pass
        
        if not elem_level_id or elem_level_id == ElementId.InvalidElementId:
            try:
                level_param = elem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
                if level_param and level_param.HasValue:
                    elem_level_id = level_param.AsElementId()
            except:
                pass
        
        if not elem_level_id or elem_level_id == ElementId.InvalidElementId:
            try:
                level_param = elem.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
                if level_param and level_param.HasValue:
                    elem_level_id = level_param.AsElementId()
            except:
                pass
        
        # For walls, also check base constraint
        if not elem_level_id or elem_level_id == ElementId.InvalidElementId:
            try:
                if hasattr(elem, 'WallType'):  # It's a wall
                    level_param = elem.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
                    if level_param and level_param.HasValue:
                        elem_level_id = level_param.AsElementId()
            except:
                pass
        
        # Check if this element's level matches the target level
        if elem_level_id and elem_level_id == level.Id:
            filtered.append(elem)
    
    return filtered

def group_by_type(elements):
    type_dict = {}
    for elem in elements:
        type_id = elem.GetTypeId()
        if type_id == ElementId.InvalidElementId:
            continue
        elem_type = doc.GetElement(type_id)
        if elem_type:
            type_param = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if type_param:
                type_name_str = type_param.AsString()
                if type_name_str not in type_dict:
                    type_dict[type_name_str] = []
                type_dict[type_name_str].append(elem)
    return type_dict

def get_current_mark(elem):
    mark_param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
    if mark_param and mark_param.HasValue:
        mark_value = mark_param.AsString()
        return mark_value if mark_value else ""
    return ""

class MarkManagerModern(Form):
    def __init__(self):
        self.selected_category = None
        self.selected_level = None
        self.selected_type = None
        self.filtered_elements = []
        self.type_dict = {}
        self.result = False
        self.InitializeUI()
    
    def InitializeUI(self):
        self.Text = "Mark Value Manager"
        self.Width = 1100
        self.Height = 750
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.BackColor = Color.FromArgb(18, 18, 24)
        
        title_font = Font("Segoe UI", 16, FontStyle.Bold)
        section_font = Font("Segoe UI", 10, FontStyle.Bold)
        label_font = Font("Segoe UI", 9)
        code_font = Font("Consolas", 8)
        
        bg_panel = Color.FromArgb(28, 28, 36)
        bg_input = Color.FromArgb(38, 38, 46)
        accent = Color.FromArgb(0, 120, 215)
        text_primary = Color.FromArgb(230, 230, 235)
        text_secondary = Color.FromArgb(160, 160, 170)
        
        # Header
        header = Panel()
        header.Location = Point(15, 15)
        header.Size = Size(1050, 80)
        header.BackColor = bg_panel
        self.Controls.Add(header)
        
        title = Label()
        title.Text = "MARK VALUE MANAGER"
        title.Font = title_font
        title.ForeColor = text_primary
        title.Location = Point(20, 15)
        title.AutoSize = True
        header.Controls.Add(title)
        
        subtitle = Label()
        subtitle.Text = "Assign mark values by category, level, and type"
        subtitle.Font = label_font
        subtitle.ForeColor = text_secondary
        subtitle.Location = Point(20, 45)
        subtitle.AutoSize = True
        header.Controls.Add(subtitle)
        
        # Left Panel
        left_panel = Panel()
        left_panel.Location = Point(15, 110)
        left_panel.Size = Size(520, 530)
        left_panel.BackColor = bg_panel
        self.Controls.Add(left_panel)
        
        Label(Text="STEP 1: SELECT CATEGORY", Font=section_font, ForeColor=text_secondary, Location=Point(20, 15), AutoSize=True, Parent=left_panel)
        
        self.category_combo = ComboBox()
        self.category_combo.Location = Point(20, 45)
        self.category_combo.Size = Size(480, 28)
        self.category_combo.Font = label_font
        self.category_combo.BackColor = bg_input
        self.category_combo.ForeColor = text_primary
        self.category_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.category_combo.FlatStyle = FlatStyle.Flat
        left_panel.Controls.Add(self.category_combo)
        
        self.category_map = {
            'Walls': BuiltInCategory.OST_Walls,
            'Doors': BuiltInCategory.OST_Doors,
            'Windows': BuiltInCategory.OST_Windows,
            'Structural Framing': BuiltInCategory.OST_StructuralFraming,
            'Structural Columns': BuiltInCategory.OST_StructuralColumns,
            'Floors': BuiltInCategory.OST_Floors,
            'Furniture': BuiltInCategory.OST_Furniture,
            'Generic Models': BuiltInCategory.OST_GenericModel,
        }
        for cat in sorted(self.category_map.keys()):
            self.category_combo.Items.Add(cat)
        self.category_combo.SelectedIndexChanged += self.OnCategoryChanged
        
        Label(Text="STEP 2: SELECT LEVEL", Font=section_font, ForeColor=text_secondary, Location=Point(20, 90), AutoSize=True, Parent=left_panel)
        
        self.level_combo = ComboBox()
        self.level_combo.Location = Point(20, 120)
        self.level_combo.Size = Size(480, 28)
        self.level_combo.Font = label_font
        self.level_combo.BackColor = bg_input
        self.level_combo.ForeColor = text_primary
        self.level_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.level_combo.FlatStyle = FlatStyle.Flat
        self.level_combo.Enabled = False
        left_panel.Controls.Add(self.level_combo)
        
        self.levels = get_all_levels()
        for level in self.levels:
            self.level_combo.Items.Add("{} (Elev: {:.2f})".format(level.Name, level.Elevation))
        self.level_combo.SelectedIndexChanged += self.OnLevelChanged
        
        Label(Text="STEP 3: SELECT TYPE", Font=section_font, ForeColor=text_secondary, Location=Point(20, 165), AutoSize=True, Parent=left_panel)
        
        self.type_combo = ComboBox()
        self.type_combo.Location = Point(20, 195)
        self.type_combo.Size = Size(480, 28)
        self.type_combo.Font = label_font
        self.type_combo.BackColor = bg_input
        self.type_combo.ForeColor = text_primary
        self.type_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.type_combo.FlatStyle = FlatStyle.Flat
        self.type_combo.Enabled = False
        self.type_combo.SelectedIndexChanged += self.OnTypeChanged
        left_panel.Controls.Add(self.type_combo)
        
        Label(Text="SELECTED ELEMENTS", Font=section_font, ForeColor=text_secondary, Location=Point(20, 240), AutoSize=True, Parent=left_panel)
        
        self.element_count = Label()
        self.element_count.Text = "0 elements"
        self.element_count.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.element_count.ForeColor = accent
        self.element_count.Location = Point(420, 242)
        self.element_count.AutoSize = True
        left_panel.Controls.Add(self.element_count)
        
        self.info_text = TextBox()
        self.info_text.Location = Point(20, 270)
        self.info_text.Size = Size(480, 240)
        self.info_text.Font = code_font
        self.info_text.BackColor = Color.FromArgb(15, 15, 20)
        self.info_text.ForeColor = text_secondary
        self.info_text.BorderStyle = BorderStyle.FixedSingle
        self.info_text.ReadOnly = True
        self.info_text.Multiline = True
        self.info_text.WordWrap = False
        self.info_text.ScrollBars = ScrollBars.Both
        self.info_text.Text = "Select category > level > type"
        left_panel.Controls.Add(self.info_text)
        
        # Right Panel
        right_panel = Panel()
        right_panel.Location = Point(550, 110)
        right_panel.Size = Size(515, 530)
        right_panel.BackColor = bg_panel
        self.Controls.Add(right_panel)
        
        Label(Text="STEP 4: ASSIGN VALUES", Font=section_font, ForeColor=text_secondary, Location=Point(20, 15), AutoSize=True, Parent=right_panel)
        
        Label(Text="Method:", Font=label_font, ForeColor=text_primary, Location=Point(20, 50), Size=Size(80, 20), Parent=right_panel)
        
        self.method_combo = ComboBox()
        self.method_combo.Location = Point(100, 47)
        self.method_combo.Size = Size(395, 28)
        self.method_combo.Font = label_font
        self.method_combo.BackColor = bg_input
        self.method_combo.ForeColor = text_primary
        self.method_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.method_combo.FlatStyle = FlatStyle.Flat
        self.method_combo.Items.Add("Single Value")
        self.method_combo.Items.Add("Sequential Numbers")
        self.method_combo.Items.Add("Sequential with Prefix")
        self.method_combo.Items.Add("Comma-Separated")
        self.method_combo.SelectedIndex = 0
        self.method_combo.SelectedIndexChanged += self.OnMethodChanged
        right_panel.Controls.Add(self.method_combo)
        
        self.input_container = Panel()
        self.input_container.Location = Point(20, 92)
        self.input_container.Size = Size(475, 140)
        self.input_container.BackColor = bg_input
        right_panel.Controls.Add(self.input_container)
        
        # Single value
        self.single_label = Label(Text="Mark Value:", Font=label_font, ForeColor=text_primary, Location=Point(15, 20), AutoSize=True, Parent=self.input_container)
        self.single_text = TextBox(Location=Point(15, 45), Size=Size(445, 28), Font=label_font, BackColor=Color.FromArgb(50, 50, 60), ForeColor=text_primary, BorderStyle=BorderStyle.FixedSingle, Parent=self.input_container)
        
        # Sequential
        self.seq_start_label = Label(Text="Start:", Font=label_font, ForeColor=text_primary, Location=Point(15, 20), AutoSize=True, Visible=False, Parent=self.input_container)
        self.seq_start_text = TextBox(Location=Point(15, 45), Size=Size(150, 28), Font=label_font, BackColor=Color.FromArgb(50, 50, 60), ForeColor=text_primary, BorderStyle=BorderStyle.FixedSingle, Text="1", Visible=False, Parent=self.input_container)
        self.seq_step_label = Label(Text="Step:", Font=label_font, ForeColor=text_primary, Location=Point(185, 20), AutoSize=True, Visible=False, Parent=self.input_container)
        self.seq_step_text = TextBox(Location=Point(185, 45), Size=Size(100, 28), Font=label_font, BackColor=Color.FromArgb(50, 50, 60), ForeColor=text_primary, BorderStyle=BorderStyle.FixedSingle, Text="1", Visible=False, Parent=self.input_container)
        
        # Prefix
        self.prefix_label = Label(Text="Prefix:", Font=label_font, ForeColor=text_primary, Location=Point(15, 85), AutoSize=True, Visible=False, Parent=self.input_container)
        self.prefix_text = TextBox(Location=Point(15, 110), Size=Size(200, 28), Font=label_font, BackColor=Color.FromArgb(50, 50, 60), ForeColor=text_primary, BorderStyle=BorderStyle.FixedSingle, Visible=False, Parent=self.input_container)
        
        # CSV
        self.csv_label = Label(Text="Comma-separated values:", Font=label_font, ForeColor=text_primary, Location=Point(15, 15), Size=Size(445, 20), Visible=False, Parent=self.input_container)
        self.csv_text = TextBox(Location=Point(15, 40), Size=Size(445, 90), Font=code_font, BackColor=Color.FromArgb(50, 50, 60), ForeColor=text_primary, BorderStyle=BorderStyle.FixedSingle, Multiline=True, ScrollBars=ScrollBars.Vertical, Visible=False, Parent=self.input_container)
        
        Label(Text="PREVIEW", Font=section_font, ForeColor=text_secondary, Location=Point(20, 247), AutoSize=True, Parent=right_panel)
        
        self.preview_text = TextBox()
        self.preview_text.Location = Point(20, 277)
        self.preview_text.Size = Size(475, 233)
        self.preview_text.Font = code_font
        self.preview_text.BackColor = Color.FromArgb(15, 15, 20)
        self.preview_text.ForeColor = text_secondary
        self.preview_text.BorderStyle = BorderStyle.FixedSingle
        self.preview_text.ReadOnly = True
        self.preview_text.Multiline = True
        self.preview_text.WordWrap = False
        self.preview_text.ScrollBars = ScrollBars.Both
        self.preview_text.Text = "Click 'Generate Preview'..."
        right_panel.Controls.Add(self.preview_text)
        
        # Buttons
        self.preview_btn = Button()
        self.preview_btn.Text = "GENERATE PREVIEW"
        self.preview_btn.Location = Point(550, 655)
        self.preview_btn.Size = Size(240, 45)
        self.preview_btn.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.preview_btn.BackColor = Color.FromArgb(70, 70, 80)
        self.preview_btn.ForeColor = Color.White
        self.preview_btn.FlatStyle = FlatStyle.Flat
        self.preview_btn.FlatAppearance.BorderSize = 0
        self.preview_btn.Click += self.GeneratePreview
        self.preview_btn.Enabled = False
        self.Controls.Add(self.preview_btn)
        
        self.apply_btn = Button()
        self.apply_btn.Text = "APPLY MARK VALUES"
        self.apply_btn.Location = Point(810, 655)
        self.apply_btn.Size = Size(255, 45)
        self.apply_btn.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.apply_btn.BackColor = accent
        self.apply_btn.ForeColor = Color.White
        self.apply_btn.FlatStyle = FlatStyle.Flat
        self.apply_btn.FlatAppearance.BorderSize = 0
        self.apply_btn.Click += self.ApplyMarks
        self.apply_btn.Enabled = False
        self.Controls.Add(self.apply_btn)
    
    def OnCategoryChanged(self, sender, args):
        self.level_combo.Enabled = True
        self.type_combo.Enabled = False
        self.type_combo.Items.Clear()
        self.filtered_elements = []
        self.UpdateInfoPanel()
        self.preview_btn.Enabled = False
        self.apply_btn.Enabled = False
    
    def OnLevelChanged(self, sender, args):
        if self.category_combo.SelectedIndex < 0 or self.level_combo.SelectedIndex < 0:
            return
        cat_name = self.category_combo.SelectedItem
        self.selected_category = self.category_map[cat_name]
        self.selected_level = self.levels[self.level_combo.SelectedIndex]
        elements = get_elements_by_category_and_level(self.selected_category, self.selected_level)
        if not elements:
            forms.alert("No {} on {}".format(cat_name, self.selected_level.Name))
            return
        self.type_dict = group_by_type(elements)
        if not self.type_dict:
            return
        self.type_combo.Items.Clear()
        for type_name in sorted(self.type_dict.keys()):
            count = len(self.type_dict[type_name])
            self.type_combo.Items.Add("{} ({})".format(type_name, count))
        self.type_combo.Enabled = True
        self.UpdateInfoPanel()
    
    def OnTypeChanged(self, sender, args):
        if self.type_combo.SelectedIndex < 0:
            return
        selected = self.type_combo.SelectedItem
        type_name = selected.split(' (')[0]
        self.selected_type = type_name
        self.filtered_elements = self.type_dict[type_name]
        self.UpdateInfoPanel()
        self.preview_btn.Enabled = True
        self.apply_btn.Enabled = True
    
    def UpdateInfoPanel(self):
        if not self.filtered_elements:
            self.info_text.Text = "Select category > level > type"
            self.element_count.Text = "0 elements"
            return
        count = len(self.filtered_elements)
        self.element_count.Text = "{} element{}".format(count, "s" if count != 1 else "")
        info = "FOUND {} ELEMENTS\n{}\n\n".format(count, "=" * 60)
        for i, elem in enumerate(self.filtered_elements[:20], 1):
            mark = get_current_mark(elem)
            info += "{}. ID: {} | Mark: {}\n".format(i, elem.Id, mark if mark else "(empty)")
        if len(self.filtered_elements) > 20:
            info += "\n... {} more".format(len(self.filtered_elements) - 20)
        self.info_text.Text = info
    
    def OnMethodChanged(self, sender, args):
        method = self.method_combo.SelectedItem
        self.single_label.Visible = self.single_text.Visible = (method == "Single Value")
        self.seq_start_label.Visible = self.seq_start_text.Visible = self.seq_step_label.Visible = self.seq_step_text.Visible = (method in ["Sequential Numbers", "Sequential with Prefix"])
        self.prefix_label.Visible = self.prefix_text.Visible = (method == "Sequential with Prefix")
        self.csv_label.Visible = self.csv_text.Visible = (method == "Comma-Separated")
    
    def GeneratePreview(self, sender, args):
        if not self.filtered_elements:
            return
        new_marks = self.GetNewMarks()
        if not new_marks:
            return
        preview = "PREVIEW: {} ELEMENTS\n{}\n\n".format(len(self.filtered_elements), "=" * 60)
        for i, (elem, new_mark) in enumerate(zip(self.filtered_elements[:15], new_marks[:15]), 1):
            old = get_current_mark(elem)
            preview += "{}. ID: {} | {} -> {}\n".format(i, elem.Id, old if old else "(empty)", new_mark)
        if len(self.filtered_elements) > 15:
            preview += "\n... {} more".format(len(self.filtered_elements) - 15)
        self.preview_text.Text = preview
    
    def GetNewMarks(self):
        method = self.method_combo.SelectedItem
        count = len(self.filtered_elements)
        if method == "Single Value":
            value = self.single_text.Text.strip()
            if not value:
                forms.alert("Enter a value")
                return None
            return [value] * count
        elif method == "Sequential Numbers":
            try:
                start = int(self.seq_start_text.Text)
                step = int(self.seq_step_text.Text)
                return [str(start + i * step) for i in range(count)]
            except:
                forms.alert("Invalid number")
                return None
        elif method == "Sequential with Prefix":
            try:
                prefix = self.prefix_text.Text.strip()
                start = int(self.seq_start_text.Text)
                step = int(self.seq_step_text.Text)
                return ["{}-{}".format(prefix, start + i * step) for i in range(count)]
            except:
                forms.alert("Invalid input")
                return None
        elif method == "Comma-Separated":
            text = self.csv_text.Text.strip()
            if not text:
                forms.alert("Enter values")
                return None
            values = [v.strip() for v in text.replace('\n', ',').split(',') if v.strip()]
            if len(values) != count:
                forms.alert("Mismatch: {} values vs {} elements".format(len(values), count))
                return None
            return values
        return None
    
    def ApplyMarks(self, sender, args):
        if not self.filtered_elements:
            return
        new_marks = self.GetNewMarks()
        if not new_marks:
            return
        if not forms.alert("Apply to {} elements?".format(len(self.filtered_elements)), yes=True, no=True):
            return
        self.apply_btn.Enabled = False
        self.apply_btn.Text = "APPLYING..."
        success = 0
        errors = 0
        t = Transaction(doc, "Set Mark Values")
        t.Start()
        try:
            for elem, mark in zip(self.filtered_elements, new_marks):
                try:
                    param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                    if param and not param.IsReadOnly:
                        param.Set(mark)
                        success += 1
                    else:
                        errors += 1
                except:
                    errors += 1
            t.Commit()
            forms.alert("Updated!\nSuccess: {}\nErrors: {}".format(success, errors))
            self.result = True
            self.Close()
        except Exception as e:
            t.RollBack()
            forms.alert("Failed: {}".format(str(e)))
            self.apply_btn.Enabled = True
            self.apply_btn.Text = "APPLY MARK VALUES"

if __name__ == '__main__':
    form = MarkManagerModern()
    form.ShowDialog()
    if form.result:
        output.print_md("## Mark Values Updated")
        output.print_md("**Elements:** {}".format(len(form.filtered_elements)))