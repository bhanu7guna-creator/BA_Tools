# -*- coding: utf-8 -*-
"""
Export NWC Files - Modern Windows Forms UI
"""
__title__ = 'Export\nNWC'
__doc__ = 'Export Navisworks Cache files with modern UI'

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
import os
import System
from System.Windows.Forms import (
    Form, Label, ComboBox, TextBox, Button, CheckBox, 
    FolderBrowserDialog, DialogResult, Panel,
    FormBorderStyle, DockStyle, AnchorStyles, BorderStyle,
    ComboBoxStyle, FlatStyle
)
from System.Drawing import (
    Color, Font, FontStyle, Size, Point, ContentAlignment, SystemFonts
)

doc = revit.doc
output = script.get_output()


# ==============================================================================
# COLLECT 3D VIEWS
# ==============================================================================

all_views = FilteredElementCollector(doc)\
    .OfClass(View3D)\
    .WhereElementIsNotElementType()\
    .ToElements()

valid_3d_views = [view for view in all_views if not view.IsTemplate]

if not valid_3d_views:
    forms.alert('No 3D views found in the project', exitscript=True)


# ==============================================================================
# MODERN WINDOWS FORMS UI
# ==============================================================================

class ModernNWCExporter(Form):
    def __init__(self, views):
        self.views = views
        self.view_dict = {view.Name: view for view in views}
        self.selected_view = None
        self.export_path = "C:\\Projects\\BIM\\Exports\\NWC"
        self.file_name = doc.Title or "Untitled"
        self.result = None
        
        self.InitializeUI()
    
    def InitializeUI(self):
        # Form properties
        self.Text = "NWC EXTRACTOR - Revit Plugin v2.4.0"
        self.Width = 900
        self.Height = 600
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
        self.bg_dark = Color.FromArgb(20, 20, 28)
        self.bg_panel = Color.FromArgb(30, 30, 40)
        self.border_color = Color.FromArgb(100, 150, 200)
        self.accent_color = Color.FromArgb(74, 144, 226)
        self.text_color = Color.FromArgb(220, 220, 220)
        self.text_muted = Color.FromArgb(140, 140, 150)
        
        y_offset = 20
        
        # ===== HEADER =====
        header_panel = Panel()
        header_panel.Location = Point(20, y_offset)
        header_panel.Size = Size(520, 80)
        header_panel.BackColor = self.bg_panel
        self.Controls.Add(header_panel)
        
        # Title
        title = Label()
        title.Text = "NWC EXTRACTOR"
        title.Font = self.title_font
        title.ForeColor = Color.White
        title.Location = Point(20, 15)
        title.Size = Size(300, 30)
        header_panel.Controls.Add(title)
        
        subtitle = Label()
        subtitle.Text = "REVIT PLUGIN v2.4.0"
        subtitle.Font = Font("Consolas", 8, FontStyle.Regular)
        subtitle.ForeColor = self.text_muted
        subtitle.Location = Point(20, 45)
        subtitle.Size = Size(300, 20)
        header_panel.Controls.Add(subtitle)
        
        y_offset += 100
        
        # ===== SOURCE VIEW =====
        view_label = Label()
        view_label.Text = "SOURCE VIEW"
        view_label.Font = self.label_font
        view_label.ForeColor = self.text_muted
        view_label.Location = Point(20, y_offset)
        view_label.Size = Size(520, 20)
        self.Controls.Add(view_label)
        
        y_offset += 25
        
        self.view_combo = ComboBox()
        self.view_combo.Location = Point(20, y_offset)
        self.view_combo.Size = Size(520, 30)
        self.view_combo.Font = self.input_font
        self.view_combo.BackColor = Color.FromArgb(15, 15, 20)
        self.view_combo.ForeColor = self.text_color
        self.view_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.view_combo.FlatStyle = FlatStyle.Flat
        
        self.view_combo.Items.Add("-- Select 3D View --")
        for view in sorted(self.view_dict.keys()):
            self.view_combo.Items.Add(view)
        self.view_combo.SelectedIndex = 0
        
        self.Controls.Add(self.view_combo)
        
        y_offset += 50
        
        # ===== OUTPUT DIRECTORY =====
        path_label = Label()
        path_label.Text = "OUTPUT DIRECTORY"
        path_label.Font = self.label_font
        path_label.ForeColor = self.text_muted
        path_label.Location = Point(20, y_offset)
        path_label.Size = Size(520, 20)
        self.Controls.Add(path_label)
        
        y_offset += 25
        
        self.path_textbox = TextBox()
        self.path_textbox.Location = Point(20, y_offset)
        self.path_textbox.Size = Size(410, 30)
        self.path_textbox.Font = self.input_font
        self.path_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.path_textbox.ForeColor = self.text_color
        self.path_textbox.BorderStyle = BorderStyle.FixedSingle
        self.path_textbox.Text = self.export_path
        self.path_textbox.ReadOnly = True
        self.Controls.Add(self.path_textbox)
        
        browse_btn = Button()
        browse_btn.Text = "BROWSE..."
        browse_btn.Location = Point(440, y_offset)
        browse_btn.Size = Size(100, 30)
        browse_btn.Font = Font("Consolas", 9, FontStyle.Bold)
        browse_btn.BackColor = Color.FromArgb(60, 60, 80)
        browse_btn.ForeColor = Color.White
        browse_btn.FlatStyle = FlatStyle.Flat
        browse_btn.FlatAppearance.BorderColor = self.border_color
        browse_btn.Click += self.BrowsePath
        self.Controls.Add(browse_btn)
        
        y_offset += 50
        
        # ===== FILENAME =====
        file_label = Label()
        file_label.Text = "FILENAME"
        file_label.Font = self.label_font
        file_label.ForeColor = self.text_muted
        file_label.Location = Point(20, y_offset)
        file_label.Size = Size(520, 20)
        self.Controls.Add(file_label)
        
        y_offset += 25
        
        self.filename_textbox = TextBox()
        self.filename_textbox.Location = Point(20, y_offset)
        self.filename_textbox.Size = Size(520, 30)
        self.filename_textbox.Font = self.input_font
        self.filename_textbox.BackColor = Color.FromArgb(15, 15, 20)
        self.filename_textbox.ForeColor = self.text_color
        self.filename_textbox.BorderStyle = BorderStyle.FixedSingle
        self.filename_textbox.Text = self.file_name
        self.Controls.Add(self.filename_textbox)
        
        y_offset += 40
        
        self.use_project_name = CheckBox()
        self.use_project_name.Text = "Use Project Name ({})".format(self.file_name)
        self.use_project_name.Location = Point(20, y_offset)
        self.use_project_name.Size = Size(520, 25)
        self.use_project_name.Font = Font("Consolas", 9, FontStyle.Regular)
        self.use_project_name.ForeColor = self.text_muted
        self.use_project_name.Checked = True
        self.use_project_name.CheckedChanged += self.ToggleProjectName
        self.Controls.Add(self.use_project_name)
        
        y_offset += 45
        
        # ===== EXPORT BUTTON =====
        self.export_btn = Button()
        self.export_btn.Text = "EXPORT NWC FILE"
        self.export_btn.Location = Point(20, y_offset)
        self.export_btn.Size = Size(520, 50)
        self.export_btn.Font = Font("Consolas", 11, FontStyle.Bold)
        self.export_btn.BackColor = self.accent_color
        self.export_btn.ForeColor = Color.White
        self.export_btn.FlatStyle = FlatStyle.Flat
        self.export_btn.FlatAppearance.BorderSize = 0
        self.export_btn.Click += self.ExportNWC
        self.Controls.Add(self.export_btn)
        
        # ===== LOG PANEL =====
        log_panel = Panel()
        log_panel.Location = Point(560, 20)
        log_panel.Size = Size(300, 500)
        log_panel.BackColor = Color.FromArgb(15, 15, 20)
        log_panel.BorderStyle = BorderStyle.FixedSingle
        self.Controls.Add(log_panel)
        
        log_header = Label()
        log_header.Text = "SYSTEM LOG"
        log_header.Font = Font("Consolas", 9, FontStyle.Bold)
        log_header.ForeColor = self.text_muted
        log_header.Location = Point(10, 10)
        log_header.Size = Size(280, 20)
        log_panel.Controls.Add(log_header)
        
        self.log_textbox = TextBox()
        self.log_textbox.Location = Point(10, 40)
        self.log_textbox.Size = Size(280, 420)
        self.log_textbox.Font = Font("Consolas", 9, FontStyle.Regular)
        self.log_textbox.BackColor = Color.FromArgb(10, 10, 15)
        self.log_textbox.ForeColor = self.text_muted
        self.log_textbox.BorderStyle = BorderStyle.None
        self.log_textbox.ReadOnly = True
        self.log_textbox.Multiline = True
        self.log_textbox.WordWrap = True
        self.log_textbox.Text = "Ready for export..."
        log_panel.Controls.Add(self.log_textbox)
        
        status_label = Label()
        status_label.Text = "STATUS: IDLE"
        status_label.Font = Font("Consolas", 8, FontStyle.Regular)
        status_label.ForeColor = self.text_muted
        status_label.Location = Point(10, 470)
        status_label.Size = Size(280, 20)
        log_panel.Controls.Add(status_label)
        self.status_label = status_label
    
    def AddLog(self, message):
        timestamp = System.DateTime.Now.ToString("HH:mm:ss")
        log_line = "[{}] {}\r\n".format(timestamp, message)
        self.log_textbox.AppendText(log_line)
        # Auto scroll to bottom
        self.log_textbox.SelectionStart = len(self.log_textbox.Text)
        self.log_textbox.ScrollToCaret()
    
    def BrowsePath(self, sender, args):
        dialog = FolderBrowserDialog()
        dialog.Description = "Select Export Folder"
        dialog.SelectedPath = self.export_path
        
        if dialog.ShowDialog() == DialogResult.OK:
            self.export_path = dialog.SelectedPath
            self.path_textbox.Text = self.export_path
            self.AddLog("Export path updated: {}".format(self.export_path))
    
    def ToggleProjectName(self, sender, args):
        if self.use_project_name.Checked:
            self.filename_textbox.Text = self.file_name
            self.filename_textbox.ReadOnly = True
            self.filename_textbox.BackColor = Color.FromArgb(25, 25, 30)
        else:
            self.filename_textbox.ReadOnly = False
            self.filename_textbox.BackColor = Color.FromArgb(15, 15, 20)
    
    def ExportNWC(self, sender, args):
        # Validate
        if self.view_combo.SelectedIndex == 0:
            forms.alert("Please select a 3D view", title="Validation Error")
            return
        
        selected_view_name = self.view_combo.SelectedItem.ToString()
        self.selected_view = self.view_dict[selected_view_name]
        
        file_name = self.filename_textbox.Text.strip()
        if not file_name:
            forms.alert("Please enter a file name", title="Validation Error")
            return
        
        # Remove invalid characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            file_name = file_name.replace(char, '_')
        
        nwc_file_path = os.path.join(self.export_path, file_name + '.nwc')
        
        # Update UI
        self.export_btn.Enabled = False
        self.export_btn.Text = "EXPORTING..."
        self.status_label.Text = "STATUS: BUSY"
        self.status_label.ForeColor = Color.FromArgb(74, 222, 128)
        
        self.AddLog("Starting export process...")
        self.AddLog("View: {}".format(selected_view_name))
        self.AddLog("Output: {}".format(nwc_file_path))
        
        try:
            # Configure NWC export options
            nwc_options = NavisworksExportOptions()
            nwc_options.ExportScope = NavisworksExportScope.View
            nwc_options.ViewId = self.selected_view.Id
            nwc_options.ExportLinks = False
            nwc_options.ExportRoomAsAttribute = True
            nwc_options.ExportRoomGeometry = True
            nwc_options.ConvertElementProperties = True
            nwc_options.Coordinates = NavisworksCoordinates.Shared
            nwc_options.FacetingFactor = 1.0
            nwc_options.ExportUrls = True
            nwc_options.ConvertLights = True
            
            self.AddLog("Configuring export settings...")
            self.AddLog("Export scope: Selected View")
            self.AddLog("Coordinates: Shared")
            
            # Perform export
            self.AddLog("Processing geometry...")
            result = doc.Export(self.export_path, file_name + '.nwc', nwc_options)
            
            if result and os.path.exists(nwc_file_path):
                file_size = os.path.getsize(nwc_file_path)
                file_size_mb = file_size / (1024.0 * 1024.0)
                
                self.AddLog("SUCCESS: Export completed!")
                self.AddLog("File size: {:.2f} MB".format(file_size_mb))
                
                forms.alert(
                    "NWC file exported successfully!\n\n"
                    "File: {}.nwc\n"
                    "Location: {}\n"
                    "Size: {:.2f} MB".format(file_name, self.export_path, file_size_mb),
                    title="Export Complete"
                )
                
                self.result = True
                self.Close()
            else:
                self.AddLog("ERROR: Export failed")
                forms.alert("Export failed. Please check the log.", title="Export Error")
        
        except Exception as e:
            self.AddLog("ERROR: {}".format(str(e)))
            forms.alert("Export failed:\n\n{}".format(str(e)), title="Export Error")
        
        finally:
            self.export_btn.Enabled = True
            self.export_btn.Text = "EXPORT NWC FILE"
            self.status_label.Text = "STATUS: IDLE"
            self.status_label.ForeColor = self.text_muted


# ==============================================================================
# SHOW FORM
# ==============================================================================

form = ModernNWCExporter(valid_3d_views)
form.ShowDialog()

if form.result:
    output.print_md("### ✅ Export Completed Successfully")
else:
    output.print_md("### Export Cancelled or Failed")