# -*- coding: utf-8 -*-
"""
PDF & DWG Export Manager - WPF Modern UI
Export production and shop drawing sets to PDF and DWG
"""
__title__ = 'Export\nPDF/DWG'
__doc__ = 'Export sheets to PDF and DWG with custom naming and settings'

import clr
import os
clr.AddReference("RevitAPI")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
import System
from System.Windows.Markup import XamlReader
from System.Windows import Window, Visibility
from System.IO import StreamReader
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List
from System.Windows.Forms import FolderBrowserDialog, DialogResult

doc = revit.doc
output = script.get_output()


# ==============================================================================
# DATA CLASSES
# ==============================================================================

class SheetItem(System.ComponentModel.INotifyPropertyChanged):
    """Observable sheet item for checkbox binding"""
    
    def __init__(self, sheet):
        self._sheet = sheet
        self._sheet_number = sheet.SheetNumber
        self._sheet_name = sheet.Name
        self._display_text = "{} - {}".format(self._sheet_number, self._sheet_name)
        self._is_selected = False
        self._property_changed = None
    
    @property
    def Sheet(self):
        return self._sheet
    
    @property
    def SheetNumber(self):
        return self._sheet_number
    
    @property
    def SheetName(self):
        return self._sheet_name
    
    @property
    def DisplayText(self):
        return self._display_text
    
    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("IsSelected")
    
    def add_PropertyChanged(self, handler):
        self._property_changed = System.Delegate.Combine(self._property_changed, handler)
    
    def remove_PropertyChanged(self, handler):
        self._property_changed = System.Delegate.Remove(self._property_changed, handler)
    
    def OnPropertyChanged(self, property_name):
        if self._property_changed:
            args = System.ComponentModel.PropertyChangedEventArgs(property_name)
            self._property_changed(self, args)


# ==============================================================================
# DATA COLLECTION FUNCTIONS
# ==============================================================================

def get_all_sheets():
    """Get all sheets in the project"""
    sheets = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Sheets)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Sort by sheet number
    return sorted(sheets, key=lambda s: s.SheetNumber)


def get_view_sets():
    """Get all view sheet sets"""
    view_sets = FilteredElementCollector(doc)\
        .OfClass(ViewSheetSet)\
        .ToElements()
    
    return sorted(view_sets, key=lambda v: v.Name)


def get_dwg_export_setups():
    """Get all DWG export setups"""
    try:
        export_setups = BaseExportOptions.GetPredefinedSetupNames(doc)
        return list(export_setups) if export_setups else []
    except:
        return []


# Collect data
all_sheets = get_all_sheets()
view_sets = get_view_sets()
dwg_setups = get_dwg_export_setups()

if not all_sheets:
    forms.alert('No sheets found in the project', exitscript=True)


# ==============================================================================
# WPF WINDOW CLASS
# ==============================================================================

class ExportPDFDWGWindow(Window):
    def __init__(self, xaml_path, sheets, viewsets, dwg_setups, document):
        # Load XAML
        stream = StreamReader(xaml_path)
        self._window = XamlReader.Load(stream.BaseStream)
        stream.Close()
        
        # Data
        self.all_sheets = sheets
        self.viewsets = viewsets
        self.dwg_setups = dwg_setups
        self.doc = document
        self.sheet_items = ObservableCollection[SheetItem]()
        self.filtered_sheet_items = ObservableCollection[SheetItem]()
        self.result = False
        
        # Default paths
        self.pdf_path = "C:\\Projects\\BIM\\Exports\\PDF"
        self.dwg_path = "C:\\Projects\\BIM\\Exports\\DWG"
        
        # Get controls
        self.btnClose = self._window.FindName("btnClose")
        self.cmbViewSet = self._window.FindName("cmbViewSet")
        self.txtSheetFilter = self._window.FindName("txtSheetFilter")
        self.btnSelectAll = self._window.FindName("btnSelectAll")
        self.btnClearAll = self._window.FindName("btnClearAll")
        self.txtSheetCount = self._window.FindName("txtSheetCount")
        self.lstSheets = self._window.FindName("lstSheets")
        
        # Export format checkboxes
        self.chkExportPDF = self._window.FindName("chkExportPDF")
        self.chkExportDWG = self._window.FindName("chkExportDWG")
        
        # PDF settings
        self.txtPDFPath = self._window.FindName("txtPDFPath")
        self.btnBrowsePDF = self._window.FindName("btnBrowsePDF")
        self.cmbPDFNaming = self._window.FindName("cmbPDFNaming")
        self.chkPDFCombine = self._window.FindName("chkPDFCombine")
        
        # DWG settings
        self.txtDWGPath = self._window.FindName("txtDWGPath")
        self.btnBrowseDWG = self._window.FindName("btnBrowseDWG")
        self.cmbDWGNaming = self._window.FindName("cmbDWGNaming")
        self.cmbDWGSetup = self._window.FindName("cmbDWGSetup")
        self.cmbDWGVersion = self._window.FindName("cmbDWGVersion")
        
        # Preview and actions
        self.txtNamingPreview = self._window.FindName("txtNamingPreview")
        self.txtStatus = self._window.FindName("txtStatus")
        self.btnExport = self._window.FindName("btnExport")
        
        # Initialize
        self.InitializeData()
        self.SetupEventHandlers()
    
    def InitializeData(self):
        """Initialize all dropdowns and data"""
        # View sets
        self.cmbViewSet.Items.Add("(All Sheets)")
        for viewset in self.viewsets:
            self.cmbViewSet.Items.Add(viewset.Name)
        self.cmbViewSet.SelectedIndex = 0
        
        # File naming options
        naming_options = [
            "Sheet Number Only",
            "Sheet Number - Sheet Name",
            "Sheet Name - Sheet Number",
            "Sheet Number_Sheet Name",
            "Project Number-Sheet Number_Revision"
        ]
        for option in naming_options:
            self.cmbPDFNaming.Items.Add(option)
            self.cmbDWGNaming.Items.Add(option)
        self.cmbPDFNaming.SelectedIndex = 1  # Default: Number - Name
        self.cmbDWGNaming.SelectedIndex = 1
        
        # DWG export setups
        if self.dwg_setups:
            for setup in self.dwg_setups:
                self.cmbDWGSetup.Items.Add(setup)
            self.cmbDWGSetup.SelectedIndex = 0
        else:
            self.cmbDWGSetup.Items.Add("(Default)")
            self.cmbDWGSetup.SelectedIndex = 0
        
        # AutoCAD versions
        versions = [
            "AutoCAD 2018",
            "AutoCAD 2013",
            "AutoCAD 2010",
            "AutoCAD 2007"
        ]
        for version in versions:
            self.cmbDWGVersion.Items.Add(version)
        self.cmbDWGVersion.SelectedIndex = 0
        
        # Set default paths
        self.txtPDFPath.Text = self.pdf_path
        self.txtDWGPath.Text = self.dwg_path
        
        # Populate sheets
        for sheet in self.all_sheets:
            sheet_item = SheetItem(sheet)
            sheet_item.PropertyChanged += self.OnSheetSelectionChanged
            self.sheet_items.Add(sheet_item)
            self.filtered_sheet_items.Add(sheet_item)
        
        self.lstSheets.ItemsSource = self.filtered_sheet_items
        self.UpdateSheetCount()
    
    def SetupEventHandlers(self):
        """Setup event handlers"""
        self.btnClose.Click += self.OnClose
        self.cmbViewSet.SelectionChanged += self.OnViewSetChanged
        self.txtSheetFilter.TextChanged += self.OnFilterChanged
        self.btnSelectAll.Click += self.OnSelectAll
        self.btnClearAll.Click += self.OnClearAll
        self.btnBrowsePDF.Click += self.OnBrowsePDF
        self.btnBrowseDWG.Click += self.OnBrowseDWG
        self.cmbPDFNaming.SelectionChanged += self.UpdateNamingPreview
        self.cmbDWGNaming.SelectionChanged += self.UpdateNamingPreview
        self.btnExport.Click += self.OnExport
    
    def OnClose(self, sender, args):
        """Close window"""
        self._window.DialogResult = False
        self._window.Close()
    
    def OnViewSetChanged(self, sender, args):
        """Filter sheets by view set"""
        if self.cmbViewSet.SelectedIndex == 0:
            # All sheets
            self.FilterSheets()
        else:
            # Specific view set
            viewset_name = self.cmbViewSet.SelectedItem.ToString()
            viewset = next((v for v in self.viewsets if v.Name == viewset_name), None)
            
            if viewset:
                viewset_sheet_ids = set([vs.Id for vs in viewset.Views])
                self.FilterSheets(viewset_sheet_ids)
    
    def FilterSheets(self, viewset_sheet_ids=None):
        """Filter sheets by view set and search text"""
        filter_text = self.txtSheetFilter.Text.lower() if self.txtSheetFilter.Text else ""
        
        self.filtered_sheet_items.Clear()
        
        for sheet_item in self.sheet_items:
            # Check view set filter
            if viewset_sheet_ids is not None and sheet_item.Sheet.Id not in viewset_sheet_ids:
                continue
            
            # Check text filter
            if filter_text and filter_text not in sheet_item.DisplayText.lower():
                continue
            
            self.filtered_sheet_items.Add(sheet_item)
        
        self.UpdateSheetCount()
    
    def OnFilterChanged(self, sender, args):
        """When filter text changes"""
        self.OnViewSetChanged(None, None)  # Re-apply all filters
    
    def OnSelectAll(self, sender, args):
        """Select all visible sheets"""
        for item in self.filtered_sheet_items:
            item.IsSelected = True
        self.UpdateSheetCount()
        self.UpdateNamingPreview(None, None)
    
    def OnClearAll(self, sender, args):
        """Clear all selections"""
        for item in self.sheet_items:
            item.IsSelected = False
        self.UpdateSheetCount()
        self.UpdateNamingPreview(None, None)
    
    def OnSheetSelectionChanged(self, sender, args):
        """When sheet selection changes"""
        self.UpdateSheetCount()
        self.UpdateNamingPreview(None, None)
    
    def UpdateSheetCount(self):
        """Update sheet count label"""
        selected_count = sum(1 for item in self.sheet_items if item.IsSelected)
        total_count = len(self.filtered_sheet_items)
        self.txtSheetCount.Text = "{} / {} selected".format(selected_count, total_count)
        
        if selected_count > 0:
            self.txtStatus.Text = "Ready to export {} sheet(s)".format(selected_count)
        else:
            self.txtStatus.Text = "Select sheets to export"
    
    def OnBrowsePDF(self, sender, args):
        """Browse for PDF output folder"""
        dialog = FolderBrowserDialog()
        dialog.Description = "Select PDF Export Folder"
        dialog.SelectedPath = self.pdf_path
        
        if dialog.ShowDialog() == DialogResult.OK:
            self.pdf_path = dialog.SelectedPath
            self.txtPDFPath.Text = self.pdf_path
    
    def OnBrowseDWG(self, sender, args):
        """Browse for DWG output folder"""
        dialog = FolderBrowserDialog()
        dialog.Description = "Select DWG Export Folder"
        dialog.SelectedPath = self.dwg_path
        
        if dialog.ShowDialog() == DialogResult.OK:
            self.dwg_path = dialog.SelectedPath
            self.txtDWGPath.Text = self.dwg_path
    
    def GetFileName(self, sheet, naming_format):
        """Generate filename based on naming format"""
        sheet_num = sheet.SheetNumber
        sheet_name = sheet.Name

        # Get project number
        project_num = ""
        try:
            project_info = FilteredElementCollector(self.doc)\
                .OfCategory(BuiltInCategory.OST_ProjectInformation)\
                .FirstElement()
            if project_info:
                project_num_param = project_info.LookupParameter("Project Number")
                if project_num_param and project_num_param.HasValue:
                    project_num = project_num_param.AsString()
        except:
            pass

        # Get revision information
        revision = ""
        try:
            # Try to get the current revision from the sheet
            current_revision_param = sheet.LookupParameter("Current Revision")
            if current_revision_param and current_revision_param.HasValue:
                revision_value = current_revision_param.AsString()
                if revision_value:
                    # Extract only the last digit from revision (e.g., "00" -> "0", "01" -> "1", "03" -> "3")
                    import re
                    digits = re.findall(r'\d', revision_value)
                    if digits:
                        revision = "_R{}".format(digits[-1])

            # Alternative: try "Sheet Issue Date" or "Revision Number" parameters
            if not revision:
                revision_param = sheet.LookupParameter("Revision Number")
                if revision_param and revision_param.HasValue:
                    revision_value = revision_param.AsString()
                    if revision_value:
                        # Extract only the last digit from revision
                        import re
                        digits = re.findall(r'\d', revision_value)
                        if digits:
                            revision = "_R{}".format(digits[-1])
        except:
            pass  # If revision lookup fails, continue without it

        # Remove invalid filename characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            sheet_num = sheet_num.replace(char, '_')
            sheet_name = sheet_name.replace(char, '_')
            project_num = project_num.replace(char, '_')
            revision = revision.replace(char, '_')

        if naming_format == "Sheet Number Only":
            return "{}{}".format(sheet_num, revision)
        elif naming_format == "Sheet Number - Sheet Name":
            return "{} - {}{}".format(sheet_num, sheet_name, revision)
        elif naming_format == "Sheet Name - Sheet Number":
            return "{} - {}{}".format(sheet_name, sheet_num, revision)
        elif naming_format == "Sheet Number_Sheet Name":
            return "{}_{}{}".format(sheet_num, sheet_name, revision)
        elif naming_format == "Project Number-Sheet Number_Revision":
            return "{}-{}{}".format(project_num, sheet_num, revision)
        else:
            return "{}{}".format(sheet_num, revision)
    
    def UpdateNamingPreview(self, sender, args):
        """Update file naming preview"""
        selected_items = [item for item in self.sheet_items if item.IsSelected]
        
        if not selected_items:
            self.txtNamingPreview.Text = "Select sheets to see naming preview..."
            return
        
        pdf_naming = self.cmbPDFNaming.SelectedItem.ToString() if self.cmbPDFNaming.SelectedItem else ""
        dwg_naming = self.cmbDWGNaming.SelectedItem.ToString() if self.cmbDWGNaming.SelectedItem else ""
        
        preview = "NAMING PREVIEW (first 3 sheets):\n\n"
        
        for i, item in enumerate(selected_items[:3], 1):
            preview += "Sheet: {}\n".format(item.DisplayText)
            
            if self.chkExportPDF.IsChecked:
                pdf_name = self.GetFileName(item.Sheet, pdf_naming)
                preview += "  PDF: {}.pdf\n".format(pdf_name)
            
            if self.chkExportDWG.IsChecked:
                dwg_name = self.GetFileName(item.Sheet, dwg_naming)
                preview += "  DWG: {}.dwg\n".format(dwg_name)
            
            preview += "\n"
        
        if len(selected_items) > 3:
            preview += "... and {} more sheet(s)".format(len(selected_items) - 3)
        
        self.txtNamingPreview.Text = preview
    
    def OnExport(self, sender, args):
        """Export selected sheets"""
        # Validation
        selected_sheets = [item for item in self.sheet_items if item.IsSelected]
        
        if not selected_sheets:
            forms.alert("Please select at least one sheet", title="Validation Error")
            return
        
        if not self.chkExportPDF.IsChecked and not self.chkExportDWG.IsChecked:
            forms.alert("Please select at least one export format (PDF or DWG)", 
                       title="Validation Error")
            return
        
        # Confirm
        confirm_msg = "Export {} sheet(s)?\n\n".format(len(selected_sheets))
        if self.chkExportPDF.IsChecked:
            confirm_msg += "• PDF to: {}\n".format(self.pdf_path)
        if self.chkExportDWG.IsChecked:
            confirm_msg += "• DWG to: {}\n".format(self.dwg_path)
        
        if not forms.alert(confirm_msg, yes=True, no=True, title="Confirm Export"):
            return
        
        # Update UI
        self.btnExport.IsEnabled = False
        self.btnExport.Content = "EXPORTING..."
        self.txtStatus.Text = "Exporting files..."
        
        pdf_success = 0
        pdf_errors = 0
        dwg_success = 0
        dwg_errors = 0
        
        try:
            # Export PDF
            if self.chkExportPDF.IsChecked:
                pdf_success, pdf_errors = self.ExportPDF(selected_sheets)
            
            # Export DWG
            if self.chkExportDWG.IsChecked:
                dwg_success, dwg_errors = self.ExportDWG(selected_sheets)
            
            # Summary
            summary = "Export complete!\n\n"
            if self.chkExportPDF.IsChecked:
                summary += "PDF - Success: {}, Errors: {}\n".format(pdf_success, pdf_errors)
            if self.chkExportDWG.IsChecked:
                summary += "DWG - Success: {}, Errors: {}\n".format(dwg_success, dwg_errors)
            summary += "\nCheck output window for details."
            
            forms.alert(summary, title="Export Complete")
            
            output.print_md("### ✅ Export Complete")
            if self.chkExportPDF.IsChecked:
                output.print_md("**PDF** - Success: {}, Errors: {}".format(pdf_success, pdf_errors))
            if self.chkExportDWG.IsChecked:
                output.print_md("**DWG** - Success: {}, Errors: {}".format(dwg_success, dwg_errors))
            
            self.result = True
            self._window.Close()
        
        except Exception as e:
            forms.alert("Export failed:\n\n{}".format(str(e)), title="Export Error")
            output.print_md("### ✗ Export Failed")
            output.print_md("**Error**: {}".format(str(e)))
        
        finally:
            self.btnExport.IsEnabled = True
            self.btnExport.Content = "EXPORT FILES"
    
    def ExportPDF(self, selected_sheets):
        """Export sheets to PDF using PDFExportOptions (same as working Dynamo script)"""
        success = 0
        errors = 0
        
        pdf_naming = self.cmbPDFNaming.SelectedItem.ToString()
        
        # Create export directory if it doesn't exist
        if not os.path.exists(self.pdf_path):
            try:
                os.makedirs(self.pdf_path)
            except Exception as e:
                output.print_md("✗ Failed to create PDF directory: {}".format(str(e)))
                return 0, len(selected_sheets)
        
        if self.chkPDFCombine.IsChecked:
            # Combined PDF - all sheets in one file
            try:
                # Create fresh PDF Export Options for combined
                opts = PDFExportOptions()
                opts.ExportQuality = PDFExportQualityType.DPI600
                opts.PaperFormat = ExportPaperFormat.Default
                opts.ZoomType = ZoomType.Zoom
                opts.ZoomPercentage = 100
                opts.HideCropBoundaries = True
                opts.HideReferencePlane = True
                opts.HideScopeBoxes = True
                opts.HideUnreferencedViewTags = True
                opts.MaskCoincidentLines = True
                opts.ViewLinksInBlue = False
                opts.ColorDepth = ColorDepthType.Color
                opts.StopOnError = False
                
                combined_name = "Combined_Sheets_{}".format(
                    System.DateTime.Now.ToString("yyyyMMdd_HHmmss")
                )
                
                opts.Combine = True
                opts.FileName = combined_name
                
                # Create list of sheet IDs - EXACTLY like Dynamo
                sheetId = List[ElementId](item.Sheet.Id for item in selected_sheets)
                
                # Export
                result = self.doc.Export(self.pdf_path, sheetId, opts)
                
                if result:
                    output.print_md("✓ PDF: {}.pdf (combined {} sheets)".format(combined_name, len(selected_sheets)))
                    success = len(selected_sheets)
                else:
                    output.print_md("✗ PDF Combined: Export returned False")
                    errors = len(selected_sheets)
                
            except Exception as e:
                output.print_md("✗ PDF Combined: {}".format(str(e)))
                import traceback
                output.print_md("Traceback: {}".format(traceback.format_exc()))
                errors = len(selected_sheets)
        else:
            # Individual PDFs - one file per sheet
            for item in selected_sheets:
                try:
                    # Create FRESH PDF Export Options for EACH sheet (critical!)
                    opts = PDFExportOptions()
                    opts.ExportQuality = PDFExportQualityType.DPI600
                    opts.PaperFormat = ExportPaperFormat.Default
                    opts.ZoomType = ZoomType.Zoom
                    opts.ZoomPercentage = 100
                    opts.HideCropBoundaries = True
                    opts.HideReferencePlane = True
                    opts.HideScopeBoxes = True
                    opts.HideUnreferencedViewTags = True
                    opts.MaskCoincidentLines = True
                    opts.ViewLinksInBlue = False
                    opts.ColorDepth = ColorDepthType.Color
                    opts.StopOnError = False
                    
                    filename = self.GetFileName(item.Sheet, pdf_naming)
                    
                    # Remove .pdf extension if present (Revit adds it automatically)
                    if filename.lower().endswith('.pdf'):
                        filename = filename[:-4]
                    
                    opts.FileName = filename
                    # Don't set Combine property for individual exports
                    
                    # Create list with single sheet ID
                    sheetId = List[ElementId]()
                    sheetId.Add(item.Sheet.Id)
                    
                    # Export
                    result = self.doc.Export(self.pdf_path, sheetId, opts)
                    
                    if result:
                        output.print_md("✓ PDF: {} - {}.pdf".format(item.SheetNumber, filename))
                        success += 1
                    else:
                        output.print_md("✗ PDF: {} - Export returned False".format(item.SheetNumber))
                        errors += 1
                    
                except Exception as e:
                    output.print_md("✗ PDF: {} - {}".format(item.SheetNumber, str(e)))
                    import traceback
                    output.print_md("Details: {}".format(traceback.format_exc()))
                    errors += 1
        
        return success, errors
    
    def ExportDWG(self, selected_sheets):
        """Export sheets to DWG"""
        success = 0
        errors = 0
        
        dwg_naming = self.cmbDWGNaming.SelectedItem.ToString()
        
        # Create export directory if it doesn't exist
        if not os.path.exists(self.dwg_path):
            os.makedirs(self.dwg_path)
        
        # Get selected export setup name
        setup_name = self.cmbDWGSetup.SelectedItem.ToString() if self.cmbDWGSetup.SelectedItem else "(Default)"
        
        # Check if using a custom export setup
        use_custom_setup = setup_name != "(Default)" and setup_name in self.dwg_setups
        
        # DWG export options
        dwg_options = DWGExportOptions()
        
        # If using custom setup, load it
        if use_custom_setup:
            try:
                # Load the predefined setup
                dwg_options = DWGExportOptions.GetPredefinedOptions(self.doc, setup_name)
                output.print_md("Using DWG export setup: {}".format(setup_name))
            except Exception as e:
                output.print_md("Warning: Could not load setup '{}', using defaults: {}".format(setup_name, str(e)))
                dwg_options = DWGExportOptions()
        
        # Try to set AutoCAD version
        try:
            version_map = {
                "AutoCAD 2018": ACADVersion.R2018,
                "AutoCAD 2013": ACADVersion.R2013,
                "AutoCAD 2010": ACADVersion.R2010,
                "AutoCAD 2007": ACADVersion.R2007
            }
            selected_version = self.cmbDWGVersion.SelectedItem.ToString()
            if selected_version in version_map:
                dwg_options.FileVersion = version_map[selected_version]
                output.print_md("AutoCAD version: {}".format(selected_version))
        except Exception as e:
            output.print_md("Note: Could not set AutoCAD version, using default: {}".format(str(e)))
        
        # Export each sheet individually
        for item in selected_sheets:
            try:
                sheet_ids = List[ElementId]()
                sheet_ids.Add(item.Sheet.Id)
                
                filename = self.GetFileName(item.Sheet, dwg_naming)
                
                self.doc.Export(self.dwg_path, filename, sheet_ids, dwg_options)
                
                output.print_md("✓ DWG: {} - {}.dwg".format(item.SheetNumber, filename))
                success += 1
            except Exception as e:
                output.print_md("✗ DWG: {} - {}".format(item.SheetNumber, str(e)))
                errors += 1
        
        return success, errors
    
    def ShowDialog(self):
        """Show dialog"""
        return self._window.ShowDialog()


# ==============================================================================
# MAIN
# ==============================================================================

# Get XAML file path
script_dir = os.path.dirname(__file__)
xaml_path = os.path.join(script_dir, "ExportPDFDWG.xaml")

# Check if XAML file exists
if not os.path.exists(xaml_path):
    forms.alert(
        "XAML file not found!\n\nExpected location:\n{}".format(xaml_path),
        title="File Not Found",
        exitscript=True
    )

output.print_md("**Found {} sheets**".format(len(all_sheets)))
output.print_md("**View Sets**: {}".format(len(view_sets)))
output.print_md("**DWG Export Setups**: {}".format(len(dwg_setups)))

# Show window
window = ExportPDFDWGWindow(xaml_path, all_sheets, view_sets, dwg_setups, doc)
window.ShowDialog()

if window.result:
    output.print_md("### ✅ Export Completed Successfully")
else:
    output.print_md("### Export Cancelled or Closed")