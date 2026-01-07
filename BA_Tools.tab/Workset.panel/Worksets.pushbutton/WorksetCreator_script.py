# -*- coding: utf-8 -*-
"""
Workset Creator - WPF Modern UI
Create multiple worksets from predefined list
"""
__title__ = 'Create\nWorksets'
__doc__ = 'Create multiple worksets from predefined list'

import clr
import os
clr.AddReference("RevitAPI")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from Autodesk.Revit.DB import Workset, FilteredWorksetCollector, WorksetKind
from pyrevit import revit, script, forms
import System
from System.Windows.Markup import XamlReader
from System.Windows import Window
from System.IO import StreamReader
from System.Collections.ObjectModel import ObservableCollection

doc = revit.doc
output = script.get_output()


# ==============================================================================
# WORKSET MASTER LIST
# ==============================================================================

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


# ==============================================================================
# DATA CLASSES
# ==============================================================================

class WorksetItem(System.ComponentModel.INotifyPropertyChanged):
    """Observable workset item for checkbox binding"""
    
    def __init__(self, workset_name):
        self._workset_name = workset_name
        self._display_text = workset_name
        self._is_selected = False
        self._property_changed = None
    
    @property
    def WorksetName(self):
        return self._workset_name
    
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
# CHECK WORKSHARING
# ==============================================================================

if not doc.IsWorkshared:
    forms.alert('Project is not workshared.\n\nThis tool only works with workshared projects.', 
                title='Not Workshared', exitscript=True)

# Get existing worksets
existing_worksets = [ws.Name for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)]
available_worksets = [w for w in WORKSETS if w not in existing_worksets]

output.print_md("**Existing worksets**: {}".format(len(existing_worksets)))
output.print_md("**Available to create**: {}".format(len(available_worksets)))


# ==============================================================================
# WPF WINDOW CLASS
# ==============================================================================

class WorksetCreatorWindow(Window):
    def __init__(self, xaml_path, available_worksets, existing_count, document):
        # Load XAML
        stream = StreamReader(xaml_path)
        self._window = XamlReader.Load(stream.BaseStream)
        stream.Close()
        
        # Data
        self.available_worksets = available_worksets
        self.existing_count = existing_count
        self.workset_items = ObservableCollection[WorksetItem]()
        self.doc = document
        self.result = False
        self.selected_worksets = []
        
        # Get controls
        self.btnClose = self._window.FindName("btnClose")
        self.txtExistingCount = self._window.FindName("txtExistingCount")
        self.txtAvailableCount = self._window.FindName("txtAvailableCount")
        self.btnSelectAll = self._window.FindName("btnSelectAll")
        self.btnClearAll = self._window.FindName("btnClearAll")
        self.txtSelectedCount = self._window.FindName("txtSelectedCount")
        self.lstWorksets = self._window.FindName("lstWorksets")
        self.borderNoWorksets = self._window.FindName("borderNoWorksets")
        self.txtStatus = self._window.FindName("txtStatus")
        self.btnCancel = self._window.FindName("btnCancel")
        self.btnCreate = self._window.FindName("btnCreate")
        
        # Initialize
        self.InitializeData()
        self.SetupEventHandlers()
    
    def InitializeData(self):
        """Initialize workset items"""
        # Update counts
        self.txtExistingCount.Text = "Existing: {} worksets".format(self.existing_count)
        self.txtAvailableCount.Text = "Available to create: {} worksets".format(len(self.available_worksets))
        
        # Check if no worksets available
        if not self.available_worksets:
            self.lstWorksets.Visibility = System.Windows.Visibility.Collapsed
            self.borderNoWorksets.Visibility = System.Windows.Visibility.Visible
            self.btnCreate.IsEnabled = False
            self.btnSelectAll.IsEnabled = False
            self.btnClearAll.IsEnabled = False
            self.txtStatus.Text = "All worksets already exist"
            return
        
        # Populate workset items
        for workset_name in sorted(self.available_worksets):
            workset_item = WorksetItem(workset_name)
            workset_item.PropertyChanged += self.OnWorksetSelectionChanged
            self.workset_items.Add(workset_item)
        
        self.lstWorksets.ItemsSource = self.workset_items
        self.UpdateSelectedCount()
    
    def SetupEventHandlers(self):
        """Setup event handlers"""
        self.btnClose.Click += self.OnClose
        self.btnSelectAll.Click += self.OnSelectAll
        self.btnClearAll.Click += self.OnClearAll
        self.btnCancel.Click += self.OnCancel
        self.btnCreate.Click += self.OnCreate
    
    def OnClose(self, sender, args):
        """Close window"""
        self._window.DialogResult = False
        self._window.Close()
    
    def OnCancel(self, sender, args):
        """Cancel and close"""
        self._window.DialogResult = False
        self._window.Close()
    
    def OnSelectAll(self, sender, args):
        """Select all worksets"""
        for item in self.workset_items:
            item.IsSelected = True
        self.UpdateSelectedCount()
    
    def OnClearAll(self, sender, args):
        """Clear all selections"""
        for item in self.workset_items:
            item.IsSelected = False
        self.UpdateSelectedCount()
    
    def OnWorksetSelectionChanged(self, sender, args):
        """When workset selection changes"""
        self.UpdateSelectedCount()
    
    def UpdateSelectedCount(self):
        """Update selected count label"""
        selected_count = sum(1 for item in self.workset_items if item.IsSelected)
        total_count = len(self.workset_items)
        self.txtSelectedCount.Text = "{} / {} selected".format(selected_count, total_count)
        
        # Update status
        if selected_count == 0:
            self.txtStatus.Text = "Select worksets to create"
        else:
            self.txtStatus.Text = "Ready to create {} workset(s)".format(selected_count)
    
    def OnCreate(self, sender, args):
        """Create selected worksets"""
        # Get selected worksets
        selected_items = [item for item in self.workset_items if item.IsSelected]
        
        if not selected_items:
            forms.alert("Please select at least one workset", title="Validation Error")
            return
        
        self.selected_worksets = [item.WorksetName for item in selected_items]
        
        # Confirm
        confirm_msg = "Create {} workset(s)?\n\n".format(len(self.selected_worksets))
        confirm_msg += "Worksets:\n"
        for i, ws in enumerate(self.selected_worksets[:5]):
            confirm_msg += "• {}\n".format(ws)
        if len(self.selected_worksets) > 5:
            confirm_msg += "• ... and {} more\n".format(len(self.selected_worksets) - 5)
        
        if not forms.alert(confirm_msg, yes=True, no=True, title="Confirm Creation"):
            return
        
        # Update UI
        self.btnCreate.IsEnabled = False
        self.btnCreate.Content = "CREATING..."
        self.txtStatus.Text = "Creating worksets..."
        
        success_count = 0
        error_count = 0
        
        # Create worksets
        with revit.Transaction("Create Worksets"):
            for workset_name in self.selected_worksets:
                try:
                    Workset.Create(self.doc, workset_name)
                    success_count += 1
                    output.print_md("✓ Created: {}".format(workset_name))
                except Exception as e:
                    error_count += 1
                    output.print_md("✗ Failed: {} - {}".format(workset_name, str(e)))
        
        # Summary
        summary_msg = "Workset creation complete!\n\n"
        summary_msg += "Success: {}\n".format(success_count)
        summary_msg += "Errors: {}\n\n".format(error_count)
        summary_msg += "Check output window for details."
        
        forms.alert(summary_msg, title="Complete")
        
        output.print_md("### ✅ Workset Creation Complete")
        output.print_md("**Success**: {}".format(success_count))
        output.print_md("**Errors**: {}".format(error_count))
        
        self.result = True
        self._window.Close()
    
    def ShowDialog(self):
        """Show dialog"""
        return self._window.ShowDialog()


# ==============================================================================
# MAIN
# ==============================================================================

# Get XAML file path
script_dir = os.path.dirname(__file__)
xaml_path = os.path.join(script_dir, "WorksetCreator.xaml")

# Check if XAML file exists
if not os.path.exists(xaml_path):
    forms.alert(
        "XAML file not found!\n\nExpected location:\n{}".format(xaml_path),
        title="File Not Found",
        exitscript=True
    )

# Show window
window = WorksetCreatorWindow(xaml_path, available_worksets, len(existing_worksets), doc)
window.ShowDialog()

if window.result:
    output.print_md("### ✅ Worksets Created Successfully")
else:
    output.print_md("### Operation Cancelled or Closed")