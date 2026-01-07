# -*- coding: utf-8 -*-
"""
Title Block Batch Update - Compact XAML UI
Select multiple sheets and update title block parameters at once
"""
__title__ = 'Title Block\nUpdate'
__doc__ = 'Batch update title block parameters on multiple sheets'

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import *
from pyrevit import revit, script, forms
from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import MemoryStream
from System.Text import Encoding
from System import DateTime
import os

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# ==============================================================================
# EMBEDDED XAML (COMPACT VERSION)
# ==============================================================================

XAML_STRING = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Title Block Update" 
        Width="900" Height="600"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        WindowStyle="None"
        Background="#1A1A1A"
        AllowsTransparency="False">

    <Border BorderBrush="#0078D4" BorderThickness="2" Background="#1A1A1A">
        <Grid Margin="0">
            <Grid.RowDefinitions>
                <RowDefinition Height="60"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="60"/>
            </Grid.RowDefinitions>

            <!-- HEADER BAR -->
            <Border Grid.Row="0" Background="#000000" BorderBrush="#0078D4" BorderThickness="0,0,0,2">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- BA LOGO -->
                    <Border Grid.Column="0" 
                            Background="#0078D4" 
                            Width="50" 
                            Height="50" 
                            CornerRadius="5"
                            Margin="10,5,15,5"
                            VerticalAlignment="Center">
                        <TextBlock Text="BA" 
                                   FontFamily="Segoe UI" 
                                   FontSize="20" 
                                   FontWeight="Bold" 
                                   Foreground="White"
                                   HorizontalAlignment="Center"
                                   VerticalAlignment="Center"/>
                    </Border>

                    <!-- TITLE -->
                    <TextBlock Grid.Column="1" 
                               Text="TITLE BLOCK UPDATE" 
                               FontFamily="Segoe UI" 
                               FontSize="20" 
                               FontWeight="Bold" 
                               Foreground="White"
                               VerticalAlignment="Center"
                               Margin="20,0,0,0"/>

                    <!-- CLOSE BUTTON -->
                    <Button Grid.Column="2" 
                            x:Name="btn_close_window"
                            Content="Close" 
                            Width="80" 
                            Height="30" 
                            Background="Transparent" 
                            Foreground="White" 
                            FontFamily="Segoe UI"
                            FontSize="13"
                            FontWeight="Bold"
                            BorderBrush="#0078D4"
                            BorderThickness="1,0,0,0"
                            Cursor="Hand">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Background" Value="Transparent"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border Background="{TemplateBinding Background}" 
                                                    BorderBrush="{TemplateBinding BorderBrush}"
                                                    BorderThickness="{TemplateBinding BorderThickness}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#0078D4"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>
            </Border>

            <!-- MAIN CONTENT -->
            <Grid Grid.Row="1" Margin="10,10,10,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="320"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- LEFT - SHEETS -->
                <Border Grid.Column="0" Background="#2D2D30" BorderBrush="#0078D4" BorderThickness="1" CornerRadius="5" Margin="0,0,5,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="40"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="55"/>
                        </Grid.RowDefinitions>

                        <TextBox x:Name="txt_sheet_filter" 
                                 Grid.Row="0"
                                 Height="28" 
                                 Background="#1A1A1A" 
                                 Foreground="White" 
                                 BorderBrush="#3F3F46"
                                 FontFamily="Segoe UI" 
                                 FontSize="11" 
                                 Padding="8,5"
                                 Margin="10,6"/>

                        <Border Grid.Row="1" Background="#1A1A1A" BorderBrush="#3F3F46" BorderThickness="1" Margin="10,0" CornerRadius="3">
                            <ListBox x:Name="list_sheets" 
                                     Background="Transparent" 
                                     BorderThickness="0" 
                                     FontFamily="Consolas" 
                                     FontSize="10" 
                                     Foreground="White"
                                     SelectionMode="Multiple">
                                <ListBox.ItemContainerStyle>
                                    <Style TargetType="ListBoxItem">
                                        <Setter Property="Padding" Value="8,5"/>
                                        <Setter Property="Background" Value="Transparent"/>
                                        <Setter Property="Foreground" Value="White"/>
                                        <Setter Property="BorderThickness" Value="0"/>
                                        <Setter Property="Margin" Value="2"/>
                                        <Style.Triggers>
                                            <Trigger Property="IsSelected" Value="True">
                                                <Setter Property="Background" Value="#0078D4"/>
                                                <Setter Property="Foreground" Value="White"/>
                                            </Trigger>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter Property="Background" Value="#3F3F46"/>
                                            </Trigger>
                                        </Style.Triggers>
                                    </Style>
                                </ListBox.ItemContainerStyle>
                            </ListBox>
                        </Border>

                        <StackPanel Grid.Row="2" Orientation="Vertical" Margin="10,5">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,5">
                                <Button x:Name="btn_select_all" Content="All" Width="70" Height="25" Background="#0078D4" Foreground="White" FontSize="10" FontWeight="Bold" BorderThickness="0" Margin="0,0,5,0" Cursor="Hand"/>
                                <Button x:Name="btn_clear_all" Content="None" Width="70" Height="25" Background="#3F3F46" Foreground="White" FontSize="10" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                            </StackPanel>
                            <TextBlock x:Name="txt_count" Text="0 / 0 selected" FontSize="9" Foreground="#0078D4" HorizontalAlignment="Center" FontWeight="Bold"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- RIGHT - PARAMETERS -->
                <Border Grid.Column="1" Background="#2D2D30" BorderBrush="#0078D4" BorderThickness="1" CornerRadius="5" Margin="5,0,0,0">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="15,10">
                        <StackPanel>
                            <TextBlock Text="PARAMETERS" FontSize="12" FontWeight="Bold" Foreground="#0078D4" Margin="0,0,0,10"/>
                            
                            <TextBlock Text="Drawn By:" FontSize="9" Foreground="#AAA" Margin="0,5,0,3"/>
                            <TextBox x:Name="txt_drawn_by" Height="26" Background="#1A1A1A" Foreground="White" BorderBrush="#3F3F46" FontSize="11" Padding="5"/>
                            
                            <TextBlock Text="Checked By:" FontSize="9" Foreground="#AAA" Margin="0,8,0,3"/>
                            <TextBox x:Name="txt_checked_by" Height="26" Background="#1A1A1A" Foreground="White" BorderBrush="#3F3F46" FontSize="11" Padding="5"/>
                            
                            <TextBlock Text="Designed By:" FontSize="9" Foreground="#AAA" Margin="0,8,0,3"/>
                            <TextBox x:Name="txt_designed_by" Height="26" Background="#1A1A1A" Foreground="White" BorderBrush="#3F3F46" FontSize="11" Padding="5"/>
                            
                            <TextBlock Text="Approved By:" FontSize="9" Foreground="#AAA" Margin="0,8,0,3"/>
                            <TextBox x:Name="txt_approved_by" Height="26" Background="#1A1A1A" Foreground="White" BorderBrush="#3F3F46" FontSize="11" Padding="5"/>
                            
                            <TextBlock Text="Issue Date:" FontSize="9" Foreground="#AAA" Margin="0,8,0,3"/>
                            <DatePicker x:Name="date_picker" 
                                        Height="26" 
                                        FontSize="11"
                                        BorderBrush="#3F3F46"
                                        BorderThickness="1">
                                <DatePicker.Resources>
                                    <Style TargetType="DatePickerTextBox">
                                        <Setter Property="Background" Value="#1A1A1A"/>
                                        <Setter Property="Foreground" Value="White"/>
                                        <Setter Property="BorderThickness" Value="0"/>
                                        <Setter Property="Padding" Value="5,3"/>
                                    </Style>
                                </DatePicker.Resources>
                            </DatePicker>
                            
                            <TextBlock Text="Project Number:" FontSize="9" Foreground="#AAA" Margin="0,8,0,3"/>
                            <TextBox x:Name="txt_project_number" Height="26" Background="#1A1A1A" Foreground="White" BorderBrush="#3F3F46" FontSize="11" Padding="5"/>
                            
                            <TextBlock Text="Project Name:" FontSize="9" Foreground="#AAA" Margin="0,8,0,3"/>
                            <TextBox x:Name="txt_project_name" Height="26" Background="#1A1A1A" Foreground="White" BorderBrush="#3F3F46" FontSize="11" Padding="5"/>
                            
                            <Button x:Name="btn_clear_fields" Content="Clear All Fields" Height="28" Background="#505050" Foreground="White" FontSize="10" FontWeight="Bold" BorderThickness="0" Margin="0,15,0,10" Cursor="Hand"/>
                        </StackPanel>
                    </ScrollViewer>
                </Border>
            </Grid>

            <!-- BOTTOM BUTTONS -->
            <Border Grid.Row="2" Background="#000000" BorderBrush="#0078D4" BorderThickness="0,2,0,0">
                <Grid>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                        <Button x:Name="btn_cancel" Content="Cancel" Width="150" Height="35" Background="#505050" Foreground="White" FontSize="12" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,0" Cursor="Hand"/>
                        <Button x:Name="btn_update" Content="Update Sheets" Width="200" Height="35" Background="#0078D4" Foreground="White" FontSize="12" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>
    </Border>
</Window>
'''


# ==============================================================================
# DATA COLLECTION
# ==============================================================================

def get_all_sheets_with_titleblocks():
    """Get all sheets with their title blocks"""
    sheets = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Sheets)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    sheet_data = []
    for sheet in sheets:
        titleblocks = FilteredElementCollector(doc, sheet.Id)\
            .OfCategory(BuiltInCategory.OST_TitleBlocks)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        if titleblocks:
            sheet_data.append({
                'sheet': sheet,
                'titleblock': titleblocks[0],
                'sheet_number': sheet.SheetNumber,
                'sheet_name': sheet.Name
            })
    
    return sorted(sheet_data, key=lambda x: x['sheet_number'])


# ==============================================================================
# WINDOW CLASS
# ==============================================================================

class TitleBlockWindow:
    def __init__(self, sheet_data, document):
        self.sheet_data = sheet_data
        self.all_sheets = sheet_data
        self.doc = document
        self.result = None
        
        # Load XAML
        xaml_bytes = Encoding.UTF8.GetBytes(XAML_STRING)
        stream = MemoryStream(xaml_bytes)
        self.window = XamlReader.Load(stream)
        
        # Get controls - search in the loaded window
        self.list_sheets = self.find_control("list_sheets")
        self.txt_sheet_filter = self.find_control("txt_sheet_filter")
        self.txt_count = self.find_control("txt_count")
        self.btn_select_all = self.find_control("btn_select_all")
        self.btn_clear_all = self.find_control("btn_clear_all")
        
        self.txt_drawn_by = self.find_control("txt_drawn_by")
        self.txt_checked_by = self.find_control("txt_checked_by")
        self.txt_designed_by = self.find_control("txt_designed_by")
        self.txt_approved_by = self.find_control("txt_approved_by")
        self.txt_project_number = self.find_control("txt_project_number")
        self.txt_project_name = self.find_control("txt_project_name")
        self.date_picker = self.find_control("date_picker")
        
        self.btn_clear_fields = self.find_control("btn_clear_fields")
        self.btn_cancel = self.find_control("btn_cancel")
        self.btn_update = self.find_control("btn_update")
    
    def find_control(self, name):
        """Find control by name in the window"""
        return self.window.FindName(name)
    
    def show(self):
        """Show the window"""
        # Get close button
        self.btn_close_window = self.find_control("btn_close_window")
        
        # Events
        self.txt_sheet_filter.TextChanged += self.OnFilterChanged
        self.list_sheets.SelectionChanged += self.OnSelectionChanged
        self.btn_select_all.Click += self.SelectAll
        self.btn_clear_all.Click += self.ClearAll
        self.btn_clear_fields.Click += self.ClearFields
        self.btn_close_window.Click += self.OnCancel
        self.btn_cancel.Click += self.OnCancel
        self.btn_update.Click += self.Update
        
        self.populate_sheets()
        self.update_count()
        
        self.window.ShowDialog()
    
    def OnCancel(self, s, e):
        """Close window"""
        self.window.Close()
    
    def populate_sheets(self):
        self.list_sheets.Items.Clear()
        filter_text = (self.txt_sheet_filter.Text or "").lower()
        
        self.sheet_data = []
        for s in self.all_sheets:
            text = "{} - {}".format(s['sheet_number'], s['sheet_name'])
            if not filter_text or filter_text in text.lower():
                self.list_sheets.Items.Add(text)
                self.sheet_data.append(s)
        
        self.update_count()
    
    def OnFilterChanged(self, s, e):
        self.populate_sheets()
    
    def OnSelectionChanged(self, s, e):
        self.update_count()
    
    def update_count(self):
        sel = self.list_sheets.SelectedItems.Count
        total = self.list_sheets.Items.Count
        self.txt_count.Text = "{} / {} selected".format(sel, total)
    
    def SelectAll(self, s, e):
        for i in range(self.list_sheets.Items.Count):
            self.list_sheets.SelectedItems.Add(self.list_sheets.Items[i])
    
    def ClearAll(self, s, e):
        self.list_sheets.SelectedItems.Clear()
    
    def ClearFields(self, s, e):
        self.txt_drawn_by.Text = ""
        self.txt_checked_by.Text = ""
        self.txt_designed_by.Text = ""
        self.txt_approved_by.Text = ""
        self.txt_project_number.Text = ""
        self.txt_project_name.Text = ""
        self.date_picker.SelectedDate = None
    
    def Update(self, s, e):
        indices = [self.list_sheets.Items.IndexOf(i) for i in self.list_sheets.SelectedItems]
        
        if not indices:
            forms.alert("Please select at least one sheet", title="Error")
            return
        
        selected = [self.sheet_data[i] for i in indices]
        
        params = {}
        if self.txt_drawn_by.Text.strip():
            params['Drawn By'] = self.txt_drawn_by.Text.strip()
        if self.txt_checked_by.Text.strip():
            params['Checked By'] = self.txt_checked_by.Text.strip()
        if self.txt_designed_by.Text.strip():
            params['Designed By'] = self.txt_designed_by.Text.strip()
        if self.txt_approved_by.Text.strip():
            params['Approved By'] = self.txt_approved_by.Text.strip()
        if self.txt_project_number.Text.strip():
            params['Project Number'] = self.txt_project_number.Text.strip()
        if self.txt_project_name.Text.strip():
            params['Project Name'] = self.txt_project_name.Text.strip()
        if self.date_picker.SelectedDate:
            params['Sheet Issue Date'] = self.date_picker.SelectedDate.ToString("MM/dd/yyyy")
        
        if not params:
            forms.alert("Please fill in at least one parameter", title="Error")
            return
        
        msg = "Update {} sheets with {} parameters?".format(len(selected), len(params))
        if not forms.alert(msg, yes=True, no=True):
            return
        
        self.btn_update.IsEnabled = False
        self.btn_update.Content = "UPDATING..."
        
        success = 0
        errors = 0
        
        t = Transaction(doc, "Update Title Blocks")
        t.Start()
        
        try:
            for sd in selected:
                try:
                    updated = []
                    for pname, pvalue in params.items():
                        param = sd['titleblock'].LookupParameter(pname)
                        if not param:
                            param = sd['sheet'].LookupParameter(pname)
                        
                        if param and not param.IsReadOnly:
                            try:
                                param.Set(pvalue)
                                updated.append(pname)
                            except:
                                pass
                    
                    if updated:
                        success += 1
                        output.print_md("✓ {} - {}".format(sd['sheet_number'], ", ".join(updated)))
                    else:
                        errors += 1
                except:
                    errors += 1
            
            t.Commit()
        except:
            t.RollBack()
        
        forms.alert("Complete!\n\nSuccess: {}\nErrors: {}".format(success, errors), title="Done")
        
        self.result = True
        self.btn_update.IsEnabled = True
        self.btn_update.Content = "Update Sheets"


# ==============================================================================
# MAIN
# ==============================================================================

try:
    all_sheets = get_all_sheets_with_titleblocks()
    
    if not all_sheets:
        forms.alert('No sheets with title blocks found', exitscript=True)
    
    output.print_md("**Found {} sheets**".format(len(all_sheets)))
    
    window = TitleBlockWindow(all_sheets, doc)
    window.show()
    
    if window.result:
        output.print_md("### ✅ Complete")
    else:
        output.print_md("### Cancelled")
        
except Exception as e:
    output.print_md("**ERROR: {}**".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))