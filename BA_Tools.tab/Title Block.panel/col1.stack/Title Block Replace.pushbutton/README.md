# Title Block Replacer - WPF Modern UI

A modern, dark-themed WPF interface for replacing title blocks on Revit sheets with an intuitive workflow.

## Features

- ✅ Modern dark-themed interface matching your existing tools
- ✅ Separate XAML and Python files for easy customization
- ✅ Real-time sheet filtering
- ✅ Current title block preview for selected sheets
- ✅ Live replacement log
- ✅ Bulk operations (Select All, Clear All, Select Filtered)
- ✅ Filter by "All" or "Unissued" (placeholder for future implementation)
- ✅ Maintains title block position on replacement

## File Structure

```
TitleBlockReplacer/
├── TitleBlockReplacer.py      # Main Python script
└── TitleBlockReplacer.xaml    # WPF UI definition
```

## Installation

### Method 1: pyRevit Extension (Recommended)

1. Create a new folder in your pyRevit extension directory:
   ```
   %APPDATA%\pyRevit\Extensions\[YourCompany].extension\[YourTab].tab\TitleBlocks.panel\
   ```

2. Create a folder called `TitleBlockReplacer.pushbutton` inside the panel folder

3. Copy both files into this folder:
   ```
   TitleBlockReplacer.pushbutton/
   ├── TitleBlockReplacer.py
   └── TitleBlockReplacer.xaml
   ```

4. Reload pyRevit (pyRevit → Reload)

5. The button will appear on your panel

### Method 2: Standalone Script

1. Create a folder on your computer (e.g., `C:\RevitScripts\TitleBlockReplacer\`)

2. Copy both files into this folder

3. In pyRevit, go to **pyRevit → Script Loader** and browse to the folder

4. Run the script

## Usage

### Step 1: Select Title Block
1. Choose a **Title Block Family** from the dropdown
2. Select the desired **Title Block Type**
3. The "Current Title Blocks" panel will show what's currently on selected sheets

### Step 2: Select Sheets
1. Use the **Filter** textbox to search for specific sheets
2. Click **Select All** to select all sheets
3. Click **Clear All** to deselect all sheets
4. Click **Select Filtered** to select only visible (filtered) sheets
5. Or manually check individual sheets

### Step 3: Replace
1. Click **REPLACE TITLE BLOCKS** button
2. Confirm the operation
3. Monitor progress in the "REPLACEMENT LOG" panel
4. Review the completion summary

## UI Components

### Header
- **TB Icon**: Title Block Replacer branding
- **Close Button**: Exit the tool

### Left Panel
- **Filter Radio Buttons**: All / Unissued (for future filtering)
- **Family Dropdown**: Select title block family
- **Type Dropdown**: Select title block type
- **Current Title Blocks**: Shows existing title blocks on selected sheets

### Right Panel
- **Sheet Filter**: Search/filter sheets by number or name
- **Action Buttons**: Select All, Clear All, Select Filtered
- **Sheet Count**: Shows selected/total sheets
- **Sheet List**: Checkboxes for each sheet
- **Replacement Log**: Real-time operation log

### Footer
- **Status Text**: Current operation status
- **View Report**: (Placeholder for future report feature)
- **REPLACE TITLE BLOCKS**: Main action button

## Customization

### Modifying Colors (XAML)

Edit the XAML file to change colors:

```xml
<!-- Primary Blue -->
<Setter Property="Background" Value="#1976D2"/>

<!-- Background Dark -->
Background="#14141C"

<!-- Panel Background -->
Background="#1E1E28"

<!-- Input Background -->
Background="#0A0A0F"
```

### Adding Features (Python)

The Python script is well-organized with clear sections:

- **DATA CLASSES**: Observable objects for data binding
- **COLLECT DATA**: Revit data collection functions
- **WPF WINDOW CLASS**: Main window logic and event handlers
- **MAIN**: Entry point and initialization

## Requirements

- Autodesk Revit 2019 or later
- pyRevit 4.8 or later
- .NET Framework 4.7 or later

## Troubleshooting

### "XAML file not found" Error
- Ensure both `.py` and `.xaml` files are in the same directory
- Check file names match exactly (case-sensitive)

### Title Blocks Not Replacing
- Ensure the selected title block family is loaded in the project
- Check that sheets are not on worksharing worksets you don't own
- Verify title blocks aren't pinned

### UI Not Displaying Correctly
- Check if .NET Framework 4.7+ is installed
- Try reloading pyRevit
- Check Windows scaling settings (100% recommended)

## Future Enhancements

- [ ] Implement "Unissued" filter functionality
- [ ] Add revision management integration
- [ ] Generate detailed replacement report
- [ ] Add undo functionality
- [ ] Export operation log to file
- [ ] Add title block comparison preview
- [ ] Batch operations history

## Version History

### v1.0.0 (2024)
- Initial release with WPF UI
- Separate XAML and Python files
- Modern dark theme
- Real-time filtering and logging
- Bulk selection operations

## Support

For issues or feature requests, contact your BIM team or modify the script to fit your specific workflow needs.

## License

This script is provided as-is for internal use. Modify as needed for your organization's requirements.

---

**Created for**: BIM Modelers working with Revit, Revit API, Dynamo, and Python
**Interface Style**: Modern Dark Theme (matching Revision Manager)
**Technology**: WPF + Python + pyRevit
