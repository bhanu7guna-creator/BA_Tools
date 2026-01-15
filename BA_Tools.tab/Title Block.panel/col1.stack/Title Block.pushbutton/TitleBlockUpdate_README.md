# Title Block Update - WPF Modern UI

A modern, dark-themed WPF interface for batch updating title block parameters across multiple sheets in Revit.

## Features

- ✅ Modern dark-themed interface matching your workflow tools
- ✅ Separate XAML and Python files for easy customization
- ✅ Batch update multiple sheets simultaneously
- ✅ Real-time sheet filtering
- ✅ Multiple parameter support
- ✅ Date picker for issue dates
- ✅ Clear field functionality
- ✅ Detailed output logging

## File Structure

```
TitleBlockUpdate/
├── TitleBlockUpdate.py      # Main Python script
└── TitleBlockUpdate.xaml    # WPF UI definition
```

## Installation

### Method 1: pyRevit Extension (Recommended)

1. Create a new folder in your pyRevit extension directory:
   ```
   %APPDATA%\pyRevit\Extensions\[YourCompany].extension\[YourTab].tab\Sheets.panel\
   ```

2. Create a folder called `TitleBlockUpdate.pushbutton` inside the panel folder

3. Copy both files into this folder:
   ```
   TitleBlockUpdate.pushbutton/
   ├── TitleBlockUpdate.py
   └── TitleBlockUpdate.xaml
   ```

4. Reload pyRevit (pyRevit → Reload)

5. The button will appear on your panel

## Usage Guide

### Step 1: Select Sheets

1. Use the **Filter** textbox to search for specific sheets
   - Search by sheet number or sheet name
   - Example: "A-" finds all A-series sheets

2. Use selection buttons:
   - **Select All**: Select all visible (filtered) sheets
   - **Clear All**: Deselect all sheets

3. Or manually select sheets from the list (Ctrl+Click for multiple)

### Step 2: Fill in Parameters

Enter values for any parameters you want to update:

#### Personnel Parameters
- **Drawn By**: Name of person who drew the sheets
- **Checked By**: Name of person who checked the work
- **Designed By**: Name of designer
- **Approved By**: Name of approver

#### Project Information
- **Project Number**: Project/job number
- **Project Name**: Project name or description

#### Date
- **Issue Date**: Sheet issue date (use date picker)

**Note**: You only need to fill in the parameters you want to update. Leave others blank.

### Step 3: Update Sheets

1. Review your selections and parameter values
2. Click **UPDATE SHEETS**
3. Confirm the operation
4. Monitor progress in pyRevit output window
5. Review completion summary

## Supported Parameters

The script looks for these parameters in title blocks:

| Parameter Name | Type | Common Use |
|----------------|------|------------|
| Drawn By | Text | Draftsperson name |
| Checked By | Text | Checker name |
| Designed By | Text | Designer name |
| Approved By | Text | Approver name |
| Project Number | Text | Job/project number |
| Project Name | Text | Project description |
| Sheet Issue Date | Date | Date issued to client |

**Parameter Location**: The script checks both title block instance and sheet parameters.

## Common Use Cases

### Use Case 1: Update Drawn By for All Sheets
```
Steps:
1. Click "Select All" to select all sheets
2. Enter name in "Drawn By" field: "John Smith"
3. Click "UPDATE SHEETS"

Result: All sheets updated with draftsperson name
```

### Use Case 2: Update Issue Date for Specific Discipline
```
Steps:
1. Filter by "A-" for architectural sheets
2. Click "Select All" (selects filtered results)
3. Select issue date from date picker
4. Click "UPDATE SHEETS"

Result: All architectural sheets have same issue date
```

### Use Case 3: Update Multiple Parameters
```
Steps:
1. Select desired sheets
2. Fill in multiple fields:
   - Drawn By: "Jane Doe"
   - Checked By: "Mike Johnson"
   - Project Number: "2024-001"
   - Issue Date: 01/15/2024
3. Click "UPDATE SHEETS"

Result: All parameters updated on selected sheets
```

### Use Case 4: Update Specific Sheets Only
```
Steps:
1. Manually select specific sheets (Ctrl+Click)
2. Fill in parameters
3. Click "UPDATE SHEETS"

Result: Only selected sheets are updated
```

## UI Components

### Header
- **TB Icon**: Title Block Update branding
- **Title**: Tool name and description
- **Close Button**: Exit the tool

### Left Panel (Sheet Selection)
- **Filter Search**: Real-time sheet filtering
- **Select All/Clear All**: Bulk selection controls
- **Sheet Count**: Shows selected vs total sheets
- **Sheet List**: Multi-select list of sheets

### Right Panel (Parameters)
- **Parameter Fields**: Text inputs for various parameters
- **Date Picker**: Calendar widget for issue date
- **Clear All Fields**: Reset all parameter inputs

### Footer
- **Status Text**: Current operation status
- **Cancel**: Close without updating
- **UPDATE SHEETS**: Main action button

## Parameter Matching Logic

The script uses this logic to find and update parameters:

1. **Check Title Block First**: Looks for parameter on title block instance
2. **Check Sheet Second**: If not found on title block, checks sheet
3. **Skip Read-Only**: Ignores read-only parameters
4. **Update Available**: Sets value on first available, writable parameter

### Example Flow
```
Looking for "Drawn By":
1. Check title block → Found and writable? → Update ✓
2. If not found → Check sheet → Found? → Update ✓
3. If not found → Skip parameter → Report in log
```

## Output Window Details

The script provides detailed logging:

```
✓ A-101 - FLOOR PLAN LEVEL 1 - Updated: Drawn By, Checked By
✓ A-102 - FLOOR PLAN LEVEL 2 - Updated: Drawn By, Checked By
✗ A-103 - FLOOR PLAN LEVEL 3 - No parameters updated
⚠ A-104 - Drawn By: Parameter is read-only

### ✅ Update Complete
**Success**: 2
**Errors**: 2
```

## Date Format

The date picker uses standard US date format:
- **Format**: MM/dd/yyyy
- **Example**: 01/15/2024

If your title block uses a different format, you may need to modify the parameter or adjust the script.

## Troubleshooting

### Problem: Parameter not updating

**Possible causes:**
1. Parameter doesn't exist in title block or sheet
2. Parameter is read-only
3. Parameter name doesn't match exactly

**Solutions:**
- Check parameter name in title block family
- Ensure parameter is not read-only
- Check for extra spaces in parameter names

### Problem: Some sheets don't update

**Causes:**
- Title block doesn't have the parameter
- Parameter is locked/read-only
- Sheet is in a workset you don't own

**Solutions:**
- Check output window for specific errors
- Verify workset ownership
- Update title block family if needed

### Problem: Date not setting correctly

**Causes:**
- Parameter is not a text parameter
- Date format doesn't match expected format

**Solutions:**
- Ensure date parameter is text type
- Check if parameter expects different format

### Problem: Can't select multiple sheets

**Solution:**
- Use Ctrl+Click to select multiple individual sheets
- Use Shift+Click to select a range
- Or use "Select All" button

## Best Practices

### Before Updating
- [ ] Filter to relevant sheets first
- [ ] Verify parameter names match your title blocks
- [ ] Test on a few sheets before bulk update
- [ ] Save project before large updates

### During Update
- [ ] Double-check parameter values
- [ ] Review selected sheet count
- [ ] Use meaningful, consistent values
- [ ] Check output window for errors

### After Update
- [ ] Review updated sheets visually
- [ ] Check output window for any errors
- [ ] Verify date formats display correctly
- [ ] Save project

## Customization

### Adding New Parameters (Python)

To add support for new parameters, edit the `OnUpdate` method:

```python
# Add new parameter
if self.txtCustomParam.Text.strip():
    params['Custom Parameter Name'] = self.txtCustomParam.Text.strip()
```

### Adding New Parameter Fields (XAML)

Add new input fields in the XAML:

```xml
<Label Content="Custom Parameter:" Style="{StaticResource ValueLabel}"/>
<TextBox x:Name="txtCustomParam" Style="{StaticResource DarkTextBox}" Height="32"/>
```

### Changing Colors

Edit colors in XAML `<Window.Resources>`:

```xml
<Setter Property="Background" Value="#1976D2"/> <!-- Change this -->
```

## Requirements

- Autodesk Revit 2019 or later
- pyRevit 4.8 or later
- .NET Framework 4.7 or later
- Sheets with title blocks in the project
- Write permissions to modify sheets

## Advanced Features

### Multi-Sheet Updates
- Update hundreds of sheets in seconds
- Consistent parameter values across project
- Bulk operations save significant time

### Intelligent Parameter Lookup
- Checks title block first, then sheet
- Skips read-only parameters automatically
- Reports which parameters were updated

### Comprehensive Logging
- Success/failure for each sheet
- Specific error messages
- Summary statistics

## Keyboard Shortcuts

- **Ctrl+A** (in sheet list): Select all sheets
- **Ctrl+Click**: Add/remove individual sheets
- **Shift+Click**: Select range of sheets

## Version History

### v1.0.0 (2024)
- Initial WPF release
- Separate XAML and Python files
- Modern dark theme
- Real-time filtering
- Multiple parameter support
- Date picker integration
- Comprehensive logging

## Support

For issues or feature requests, contact your BIM team or modify the script to fit your specific workflow needs.

## License

This script is provided as-is for internal use. Modify as needed for your organization's requirements.

---

**Created for**: BIM Modelers working with Revit, Revit API, Dynamo, and Python
**Interface Style**: Modern Dark Theme (matching Revision Manager)
**Technology**: WPF + Python + pyRevit
