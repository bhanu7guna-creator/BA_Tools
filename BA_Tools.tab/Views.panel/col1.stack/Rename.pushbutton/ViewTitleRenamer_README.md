# View Title Renamer - WPF Modern UI

A modern, dark-themed WPF interface for renaming view titles on Revit sheets with powerful find/replace, prefix, and suffix options.

## Features

- ✅ Modern dark-themed interface matching your workflow tools
- ✅ Separate XAML and Python files for easy customization
- ✅ Real-time view filtering by sheet number, view name, or title
- ✅ Live preview of rename operations
- ✅ Find and replace with case-sensitive option
- ✅ Add custom prefix and suffix
- ✅ Option to use View Name instead of current title
- ✅ Add sheet number as automatic prefix
- ✅ Bulk operations (Select All, Clear All, Select Filtered)
- ✅ Real-time preview showing before/after for first 3 selections

## File Structure

```
ViewTitleRenamer/
├── ViewTitleRenamer.py      # Main Python script
└── ViewTitleRenamer.xaml    # WPF UI definition
```

## Installation

### Method 1: pyRevit Extension (Recommended)

1. Create a new folder in your pyRevit extension directory:
   ```
   %APPDATA%\pyRevit\Extensions\[YourCompany].extension\[YourTab].tab\Sheets.panel\
   ```

2. Create a folder called `ViewTitleRenamer.pushbutton` inside the panel folder

3. Copy both files into this folder:
   ```
   ViewTitleRenamer.pushbutton/
   ├── ViewTitleRenamer.py
   └── ViewTitleRenamer.xaml
   ```

4. Reload pyRevit (pyRevit → Reload)

5. The button will appear on your panel

### Method 2: Standalone Script

1. Create a folder on your computer (e.g., `C:\RevitScripts\ViewTitleRenamer\`)

2. Copy both files into this folder

3. In pyRevit, go to **pyRevit → Script Loader** and browse to the folder

4. Run the script

## Usage Guide

### Step 1: Select Views on Sheets

1. Use the **Filter** textbox to search for specific views
   - Search by sheet number, view name, or current title
   - Example: "FLOOR" finds all views containing "FLOOR"

2. Use selection buttons:
   - **Select All**: Select all views in the project
   - **Clear All**: Deselect all views
   - **Select Filtered**: Select only currently visible (filtered) views

3. Or manually check individual view checkboxes

### Step 2: Configure Rename Options

#### Find and Replace
- **Find Text**: Text to search for in current titles
- **Replace With**: Text to replace found text with
- **Case Sensitive**: Enable for exact case matching

#### Prefixes and Suffixes
- **Use View Name**: Use the view's name instead of current title as base
- **Add Sheet Number**: Automatically add sheet number as prefix (e.g., "A-101 - ")
- **Custom Prefix**: Add your own prefix text
- **Custom Suffix**: Add your own suffix text

#### Processing Order
1. Start with current title or view name (based on checkbox)
2. Apply find/replace
3. Add sheet number prefix (if enabled)
4. Add custom prefix
5. Add custom suffix

### Step 3: Preview and Apply

1. The **PREVIEW** panel shows:
   - Number of views to be renamed
   - Before → After for first 3 selections
   - Indication if more views are affected

2. Click **APPLY RENAME** to execute

3. Confirm the operation

4. Review results in pyRevit output window

## Common Use Cases

### Use Case 1: Add Prefix to All View Titles
```
Options:
- Custom Prefix: "TYPICAL - "
- Select: All views on floor plans

Result:
"LEVEL 1" → "TYPICAL - LEVEL 1"
```

### Use Case 2: Replace Department Codes
```
Options:
- Find Text: "ARCH"
- Replace With: "ARCHITECTURAL"
- Select: All views

Result:
"ARCH PLAN" → "ARCHITECTURAL PLAN"
```

### Use Case 3: Add Sheet Numbers to Titles
```
Options:
- Add Sheet Number as Prefix: ✓
- Select: All views on sheets A-101 through A-105

Result:
"FLOOR PLAN" → "A-101 - FLOOR PLAN"
```

### Use Case 4: Standardize View Titles from View Names
```
Options:
- Use View Name instead of Current Title: ✓
- Custom Suffix: " - ISSUED"
- Select: All views

Result:
Current Title: "Old Title"
View Name: "Level 1 Floor Plan"
New Title: "Level 1 Floor Plan - ISSUED"
```

### Use Case 5: Batch Rename with Complex Pattern
```
Options:
- Find Text: "REV"
- Replace With: "REVISION"
- Custom Prefix: "SHEET - "
- Custom Suffix: " - PRELIMINARY"
- Select: Filtered views

Result:
"REV A PLAN" → "SHEET - REVISION A PLAN - PRELIMINARY"
```

## UI Components

### Header
- **VT Icon**: View Title Renamer branding
- **Title & Description**: Tool name and purpose
- **Close Button**: Exit the tool

### Left Panel (View Selection)
- **Filter Search**: Real-time search/filter
- **Action Buttons**: Select All, Clear All, Select Filtered
- **View Count**: Shows selected vs total visible views
- **View List**: Checkboxes showing "Sheet | View Name | Current Title"

### Right Panel (Rename Options)
- **Find/Replace**: Text find and replace functionality
- **Checkboxes**: Use View Name, Add Sheet Number, Case Sensitive
- **Prefix/Suffix**: Custom text to add before/after
- **Preview Panel**: Live preview of rename operation

### Footer
- **Status Text**: Current operation status
- **APPLY RENAME**: Main action button

## Tips and Best Practices

### Filtering Tips
- Filter by sheet range: "A-1" shows all A-100 series sheets
- Filter by view type: "PLAN" or "SECTION"
- Filter by existing title: "TYPICAL"

### Naming Conventions
- Use consistent prefixes for easy filtering later
- Include sheet numbers for cross-referencing
- Keep titles concise but descriptive

### Safety Tips
- Always preview before applying
- Test on a few views first
- Save your project before bulk operations
- Use Revit's Undo if needed (Ctrl+Z)

## Customization

### Modifying Colors (XAML)

Edit the XAML file to change the color scheme:

```xml
<!-- Primary Blue -->
<Setter Property="Background" Value="#1976D2"/>

<!-- Preview Text Color (Green) -->
Foreground="#64C864"

<!-- Background Dark -->
Background="#14141C"
```

### Adding Custom Features (Python)

The Python script is organized into clear sections:

- **DATA CLASSES**: Observable objects for WPF binding
- **COLLECT DATA**: Revit viewport data collection
- **WPF WINDOW CLASS**: UI logic and event handlers
- **MAIN**: Entry point and initialization

## Requirements

- Autodesk Revit 2019 or later
- pyRevit 4.8 or later
- .NET Framework 4.7 or later
- Views must be placed on sheets

## Troubleshooting

### "XAML file not found" Error
- Ensure both `.py` and `.xaml` files are in the same directory
- Check file names match exactly (case-sensitive)

### No Views Shown
- Ensure views are actually placed on sheets
- Check if views are on worksharing worksets you can access

### Titles Not Changing
- Verify VIEW_DESCRIPTION parameter is not read-only
- Check if views are in design options you don't have access to
- Ensure you have edit rights to the sheets

### Preview Not Updating
- Try changing an option to trigger refresh
- Deselect and reselect a view

### Regular Expression Errors
- Escape special characters in find text
- Use simpler patterns if complex regex fails

## Keyboard Shortcuts

While in the filter textbox:
- **Enter**: Updates filter (automatic)
- **Escape**: Clears filter

## Advanced Features

### Case-Insensitive Replace
With "Case Sensitive" unchecked:
- Find: "level" matches "Level", "LEVEL", "level"
- Useful for standardizing capitalization

### Combining Options
All options work together:
1. Find/Replace happens first
2. Then sheet number prefix (if enabled)
3. Then custom prefix
4. Finally custom suffix

This allows complex transformations like:
```
Original: "arch plan rev a"
Find: "arch" → Replace: "ARCHITECTURAL"
Add Sheet Number: ✓ (Sheet A-101)
Prefix: "PROJECT - "
Suffix: " - ISSUED"

Result: "PROJECT - A-101 - ARCHITECTURAL plan rev a - ISSUED"
```

## Version History

### v1.0.0 (2024)
- Initial WPF release
- Separate XAML and Python files
- Modern dark theme
- Real-time filtering and preview
- Find/replace functionality
- Prefix/suffix support
- Case-sensitive option
- Use view name option
- Sheet number prefix option

## Support

For issues or feature requests, contact your BIM team or modify the script to fit your specific workflow needs.

## License

This script is provided as-is for internal use. Modify as needed for your organization's requirements.

---

**Created for**: BIM Modelers working with Revit, Revit API, Dynamo, and Python
**Interface Style**: Modern Dark Theme (matching Revision Manager)
**Technology**: WPF + Python + pyRevit
