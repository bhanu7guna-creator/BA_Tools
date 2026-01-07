# NWC Exporter - WPF Modern UI

A modern, dark-themed WPF interface for exporting Navisworks Cache (.nwc) files from Revit with comprehensive export options and live logging.

## Features

- ✅ Modern dark-themed interface matching your workflow tools
- ✅ Separate XAML and Python files for easy customization
- ✅ Export from any 3D view in the project
- ✅ Comprehensive export options (rooms, links, properties, lights)
- ✅ Real-time system log with timestamps
- ✅ Live status indicator
- ✅ Automatic or custom filename
- ✅ File size reporting
- ✅ Path validation and directory creation

## File Structure

```
NWCExporter/
├── NWCExporter.py      # Main Python script
└── NWCExporter.xaml    # WPF UI definition
```

## Installation

### Method 1: pyRevit Extension (Recommended)

1. Create a new folder in your pyRevit extension directory:
   ```
   %APPDATA%\pyRevit\Extensions\[YourCompany].extension\[YourTab].tab\Export.panel\
   ```

2. Create a folder called `NWCExporter.pushbutton` inside the panel folder

3. Copy both files into this folder:
   ```
   NWCExporter.pushbutton/
   ├── NWCExporter.py
   └── NWCExporter.xaml
   ```

4. Reload pyRevit (pyRevit → Reload)

5. The button will appear on your panel

### Method 2: Standalone Script

1. Create a folder on your computer (e.g., `C:\RevitScripts\NWCExporter\`)

2. Copy both files into this folder

3. In pyRevit, go to **pyRevit → Script Loader** and browse to the folder

4. Run the script

## Usage Guide

### Step 1: Select Source 3D View

1. Open the dropdown menu
2. Select the 3D view you want to export
3. Any non-template 3D view in the project is available

**Tips:**
- Use a 3D view with appropriate visibility settings
- Ensure section boxes are set correctly if needed
- View filters will be applied in the export

### Step 2: Choose Output Directory

1. Click **Browse...** button
2. Navigate to your desired export location
3. Or use the default path: `C:\Projects\BIM\Exports\NWC`

**Default Directory:**
```
C:\Projects\BIM\Exports\NWC\
```

### Step 3: Set Filename

**Option A: Use Project Name (Recommended)**
- Check "Use Project Name"
- Filename automatically matches your Revit project name
- Filename field becomes read-only

**Option B: Custom Filename**
- Uncheck "Use Project Name"
- Enter your custom filename
- Extension (.nwc) is added automatically
- Invalid characters are automatically replaced with underscores

**Invalid Characters:** `< > : " / \ | ? *`

### Step 4: Configure Export Options

#### Room Settings
- **Export Room Geometry** ✓ (Recommended)
  - Includes 3D room volumes in export
  - Useful for space planning and clash detection

- **Export Room as Attribute** ✓ (Recommended)
  - Includes room data as element properties
  - Preserves room names, numbers, and parameters

#### Link Settings
- **Export Links** ☐ (Default: Off)
  - Includes linked Revit models
  - Increases file size and export time
  - Enable for federated coordination

#### Property Settings
- **Convert Element Properties** ✓ (Recommended)
  - Exports Revit parameters to Navisworks
  - Essential for clash detection rules
  - Enables property-based searches

#### Additional Options
- **Export URLs** ✓ (Recommended)
  - Maintains hyperlinks from Revit
  - Useful for documentation links

- **Convert Lights** ✓ (Recommended)
  - Exports lighting fixtures
  - Useful for rendering in Navisworks

### Step 5: Export

1. Review settings in System Log
2. Click **EXPORT NWC FILE**
3. Monitor progress in the log panel
4. Wait for completion message
5. File is automatically saved to chosen directory

## System Log

The right panel shows real-time export progress:

```
[14:32:15] Selected view: 3D View - Coordination
[14:32:18] Export path updated: C:\Projects\BIM\Exports\NWC
[14:32:20] Starting export process...
[14:32:20] View: 3D View - Coordination
[14:32:20] Output: C:\Projects\BIM\Exports\NWC\Project.nwc
[14:32:20] Configuring export settings...
[14:32:20]   Export scope: Selected View
[14:32:20]   Coordinates: Shared
[14:32:20]   Export links: False
[14:32:20]   Room geometry: True
[14:32:21] Processing geometry...
[14:32:45] SUCCESS: Export completed!
[14:32:45] File size: 42.35 MB
```

## Export Options Explained

### Coordinates: Shared
- Uses project shared coordinates
- Ensures proper alignment in Navisworks
- Essential for multi-model coordination

### Faceting Factor: 1.0
- Controls geometry tessellation quality
- 1.0 = Maximum quality (default)
- Lower values reduce file size but decrease quality

### Export Scope: View
- Exports only elements visible in selected 3D view
- Respects view filters and visibility settings
- Honors section boxes and crop regions

## Common Use Cases

### Use Case 1: Basic Coordination Model
```
Settings:
- View: 3D View - Coordination
- Export Links: No
- Room Geometry: Yes
- Room as Attribute: Yes
- Element Properties: Yes

Result: Clean coordination model with rooms
File Size: ~30-50 MB (typical)
```

### Use Case 2: Federated Model with Links
```
Settings:
- View: 3D View - Full Model
- Export Links: Yes ✓
- Room Geometry: Yes
- Room as Attribute: Yes
- Element Properties: Yes

Result: Complete federated model
File Size: ~100-300 MB (depending on links)
```

### Use Case 3: Lightweight Geometry Only
```
Settings:
- View: 3D View - Simple
- Export Links: No
- Room Geometry: No
- Room as Attribute: No
- Element Properties: No
- URLs: No
- Lights: No

Result: Minimal geometry export
File Size: ~15-25 MB
```

### Use Case 4: Design Review
```
Settings:
- View: 3D View - Presentation
- Export Links: No
- Room Geometry: Yes
- Room as Attribute: Yes
- Element Properties: Yes
- URLs: Yes
- Lights: Yes ✓

Result: Full presentation model
File Size: ~40-60 MB
```

## UI Components

### Header
- **NW Icon**: NWC Exporter branding
- **Title**: Tool name and description
- **Close Button**: Exit the tool

### Left Panel (Export Settings)
- **Source View**: 3D view selection dropdown
- **Output Directory**: Path display with browse button
- **Filename**: Custom or automatic naming
- **Export Options**: Checkboxes for various settings

### Right Panel (System Log)
- **Log Area**: Real-time export progress
- **Status Bar**: Current operation status (IDLE/BUSY)

### Footer
- **File Info**: Shows complete output path
- **EXPORT Button**: Main action button

## File Naming

### Automatic (Use Project Name)
```
Project: Healthcare_Building_2024.rvt
Output: Healthcare_Building_2024.nwc
```

### Custom Naming
```
Input: Site Model - Rev A
Output: Site_Model_-_Rev_A.nwc (invalid chars replaced)
```

### Best Practices
- Include revision in filename: `Project_RevA`
- Add date for versioning: `Project_2024-01-15`
- Use descriptive names: `Arch_Coordination_Model`
- Keep names under 50 characters

## Troubleshooting

### Problem: "No 3D views found"

**Cause**: No 3D views exist in the project

**Solution**:
- Create a 3D view in Revit
- Ensure the view is not a template
- Check if views are in a workset you don't own

### Problem: Export fails with no error

**Causes:**
- Directory doesn't exist
- No write permissions
- Disk space full

**Solutions**:
- Check folder permissions
- Verify sufficient disk space (2-3x model size)
- Try exporting to a different directory

### Problem: File size is very large

**Causes:**
- Links are included
- High geometry detail
- Many rooms included

**Solutions**:
- Disable "Export Links"
- Use a simplified 3D view
- Reduce visible elements in source view

### Problem: Missing elements in Navisworks

**Causes:**
- Elements hidden in source view
- Elements on hidden categories
- Section box excluding elements

**Solutions**:
- Check visibility/graphics in source view
- Verify category visibility
- Adjust or remove section box

### Problem: Can't find exported file

**Solution**:
- Check the path in the file info (footer)
- Look in Windows Explorer at the exact path
- Verify filename doesn't have invalid characters

## Performance Tips

### For Faster Exports
1. Use a view with fewer visible elements
2. Disable "Export Links"
3. Turn off "Export Room Geometry" if not needed
4. Use a local drive (not network)

### For Smaller Files
1. Simplify the source 3D view
2. Hide unnecessary categories
3. Use detail levels appropriately
4. Disable unused export options

### For Better Quality
1. Keep "Convert Element Properties" enabled
2. Use "Export Room as Attribute"
3. Enable all relevant options
4. Accept larger file size

## Integration with Navisworks

### Opening in Navisworks
1. Launch Navisworks Manage/Simulate
2. File → Open → Select .nwc file
3. File loads with Revit properties intact

### Clash Detection Setup
1. Ensure "Convert Element Properties" was enabled
2. Properties are available in Selection Tree
3. Use for clash rules and filters

### Timeliner Animation
1. Room attributes support phasing
2. Element properties include phase data
3. Create 4D simulations

## Best Practices

### Before Exporting
- [ ] Clean up the model (delete unused elements)
- [ ] Set up appropriate 3D view
- [ ] Check visibility/graphics settings
- [ ] Test section box if using one
- [ ] Verify workset visibility

### During Export
- [ ] Monitor system log for errors
- [ ] Don't close Revit during export
- [ ] Avoid working in the model
- [ ] Wait for completion message

### After Export
- [ ] Verify file was created
- [ ] Check file size is reasonable
- [ ] Test open in Navisworks
- [ ] Verify element properties
- [ ] Document export settings used

## Customization

### Changing Colors (XAML)

Edit colors in the `<Window.Resources>` section:

```xml
<!-- Primary Blue -->
<Setter Property="Background" Value="#1976D2"/>

<!-- Log Text Color -->
Foreground="#808090"

<!-- Background -->
Background="#14141C"
```

### Adding Export Presets (Python)

Add preset configurations:

```python
def ApplyPreset(self, preset_name):
    presets = {
        'coordination': {
            'links': False,
            'rooms': True,
            'properties': True
        },
        'lightweight': {
            'links': False,
            'rooms': False,
            'properties': False
        }
    }
    # Apply settings...
```

## Requirements

- Autodesk Revit 2019 or later
- pyRevit 4.8 or later
- .NET Framework 4.7 or later
- At least one 3D view in the project
- Write permissions to export directory

## Advanced Features

### Automatic Directory Creation
- If export directory doesn't exist, it's created automatically
- Parent directories are also created as needed
- Failed creation shows error message

### Invalid Character Replacement
- Automatically replaces invalid filename characters
- Preserves readable filename
- Example: `Site/Plan*2024` → `Site_Plan_2024`

### File Size Reporting
- Shows exact file size after export
- Helps track export efficiency
- Compare with different settings

## Keyboard Shortcuts

While window is active:
- **Escape**: Close window (same as Close button)
- **Enter** (in filename field): Move to export

## Version History

### v1.0.0 (2024)
- Initial WPF release
- Separate XAML and Python files
- Modern dark theme
- Real-time logging
- Comprehensive export options
- Status indicator
- File info display

## Support

For issues or feature requests, contact your BIM team or modify the script to fit your specific workflow needs.

## License

This script is provided as-is for internal use. Modify as needed for your organization's requirements.

---

**Created for**: BIM Modelers working with Revit, Navisworks, and Coordination workflows
**Interface Style**: Modern Dark Theme (matching Revision Manager)
**Technology**: WPF + Python + pyRevit
