# Workset Creator - WPF Modern UI

A modern, dark-themed WPF interface for creating multiple worksets from a predefined master list in Revit workshared projects.

## Features

- ✅ Modern dark-themed interface matching your workflow tools
- ✅ Separate XAML and Python files for easy customization
- ✅ Predefined master list of common worksets
- ✅ Batch create multiple worksets simultaneously
- ✅ Visual checkboxes (18x18px with blue borders)
- ✅ Shows existing vs available worksets
- ✅ Select All / Clear All functionality
- ✅ Detailed output logging

## File Structure

```
WorksetCreator/
├── WorksetCreator.py      # Main Python script
└── WorksetCreator.xaml    # WPF UI definition
```

## Installation

### Method 1: pyRevit Extension (Recommended)

1. Create a new folder in your pyRevit extension directory:
   ```
   %APPDATA%\pyRevit\Extensions\[YourCompany].extension\[YourTab].tab\Manage.panel\
   ```

2. Create a folder called `WorksetCreator.pushbutton` inside the panel folder

3. Copy both files into this folder:
   ```
   WorksetCreator.pushbutton/
   ├── WorksetCreator.py
   └── WorksetCreator.xaml
   ```

4. Reload pyRevit (pyRevit → Reload)

5. The button will appear on your panel

## Requirements

- Project must be **workshared** (File → Worksharing → Enable Worksharing)
- User must have permission to create worksets
- Revit 2019 or later
- pyRevit 4.8 or later

## Predefined Workset List

The script includes a comprehensive list of common structural and link worksets:

### Architecture & Civil
- AR-Civil Works-GF
- AR-Civil Works-FF

### Links
- LINK-CAD-ST
- LINK-RVT-AR (Architecture)
- LINK-RVT-EL (Electrical)
- LINK-RVT-GN (General)
- LINK-RVT-ID (Interior Design)
- LINK-RVT-LA (Landscape)
- LINK-RVT-ME (Mechanical)
- LINK-RVT-PL (Plumbing)
- LINK-RVT-ST (Structural)

### Structural Elements
- **Beams**: ST-Beam-FF, ST-Beam-RF
- **Block Walls**: ST-BLOCK WALL-FF, ST-BLOCK WALL-GF, ST-BLOCK WALL-PA
- **Cast-in-Place**: ST-CIP-FF, ST-CIP-RF
- **Columns**: ST-Column-FF, ST-Column-GF
- **Erection Marks**: ST-Erection Mark-FF, ST-Erection Mark-GF, ST-Erection Mark-RF
- **Foundation**: ST-Foundation
- **Hollowcore**: ST-Hollowcore-FF, ST-Hollowcore-RF
- **Reveals**: ST-Reveals-FF, ST-Reveals-GF, ST-Reveals-RF
- **Slabs**: ST-SLAB-FF, ST-SLAB-GF, ST-SLAB-RF
- **Stairs**: ST-STAIR-FF, ST-STAIR-GF
- **Openings**: ST-Stru Opngs-FF, ST-Stru Opngs-RF
- **Walls**: ST-WALL-FF, ST-WALL-GF, ST-WALL-RF

**Total**: 39 worksets in master list

## Usage Guide

### Step 1: Open Workshared Project

The script only works with workshared projects. If your project isn't workshared:
1. File → Worksharing → Enable Worksharing
2. Save and reopen the project
3. Run the script

### Step 2: Review Available Worksets

The tool automatically:
- Scans existing worksets in your project
- Shows count of existing worksets
- Displays only worksets that don't exist yet
- Sorts available worksets alphabetically

### Step 3: Select Worksets

**Option A: Select All**
- Click "Select All" to create all available worksets

**Option B: Select Specific**
- Check individual worksets you need
- Use Ctrl+Click for multiple selections

**Option C: Clear and Reselect**
- Click "Clear All" to deselect everything
- Manually select only the ones you want

### Step 4: Create Worksets

1. Click **CREATE WORKSETS**
2. Review confirmation dialog
3. Confirm creation
4. Monitor progress in output window

## UI Components

### Header
- **WS Icon**: Workset Creator branding
- **Title**: Tool name and description
- **Close Button**: Exit without creating

### Info Panel
- **Existing Count**: Number of worksets already in project
- **Available Count**: Number of worksets that can be created

### Selection Area
- **Select All Button**: Check all available worksets
- **Clear All Button**: Uncheck all worksets
- **Selected Count**: Shows X / Y selected
- **Workset List**: Checkboxes for each available workset

### Footer
- **Status**: Current operation status
- **Cancel**: Close without creating
- **CREATE WORKSETS**: Main action button

## Naming Convention

### Prefix System
- **AR-**: Architecture
- **LINK-**: Linked files (CAD or RVT)
- **ST-**: Structural

### Level Suffixes
- **-GF**: Ground Floor
- **-FF**: First Floor (or typical floors)
- **-RF**: Roof
- **-PA**: Parking/Podium Area

### Examples
```
ST-Column-GF     = Structural Columns on Ground Floor
ST-SLAB-FF       = Structural Slabs on First Floor
LINK-RVT-AR      = Linked Revit Architecture Model
AR-Civil Works-GF = Architecture Civil Works Ground Floor
```

## Customizing the Workset List

To add or modify worksets, edit the `WORKSETS` list in the Python file:

```python
WORKSETS = [
    # Add your custom worksets here
    "Custom-Workset-Name",
    "Another-Workset",
    
    # Or modify existing ones
    "AR-Civil Works-GF",
    # ... rest of list
]
```

### Naming Best Practices
- Use consistent prefixes (discipline codes)
- Include level information
- Keep names under 50 characters
- Use hyphens for readability
- Avoid special characters
- Be descriptive but concise

## Common Use Cases

### Use Case 1: New Structural Project Setup
```
1. Enable worksharing on new project
2. Run Workset Creator
3. Select All structural worksets (ST-*)
4. Create
Result: All structural worksets ready
```

### Use Case 2: Add Missing Link Worksets
```
1. Open project with models linked
2. Run Workset Creator
3. Select only LINK-* worksets
4. Create
Result: Link worksets organized by discipline
```

### Use Case 3: Specific Discipline Setup
```
1. Filter visually for specific prefix
2. Select only those worksets
3. Create
Result: Just the worksets you need
```

### Use Case 4: Multi-Level Building
```
1. Select all level-specific worksets
   (-GF, -FF, -RF)
2. Create
Result: Organized by building level
```

## Output Window Details

The script provides detailed logging:

```markdown
**Existing worksets**: 12
**Available to create**: 27

✓ Created: ST-Column-GF
✓ Created: ST-Column-FF
✓ Created: ST-SLAB-GF
✓ Created: ST-SLAB-FF
✗ Failed: ST-Foundation - Workset already exists

### ✅ Workset Creation Complete
**Success**: 4
**Errors**: 1
```

## Troubleshooting

### Problem: "Project is not workshared"

**Cause**: Project doesn't have worksharing enabled

**Solution**:
1. File → Worksharing → Enable Worksharing
2. Save project
3. Reopen project
4. Run script again

### Problem: "All worksets already exist"

**Cause**: All worksets in master list are already created

**Solutions**:
- Add new worksets to master list in Python file
- Or manually create additional worksets
- Script shows warning message automatically

### Problem: Can't create worksets

**Causes**:
- Don't own Shared Levels and Grids workset
- Don't have permission
- Project is read-only

**Solutions**:
- Make yourself owner of Shared Levels and Grids
- Check with BIM manager for permissions
- Ensure project is not read-only

### Problem: Workset creation fails

**Causes**:
- Workset name conflicts with existing
- Invalid characters in workset name
- Network issues (central model)

**Solutions**:
- Check output window for specific error
- Verify workset names are unique
- Ensure stable network connection to central

## Best Practices

### Before Creating
- [ ] Project is workshared
- [ ] You own necessary worksets
- [ ] Review master list for needed worksets
- [ ] Consider project requirements

### During Creation
- [ ] Select only worksets you need
- [ ] Don't create unnecessary worksets
- [ ] Review confirmation dialog
- [ ] Monitor output window

### After Creation
- [ ] Verify worksets were created
- [ ] Assign elements to appropriate worksets
- [ ] Set workset visibility settings
- [ ] Document workset usage for team

## Workset Management Tips

### Organization
- Create worksets by discipline
- Use consistent naming
- Include level information
- Group related elements

### Performance
- Don't create too many worksets
- Group similar elements together
- Balance between organization and performance
- Typical range: 20-50 worksets for most projects

### Team Coordination
- Document workset naming convention
- Share workset list with team
- Establish ownership rules
- Create workset usage guide

## Advanced Customization

### Adding Category-Specific Worksets

```python
WORKSETS = [
    # Walls by level
    "ST-WALL-GF",
    "ST-WALL-FF", 
    "ST-WALL-RF",
    
    # Walls by type
    "ST-WALL-Concrete",
    "ST-WALL-Masonry",
    "ST-WALL-Precast",
]
```

### Adding Project Phase Worksets

```python
WORKSETS = [
    # Phase-based
    "ST-Existing",
    "ST-Demo",
    "ST-New-Phase1",
    "ST-New-Phase2",
    "ST-Future",
]
```

### Adding Zone-Based Worksets

```python
WORKSETS = [
    # By building zone
    "ST-Tower-A",
    "ST-Tower-B",
    "ST-Podium",
    "ST-Basement",
]
```

## Integration with Other Tools

### With Workset Manager
1. Create worksets with this tool
2. Use Workset Manager to assign elements
3. Set visibility and ownership

### With View Templates
1. Create worksets first
2. Set up view templates
3. Control workset visibility per template

### With Filters
1. Create worksets
2. Create filters based on workset
3. Apply filters to views

## Keyboard Shortcuts

While window is active:
- **Escape**: Close window
- **Enter**: Create worksets (when button focused)

## Version History

### v1.0.0 (2024)
- Initial WPF release
- Separate XAML and Python files
- Modern dark theme
- Visible checkboxes (18x18px)
- Blue borders and styling
- Observable collections
- Info panel with counts
- 39 predefined worksets

## Support

For issues or feature requests, contact your BIM team or modify the script to fit your specific workflow needs.

## License

This script is provided as-is for internal use. Modify as needed for your organization's requirements.

---

**Created for**: BIM Modelers working with Revit Workshared Projects
**Interface Style**: Modern Dark Theme (matching Revision Manager)
**Technology**: WPF + Python + pyRevit
