# Set Comment Values - WPF Modern UI

A modern, dark-themed WPF interface for assigning comment values to Revit elements filtered by category, level, and type with multiple assignment methods.

## Features

- ✅ Modern dark-themed interface matching your workflow tools
- ✅ Separate XAML and Python files for easy customization
- ✅ Four-step workflow (Category → Level → Type → Method)
- ✅ Five assignment methods
- ✅ Live preview before applying
- ✅ Detailed element information
- ✅ Mark-based sequential numbering
- ✅ Comprehensive logging

## File Structure

```
SetCommentValues/
├── SetCommentValues.py      # Main Python script
└── SetCommentValues.xaml    # WPF UI definition
```

## Installation

### Method 1: pyRevit Extension (Recommended)

1. Create a new folder in your pyRevit extension directory:
   ```
   %APPDATA%\pyRevit\Extensions\[YourCompany].extension\[YourTab].tab\Modify.panel\
   ```

2. Create a folder called `SetCommentValues.pushbutton` inside the panel folder

3. Copy both files into this folder:
   ```
   SetCommentValues.pushbutton/
   ├── SetCommentValues.py
   └── SetCommentValues.xaml
   ```

4. Reload pyRevit (pyRevit → Reload)

5. The button will appear on your panel

## Usage Guide

### Four-Step Workflow

#### STEP 1: Select Category
Choose the element category to filter:
- Walls
- Doors
- Windows
- Structural Framing
- Structural Columns
- Floors
- Furniture
- Generic Models

#### STEP 2: Select Level
Choose the level where elements are located. Displays:
- Level name
- Elevation value

#### STEP 3: Select Type
Choose the specific type from available types on that level. Shows:
- Type name
- Element count for that type

#### STEP 4: Assign Comments
Choose assignment method and configure parameters.

### Assignment Methods

#### 1. Single Value
Apply the same comment value to all selected elements.

**Use Case**: Mark all elements with same category
```
Example: "STRUCTURAL STEEL"
Result: All elements get "STRUCTURAL STEEL"
```

**Input:**
- Comment Value: Text to apply to all elements

#### 2. Sequential Numbers
Number elements sequentially starting from a number.

**Use Case**: Simple numbering scheme
```
Start: 1, Step: 1
Result: 1, 2, 3, 4, 5...

Start: 10, Step: 5
Result: 10, 15, 20, 25, 30...
```

**Inputs:**
- Start Number: Initial number
- Step: Increment value

#### 3. Sequential with Prefix
Add a prefix before sequential numbers.

**Use Case**: Department or zone coding
```
Prefix: "ZONE-A", Start: 1, Step: 1
Result: ZONE-A-1, ZONE-A-2, ZONE-A-3...

Prefix: "COL", Start: 101, Step: 1
Result: COL-101, COL-102, COL-103...
```

**Inputs:**
- Start Number: Initial number
- Step: Increment value
- Prefix: Text before number

#### 4. Comma-Separated
Provide exact values for each element (comma or line-separated).

**Use Case**: Custom values from external source
```
Input: 
ABC, DEF, GHI
or
ABC
DEF
GHI

Result: Element 1 = ABC, Element 2 = DEF, Element 3 = GHI
```

**Requirements:**
- Must provide exactly as many values as selected elements
- Can use commas or newlines as separators

#### 5. Based on Mark Order
Sort elements by Mark parameter and number sequentially.

**Use Case**: Follow existing Mark numbering scheme
```
Elements with Marks: B1, B3, B5, B7
Prefix: "STRUCT"
Result: 
  Mark B1 → STRUCT-001
  Mark B3 → STRUCT-003
  Mark B5 → STRUCT-005
  Mark B7 → STRUCT-007
```

**Features:**
- Natural sorting (B1, B2, B10 not B1, B10, B2)
- Extracts last number from Mark
- Pads numbers to 3 digits (001, 002, etc.)
- Only processes elements that have Mark values

**Input:**
- Prefix: Text before padded number

## UI Components

### Header
- **CM Icon**: Comment Manager branding
- **Title**: Tool name and description
- **Close Button**: Exit without applying

### Left Panel (Element Selection)
- **STEP 1**: Category dropdown
- **STEP 2**: Level dropdown (enabled after category)
- **STEP 3**: Type dropdown (enabled after level)
- **Element Count**: Shows selected element count
- **Element Info**: Shows first 20 elements with IDs and current comments

### Right Panel (Comment Assignment)
- **STEP 4**: Method selection
- **Input Area**: Method-specific inputs (changes based on method)
- **Preview Area**: Shows before/after for first 15 elements

### Footer
- **Status**: Current operation status
- **Generate Preview**: Preview assignments before applying
- **Apply Comments**: Apply comment values to elements

## Workflow Examples

### Example 1: Number Structural Columns
```
Step 1: Select "Structural Columns"
Step 2: Select "Level 1"
Step 3: Select column type (e.g., "300x300mm (12)")
Step 4: Choose "Sequential with Prefix"
  - Prefix: "COL"
  - Start: 1
  - Step: 1
Preview: COL-1, COL-2, COL-3...
Apply: Comments updated
```

### Example 2: Mark Furniture by Zone
```
Step 1: Select "Furniture"
Step 2: Select "Level 2"  
Step 3: Select furniture type
Step 4: Choose "Comma-Separated"
Input: ZONE-A, ZONE-A, ZONE-B, ZONE-B...
Preview: Check assignments
Apply: Zone labels applied
```

### Example 3: Sequential Beam Numbering
```
Step 1: Select "Structural Framing"
Step 2: Select "Level 1"
Step 3: Select beam type
Step 4: Choose "Based on Mark Order"
  - Prefix: "BEAM"
(Elements sorted by Mark: B1, B2, B3...)
Preview: BEAM-001, BEAM-002, BEAM-003...
Apply: Numbered by mark order
```

### Example 4: Same Value for All Doors
```
Step 1: Select "Doors"
Step 2: Select "Ground Floor"
Step 3: Select door type
Step 4: Choose "Single Value"
  - Value: "FIRE RATED"
Preview: All doors show "FIRE RATED"
Apply: Value applied to all
```

## Preview Feature

Before applying, use **GENERATE PREVIEW** to see:
- Element ID
- Current comment (or "(empty)")
- New comment value
- First 15 elements displayed

**Example Preview:**
```
PREVIEW: 25 ELEMENTS
============================================================

1. ID: 123456 | (empty) -> COL-1
2. ID: 123457 | Old value -> COL-2
3. ID: 123458 | (empty) -> COL-3
...
15. ID: 123470 | (empty) -> COL-15

... 10 more elements
```

## Level Detection Logic

The script tries multiple parameters to find element levels (in priority order):

1. **Family Level** (`FAMILY_LEVEL_PARAM`)
2. **Reference Level** (`INSTANCE_REFERENCE_LEVEL_PARAM`)
3. **Schedule Level** (`SCHEDULE_LEVEL_PARAM`)
4. **Base Constraint** (for walls - `WALL_BASE_CONSTRAINT`)

This ensures most elements are correctly filtered by level.

## Natural Sorting in Mark Order

When using "Based on Mark Order", the script uses natural sorting:

**Standard Sort** (wrong):
```
B1, B10, B2, B20, B3
```

**Natural Sort** (correct):
```
B1, B2, B3, B10, B20
```

This ensures logical ordering for construction sequences.

## Output Window Logging

Detailed logging for each element:

```markdown
✓ ID: 123456 → COL-1
✓ ID: 123457 → COL-2
✗ ID: 123458 - Parameter read-only or not found

### ✅ Comment Values Applied
**Success**: 24
**Errors**: 1
```

## Troubleshooting

### Problem: No elements found on level

**Causes:**
- Elements don't have level parameter set
- Elements are on different level than selected
- Wrong category selected

**Solutions:**
- Check element properties in Revit
- Try different level parameter
- Verify element category

### Problem: Can't apply to some elements

**Causes:**
- Comments parameter is read-only
- Comments parameter doesn't exist
- Element is in linked model

**Solutions:**
- Check if Comments parameter exists in element
- Verify parameter isn't read-only
- Only works on model elements, not linked

### Problem: Mark order doesn't work

**Causes:**
- Elements don't have Mark values
- Mark parameter is empty
- Wrong Mark format

**Solutions:**
- Assign Mark values to elements first
- Check Mark parameter has values
- Use standard Revit Mark parameter

### Problem: Comma-separated count mismatch

**Cause:** Number of values doesn't match element count

**Solution:**
- Count elements: shows "X elements"
- Provide exactly X values
- Check for extra commas or line breaks

## Best Practices

### Before Applying
- [ ] Preview first to check assignments
- [ ] Verify element count matches expectations
- [ ] Check current comment values
- [ ] Test on small subset first

### During Application
- [ ] Use meaningful prefixes
- [ ] Follow project naming conventions
- [ ] Document numbering schemes
- [ ] Consider construction sequence

### After Application
- [ ] Verify in schedules
- [ ] Check sample elements in model
- [ ] Document the scheme used
- [ ] Create schedule filters

## Tips & Tricks

### Tip 1: Export to Schedule
After applying comments, create a schedule:
- Add Comment parameter column
- Sort by Comment
- Export to Excel for tracking

### Tip 2: Use Consistent Prefixes
Establish project standards:
```
Structural: ST-, STR-, STRUCT-
Architecture: AR-, ARCH-
MEP: ME-, MEP-
```

### Tip 3: Pad Numbers
Use Mark Order method for:
- 3-digit padding (001, 002...)
- Consistent width in schedules
- Better sorting

### Tip 4: Combine with Filters
1. Apply comment values
2. Create view filters based on comments
3. Control visibility by comment

### Tip 5: Incremental Numbering
For additions to existing numbering:
```
Existing: COL-1 to COL-20
New: Start at 21, not 1
```

## Advanced Usage

### Custom Category Support

To add more categories, edit the Python file:

```python
self.category_map = {
    # Existing categories...
    'Casework': BuiltInCategory.OST_Casework,
    'Plumbing Fixtures': BuiltInCategory.OST_PlumbingFixtures,
    'Lighting Fixtures': BuiltInCategory.OST_LightingFixtures,
}
```

### Custom Mark Formatting

Edit the Mark Order logic in Python:

```python
# Change padding width (default 3)
num_padded = str(num_int).zfill(4)  # 4-digit: 0001

# Change separator
elem_to_comment[elem.Id] = "{}.{}".format(prefix, num_padded)  # Dot separator
```

## Requirements

- Autodesk Revit 2019 or later
- pyRevit 4.8 or later
- .NET Framework 4.7 or later
- Elements must have Comments parameter (standard Revit parameter)

## Version History

### v1.0.0 (2024)
- Initial WPF release
- Separate XAML and Python files
- Modern dark theme
- Five assignment methods
- Preview functionality
- Natural sorting for marks
- Comprehensive logging

## Support

For issues or feature requests, contact your BIM team or modify the script to fit your specific workflow needs.

## License

This script is provided as-is for internal use. Modify as needed for your organization's requirements.

---

**Created for**: BIM Modelers working with Revit, Revit API, Dynamo, and Python
**Interface Style**: Modern Dark Theme (matching Revision Manager)
**Technology**: WPF + Python + pyRevit
