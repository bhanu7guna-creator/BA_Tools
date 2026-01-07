# Modern WPF Scripts - Complete Setup Guide

Complete guide for setting up both **Title Block Replacer** and **View Title Renamer** with modern WPF interfaces.

## 📦 Package Contents

This package includes two modernized pyRevit scripts with WPF interfaces:

### 1. Title Block Replacer
```
TitleBlockReplacer.pushbutton/
├── TitleBlockReplacer.py      # Main logic
├── TitleBlockReplacer.xaml    # UI definition
└── README.md                   # Documentation
```

**Purpose**: Replace title blocks on multiple sheets at once
**Key Features**: Family/type selection, sheet filtering, current title block preview

### 2. View Title Renamer
```
ViewTitleRenamer.pushbutton/
├── ViewTitleRenamer.py        # Main logic
├── ViewTitleRenamer.xaml      # UI definition
└── README.md                   # Documentation
```

**Purpose**: Batch rename view titles on sheets with find/replace
**Key Features**: Find/replace, prefix/suffix, use view name, sheet number prefix

## 🎨 Design Features

Both scripts share a consistent modern dark theme:

- **Color Scheme**: Dark backgrounds (#14141C) with blue accents (#1976D2)
- **Typography**: Segoe UI for labels, Consolas for data
- **Layout**: Two-panel design (selection + options)
- **Styling**: Flat modern buttons, clean borders, professional appearance
- **Icons**: Custom logo badges (TB and VT)

## 🚀 Quick Setup (Recommended)

### Step 1: Locate Your pyRevit Extensions Folder
```
%APPDATA%\pyRevit\Extensions\
```
Or in Windows Explorer:
```
C:\Users\[YourName]\AppData\Roaming\pyRevit\Extensions\
```

### Step 2: Create/Navigate to Your Extension Structure
```
Extensions\
└── [YourCompany].extension\
    └── [YourTab].tab\
        └── Sheets.panel\          # Or any panel you prefer
```

If these don't exist, create them:
```
Extensions\
└── MyCompany.extension\
    └── BIM.tab\
        └── Sheets.panel\
```

### Step 3: Create Button Folders

Inside your panel folder, create two folders:
```
Sheets.panel\
├── TitleBlockReplacer.pushbutton\
└── ViewTitleRenamer.pushbutton\
```

### Step 4: Copy Files

**Title Block Replacer:**
```
TitleBlockReplacer.pushbutton\
├── TitleBlockReplacer.py
└── TitleBlockReplacer.xaml
```

**View Title Renamer:**
```
ViewTitleRenamer.pushbutton\
├── ViewTitleRenamer.py
└── ViewTitleRenamer.xaml
```

### Step 5: Reload pyRevit

1. Open Revit
2. Go to **pyRevit** tab
3. Click **Reload**
4. Look for your new buttons in the panel

### Step 6: Test

1. Open a Revit project with sheets
2. Click each button to verify they load
3. Test basic functionality

## 📁 Complete Folder Structure Example

```
%APPDATA%\pyRevit\Extensions\
└── MyCompany.extension\
    ├── lib\                                  # Optional: shared libraries
    └── BIM.tab\
        ├── Sheets.panel\
        │   ├── TitleBlockReplacer.pushbutton\
        │   │   ├── TitleBlockReplacer.py
        │   │   └── TitleBlockReplacer.xaml
        │   └── ViewTitleRenamer.pushbutton\
        │       ├── ViewTitleRenamer.py
        │       └── ViewTitleRenamer.xaml
        ├── Views.panel\
        └── Modeling.panel\
```

## 🎯 Button Icons (Optional)

To add custom icons to your buttons:

### Create Icons
1. Create 32x32 PNG images
2. Name them `icon.png`
3. Place in each button folder:

```
TitleBlockReplacer.pushbutton\
├── icon.png                    # Your custom icon
├── TitleBlockReplacer.py
└── TitleBlockReplacer.xaml
```

### Icon Design Tips
- Use 32x32 pixels
- Keep it simple and recognizable
- Use colors that match your theme
- PNG format with transparency

## ⚙️ Configuration Options

### Customizing Colors

Both scripts use consistent colors defined in XAML. To change the color scheme:

**Edit the XAML file, find `<Window.Resources>` section:**

```xml
<Window.Resources>
    <!-- Primary Button Color -->
    <Style x:Key="PrimaryButton" TargetType="Button">
        <Setter Property="Background" Value="#1976D2"/> <!-- Change this -->
    </Style>
    
    <!-- Background Colors -->
    Background="#14141C"  <!-- Main background -->
    Background="#1E1E28"  <!-- Panel background -->
    Background="#0A0A0F"  <!-- Input background -->
</Window.Resources>
```

### Common Color Schemes

**Blue Theme (Current):**
```xml
Primary: #1976D2
Background: #14141C
```

**Green Theme:**
```xml
Primary: #388E3C
Background: #1B1B1B
```

**Purple Theme:**
```xml
Primary: #7B1FA2
Background: #1A1A1A
```

**Orange Theme:**
```xml
Primary: #F57C00
Background: #1C1C1C
```

## 🔧 Advanced Setup

### Company Branding

Add your company logo to the header:

**In XAML, find the logo section:**
```xml
<Border Background="#1976D2" ...>
    <TextBlock Text="TB" ... /> <!-- Replace with Image control -->
</Border>
```

**Replace with:**
```xml
<Border>
    <Image Source="pack://application:,,,/company_logo.png" />
</Border>
```

### Adding Keyboard Shortcuts

**In Python, add to SetupEventHandlers:**
```python
def SetupEventHandlers(self):
    # ... existing code ...
    
    # Add keyboard shortcuts
    self._window.KeyDown += self.OnKeyDown

def OnKeyDown(self, sender, args):
    # Ctrl+A = Select All
    if args.Key == System.Windows.Input.Key.A and \
       args.KeyboardDevice.Modifiers == System.Windows.Input.ModifierKeys.Control:
        self.OnSelectAllSheets(None, None)
```

### Multi-Monitor Support

The windows are already set to `WindowStartupLocation="CenterScreen"`, but you can change this:

```xml
<!-- In XAML -->
WindowStartupLocation="CenterOwner"  <!-- Center on Revit window -->
WindowStartupLocation="Manual"       <!-- Custom position -->
```

## 🐛 Troubleshooting

### Problem: "XAML file not found"

**Cause**: Python and XAML files not in same folder

**Solution**: 
```bash
# Verify both files exist:
TitleBlockReplacer.pushbutton\
├── TitleBlockReplacer.py      ✓
└── TitleBlockReplacer.xaml    ✓
```

### Problem: Buttons don't appear after reload

**Possible causes:**
1. Folder structure incorrect
2. Files named incorrectly
3. pyRevit cache issue

**Solutions:**
```
1. Check folder names end with .pushbutton
2. Check .py files are named exactly like folder
3. Try: pyRevit → Settings → Clear Cache
```

### Problem: UI looks wrong or pixelated

**Cause**: Windows DPI scaling

**Solution**:
- Right-click Revit shortcut
- Properties → Compatibility
- Override high DPI scaling behavior
- System (Enhanced)

### Problem: Preview not updating

**Cause**: Events not connected

**Solution**: Check that all TextChanged and Checked events are wired in `SetupEventHandlers()`

### Problem: Can't modify parameters

**Cause**: Worksharing/permissions

**Solutions:**
- Ensure worksharing elements aren't owned by others
- Check workset permissions
- Verify elements aren't in design options

## 📊 Testing Checklist

### Title Block Replacer
- [ ] Window opens without errors
- [ ] Title block families populate
- [ ] Types populate when family selected
- [ ] Sheet filter works
- [ ] Select All/Clear All work
- [ ] Current title blocks display
- [ ] Replace operation succeeds
- [ ] Transaction commits properly

### View Title Renamer
- [ ] Window opens without errors
- [ ] Views populate in list
- [ ] View filter works
- [ ] Selection buttons work
- [ ] Find/replace works
- [ ] Prefix/suffix work
- [ ] Preview updates correctly
- [ ] Rename operation succeeds

## 📚 Additional Resources

### Documentation
- `README.md` files for each script (detailed usage)
- `MIGRATION_GUIDE.md` (Title Block Replacer)
- `COMPARISON.md` (View Title Renamer)

### Learning Resources
- **WPF**: https://docs.microsoft.com/en-us/dotnet/desktop/wpf/
- **pyRevit**: https://pyrevitlabs.notion.site/
- **Revit API**: https://www.revitapidocs.com/

### Community
- pyRevit Forum: https://discourse.pyrevitlabs.io/
- Revit API Forum: https://forums.autodesk.com/t5/revit-api-forum/bd-p/160

## 🔄 Updating Scripts

To update scripts later:

1. **Backup current version** (copy folder)
2. **Replace files** with new versions
3. **Keep XAML if customized** (may need manual merge)
4. **Test** before deploying to team
5. **Update documentation** if you made changes

## 👥 Team Deployment

### Small Team (< 10 people)
1. Set up scripts on network drive
2. Share folder path with team
3. Each user adds to their extensions manually

### Large Team (> 10 people)
1. Create company extension
2. Use pyRevit extension installer
3. Deploy via network or shared drive
4. Document in team wiki

### Best Practices
- Version control your extensions (Git)
- Document customizations
- Test updates before deployment
- Train users on new features
- Collect feedback for improvements

## 🎓 Training Users

### Quick Start Guide
1. Show where buttons are located
2. Demonstrate basic workflow
3. Show common use cases
4. Explain preview feature
5. Emphasize undo capability

### Advanced Training
- Custom color schemes
- Keyboard shortcuts
- Batch operations
- Filtering techniques
- Troubleshooting common issues

## 📞 Support

### Internal Support
- Create internal wiki page
- Document common issues
- Assign BIM coordinator for questions

### External Support
- Original script documentation
- pyRevit community forums
- Revit API documentation

## ✅ Post-Installation Checklist

- [ ] Both scripts installed correctly
- [ ] Buttons visible in Revit
- [ ] XAML files in correct location
- [ ] Colors match company standards (if customized)
- [ ] Icons added (if desired)
- [ ] Tested on sample project
- [ ] Team trained on usage
- [ ] Documentation updated
- [ ] Backup of original scripts saved

## 🎉 Success!

Your modern WPF scripts are now ready to use. Enjoy the improved interface and easier maintenance!

---

**Package Version**: 1.0.0
**Compatible with**: Revit 2019+, pyRevit 4.8+
**Last Updated**: 2024
**Created for**: BIM Modelers using Revit API and Python
