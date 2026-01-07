# Modern WPF Scripts Collection

Complete collection of three modernized pyRevit scripts with consistent WPF interfaces and dark theme styling.

## 📦 Package Overview

This package contains three professional BIM tools with modern WPF interfaces:

### 1. Title Block Replacer
**Purpose**: Batch replace title blocks on multiple sheets  
**Icon**: TB  
**Color**: Blue (#1976D2)

**Key Features:**
- Select title block family and type
- Filter and select sheets
- Preview current title blocks
- Bulk replacement operations
- Maintains title block position

### 2. View Title Renamer  
**Purpose**: Batch rename view titles on sheets with find/replace  
**Icon**: VT  
**Color**: Blue (#1976D2)

**Key Features:**
- Find and replace text
- Add custom prefix and suffix
- Use view name or current title
- Add sheet number automatically
- Live preview of changes

### 3. NWC Exporter
**Purpose**: Export Navisworks Cache files from Revit  
**Icon**: NW  
**Color**: Blue (#1976D2)

**Key Features:**
- Export from any 3D view
- Comprehensive export options
- Real-time system log
- Automatic or custom filename
- File size reporting

## 🎨 Design System

All three scripts share a consistent design language:

### Color Palette
```
Primary Blue:    #1976D2
Background:      #14141C
Panel:           #1E1E28
Input:           #0A0A0F
Border:          #2C2C36
Text:            #DCDCDC
Muted Text:      #808090
Success Green:   #4ADE80
```

### Typography
- **Headers**: Segoe UI, Bold
- **Labels**: Segoe UI, Regular
- **Data/Logs**: Consolas, Monospace
- **Buttons**: Segoe UI, SemiBold

### Layout Pattern
```
┌─────────────────────────────────┐
│  HEADER (Logo + Title + Close)  │
├───────────────┬─────────────────┤
│               │                 │
│  LEFT PANEL   │  RIGHT PANEL    │
│  (Selection)  │  (Options/Log)  │
│               │                 │
├───────────────┴─────────────────┤
│  FOOTER (Status + Action Button)│
└─────────────────────────────────┘
```

## 📁 Complete File Structure

```
ModernWPFScripts/
├── TitleBlockReplacer/
│   ├── TitleBlockReplacer.py
│   ├── TitleBlockReplacer.xaml
│   ├── README.md
│   └── MIGRATION_GUIDE.md
│
├── ViewTitleRenamer/
│   ├── ViewTitleRenamer.py
│   ├── ViewTitleRenamer.xaml
│   ├── ViewTitleRenamer_README.md
│   └── ViewTitleRenamer_COMPARISON.md
│
├── NWCExporter/
│   ├── NWCExporter.py
│   ├── NWCExporter.xaml
│   └── NWCExporter_README.md
│
└── Documentation/
    ├── MASTER_SETUP_GUIDE.md
    └── COLLECTION_README.md (this file)
```

## 🚀 Quick Setup

### Step 1: Prepare Extension Folder

Navigate to your pyRevit extensions directory:
```
%APPDATA%\pyRevit\Extensions\
```

Create your company extension structure:
```
Extensions\
└── [YourCompany].extension\
    └── [YourTab].tab\
        ├── Sheets.panel\
        ├── Views.panel\
        └── Export.panel\
```

### Step 2: Install Scripts

**Title Block Replacer** → `Sheets.panel\`
```
TitleBlockReplacer.pushbutton\
├── TitleBlockReplacer.py
└── TitleBlockReplacer.xaml
```

**View Title Renamer** → `Sheets.panel\` or `Views.panel\`
```
ViewTitleRenamer.pushbutton\
├── ViewTitleRenamer.py
└── ViewTitleRenamer.xaml
```

**NWC Exporter** → `Export.panel\`
```
NWCExporter.pushbutton\
├── NWCExporter.py
└── NWCExporter.xaml
```

### Step 3: Reload pyRevit

1. Open Revit
2. Go to **pyRevit** tab  
3. Click **Reload**
4. Look for buttons in respective panels

## 🎯 Usage Workflow

### For Title Block Replacement
```
1. Click Title Block Replacer button
2. Select new title block family & type
3. Filter and select sheets
4. Preview current title blocks
5. Click REPLACE TITLE BLOCKS
6. Confirm and execute
```

### For View Title Renaming
```
1. Click View Title Renamer button
2. Select views on sheets
3. Configure rename options:
   - Find/Replace text
   - Add prefix/suffix
   - Use view name option
4. Preview changes
5. Click APPLY RENAME
6. Confirm and execute
```

### For NWC Export
```
1. Click NWC Exporter button
2. Select source 3D view
3. Choose output directory
4. Set filename (auto or custom)
5. Configure export options
6. Click EXPORT NWC FILE
7. Monitor progress in log
```

## 📊 Comparison Table

| Feature | Title Block | View Title | NWC Export |
|---------|------------|------------|------------|
| **Primary Use** | Replace blocks | Rename titles | Export files |
| **Left Panel** | Title block selection | View selection | Export settings |
| **Right Panel** | Sheet selection + Preview | Options + Preview | System log |
| **Main Action** | Replace | Rename | Export |
| **Preview** | Current blocks | Before/After | Real-time log |
| **Batch Operation** | ✓ Multiple sheets | ✓ Multiple views | ✗ Single view |
| **Filter** | ✓ By sheet | ✓ By view/sheet | N/A |

## 🔧 Customization Guide

### Changing Primary Color

All scripts use the same primary blue. To change globally:

**In each XAML file, find:**
```xml
<Style x:Key="PrimaryButton">
    <Setter Property="Background" Value="#1976D2"/> <!-- Change here -->
</Style>
```

**Popular alternatives:**
```
Green:  #388E3C
Purple: #7B1FA2
Orange: #F57C00
Red:    #D32F2F
Teal:   #00897B
```

### Adding Company Logo

**In XAML, replace the icon:**
```xml
<!-- Find this: -->
<TextBlock Text="TB" ... /> <!-- or VT or NW -->

<!-- Replace with: -->
<Image Source="pack://application:,,,/logo.png" 
       Width="30" Height="30"/>
```

### Modifying Layout

Each script follows the same grid structure in XAML:
```xml
<Grid.ColumnDefinitions>
    <ColumnDefinition Width="[LeftWidth]"/>
    <ColumnDefinition Width="*"/>
</Grid.ColumnDefinitions>
```

Adjust column widths to rebalance panels.

## 🐛 Common Issues

### Issue: XAML Not Found

**Symptoms:** Error message on startup  
**Cause:** Files not in same folder  
**Solution:**
```
✓ TitleBlockReplacer.py
✓ TitleBlockReplacer.xaml
✓ Both in same .pushbutton folder
```

### Issue: Buttons Don't Appear

**Symptoms:** No buttons after reload  
**Causes:**
- Folder name doesn't end with `.pushbutton`
- Python file name doesn't match folder name
- Files not in correct extension structure

**Solution:** Verify folder structure exactly matches requirements

### Issue: UI Looks Pixelated

**Symptoms:** Blurry or pixelated interface  
**Cause:** Windows DPI scaling  
**Solution:**
1. Right-click Revit shortcut
2. Properties → Compatibility
3. Check "Override high DPI scaling behavior"
4. Select "System (Enhanced)"

### Issue: Script Works but UI Doesn't Match

**Cause:** Old Windows Forms version still running  
**Solution:**
1. Ensure XAML file is present
2. Check Python script is new WPF version
3. Clear pyRevit cache: pyRevit → Settings → Clear Cache
4. Reload pyRevit

## 📚 Documentation Index

### Getting Started
- **MASTER_SETUP_GUIDE.md** - Complete installation guide for all scripts
- **Individual README.md files** - Detailed usage for each script

### Migration Guides
- **MIGRATION_GUIDE.md** - Upgrading from Windows Forms to WPF
- **COMPARISON.md** - Before/after feature comparison

### Technical
- **XAML files** - UI layout and styling
- **Python files** - Business logic and Revit API integration

## 🎓 Training Resources

### For Users
1. **Quick Start Guide**: 5-minute overview of each tool
2. **Video Tutorials**: Screen recordings of common workflows
3. **Best Practices**: Tips for efficient usage
4. **FAQ Document**: Common questions and answers

### For Administrators
1. **Installation Guide**: Deploying to team
2. **Customization Guide**: Branding and modifications
3. **Troubleshooting**: Common issues and solutions
4. **Version Control**: Managing updates

### For Developers
1. **Code Structure**: How the scripts are organized
2. **WPF Patterns**: Data binding and MVVM concepts
3. **Revit API**: Integration points
4. **Extension Development**: Creating new tools

## 🔄 Update Process

### Updating Individual Scripts

1. **Backup Current Version**
   ```
   TitleBlockReplacer_OLD.pushbutton\
   ```

2. **Replace Files**
   - Keep XAML if you customized it
   - Update Python file
   - Merge XAML changes manually if needed

3. **Test**
   - Run script
   - Verify all features work
   - Check customizations still apply

4. **Document**
   - Note changes made
   - Update internal wiki
   - Inform team

### Version Tracking

Recommended version format in script header:
```python
__version__ = "2.0.0"  # WPF Version
__updated__ = "2024-01-15"
```

## 🌟 Best Practices

### File Organization
```
✓ Each script in own .pushbutton folder
✓ XAML file alongside Python file
✓ Optional: icon.png for button icon
✓ Optional: readme.txt for quick reference
```

### Naming Conventions
```
✓ Folder: ToolName.pushbutton
✓ Python: ToolName.py
✓ XAML: ToolName.xaml
✓ Consistent casing and naming
```

### Team Deployment
```
✓ Test locally first
✓ Deploy to test environment
✓ Train power users
✓ Roll out to team
✓ Collect feedback
✓ Iterate and improve
```

## 📈 Success Metrics

Track these metrics to measure adoption:

### Usage Metrics
- Number of users
- Frequency of use
- Time saved per operation
- Error rate

### Quality Metrics
- User satisfaction scores
- Bug reports
- Feature requests
- Training completion

### Business Impact
- Productivity improvement
- Error reduction
- Standardization improvement
- Training time reduction

## 🤝 Contributing

### Reporting Issues
1. Document the problem
2. Include steps to reproduce
3. Attach screenshots if relevant
4. Note Revit and pyRevit versions

### Suggesting Features
1. Describe the feature
2. Explain the use case
3. Provide examples
4. Consider implementation complexity

### Sharing Improvements
1. Test thoroughly
2. Document changes
3. Follow existing code style
4. Update relevant documentation

## 📞 Support Channels

### Internal Support
- BIM Coordinator
- Internal wiki/documentation
- Team chat (Slack/Teams)
- Monthly user group meetings

### External Resources
- pyRevit Forum: https://discourse.pyrevitlabs.io/
- Revit API Forum: https://forums.autodesk.com/
- Stack Overflow: revit-api tag

## ✅ Deployment Checklist

### Pre-Deployment
- [ ] All scripts tested individually
- [ ] XAML files validated
- [ ] Buttons appear correctly
- [ ] Icons added (optional)
- [ ] Colors customized (if needed)
- [ ] Documentation updated

### Deployment
- [ ] Files copied to extension folder
- [ ] Folder structure verified
- [ ] pyRevit reloaded successfully
- [ ] Buttons visible in Revit
- [ ] Scripts launch without errors

### Post-Deployment
- [ ] Quick functionality test
- [ ] User training scheduled
- [ ] Documentation distributed
- [ ] Feedback mechanism in place
- [ ] Success metrics defined

## 🎉 Benefits Summary

### For Users
- ✅ Modern, intuitive interface
- ✅ Consistent experience across tools
- ✅ Real-time feedback and preview
- ✅ Reduced errors
- ✅ Faster workflows

### For Administrators
- ✅ Easy to customize
- ✅ Simple to maintain
- ✅ Clear separation of UI and logic
- ✅ Professional appearance
- ✅ Scalable architecture

### For the Organization
- ✅ Increased productivity
- ✅ Better standardization
- ✅ Reduced training time
- ✅ Improved quality
- ✅ Modern toolset

## 📝 License & Credits

**License**: Internal use - modify as needed for your organization

**Technology Stack**:
- Windows Presentation Foundation (WPF)
- IronPython 2.7
- Revit API
- pyRevit Framework

**Design Inspiration**: Modern dark themes from VS Code, GitHub Desktop, and other contemporary developer tools

---

## 🚀 Getting Started

Ready to modernize your Revit workflow?

1. **Read the MASTER_SETUP_GUIDE.md**
2. **Install all three scripts**
3. **Test each tool with sample data**
4. **Customize colors if desired**
5. **Train your team**
6. **Enjoy the productivity boost!**

---

**Package Version**: 1.0.0  
**Compatible with**: Revit 2019+, pyRevit 4.8+  
**Last Updated**: 2024  
**Created for**: BIM Modelers using Revit API, Dynamo, and Python
