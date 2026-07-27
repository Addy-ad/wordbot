>Obviously this Readme.md is AI generated. But I suppose all this info is okay so far. When you run the `installWordBotTemplate.ps1` or `docLiveUpdateMacros.ps1` scripts, i tried as much as i can to ask before deleting files. so I guess those logic is well written. Besides, you can ask any AI to look into the code.

# Wordbot - PowerShell Automation for Word VBA Development

A comprehensive PowerShell toolkit for automating Microsoft Word VBA development, macro deployment, and template management with Python integration.

## Overview

Wordbot provides a complete workflow for developing, testing, and deploying VBA macros in Microsoft Word. It includes:

- **Live macro updating** without restarting Word
- **Automated template installation** to Word's STARTUP folder
- **Python server integration** for extended functionality
- **Ribbon UI customization**
- **Virtual environment management** for Python dependencies

## Prerequisites

### Required Software
- **Microsoft Word** (2010 or later)
- **PowerShell** 5.0 or later
- **Python** 3.7 or later (currently tested with 3.11)

### Windows Settings
1. **Enable VBA Object Model Trust** (Required for automated macro injection):
   ```
   Word → File → Options → Trust Center → Trust Center Settings → Macro Settings
   ✓ Trust access to the VBA project object model
   ```

2. **Enable Macros** (Recommended):
   ```
   Word → File → Options → Trust Center → Trust Center Settings → Macro Settings
   ○ Enable all macros (not recommended; potentially dangerous code can run)
   ```

## Installation

1. **Clone or download** the Wordbot repository to your local machine.

2. **Configure your settings** in `config.psd1`:
   ```powershell
   # Edit config.psd1 to match your preferences:
   # - TemplateName: Name of your .dotm template
   # - outputPath: Where to save the generated template
   # - PythonExeLocation: Path to Python environment (optional)
   ```

3. **Run the template installer**:
   ```powershell
   # From PowerShell, navigate to the Wordbot folder and run:
   .\installWordBotTemplate.ps1
   ```
   This will:
   - Check for Word installation
   - Verify Python dependencies
   - Build the VBA template with all macros
   - Install the template to Word's STARTUP folder
   - Add the custom ribbon interface

## Usage Guide

### Main Scripts

#### 1. `installWordBotTemplate.ps1` - Template Installer
**Purpose**: Build and install your VBA template into Word.

```powershell
.\installWordBotTemplate.ps1
```

**What it does**:
- Checks Word installation and VBA trust settings
- Verifies Python environment and installs dependencies
- Creates a new Word template with your VBA code
- Adds custom ribbon UI
- Installs to Word's STARTUP folder
- Cleans up automatically

**When to use**:
- First-time installation
- After making changes to VBA source files
- After updating Python dependencies
- To reset or rebuild the template

---

#### 2. `docLiveUpdateMacros.ps1` - Live Development Tool
**Purpose**: Update macros in real-time without restarting Word.

```powershell
.\docLiveUpdateMacros.ps1
```

**What it does**:
- Connects to existing Word instance or creates a new one
- Detects your active document or creates a new one
- Updates VBA code from source files
- Preserves document state and content

**When to use**:
- During active VBA development
- Testing changes to macro code
- Iterative debugging without restarting Word
- Quick updates to active documents

---

### Configuration File (`config.psd1`)

Control all aspects of Wordbot behavior:

```powershell
@{
    # Template Settings
    TemplateName    = 'Wordbot.dotm'    # Name of your template
    FileName        = 'Wordbot'          # Base filename (no extension)
    Extension       = '.dotm'            # .dotm for templates, .docm for documents
    outputPath      = ''                 # Empty = Word STARTUP, or specify path

    # File Locations (relative to install script)
    VBAFolderRelativePath   = 'Resources\WordVBAmacros'
    RibbonXmlRelativePath   = 'Resources\WordbotRibbon.xml'
    PythonServerLocation    = '..\Python\main.py'

    # Python Configuration
    PythonExeLocation = ''   # Empty = auto-detect, or specify Python path
}
```

#### Configuration Options

| Setting | Description | Valid Values |
|---------|-------------|--------------|
| `TemplateName` | Full name including extension | Any valid filename |
| `Extension` | File type for VBA macros | `.dotm` (global template) or `.docm` (document) |
| `outputPath` | Installation location | Empty (`''`) = Word STARTUP, absolute path = custom folder |
| `PythonExeLocation` | Python environment | Empty = system Python, directory = virtual env, full path = specific exe |

---

## VBA Development Workflow

### 1. Set Up Your Environment
```powershell
# Install/update the template
.\installWordBotTemplate.ps1
```

### 2. Develop Your Macros
- Edit source files in `Resources\WordVBAmacros\`
- Supported file types: `.bas` (modules), `.cls` (classes), `.frm` (forms)

### 3. Live Test Your Changes
```powershell
# Update macros in your Word document without restarting
.\docLiveUpdateMacros.ps1
```

### 4. Deploy to Production
- Run `installWordBotTemplate.ps1` one final time
- Copy the generated `.dotm` file to Word's STARTUP folder
- Restart Word to load the new template

---

## Python Integration

### Auto-Setup Features
Wordbot automatically:
1. Detects Python installation
2. Creates virtual environments when needed
3. Installs dependencies from `requirements.txt`
4. Configures paths for your Python server

### Python Configuration Options
```powershell
# In config.psd1:
PythonExeLocation = ''           # Auto-detection
PythonExeLocation = 'C:\Python39\python.exe'  # Specific Python version
PythonExeLocation = '.\venv'     # Virtual environment folder
```

### Python Path Substitution
The installer automatically updates VBA code with the correct Python paths:
```vba
' In aWordbotRibbonBtnFunctions.bas:
' {{PYTHON_EXE_PATH}} is replaced with your Python executable
' {{PYTHON_SERVER_PATH}} is replaced with your server path
```

---

## Common Use Cases

### Case 1: First-Time Setup
```powershell
# 1. Review and edit config.psd1
# 2. Ensure Python and Word are installed
# 3. Run installer
.\installWordBotTemplate.ps1
```

### Case 2: Developing Macros
```powershell
# 1. Edit your .bas/.cls/.frm files
# 2. Update macros in Word
.\docLiveUpdateMacros.ps1
# 3. Test in Word
# 4. Repeat steps 1-3 until ready
```

### Case 3: Deploying to Team
```powershell
# 1. Final build
.\installWordBotTemplate.ps1
# 2. Share the generated .dotm file
# 3. Teams can install manually or use your script
```

### Case 4: Debugging with Python Server
```powershell
# 1. Ensure requirements.txt is in the Python folder
# 2. Let Wordbot manage dependencies
# 3. Run installer - it auto-installs missing packages
.\installWordBotTemplate.ps1
```

---

## Troubleshooting

### VBA Trust Issues
```
[X] VBA project object model access is DISABLED
```
**Solution**: Enable "Trust access to the VBA project object model" in Word's Trust Center.

### Python Not Found
```
[-] Python was not found in PATH.
```
**Solution**: Install Python or specify `PythonExeLocation` in config.

### Word is Running
```
[-] Word is currently running
```
**Solution**: Close Word, or use the Force option in the installer.

### Template Already Loaded
```
IMPORTANT!!: Your project template 'Wordbot.dotm' is currently loaded in Word.
```
**Solution**: 
- Choose 'y' to edit loaded template directly
- Choose 'n' to close Word and work with a document

### File Locked/In Use
```
[-] Failed to remove: File may be open or locked
```
**Solution**: Close any programs using the file, or use Task Manager to end Word processes.


## Best Practices

1. **Always test with `docLiveUpdateMacros.ps1`** before final installation
2. **Keep VBA source files in version control** - not the compiled `.dotm`
3. **Use meaningful filenames** for VBA components
4. **Document your macros** with comments
5. **Test Python dependencies** with `pip freeze > requirements.txt`
6. **Use virtual environments** for Python development