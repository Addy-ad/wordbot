# Developer Guide - PowerShell Toolchain

This guide covers the PowerShell automation toolchain for building, installing, and developing Wordbot templates from source. This is strictly only for `Windows OS` as it needs word COM automation, windows specific libraries that are accessed from powershell.

## Overview

The PowerShell toolchain provides a complete development workflow for Wordbot:

- **Build templates** from VBA source files
- **Install templates** to Word's STARTUP folder
- **Live update macros** without restarting Word
- **Build multiple editions** in one command
- **Add ribbon XML** automatically
- **Manage Python dependencies** automatically

## Prerequisites

- Windows 10 or 11 (PowerShell toolchain is Windows-only)
- PowerShell 5.1 (the one that comes with windows preinstalled) or PowerShell 7+
- Microsoft Word (2016 or later)
- Python 3.7+ (for LLM and Research editions) (Tested with 3.11)
- VBA Object Model Trust enabled in Word

> [!IMPORTANT]
> The VBA object model trust setting is required for automated macro injection.
> Enable it in: Word > File > Options > Trust Center > Trust Center Settings > Macro Settings
> Check: Trust access to the VBA project object model

## Scripts Overview

The toolchain consists of three PowerShell scripts located in the `Powershell/` folder:

| Script | Purpose | Primary Use |
|--------|---------|-------------|
| `installWordBotTemplate.ps1` | Build and install a template | First-time setup, final builds |
| `docLiveUpdateMacros.ps1` | Update macros in a live Word session | Development, testing, iteration |
| `BuildAllEditions.ps1` | Build all three editions at once | Release builds, automation |

### Overall Parameter Explanations

> [!INFO]
> Look into each script descriptions for script specific parameters. 

- **Edition**: Selects which Wordbot version to build or update (Markdown, LLM, or Research).
- **Force**: Skips file overwrite prompts but preserves critical safety warnings like unsaved Word documents.
- **RemovePersonalInfo**: Strips personal metadata (author details, document properties) from output templates during save.
- **ContinueOnError**: Continues building remaining editions even if one fails. Only applies to BuildAllEditions.ps1.
- **Silent**: Suppresses detailed console output, showing only progress and errors. Only applies to BuildAllEditions.ps1.
- **UpdatePythonPaths**: Updates Python executable and server paths in VBA code by replacing placeholders with actual paths.
- **N/A**: Parameter is not applicable to this script.

### Getting Help

Each script includes built-in help:

```powershell
Get-Help .\installWordBotTemplate.ps1 -Full
Get-Help .\docLiveUpdateMacros.ps1 -Full
Get-Help .\BuildAllEditions.ps1 -Full
```

## Directory Structure

```
Powershell/
├── BuildAllEditions.ps1          # Build all editions
├── docLiveUpdateMacros.ps1       # Live macro updater
├── installWordBotTemplate.ps1    # Template installer
├── config.psd1                   # Configuration file
├── build/                        # Built templates output
│   ├── Wordbot_Markdown.dotm
│   ├── Wordbot_LLM.dotm
│   └── Wordbot_Research.dotm
├── Module/
│   └── Wordbot.psm1              # Core PowerShell module
└── Resources/
    ├── RibbonXML/                # Edition-specific ribbon XML
    │   ├── WordbotRibbon_Markdown.xml
    │   ├── WordbotRibbon_LLM.xml
    │   └── WordbotRibbon_Research.xml
    └── WordVBAmacros/            # VBA source files
        ├── Markdown_Main.bas
        ├── aWordbot_Main.bas
        └── ...
```

## Configuration File (config.psd1)

The `config.psd1` file controls all aspects of the build process:

```powershell
@{
    # Template Settings
    TemplateName    = 'Wordbot.dotm'    # Name of your template
    FileName        = 'Wordbot'          # Base filename (no extension)
    Extension       = '.dotm'            # .dotm for templates, .docm for documents
    outputPath      = ''                 # Empty = Word STARTUP, or specify path

    # File Locations (relative to install script)
    VBAFolderRelativePath   = 'Resources\WordVBAmacros'
    RibbonXmlRelativePath   = 'Resources\RibbonXML'
    PythonServerLocation    = '..\Python\main.py'

    # Python Configuration
    PythonExeLocation = ''   # Empty = auto-detect, or specify Python path
}
```

### Configuration Options

| Setting | Description | Valid Values |
|---------|-------------|--------------|
| `TemplateName` | Full name including extension | Any valid filename |
| `FileName` | Base filename without extension | Any valid filename |
| `Extension` | File type | `.dotm` (global template) or `.docm` (document) |
| `outputPath` | Installation location | `''` = Word STARTUP, absolute path = custom folder |
| `RibbonXmlRelativePath` | Folder containing edition-specific ribbon XML files | Relative path |
| `PythonExeLocation` | Python environment | `''` = system Python, directory = virtual env, full path = specific exe |

> [!NOTE]
> The script dynamically resolves edition-specific ribbon files using the naming pattern `WordbotRibbon_$Edition.xml` (e.g., `WordbotRibbon_Markdown.xml`). The `RibbonXmlRelativePath` setting specifies the folder containing these XML files.

## Script Reference

### 1. installWordBotTemplate.ps1 - Template Installer

**Purpose**: Build and install a VBA template into Word.

```powershell
.\installWordBotTemplate.ps1 [-Edition <Markdown|LLM|Research>] [-UpdatePythonPaths] [-RemovePersonalInfo] [-Force]
```

**Parameters**:

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `-Edition` | String | Edition to install: Markdown, LLM, or Research | Markdown |
| `-UpdatePythonPaths` | Switch | Update Python paths in VBA code | On (when run directly) |
| `-RemovePersonalInfo` | Switch | Strip personal metadata from output file | Off |
| `-Force` | Switch | Skip file overwrite prompts | Off |

**Examples**:

```powershell
# Run with defaults (Edition=Markdown, UpdatePythonPaths=On, Force=Off, RemovePersonalInfo=Off)
.\installWordBotTemplate.ps1

# Install LLM edition with Python paths updated
.\installWordBotTemplate.ps1 -Edition LLM

# Install Research edition and strip personal information
.\installWordBotTemplate.ps1 -Edition Research -RemovePersonalInfo

# Install Research edition without updating Python paths
.\installWordBotTemplate.ps1 -Edition Research -UpdatePythonPaths:$false

# Force overwrite existing files
.\installWordBotTemplate.ps1 -Edition Markdown -Force
```

> [!NOTE]
> The `-Force` parameter skips file overwrite prompts but does NOT skip safety prompts (Word running warning, .docm to startup warning). This prevents accidental data loss.

**What it does**:

1. Checks Word installation and VBA trust settings
2. Verifies Python environment and installs dependencies (LLM/Research only)
3. Creates a new Word template with your VBA code
4. Adds custom ribbon UI
5. Installs to Word's STARTUP folder or custom location
6. Compiles VBA to check for errors
7. Cleans up automatically

**When to use**:

- First-time installation
- After making changes to VBA source files
- After updating Python dependencies
- To reset or rebuild the template
- Creating production-ready templates

---

### 2. docLiveUpdateMacros.ps1 - Live Development Tool

**Purpose**: Update macros without closing Word. Good writing macros from external IDE such as VScode.

```powershell
.\docLiveUpdateMacros.ps1 [-Edition <Markdown|LLM|Research>] [-UpdatePythonPaths] [-Force]
```

**Parameters**:

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `-Edition` | String | Edition to update: Markdown, LLM, or Research | Markdown |
| `-UpdatePythonPaths` | Switch | Update Python paths in VBA code | On (when run directly) |
| `-Force` | Switch | Reserved for future use | Off |

**Examples**:

```powershell
# Update with defaults (Edition=Markdown, UpdatePythonPaths=On, Force=Off)
.\docLiveUpdateMacros.ps1

# Update LLM edition in live Word document
.\docLiveUpdateMacros.ps1 -Edition LLM

# Update Research edition without updating Python paths
.\docLiveUpdateMacros.ps1 -Edition Research -UpdatePythonPaths:$false
```

> [!CAUTION]
> This script does NOT add or update ribbon XML files. To use the updated macros, you must either:
> 1. Use an existing .dotm or .docm file that already has the ribbon XML embedded
> 2. Keep a document with the ribbon XML open in Word
> 3. Run the macros manually from the VBA editor (Alt+F8)

**What it does**:

1. Connects to existing Word instance or creates a new one
2. Detects your active document or creates a new one
3. Updates VBA code from source files
4. Compiles VBA to verify changes and check for errors
5. Preserves document state and content
6. Leaves Word open for continued work

**When to use**:

- During active VBA development
- Testing changes to macro code
- Iterative debugging without restarting Word
- Quick updates to active documents

---

### 3. BuildAllEditions.ps1 - Build All Editions

**Purpose**: Automatically build all three editions in one go.

```powershell
.\BuildAllEditions.ps1 [-Editions <string[]>] [-UpdatePythonPaths] [-RemovePersonalInfo] [-Force] [-ContinueOnError] [-Silent] [-LogPath <string>]
```

**Parameters**:

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `-Editions` | String | Editions to build: Markdown, LLM, Research | All three |
| `-UpdatePythonPaths` | Switch | Update Python paths in VBA code | Off |
| `-RemovePersonalInfo` | Switch | Strip personal metadata from built templates | On (when run directly) |
| `-Force` | Switch | Skip file overwrite prompts | On (when run directly) |
| `-ContinueOnError` | Switch | Continue if one edition fails | On (when run directly) |
| `-Silent` | Switch | Suppress detailed console output | On (when run directly) |
| `-LogPath` | String | Path to save build log | build_log.txt |

**Examples**:

```powershell
# Build all three editions with defaults (Force, ContinueOnError, Silent enabled)
.\BuildAllEditions.ps1

# Build all editions and update Python paths
.\BuildAllEditions.ps1 -UpdatePythonPaths

# Build only LLM and Research editions
.\BuildAllEditions.ps1 -Editions "LLM", "Research"

# Build Markdown with full console output (no silent mode)
.\BuildAllEditions.ps1 -Editions "Markdown" -Silent:$false

# Build with logging
.\BuildAllEditions.ps1 -LogPath "build.log"

# Build specific editions with Python updates and timestamped log
.\BuildAllEditions.ps1 -Editions "Markdown","LLM" -UpdatePythonPaths -LogPath "build_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
```

**When to use**:

- Building all editions for release
- Automated builds
- Creating distribution packages

## Default Behaviors (when running from VSCode for example)

When you run any script without parameters (e.g., pressing F5 in VSCode), it uses these defaults:

| Script | Default Edition | Force | RemovePersonalInfo | ContinueOnError | Silent | UpdatePythonPaths |
|--------|-----------------|-------|--------------------|-----------------|--------|-------------------|
| `installWordBotTemplate.ps1` | Markdown | Off | Off | N/A | N/A | On |
| `docLiveUpdateMacros.ps1` | Markdown | Off | N/A | N/A | N/A | On |
| `BuildAllEditions.ps1` | All three | On | On | On | On | Off |

These defaults can be changed inside respective scripts. for example look for these lines inside `docLiveUpdateMacros.ps1` to change defaults.

```powershell
# Default when running as script when running in VScode. 
if (-not $PSBoundParameters.ContainsKey('Edition')) { $Edition = "Markdown" }
if (-not $PSBoundParameters.ContainsKey('Force')) { $Force = $false }
if (-not $PSBoundParameters.ContainsKey('UpdatePythonPaths')) { $UpdatePythonPaths = $true }
if ($PSBoundParameters.Count -eq 0) { Clear-Host }
```

> [!NOTE]
> For BuildAllEditions.ps1, when building final templates for distribution, explicitly pass `-UpdatePythonPaths` to populate the placeholders `{{PYTHON_EXE_PATH}}` and `{{PYTHON_SERVER_PATH}}` in aWordbotRibbonLLMFunctions.bas.

## Development Workflow

The typical development loop is: configure -> build -> test -> edit -> live update -> repeat. Start by setting up your configuration file, then build the template, test it in Word, make changes to VBA source files, and use the live updater to see changes immediately without restarting Word.

### 1. Configure config.psd1

Edit `Powershell/config.psd1` to set your paths. This determines where templates are saved and where Python dependencies are located.

```powershell
@{
    # Template Settings
    TemplateName    = 'Wordbot.dotm'    # Name of your template
    FileName        = 'Wordbot'          # Base filename (no extension)
    Extension       = '.dotm'            # .dotm for templates, .docm for documents
    outputPath      = ''                 # Empty = Word STARTUP, or specify path

    # File Locations (relative to install script)
    VBAFolderRelativePath   = 'Resources\WordVBAmacros'
    RibbonXmlRelativePath   = 'Resources\RibbonXML'
    PythonServerLocation    = '..\Python\main.py'

    # Python Configuration
    PythonExeLocation = ''   # Empty = auto-detect, or specify Python path
}
```

> [!IMPORTANT]
> The `outputPath` setting determines where your built template will be saved. If left empty (`''`), it installs to Word's STARTUP folder. If you specify a custom path, templates will be saved there instead.

### 2. Build the Template

```powershell
# Navigate to the Powershell folder
cd Powershell

# Build and install the template
.\installWordBotTemplate.ps1
```

### 3. Develop Your Macros

Edit VBA source files in `Resources\WordVBAmacros\`:
- `.bas` - Standard modules
- `.cls` - Class modules (I have tested this in previous versions, but removed it. I know it works)
- `.frm` - User forms (I have not tested these. but you cant. My script just imports it.)

### 4. Live Test Your Changes

```powershell
# Update macros in Word without restarting
.\docLiveUpdateMacros.ps1
```

### 5. Build All Editions for Release

```powershell
# Build all three editions with Python path updates and logging
.\BuildAllEditions.ps1 -UpdatePythonPaths -LogPath "release_build.log"
```

### 6. Deploy to Production

- The built templates are in the `Powershell/build/` folder
- Copy the desired template to Word's STARTUP folder
- Restart Word to load the new template

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

# Option 1: Auto-detect - uses Python from system PATH, installs dependencies globally
PythonExeLocation = ''

# Option 2: Virtual environment folder - creates isolated environment at specified path, installs dependencies inside it
PythonExeLocation = 'D:\wordbotTemp'

# Option 3: Want to target a specific Python version WITH a virtual environment?
#    First create the venv: C:\Python39\python.exe -m venv D:\myproject\venv
#    Then point to the venv's python.exe:
#    PythonExeLocation = 'D:\myproject\venv\Scripts\python.exe'
PythonExeLocation = 'D:\wordbotTemp\venv\Scripts\python.exe'
```

### Python Path Substitution

The installer automatically updates VBA code with the correct Python paths by providing `-UpdatePythonPaths` switch:

```vba
' In aWordbotRibbonLLMFunctions.bas:
' {{PYTHON_EXE_PATH}} is replaced with your Python executable
' {{PYTHON_SERVER_PATH}} is replaced with your server path
```

## Troubleshooting

### VBA Trust Issues

```
[X] VBA project object model access is DISABLED
```

Enable "Trust access to the VBA project object model" in Word's Trust Center.

### Python Not Found

```
[-] Python was not found in PATH.
```

Install Python or specify `PythonExeLocation` in `config.psd1`.

### Word is Running

```
[-] Word is currently running
```

Close Word, or use the Force option in the installer.

### Template Already Loaded

```
IMPORTANT!!: Your project template 'Wordbot.dotm' is currently loaded in Word.
```

- Choose 'y' to edit loaded template directly
- Choose 'n' to close Word and work with a document

### File Locked/In Use

```
[-] Failed to remove: File may be open or locked
```

Close any programs using the file, or use Task Manager to end Word processes.