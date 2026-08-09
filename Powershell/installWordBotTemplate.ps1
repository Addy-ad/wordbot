<#
.SYNOPSIS
    Installs a Wordbot template edition (Markdown, LLM, or Research)

.DESCRIPTION
    This script builds and installs a Wordbot template (.dotm or .docm) for the specified edition.
    It creates a new Word document, imports VBA modules, compiles the VBA project, and adds the
    appropriate ribbon XML.

    The script supports three editions:
    - Markdown: Core Markdown formatting and conversion features
    - LLM: Markdown features + Large Language Model integration
    - Research: LLM features + Research and citation tools

    Key Features:
    - Automatic virtual environment setup for Python dependencies (LLM/Research)
    - Add macros according to the selected edition
    - Adds a word Ribbon to run added macros
    - Compiles VBA to check for errors

.PARAMETER Edition
    Specifies which edition to install. Valid values: "Markdown", "LLM", "Research"
    Default: "Markdown"

.PARAMETER UpdatePythonPaths
    Updates Python executable and server paths in the VBA code.
    When enabled, replaces placeholders {{PYTHON_EXE_PATH}} and {{PYTHON_SERVER_PATH}}
    in aWordbotRibbonLLMFunctions.bas with actual paths.
    Default: $true (enabled) when running the script directly

.PARAMETER Force
    Skips confirmation prompts for non-critical operations.
    Prompts skipped with -Force:
    - "File exists, overwrite?" - Automatically overwrites
    - "Remove existing file?" - Automatically removes
    
    Prompts NOT skipped with -Force (Safety First):
    - Word is running with unsaved documents - Always prompts to prevent data loss
    - Installing .docm to STARTUP folder - Always warns about potential Word startup issues
    
    Default: $false

.EXAMPLE
    .\installWordBotTemplate.ps1 -Edition Markdown
    Installs the Markdown edition with default settings. Python paths are updated (since
    UpdatePythonPaths defaults to $true). Prompts before overwriting existing files.

.EXAMPLE
    .\installWordBotTemplate.ps1 -Edition LLM -Force
    Installs the LLM edition and forces overwrite of existing files without prompting.
    Python paths are updated. Safety prompts for Word running and .docm warnings remain.

.EXAMPLE
    .\installWordBotTemplate.ps1 -Edition Research -UpdatePythonPaths:$false
    Installs the Research edition without updating Python paths. Useful when Python
    environment hasn't changed and you want to save time.

.EXAMPLE
    .\installWordBotTemplate.ps1 -Edition Markdown -Force:$false
    Installs the Markdown edition with explicit prompting for file overwrites.
    Equivalent to running without -Force.

.EXAMPLE
    .\installWordBotTemplate.ps1 -Edition LLM -UpdatePythonPaths:$false -Force
    Installs the LLM edition without updating Python paths, but forces file overwrites.
    Useful for quick rebuilds when only VBA code has changed.

.EXAMPLE
    .\installWordBotTemplate.ps1 -Edition Markdown -UpdatePythonPaths
    Installs the Markdown edition with Python paths updated (explicitly enabled).
    Note: UpdatePythonPaths has no effect on Markdown edition (no Python dependencies).

.EXAMPLE
    .\installWordBotTemplate.ps1
    Installs the Markdown edition with default settings (since Edition defaults to Markdown).
    Python paths are updated (default behavior).

.LINK
    .\BuildAllEditions.ps1
    https://github.com/Addy-ad/wordbot

.NOTES
    To view this help: Get-Help .\installWordBotTemplate.ps1 -Full

    Author: Addy-ad
    Requires: PowerShell 5.1 or later, Windows, Microsoft Word installed

    Python Dependencies:
    - Markdown edition: No Python required
    - LLM edition: Python (Tested 3.11) with packages from requirements.txt
    - Research edition: Python (Tested 3.11) with packages from requirements.txt

    File Output:
    - Output location: Configurable via config.psd1 (outputPath setting)
    - Default output: Word STARTUP folder or config-specified location
    - File format: .dotm (macro-enabled template) (default) or .docm (macro-enabled document)

    Safety Notes:
    - This script requires "Trust access to the VBA project object model" to be enabled
    - Will prompt to close Word during installation
#>

# installWordBotTemplate.ps1
param (
    [ValidateSet("Markdown", "LLM", "Research")]
    [string]$Edition,
    [switch]$UpdatePythonPaths,
    [switch]$Force
)

# Clear local variables before exiting script scope
$varsToClear = @(
    "config", "fileName", "extension", "outputPath", "paths",
    "pythonExe", "choice", "confirm", "startupFilePath",
    "word", "doc", "project", "modulePath","hasCompileError"
)

$word = $null
$doc = $null

trap {
    Write-Host "`n[!] Unhandled Exception Caught!" -ForegroundColor Red
    
    $exceptionMsg = $_.Exception.Message
    $scriptLine   = $_.InvocationInfo.ScriptLineNumber
    $commandName  = $_.InvocationInfo.MyCommand

    if ($commandName) { Write-Host "    - Command: $commandName" -ForegroundColor Yellow }
    if ($scriptLine)  { Write-Host "    - Line:    $scriptLine" -ForegroundColor Yellow }
    Write-Host "    - Details: $exceptionMsg" -ForegroundColor Red

    # Cleanup handles if an unexpected crash bypassed manual checks
    Invoke-ScriptCleanup -WordApp $word -Document $doc

    foreach ($var in $varsToClear) {
        Remove-Variable -Name $var -ErrorAction SilentlyContinue
    }

    Write-Host ""

    exit 1
}

# Default when running as script when running in VScode. 
if (-not $PSBoundParameters.ContainsKey('Edition')) { $Edition = "Research" }
if (-not $PSBoundParameters.ContainsKey('Force')) { $Force = $false }
# Default to updating Python paths when running the script directly
if (-not $PSBoundParameters.ContainsKey('UpdatePythonPaths')) { $UpdatePythonPaths = $true }
if ($PSBoundParameters.Count -eq 0) { Clear-Host }

Add-Type -AssemblyName WindowsBase

# Import the module
$modulePath = Join-Path $PSScriptRoot "Module\Wordbot.psm1"
Import-Module $modulePath -Force

$config = Get-WordbotConfig
if (-not $config) {
    Write-Host "[-] Config file not found or invalid. Please ensure config.psd1 exists." -ForegroundColor Red
    Write-Host "    Expected location: $(Join-Path (Split-Path $PSScriptRoot -Parent) 'config.psd1')" -ForegroundColor Yellow
    foreach ($var in $varsToClear) {
        Remove-Variable -Name $var -ErrorAction SilentlyContinue
    }
    exit 1
}
# 1. Target output name based on Edition
$fileName = "Wordbot_$Edition"
$extension = $config.Extension
$outputPath = $config.outputPath

Write-Host "Wordbot Template Installer - [$Edition Edition]" -ForegroundColor Cyan
Write-Host "------------------------------------------------" -ForegroundColor Cyan
Write-Host ""

# Test if Word is installed
Write-Host "Step 1: Checking Word installation..." -ForegroundColor Yellow
if (-not (Test-WordInstalled)) {
    foreach ($var in $varsToClear) {
        Remove-Variable -Name $var -ErrorAction SilentlyContinue
    }
    exit 1
}

# Step 2: Check if Word is running
Write-Host "Step 2: Checking if Word is running..." -ForegroundColor Yellow
if (Test-WordRunning) {
    Write-Host "[!] Word is currently running. Proceeding will close Word without saving unsaved changes!" -ForegroundColor Red
    $confirm = Read-Host "Do you want to force close Word? (y/n)"
    if ($confirm.Trim().ToLower() -eq 'y') {
        Write-Host "[*] Force closing Word processes..." -ForegroundColor Cyan
        Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } else {
        Write-Host "[*] Operation cancelled by user. Please save your work and close Word manually." -ForegroundColor Yellow
        exit 1
    }
}
Write-Host ""

Write-Host "Step 3: Getting paths..." -ForegroundColor Yellow

# Pass $Edition directly into Get-Paths to resolve all paths (including edition-specific XML)
$paths = Get-Paths -fileName $fileName -extension $extension -outputPath $outputPath -Config $config -Edition $Edition

Write-Host "    Script path:        $($paths.scriptPath)"
Write-Host "    File name:          $($paths.fullFileName)"
Write-Host "    Output Folder:      $($paths.outputFolder)"
Write-Host "    FullFile path:      $($paths.fullFilePath)"
Write-Host "    Word startup path:  $($paths.wordStartupPath)"
Write-Host "    VBA folder:         $($paths.VBAFolder)"
Write-Host "    Ribbon XML:         $($paths.ribbonXml)"
Write-Host "    Python server path: $($paths.pythonServerPath)"
Write-Host ""

# Step 3.14159: Python Check (Only required for LLM and Research editions)
$pythonExe = $null
if ($Edition -ne "Markdown") {
    Write-Host "Step 3.14159: Checking Python dependencies..." -ForegroundColor Yellow
    $pythonExe = Install-PythonDependencies -pythonServerPath $paths.pythonServerPath -pythonPath $config.PythonExeLocation
    if (-not $pythonExe) {
        Write-Host "    - Python setup incomplete. Exiting." -ForegroundColor Red
        foreach ($var in $varsToClear) {
            Remove-Variable -Name $var -ErrorAction SilentlyContinue
        }
        exit 1
    }
} else {
    Write-Host "Step 3.14159: Markdown Edition selected. Skipping Python check." -ForegroundColor Green
}

Write-Host ""
 
# Check if trying to install .docm to startup
if ($extension -eq ".docm" -and $paths.outputFolder -eq $paths.wordStartupPath) {
    if (-not $Force) {
        Write-Host "    `nWARNING: You are trying to install a .docm file to Word's STARTUP folder." -ForegroundColor DarkRed
        Write-Host "             Consider using .dotm for startup templates instead." -ForegroundColor DarkRed
        $choice = Read-Host "    Continue anyway? (y/n)"
        if ($choice -ne 'y') {
            Write-Host "    Installation cancelled." -ForegroundColor Red
            exit 1
        }
    }
    Write-Host "    Continuing with .docm installation..." -ForegroundColor Yellow
}

# Step 4: Check for existing files
Write-Host ""
Write-Host "Step 4: Checking for existing files..." -ForegroundColor Yellow

# Check startup path first (always need to handle this)
$startupFilePath = Join-Path $paths.wordStartupPath $paths.fullFileName

if (Test-CheckFileExists -filePath $startupFilePath) {
    Write-Host "    + IMPORTANT: Removing exisiting startup template to prevent runtime conflicts." -ForegroundColor Cyan
    if (-not (Remove-FileIfExists -FilePath $startupFilePath -Force:$Force)) {
        exit 1
    }
} else {
    Write-Host "    + No existing template found in startup" -ForegroundColor Green
}

# Check output path if different
if ($paths.outputFolder -ne $paths.wordStartupPath) {
    Write-Host ""
    Write-Host "    + Checking for existing file '$($paths.fullFileName)' in output path..." 
    if (Test-CheckFileExists -filePath $paths.fullFilePath) {
        Write-Host "    + File already exists in: $($paths.fullFilePath)" -ForegroundColor Yellow
        if (-not (Remove-FileIfExists -FilePath $paths.fullFilePath -Force:$Force)) {
            exit 1
        }
    } else {
        Write-Host "    + No existing file found in custom path" -ForegroundColor Green
    }
}

# Step 5: Create Word instance and document
Write-Host ""
Write-Host "Step 5: Creating Word instance and document..." -ForegroundColor Yellow

$word = New-WordApplication -Mode "COM"
if (-not $word) { Invoke-ScriptCleanup -WordApp $word -Document $doc; exit 1 }

$doc = New-Document -WordApp $word
if (-not $doc) { Invoke-ScriptCleanup -WordApp $word -Document $doc; exit 1 }

# Step 5.1: Check if we can programitically access VBA projects
if (-not (Test-VBAObjectModelTrust -doc $doc)) { Invoke-ScriptCleanup -WordApp $word -Document $doc; exit 1 }

$project = $doc.VBProject
$project.Name = $fileName+"_Project"
Write-Host "    + Project Name: '$($project.Name)'" -ForegroundColor Green

Write-Host "    + Word and document ready" -ForegroundColor Cyan

# Step 6: Save document first
Write-Host ""
Write-Host "Step 6: Saving document..." -ForegroundColor Yellow

if (-not (Save-Document -Document $doc -FilePath $paths.fullFilePath -WordPID $word.ProcessID)) {
    Write-Host "    - Failed to save. Cleaning up..." -ForegroundColor Red
    Invoke-ScriptCleanup -WordApp $word -Document $doc 
    exit 1
}

# Step 7: Import VBA components
Write-Host ""
Write-Host "Step 7: Importing VBA components for [$Edition]..." -ForegroundColor Yellow

# Filter VBA modules based on the selected Edition
$vbaModulesToImport = Get-VBAListForEdition -VBAFolder $paths.VBAFolder -Edition $Edition
$importResult = Import-VBAComponents -project $project -VBAList $vbaModulesToImport -pythonExe $pythonExe -pythonServerPath $paths.pythonServerPath -UpdatePythonPaths:$UpdatePythonPaths

if ($importResult) {
    Save-Document -Document $doc -FilePath $paths.fullFilePath -WordPID $word.ProcessID | Out-Null
} else {
    Write-Host "    - Failed to import VBA components." -ForegroundColor Red
    Invoke-ScriptCleanup -WordApp $word -Document $doc
    exit 1 
}

# Step 8: Compile and verify
Write-Host "`nStep 8: Compiling VBA project..." -ForegroundColor Yellow

# Make Word Visible
$word.Visible = $true

# use -Verbose for more details
$compileResult = Get-VBACompileError -WordApp $word -Project $project

if ($compileResult.Success) {
    Write-Host "    + VBA Project compiled successfully!" -ForegroundColor Green
} else {
    $hasCompileError = $true
    Write-Host "    - VBA Project compilation FAILED!" -ForegroundColor Red
    Write-Host "    - Reason:" -ForegroundColor Magenta
    Write-Host "        $($compileResult.Error)" -ForegroundColor Red
    
    # Only print module/line info if we have it
    if ($compileResult.Module) {
        Write-Host "        Module: $($compileResult.FullModuleName)" -ForegroundColor Red
    }
    if ($compileResult.ErrorLine) {
        Write-Host "        Line: $($compileResult.ErrorLine)" -ForegroundColor Red
        Write-Host "        Line Content: $($compileResult.ErrorLineContent)" -ForegroundColor Red
    }
}

try {
    $word.Visible = $false        
} catch {}


# Step 9: Clean up Word and document
Write-Host "`nStep 9: Closing up Document and Word to add XML Ribbon..." -ForegroundColor Yellow
Invoke-ScriptCleanup -WordApp $word -Document $doc

# Step 10: Add ribbon to file
Write-Host ""
Write-Host "Step 10: Adding [$Edition] ribbon to '$($paths.fullFileName)'..." -ForegroundColor Yellow

if (-not (Add-RibbonToTemplate -FilePath $paths.fullFilePath -RibbonXmlPath $paths.ribbonXml)) {
    Write-Host "    - Failed to add ribbon. '$($paths.fullFileName)'" -ForegroundColor Yellow
    Write-Host "    - Still, the macros can be called manually." -ForegroundColor Yellow
}

# Step 11: Final Status Output
if ($hasCompileError) {
    Write-Host "`n--------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[$Edition Edition] Installed with Compilation Warnings!" -ForegroundColor Yellow
    Write-Host "File saved to: $($paths.fullFilePath)" -ForegroundColor Yellow
    Write-Host "ACTION REQUIRED: Rectify the VBA compile error above and run" -ForegroundColor Red
    Write-Host "                'docLiveUpdateMacros.ps1' to update and re-compile." -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------`n" -ForegroundColor Yellow
} else {
    Write-Host "`n[$Edition Edition] Installation complete!" -ForegroundColor Cyan
    Write-Host "File saved to: $($paths.fullFilePath)`n" -ForegroundColor Green
}


foreach ($var in $varsToClear) {
    Remove-Variable -Name $var -ErrorAction SilentlyContinue
}

exit 0