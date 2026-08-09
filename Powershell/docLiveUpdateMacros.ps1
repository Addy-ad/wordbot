<#
.SYNOPSIS
    Updates VBA macros in a live Word document without closing Word

.DESCRIPTION
    This script updates Wordbot macros in an already open Word document or template.
    Unlike the installer script, this does not close Word - it injects the updated
    VBA code directly into the active document or template.

    This is useful for:
    - Rapid development and testing of VBA changes
    - Hot-fixing macros without restarting Word
    - Iterative development where you want to keep Word open

    The script supports three editions:
    - Markdown: Core Markdown formatting and conversion features
    - LLM: Markdown features + Large Language Model integration
    - Research: LLM features + Research and citation tools

    Key Features:
    - Attaches to an existing Word instance or creates a new one
    - Finds or creates the target document based on project name
    - Updates VBA components in the live document
    - Compiles VBA to check for errors
    - Does NOT close Word after update (live update)

    IMPORTANT NOTES ABOUT RIBBON:
    - This script does NOT add or update ribbon XML files
    - To use the updated macros, you must either:
      1. Use an existing .dotm or .docm file that already has the ribbon XML embedded
      2. Keep a document with the ribbon XML open in Word
      3. Run the macros manually from the VBA editor or assign them to shortcuts
    - If you need to add the ribbon, use installWordBotTemplate.ps1 instead

.PARAMETER Edition
    Specifies which edition to update. Valid values: "Markdown", "LLM", "Research"
    Default: "Markdown"

.PARAMETER UpdatePythonPaths
    Updates Python executable and server paths in the VBA code.
    When enabled, replaces placeholders {{PYTHON_EXE_PATH}} and {{PYTHON_SERVER_PATH}}
    in aWordbotRibbonLLMFunctions.bas with actual paths.
    Default: $true (enabled) when running the script directly

.PARAMETER Force
    Force mode for live updates.
    Currently reserved for future use (e.g., forcing document creation without prompts).
    Default: $false

.EXAMPLE
    .\docLiveUpdateMacros.ps1 -Edition Markdown
    Updates the Markdown edition in the active Word document. Creates a new Word
    instance if none is running. Python paths are updated (default behavior).

.EXAMPLE
    .\docLiveUpdateMacros.ps1 -Edition LLM -Force
    Updates the LLM edition in the active Word document with force mode enabled.
    Python paths are updated.

.EXAMPLE
    .\docLiveUpdateMacros.ps1 -Edition Research -UpdatePythonPaths:$false
    Updates the Research edition without updating Python paths. Useful when Python
    environment hasn't changed and you want faster updates.

.EXAMPLE
    .\docLiveUpdateMacros.ps1 -Edition Markdown -UpdatePythonPaths:$false
    Updates the Markdown edition without updating Python paths.
    Note: UpdatePythonPaths has no effect on Markdown edition (no Python dependencies).

.EXAMPLE
    .\docLiveUpdateMacros.ps1
    Updates the Markdown edition with default settings. Creates a new Word instance
    if needed. Python paths are updated (default behavior).

.LINK
    .\installWordBotTemplate.ps1
    .\BuildAllEditions.ps1
    https://github.com/Addy-ad/wordbot

.NOTES
    Author: Addy-ad
    Version: 2.0
    Last Modified: 2026-08-07
    Requires: PowerShell 5.1 or later, Windows, Microsoft Word installed

    How It Works:
    1. Checks for an active Word instance or creates a new one
    2. Finds the target document by project name or file path
    3. Creates a new document if none exists
    4. Updates VBA components in the live document
    5. Compiles VBA to verify changes
    6. Leaves Word open for continued work

    Differences from installWordBotTemplate.ps1:
    - Does NOT close Word after update
    - Attaches to existing Word instance when possible
    - Faster for development iteration
    - Does NOT add ribbon XML (assumes it already exists)

    Ribbon XML Requirements:
    Since this script does not add ribbon XML, you need one of these:
    - A .dotm or .docm with ribbon XML already embedded (from installWordBotTemplate.ps1)
    - An open document that already has the ribbon loaded
    - Or you can run macros manually via Alt+F8 in Word

    Safety Notes:
    - This script requires "Trust access to the VBA project object model" to be enabled
    - Works with unsaved documents - they will be updated in memory
    - Save your document after update to persist changes
    - VBA compilation errors will be displayed but Word remains open
#>

# docLiveUpdateMacros.ps1
# docLiveUpdateMacros.ps1
param (
    [ValidateSet("Markdown", "LLM", "Research")]
    [string]$Edition,
    [switch]$UpdatePythonPaths,
    [switch]$Force
)

# Clear local variables before exiting script scope
$varsToClear = @(
    "config", "fileName", "extension", "outputPath", "paths",
    "pythonExe", "choice", "confirm", "startupFilePath","hasCompileError",
    "word", "doc", "project", "modulePath", "vbaModulesToImport", 
    "compileResult", "fullFilePath", "fileExists", "importResult"
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

    exit 1
}

# Default when running as script when running in VScode. 
if (-not $PSBoundParameters.ContainsKey('Edition')) { $Edition = "Markdown" }
if (-not $PSBoundParameters.ContainsKey('Force')) { $Force = $false }
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

# Identify target template for active Word detection
$expectedProjectName = $fileName + "_Project"

# Resolve paths early to get full file path
$paths = Get-Paths -fileName $fileName -extension $extension -outputPath $outputPath -Config $config -Edition $Edition
$fullFilePath = $paths.fullFilePath

Write-Host "Wordbot Template Live Updater - [$Edition Edition]" -ForegroundColor Cyan
Write-Host "--------------------------------------------------" -ForegroundColor Cyan
Write-Host ""

# Step 1: Check active Word instance or attach
Write-Host "Step 1: Checking active Word instance..." -ForegroundColor Yellow
$word = Get-ActiveWordInstance

if (-not $word) {
    Write-Host "    - No active Word instance found. Creating new instance..." -ForegroundColor Yellow
    $word = New-WordApplication -Mode "COM"
    if (-not $word) { 
        foreach ($var in $varsToClear) {
            Remove-Variable -Name $var -ErrorAction SilentlyContinue
        }
        exit 1 
    }
} else {
    Write-Host "    + Using existing Word instance" -ForegroundColor Green
}

# Resolve the target live document cleanly
$targetResult = Get-OpenTargetDocument -WordApp $word -FullFilePath $fullFilePath -ExpectedProjectName $expectedProjectName -FileName $fileName -Force:$Force
if ($targetResult.ClearVars -eq 1) {
    foreach ($var in $varsToClear) {
        Remove-Variable -Name $var -ErrorAction SilentlyContinue
    }
    exit 1
}
$doc = $targetResult.Document
$project = $doc.VBProject

# Only flash/maximize window state if a brand new document was spawned
if ($targetResult.IsNewDocument) {
    $word.Visible = $true
    if ($word.ActiveWindow) {
        $word.ActiveWindow.WindowState = 2 # Minimize
        $word.ActiveWindow.WindowState = 1 # Maximize
    }
    $word.Activate()
}

# Step 2: Check VBA Protection
Write-Host ""
Write-Host "Step 2: Checking VBA project protection..." -ForegroundColor Yellow
if ($project.Protection -eq 1) {
    Write-Host "    - VBA Project is PROTECTED!" -ForegroundColor Red
    Write-Host "    - Cannot update macros in a protected project." -ForegroundColor Yellow
    Invoke-ScriptCleanup -WordApp $word -Document $doc
    exit 1
} else {
    Write-Host "    + Project selected: '$($project.Name)'" -ForegroundColor Green
    Write-Host "    + VBA Project is ready for live update [$Edition Edition]" -ForegroundColor Green
}

Write-Host ""

# Step 3: Python Check (Only required for LLM and Research editions)
$pythonExe = $null
if ($Edition -ne "Markdown") {
    Write-Host "Step 3: Checking Python dependencies..." -ForegroundColor Yellow
    $pythonExe = Install-PythonDependencies -pythonServerPath $paths.pythonServerPath -pythonPath $config.PythonExeLocation
    if (-not $pythonExe) {
        Write-Host "    - Python setup incomplete. Exiting..." -ForegroundColor Red
        Invoke-ScriptCleanup -WordApp $word -Document $doc
        exit 1
    }
} else {
    Write-Host "Step 3: Markdown Edition selected. Skipping Python check." -ForegroundColor Green
}

Write-Host ""

# Step 4: Import VBA components
Write-Host "Step 4: Importing VBA components for [$Edition Edition]..." -ForegroundColor Yellow
$vbaModulesToImport = Get-VBAListForEdition -VBAFolder $paths.VBAFolder -Edition $Edition
$importResult = Import-VBAComponents -project $project -VBAList $vbaModulesToImport -pythonExe $pythonExe -pythonServerPath $paths.pythonServerPath -UpdatePythonPaths:$UpdatePythonPaths

if (-not $importResult) {
    Write-Host "    - Failed to import VBA components." -ForegroundColor Red
    Invoke-ScriptCleanup -WordApp $word -Document $doc
    exit 1
}

Write-Host "    + VBA components imported successfully!" -ForegroundColor Green
Write-Host ""

# Step 5: Compile and verify
Write-Host "Step 5: Compiling VBA project..." -ForegroundColor Yellow

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

if ($hasCompileError) {
    Write-Host "`n--------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "ACTION REQUIRED: Rectify the VBA compile error above and re-run this script to update" -ForegroundColor Red
    Write-Host "--------------------------------------------------------`n" -ForegroundColor Yellow
} else {
    Write-Host "`nLive update complete!" -ForegroundColor Cyan
    Write-Host "File updated: $fullFilePath`n" -ForegroundColor Green
}

foreach ($var in $varsToClear) {
    Remove-Variable -Name $var -ErrorAction SilentlyContinue
}

exit 0