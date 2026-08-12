<#
.SYNOPSIS
    Builds all Wordbot template editions (Markdown, LLM, and Research)

.DESCRIPTION
    This script automates the building of all three Wordbot template editions.
    It calls installWordBotTemplate.ps1 for each edition with the specified parameters.
    The script supports selective building, Python path updates, personal information removal, force mode, error handling, and logging.

    When run without any parameters (e.g., F5 in VSCode), it defaults to:
    - Force mode enabled (overwrites existing files without prompting)
    - RemovePersonalInfo enabled (strips metadata/personal info on save)
    - ContinueOnError enabled (continues building remaining editions even if one fails)
    - Silent mode enabled (minimal console output)
    - All three editions are built

.PARAMETER Editions
    Specifies which editions to build. Valid values: "Markdown", "LLM", "Research"
    Default: All three editions ("Markdown", "LLM", "Research")

.PARAMETER UpdatePythonPaths
    If specified, updates Python executable and server paths in the VBA code.
    If omitted, keeps existing Python paths (faster for routine builds).

.PARAMETER RemovePersonalInfo
    Strips personal metadata (author details, document properties) from the built templates.
    Default: $true when running without parameters, otherwise $false

.PARAMETER Force
    If specified, forces overwrite of existing files without prompting.
    Default: $true when running without parameters, otherwise $false

.PARAMETER ContinueOnError
    If specified, continues building remaining editions even if one fails.
    If omitted, stops at the first error.
    Default: $true when running without parameters, otherwise $false

.PARAMETER Silent
    If specified, suppresses detailed console output. Only progress messages and errors are shown.
    If omitted, shows all detailed output.
    Default: $true when running without parameters, otherwise $false

.PARAMETER LogPath
    Path to save a build log file. If provided, all console output is captured to this file.
    The log file is overwritten if it already exists.
    Default: build_log.txt in the script directory (but only created when explicitly used)

.EXAMPLE
    .\BuildAllEditions.ps1
    Builds all three editions with default debug settings: Force, RemovePersonalInfo, ContinueOnError, and Silent are enabled.

.EXAMPLE
    .\BuildAllEditions.ps1 -UpdatePythonPaths
    Builds all three editions, updates Python paths, with debug defaults (Force, RemovePersonalInfo, ContinueOnError, Silent).

.EXAMPLE
    .\BuildAllEditions.ps1 -RemovePersonalInfo:$false
    Builds all editions but retains personal information/metadata in the generated documents.

.EXAMPLE
    .\BuildAllEditions.ps1 -Editions "LLM", "Research"
    Builds only LLM and Research editions with debug defaults.

.EXAMPLE
    .\BuildAllEditions.ps1 -Editions "Markdown" -Silent:$false
    Builds only Markdown edition with full console output (no silent mode).

.EXAMPLE
    .\BuildAllEditions.ps1 -UpdatePythonPaths -Force:$false
    Builds all editions with Python updates, but prompts before overwriting existing files.

.EXAMPLE
    .\BuildAllEditions.ps1 -ContinueOnError:$false
    Builds all editions but stops at the first error.

.EXAMPLE
    .\BuildAllEditions.ps1 -LogPath "build.log"
    Builds all editions with debug defaults and saves everything to build.log.

.EXAMPLE
    .\BuildAllEditions.ps1 -Silent:$false -LogPath "build.log"
    Builds all editions with full console output and everything logged to build.log.

.EXAMPLE
    .\BuildAllEditions.ps1 -Editions "Markdown","LLM" -UpdatePythonPaths -RemovePersonalInfo -Force -ContinueOnError -Silent -LogPath "build_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Full automation: builds Markdown and LLM with Python updates, strips personal info, forces overwrite, continues on errors, silent console, timestamped log.

.LINK
    .\installWordBotTemplate.ps1
    https://github.com/Addy-ad/wordbot

.NOTES
    To view this help: Get-Help .\BuildAllEditions.ps1 -Full

    Author: Addy-ad
    Requires: PowerShell 5.1 or later, Windows, Microsoft Word installed

    Default behavior when run without parameters (F5 in VSCode):
    - Force: enabled
    - RemovePersonalInfo: enabled
    - ContinueOnError: enabled
    - Silent: enabled
    - All three editions are built

    The log file captures ALL console output (including colors as ANSI codes).
    To view the log file with colors, use: Get-Content build.log
#>

param(
    [ValidateSet("Markdown", "LLM", "Research")]
    [string[]]$Editions = @("Markdown", "LLM", "Research"),
    
    [switch]$UpdatePythonPaths,
    [switch]$RemovePersonalInfo,
    [switch]$Force,
    [switch]$ContinueOnError,
    [switch]$Silent,
    [string]$LogPath = (Join-Path $PSScriptRoot "build_log.txt")
)

# Default when running as script when running in VScode. 
if (-not $PSBoundParameters.ContainsKey('Force')) { $Force = $true }
if (-not $PSBoundParameters.ContainsKey('RemovePersonalInfo')) { $RemovePersonalInfo = $true }
if (-not $PSBoundParameters.ContainsKey('ContinueOnError')) { $ContinueOnError = $true }
if ($PSBoundParameters.Count -eq 0) { Clear-Host }

# Start transcript for logging if LogPath is provided (overwrite existing)
if ($LogPath) {
    if (Test-Path $LogPath) { Remove-Item $LogPath -Force }
    Start-Transcript -Path $LogPath | Out-Null
}

function Write-Log {
    param(
        $Message, 
        $Color = "White",
        [switch]$NoNewline,
        [switch]$AlwaysShow
    )
    
    # Write to console - skip if Silent mode and not AlwaysShow
    if (-not $Silent -or $AlwaysShow) {
        Write-Host $Message -ForegroundColor $Color -NoNewline:$NoNewline
    }
}

# Resolve script path relative to this file's location
$installerScript = Join-Path $PSScriptRoot "installWordBotTemplate.ps1"

if (-not (Test-Path $installerScript)) {
    Write-Log "Error: Could not find installer script at: $installerScript" "Red" -AlwaysShow
    if ($LogPath) { Stop-Transcript }
    exit 1
}

# Build header - always show these
Write-Log "Wordbot Template Builder" "Cyan" -AlwaysShow
Write-Log "========================" "Cyan" -AlwaysShow
Write-Log "Start time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "Gray" -AlwaysShow
Write-Log "" -AlwaysShow
Write-Log "Build Configuration:" "Yellow" -AlwaysShow
Write-Log "  Editions: $($Editions -join ', ')" "Gray" -AlwaysShow
Write-Log "  Update Python Paths: $UpdatePythonPaths" "Gray" -AlwaysShow
Write-Log "  Remove Personal Info: $RemovePersonalInfo" "Gray" -AlwaysShow
Write-Log "  Force mode: $Force" "Gray" -AlwaysShow
Write-Log "  Continue on error: $ContinueOnError" "Gray" -AlwaysShow
Write-Log "  Silent mode: $Silent" "Gray" -AlwaysShow
Write-Log "  Log file: $LogPath" "Gray" -AlwaysShow
Write-Log "" -AlwaysShow

# Build parameters
$failedEditions = @()
$successCount = 0
$startTime = Get-Date

foreach ($edition in $Editions) {
    try {
        # Execute the installer script - let it write directly to console (preserves colors)
        # Silent mode: suppress installer output
        # Non-silent: show installer output
        if ($Silent) {
            # Silent: suppress all installer output
            & $installerScript -Edition $edition `
                              -UpdatePythonPaths:$UpdatePythonPaths `
                              -RemovePersonalInfo:$RemovePersonalInfo `
                              -Force:$Force *>$null
        } else {
            # Non-silent: show everything with colors
            & $installerScript -Edition $edition `
                              -UpdatePythonPaths:$UpdatePythonPaths `
                              -RemovePersonalInfo:$RemovePersonalInfo `
                              -Force:$Force
        }
        
        $exitCode = $LASTEXITCODE
        
        if ($exitCode -ne 0) {
            Write-Log "Error building $edition edition (Exit Code: $exitCode)!" "Red" -AlwaysShow
            $failedEditions += $edition
            
            if (-not $ContinueOnError) {
                Write-Log "Exiting due to error (use -ContinueOnError to skip failures)..." "Red" -AlwaysShow
                if ($LogPath) { Stop-Transcript }
                exit $exitCode
            }
        } else {
            $successCount++
            Write-Log "Successfully built $edition edition" "Green" -AlwaysShow
        }
    } catch {
        Write-Log "Exception building $edition edition: $_" "Red" -AlwaysShow
        $failedEditions += $edition
        
        if (-not $ContinueOnError) {
            if ($LogPath) { Stop-Transcript }
            exit 1
        }
    }
    
    Write-Log "" -AlwaysShow
}

# Summary
$endTime = Get-Date
$duration = $endTime - $startTime

Write-Log "---------------------------------------------------" "Gray" -AlwaysShow
Write-Log "Build Summary:" "Cyan" -AlwaysShow
Write-Log "==============" "Cyan" -AlwaysShow
Write-Log "Start time:  $($startTime.ToString('HH:mm:ss'))" "Gray" -AlwaysShow
Write-Log "End time:    $($endTime.ToString('HH:mm:ss'))" "Gray" -AlwaysShow
Write-Log "Duration:    $($duration.ToString('hh\:mm\:ss'))" "Gray" -AlwaysShow
Write-Log "Success:     $successCount / $($Editions.Count) editions" "Green" -AlwaysShow

if ($failedEditions.Count -gt 0) {
    Write-Log "Failed:      $($failedEditions -join ', ')" "Red" -AlwaysShow
    Write-Log "Overall:     BUILD FAILED" "Red" -AlwaysShow
} else {
    Write-Log "Overall:     ALL BUILDS SUCCESSFUL" "Green" -AlwaysShow
}

if ($LogPath) {
    Write-Log "" -AlwaysShow
    Write-Log "Log saved to: $LogPath" "Cyan" -AlwaysShow
}

# Stop transcript
if ($LogPath) {
    Stop-Transcript
}

# Return appropriate exit code
if ($failedEditions.Count -gt 0) {
    exit 1
}

exit 0