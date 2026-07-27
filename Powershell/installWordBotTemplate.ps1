# installWordBotTemplate.ps1
# Wordbot Template Installer

Clear-Host

Add-Type -AssemblyName WindowsBase

# Import the module
$modulePath = Join-Path $PSScriptRoot "Module\Wordbot.psm1"
Import-Module $modulePath -Force

# Load configuration dynamically
$config = Get-WordbotConfig

# Definitions
$fileName = $config.FileName
$extension = $config.Extension
$outputPath = $config.outputPath

Write-Host "Wordbot Template Installer" -ForegroundColor Cyan
Write-Host "--------------------------" -ForegroundColor Cyan
Write-Host ""

# Test if Word is installed
Write-Host "Step 1: Checking Word installation..." -ForegroundColor Yellow
if (-not (Test-WordInstalled)) {
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

$paths = Get-Paths $fileName $extension $outputPath $config

Write-Host "    Script path:        $($paths.scriptPath)"
Write-Host "    File name:          $($paths.fullFileName)"
Write-Host "    Output Folder:      $($paths.outputFolder)"
Write-Host "    FullFile path:      $($paths.fullFilePath)"
Write-Host "    Word startup path:  $($paths.wordStartupPath)"
Write-Host "    VBA folder:         $($paths.VBAFolder)"
Write-Host "    Ribbon XML:         $($paths.ribbonXml)"
Write-Host "    Python server path: $($paths.pythonServerPath)"
Write-Host ""

Write-Host "Step 3.14159: Checking Python and dependencies for Wordbot server..." -ForegroundColor Yellow

$pythonExe = Install-PythonDependencies -pythonServerPath $paths.pythonServerPath -pythonPath $config.PythonExeLocation

if (-not $pythonExe) {
    Write-Host "    - Python setup incomplete or cancelled. Exiting installer." -ForegroundColor Red
    exit 1
}
Write-Host "    + Python executable resolved: $pythonExe" -ForegroundColor Cyan

Write-Host ""

# Check if trying to install .docm to startup
if ($extension -eq ".docm" -and $paths.outputFolder -eq $paths.wordStartupPath) {
    Write-Host ""
    Write-Host "    WARNING: You are trying to install a .docm file to Word's STARTUP folder." -ForegroundColor DarkRed
    Write-Host "             Consider using .dotm for startup templates instead." -ForegroundColor DarkRed
    $choice = Read-Host "    Continue anyway? (y/n)"
    if ($choice -ne 'y') {
        Write-Host "    Installation cancelled." -ForegroundColor Red
        exit 1
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
    if (-not (Remove-FileIfExists -FilePath $startupFilePath -Ask)) {
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
        if (-not (Remove-FileIfExists -FilePath $paths.fullFilePath -Ask)) {
            exit 1
        }
    } else {
        Write-Host "    + No existing file found in custom path" -ForegroundColor Green
    }
}

# Step 5: Create Word instance and document
Write-Host ""
Write-Host "Step 5: Creating Word instance and document..." -ForegroundColor Yellow

$word = New-WordApplication
if (-not $word) {
    Write-Host "    - Failed to create Word instance. Exiting." -ForegroundColor Red
    exit 1
}

$doc = New-Document -WordApp $word
if (-not $doc) {
    Write-Host "    - Failed to create document. Cleaning up..." -ForegroundColor Red
    Close-WordApplication -WordApp $word | Out-Null
    exit 1
}

# Step 5.1: Check if we can programitically access VBA projects
$trustEnabled = Test-VBAObjectModelTrust -doc $doc
if (-not $trustEnabled) {
    Close-Document -Document $doc | Out-Null
    Close-WordApplication -WordApp $word | Out-Null
    exit 1
}

$project = $doc.VBProject
$project.Name = $fileName+"Project"
Write-Host "    + Project Name: '$($project.Name)'" -ForegroundColor Green

Write-Host "    + Word and document ready" -ForegroundColor Cyan

# Step 6: Save document first
Write-Host ""
Write-Host "Step 6: Saving document..." -ForegroundColor Yellow

if (-not (Save-Document -Document $doc -FilePath $paths.fullFilePath)) {
    Write-Host "    - Failed to save. Cleaning up..." -ForegroundColor Red
    Close-Document -Document $doc | Out-Null
    Close-WordApplication -WordApp $word | Out-Null
    exit 1
}

# Step 7: Import VBA components
Write-Host ""
Write-Host "Step 7: Importing VBA components..." -ForegroundColor Yellow

if (Import-VBAComponents -project $project -VBAFolder $paths.VBAFolder -pythonExe $pythonExe) {
    Save-Document -Document $doc -FilePath $paths.fullFilePath | Out-Null
} else {
    Write-Host "    - Failed to import VBA components. Cleaning up..." -ForegroundColor Red
    Close-Document -Document $doc | Out-Null
    Close-WordApplication -WordApp $word | Out-Null
    exit 1
}

# Step 9: Clean up Word and document
Write-Host ""
Write-Host "Step 9: Closing up Document and Word to add XML Ribbon..." -ForegroundColor Yellow

if ($doc) {
    Close-Document -Document $doc | Out-Null
}

if ($word) {
    Close-WordApplication -WordApp $word | Out-Null
}

# Step 10: Add ribbon to file
Write-Host ""
Write-Host "Step 10: Adding ribbon to '$($paths.fullFileName)'..." -ForegroundColor Yellow

if (-not (Add-RibbonToTemplate -FilePath $paths.fullFilePath -RibbonXmlPath $paths.ribbonXml)) {
    Write-Host "    - Failed to add ribbon. '$($paths.fullFileName)' may still work without it." -ForegroundColor Yellow
    # Don't exit, ribbon is optional
}

Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Cyan
Write-Host "File saved to: $($paths.fullFilePath)" -ForegroundColor Green
Write-Host ""
Write-Host ""

$varsToClear = @(
    "config",
    "fileName",
    "extension",
    "outputPath",
    "paths",
    "pythonExe",
    "choice",
    "confirm",
    "startupFilePath",
    "word",
    "doc",
    "project",
    "modulePath"
)

foreach ($var in $varsToClear) {
    Remove-Variable -Name $var -ErrorAction SilentlyContinue
}