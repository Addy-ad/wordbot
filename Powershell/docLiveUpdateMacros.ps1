Clear-Host

# List all variables to clear at script exit
$varsToClear = @(
    "config",
    "templateName",
    "file",
    "fileName",
    "extension",
    "outputPath",
    "WordApp",
    "WordAppInfo",
    "result",
    "targetProject",
    "paths",
    "pythonExe"
)

foreach ($var in $varsToClear) {
    Remove-Variable -Name $var -ErrorAction SilentlyContinue
}

Import-Module "$PSScriptRoot\Module\Wordbot.psm1" -Force

# Load configuration dynamically
$config = Get-WordbotConfig

# # Three main variables to work with
$templateName = $config.TemplateName    # Name of template (if .dotm provided)
$file         = $config.File            # Full path to file user provided (if any)
  
# Get active word Instance using the function
# Try to get active Word instance
$WordApp = Get-ActiveWordInstance

if ($WordApp) {
    # Display AND capture objects in the open word instance
    $WordAppInfo = Get-WordInfo -WordApp $WordApp -ReturnObjects
    $WordAppInfo.Visible = $true
    $result = Get-TargetProject -WordApp $WordApp -WordAppInfo $WordAppInfo -TemplateName $templateName -File $file
    # Extract the returned objects
} else {
    $WordApp = New-WordApplication -Mode Process
    if (-not $WordApp) {
        Write-Host "    - Failed to create Word instance. Exiting." -ForegroundColor Red
        exit 1
    }

    # Get info for the new Word instance
    $WordAppInfo = Get-WordInfo -WordApp $WordApp -ReturnObjects
    
    # Create a new blank document
    Write-Host "    Creating new blank document..." -ForegroundColor Yellow
    $result = New-BlankDocument -WordApp $WordApp -WordAppInfo $WordAppInfo

    $WordApp = $result.WordApp
    $WordAppInfo = $result.WordAppInfo
    
    # Make Word visible
    $WordApp.Visible = $true
    if ($WordApp.Visible) {
        $WordApp.WindowState = 2 # Minimize
        $WordApp.WindowState = 1 # Maximize
        $WordApp.Activate()
    }

    $result = Get-TargetProject -WordApp $WordApp -WordAppInfo $WordAppInfo -TemplateName $templateName -File $file
}

$WordApp = $result.WordApp
$WordAppInfo = $result.WordAppInfo
$targetProject = $result.TargetProject

if ($targetProject.Protection -eq 1) {
    Write-Host "VBA Project is PROTECTED!" -ForegroundColor Red
    # $targetProject.Protection = 0
    # Write-Host "You need to unprotect it before importing components." -ForegroundColor Yellow
} else {
    Write-Host "VBA Project is ready to work" -ForegroundColor Green
}

# Load configuration dynamically
$config = Get-WordbotConfig

# Definitions
$fileName = $config.FileName
$extension = $config.Extension
$outputPath = $config.outputPath

$paths = Get-Paths $fileName $extension $outputPath $config

$pythonExe = Install-PythonDependencies -pythonServerPath $paths.pythonServerPath -pythonPath $config.PythonExeLocation

# Write-Host "Adding/overwriting VBA components..." -ForegroundColor Yellow
Import-VBAComponents -project $targetProject -VBAFolder $paths.VBAFolder  -pythonExe $pythonExe

# Compile and get error
Write-Host "Compiling VBA project..." -ForegroundColor Yellow

$compileResult = Get-VBACompileError -WordApp $WordApp

if ($compileResult.Success) {
    Write-Host "VBA Project compiled successfully!" -ForegroundColor Green
} else {
    Write-Host "VBA Project compilation FAILED!" -ForegroundColor Red
    Write-Host "Error: $($compileResult.Error)" -ForegroundColor Red
    
    # Only print module/line info if we have it
    if ($compileResult.Module) {
        Write-Host "Module: $($compileResult.FullModuleName)" -ForegroundColor Yellow
    }
    if ($compileResult.ErrorLine) {
        Write-Host "Line: $($compileResult.ErrorLine)" -ForegroundColor Yellow
        Write-Host "Line Content: $($compileResult.ErrorLineContent)" -ForegroundColor Yellow
    }
}

Write-Host ""

