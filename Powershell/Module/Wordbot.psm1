# Module/Wordbot.psm1

function New-PythonVirtualEnvironment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCmd) {
        Write-Host "[-] Base Python was not found in PATH to create virtual environment." -ForegroundColor Red
        return $null
    }

    $pipPath = Join-Path $TargetPath "Scripts\python.exe"

    if (-not (Test-Path $pipPath)) {
        Write-Host "[*] Initializing virtual environment at '$TargetPath'..." -ForegroundColor Cyan
        python -m venv "$TargetPath"
    }

    if (-not (Test-Path $pipPath)) {
        Write-Host "[-] Failed to create virtual environment executable at '$pipPath'." -ForegroundColor Red
        return $null
    }

    Write-Host "[+] Virtual environment ready at '$TargetPath'." -ForegroundColor Green
    return $pipPath
}

function Install-PipPackageList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonExe,

        [Parameter(Mandatory = $true)]
        [string[]]$Packages
    )

    $total = $Packages.Count
    $current = 0

    foreach ($pkg in $Packages) {
        $current++
        Write-Host "[*] ($current/$total) Installing '$pkg'..." -ForegroundColor Cyan -NoNewline
        
        $output = & "$PythonExe" -m pip install $pkg 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Host " Done." -ForegroundColor Green
        } else {
            Write-Host " Failed!" -ForegroundColor Red
            Write-Host "[-] Error installing $pkg`:`n$output" -ForegroundColor Red
            return $false
        }
    }

    return $true
}
function Install-PythonDependencies {
    [CmdletBinding()]
    param(
        [string]$pythonServerPath,
        [string]$pythonPath
    )

    # 1. Resolve Target Python Executable
    $targetPython = $null

    if ($pythonPath) {
        if (Test-Path -Path $pythonPath -PathType Container) {
            $candidate = Join-Path $pythonPath "Scripts\python.exe"
            if (Test-Path $candidate) {
                $targetPython = $candidate
            } else {
                Write-Host "    Virtual environment is not initialized in '$pythonPath'." -ForegroundColor Yellow
                $confirm = Read-Host "Create a new virtual environment in this folder? (y/n)"
                if ($confirm.Trim().ToLower() -eq 'y') {
                    $targetPython = New-PythonVirtualEnvironment -TargetPath $pythonPath
                } else {
                    Write-Host "[*] Operation cancelled by user." -ForegroundColor Yellow
                    return $null
                }
            }
        } elseif (Test-Path -Path $pythonPath -PathType Leaf) {
            $targetPython = $pythonPath
        }

        if (-not $targetPython) {
            Write-Host "[-] Specified Python path not found or invalid: $pythonPath" -ForegroundColor Red
            return $null
        }
    } else {
        $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
        if (-not $pythonCmd) {
            Write-Host "[-] Python was not found in PATH." -ForegroundColor Red
            Write-Host "    Please install Python and make sure to check 'Add Python to PATH' during setup." -ForegroundColor Yellow
            return $null
        }
        $targetPython = $pythonCmd.Source
    }

    Write-Host "[+] Using Python executable: $targetPython" -ForegroundColor Green

    # 2. Locate requirements.txt
    if ($pythonServerPath) {
        $serverDir = Split-Path -Path $pythonServerPath -Parent
    } else {
        $serverDir = $PSScriptRoot
    }

    $reqPath = Join-Path $serverDir "requirements.txt"
    if (-not (Test-Path $reqPath)) {
        Write-Host "[-] requirements.txt not found in $serverDir" -ForegroundColor Red
        return $null
    }

    # 3. Read package requirements
    $requiredPackages = Get-Content $reqPath | ForEach-Object {
        $line = $_.Trim()
        if ($line -and -not $line.StartsWith("#")) {
            # Strip specifiers (~=, ==, >=, <=, etc.) and environment markers
            # Enhanced regex to handle package names with dots (e.g., google.cloud)
            if ($line -match '^([a-zA-Z0-9_.\-]+)') {
                $matches[1]
            }
        }
    }

    if (-not $requiredPackages) {
        Write-Host "[!] No valid packages found in requirements.txt" -ForegroundColor Yellow
        return $targetPython
    }

    # 4. Check installed packages in target environment
    $installedOutput = & "$targetPython" -m pip list --format=freeze 2>$null
    $installedPackages = $installedOutput | ForEach-Object {
        ($_.Trim() -split '==')[0].Trim().ToLower()
    }

    $missingPackages = @()
    foreach ($pkg in $requiredPackages) {
        if ($installedPackages -notcontains $pkg.ToLower()) {
            $missingPackages += $pkg
        }
    }

    # 5. Handle missing packages
    if ($missingPackages.Count -eq 0) {
        Write-Host "[+] All required packages are already installed." -ForegroundColor Green
        return $targetPython
    }

    Write-Host "`n[!] Missing packages detected:" -ForegroundColor Yellow
    foreach ($pkg in $missingPackages) {
        Write-Host "    - $pkg" -ForegroundColor Yellow
    }

    # IF $pythonPath was supplied, ask for consent before installing directly into $targetPython.
    if ($pythonPath) {
        Write-Host "`n[*] Proceed with installing missing dependencies into '$targetPython'?" -ForegroundColor Cyan
        $confirm = Read-Host "Install missing packages? (y/n)"

        if ($confirm.Trim().ToLower() -ne 'y') {
            Write-Host "[*] Installation cancelled by user." -ForegroundColor Yellow
            return $null
        }

        if (-not (Install-PipPackageList -PythonExe $targetPython -Packages $missingPackages)) {
            return $null
        }

        Write-Host "[+] All dependencies installed successfully." -ForegroundColor Green
        return $targetPython
    }

    # ONLY IF $pythonPath was EMPTY in config do we show the menu fallback
    $defaultPipLocation = & "$targetPython" -c "import site; print(site.getsitepackages()[0])" 2>$null

    Write-Host "`nSelect installation target:" -ForegroundColor Cyan
    Write-Host "  1) System environment ($defaultPipLocation)"
    Write-Host "  2) Custom virtual environment (Select directory via dialog)"
    Write-Host "  0) Cancel and exit"
    
    $choice = Read-Host "`nEnter option (0, 1, or 2)"

    if ($choice -eq "0") {
        Write-Host "[*] Installation cancelled by user." -ForegroundColor Yellow
        return $null
    } elseif ($choice -eq "2") {
        Add-Type -AssemblyName System.Windows.Forms
        $browser = New-Object System.Windows.Forms.FolderBrowserDialog
        $browser.Description = "Select target directory for virtual environment"
        $browser.SelectedPath = $serverDir
        $browser.ShowNewFolderButton = $true

        $result = $browser.ShowDialog()

        if ($result -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($browser.SelectedPath)) {
            Write-Host "[*] No folder selected. Operation cancelled." -ForegroundColor Yellow
            return $null
        }

        $venvPath = $browser.SelectedPath
        
        Write-Host "`n[*] Proceed with creating virtual environment and installing packages in '$venvPath'?" -ForegroundColor Cyan
        $confirm = Read-Host "Proceed with setup? (y/n)"

        if ($confirm.Trim().ToLower() -ne 'y') {
            Write-Host "[*] Installation cancelled by user." -ForegroundColor Yellow
            return $null
        }

        $pipPath = New-PythonVirtualEnvironment -TargetPath $venvPath
        if (-not $pipPath) {
            return $null
        }

        if (-not (Install-PipPackageList -PythonExe $pipPath -Packages $missingPackages)) {
            return $null
        }

        Write-Host "[+] All dependencies installed successfully." -ForegroundColor Green
        return $pipPath
    } else {
        Write-Host "`n[*] Proceed with installing missing dependencies into system environment ($targetPython)?" -ForegroundColor Cyan
        $confirm = Read-Host "Install missing packages? (y/n)"

        if ($confirm.Trim().ToLower() -ne 'y') {
            Write-Host "[*] Installation cancelled by user." -ForegroundColor Yellow
            return $null
        }

        if (-not (Install-PipPackageList -PythonExe $targetPython -Packages $missingPackages)) {
            return $null
        }

        Write-Host "[+] All dependencies installed successfully." -ForegroundColor Green
        return $targetPython
    }
}

function Get-WordbotConfig {
    param([string]$ConfigPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "config.psd1"))
    
    if (Test-Path $ConfigPath) {
        try {
            return Import-PowerShellDataFile -Path $ConfigPath
        } catch {
            Write-Host "    - Failed to parse config file at $($ConfigPath): $_" -ForegroundColor Red
            return $null
        }
    } else {
        Write-Host "    - Config file not found at $($ConfigPath)" -ForegroundColor Red
        return $null
    }
    return @{}
}

function Test-VBAObjectModelTrust {
    param($doc)
    
    try {
        Write-Host "    Checking access to VBA project object model" -ForegroundColor Yellow
        
        $project = $doc.VBProject
        
        # # Just try to ACCESS a property that requires trust
        # $projectName = $project.Name  # Just read it
        
        # Try to add a temporary module (more reliable test)
        $module = $project.VBComponents.Add(1)  # Add standard module
        # $moduleName = $module.Name
        $project.VBComponents.Remove($module)  # Remove it immediately
        
        Write-Host "    VBA project object model access is ENABLED" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "    [X] VBA project object model access is DISABLED (Most Likely)" -ForegroundColor Red
        Write-Host "        Check in: Word → File → Options → Trust Center →" -ForegroundColor Yellow
        Write-Host "        Trust Center Settings → Macro Settings →" -ForegroundColor Yellow
        Write-Host "        Tick the 'Trust access to the VBA project object model'" -ForegroundColor Yellow
        return $false
    }
}

function Test-WordInstalled {
    $wordProgId = [Type]::GetTypeFromProgID("Word.Application")
    if ($wordProgId) {
        Write-Host "    + Word is installed and COM registration OK" -ForegroundColor Green
        return $true
    } else {
        Write-Host "    - Word is not installed or COM registration missing." -ForegroundColor Red
        return $false
    }
}

function Test-WordRunning {
    try {
        $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
        if ($wordProcesses) {
            Write-Host "    - Word is currently running" -ForegroundColor Yellow
            Write-Host "    - Running Word processes:"
            $wordProcesses | ForEach-Object { 
                Write-Host "        PID: $($_.Id) - $($_.ProcessName)"
            }
            Write-Host "    - Please close Word and run again" -ForegroundColor Red
            return $true
        } else {
            Write-Host "    + Word is not running" -ForegroundColor Green
            return $false
        }
    } catch {
        Write-Host "    - Could not check for Word processes" -ForegroundColor Yellow
        return $false
    }
}

function Get-Paths {
    param(
        $fileName = $null,
        $extension = $null,
        $outputPath = $null,
        $Config = $null,
        $Edition = $null,
        [switch]$wordStartupPathFlag
    )
    
    $scriptPath = Split-Path $PSScriptRoot -Parent
    if (-not $scriptPath) { $scriptPath = Get-Location }
    
    $wordStartupPath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) "Microsoft\Word\STARTUP"

    if ($wordStartupPathFlag) { return $wordStartupPath }

    # Fall back to Config if parameters are not explicitly passed
    if ($Config) {
        if (-not $fileName) { $fileName = $Config.FileName }
        if (-not $extension) { $extension = $Config.Extension }
        if (-not $outputPath -and $Config.outputFolder) { $outputPath = $Config.outputFolder }
        if (-not $Edition -and $Config.Edition) { $Edition = $Config.Edition }
    }

    # Default fallback values
    if (-not $fileName) { $fileName = "Wordbot" }
    if (-not $extension) { $extension = ".dotm" }
    if (-not $outputPath) { $outputPath = $wordStartupPath }
    if (-not $Edition) { $Edition = "Markdown" }
    
    $fullFileName = $fileName + $extension

    # FIX: Null-check for Config properties
    $vbaRel = if ($Config -and $Config.VBAFolderRelativePath) { $Config.VBAFolderRelativePath } else { "Resources\WordVBAmacros" }
    $pythonRel = if ($Config -and $Config.PythonServerLocation) { $Config.PythonServerLocation } else { "..\Python\main.py" }
    $ribbonFolderRel = if ($Config -and $Config.RibbonXmlRelativePath) { $Config.RibbonXmlRelativePath } else { "Resources\RibbonXML" }
    
    $ribbonFileName = "WordbotRibbon_$Edition.xml"
    $ribbonRel = Join-Path $ribbonFolderRel $ribbonFileName

    return @{
        scriptPath        = $scriptPath
        fullFileName      = $fullFileName
        outputFolder      = $outputPath
        fullFilePath      = [System.IO.Path]::GetFullPath((Join-Path $outputPath $fullFileName))
        wordStartupPath   = $wordStartupPath
        VBAFolder         = [System.IO.Path]::GetFullPath((Join-Path $scriptPath $vbaRel))
        ribbonXml         = [System.IO.Path]::GetFullPath((Join-Path $scriptPath $ribbonRel))
        pythonServerPath  = [System.IO.Path]::GetFullPath((Join-Path $scriptPath $pythonRel))
    }
}

function Test-CheckFileExists {
    param($filePath)
    return [bool](Test-Path $filePath)
    
    if (Test-Path $filePath) {
        Write-Host "    File exists: $filePath"
        return $true
    }
    Write-Host "    File does not exist: $filePath"
    return $false
}

function Remove-FileIfExists {
    param(
        $FilePath,
        [switch]$Force
    )
    
    if (Test-Path $FilePath) {
        if (-not $Force) {
            $choice = Read-Host "    Would you like to remove now? (y/n)"
            if ($choice -ne 'y') {
                Write-Host "    + Skipped removal and Stopping script" -ForegroundColor Red
                return $false
            }
        }
        
        Write-Host "    - Removing: $FilePath" -ForegroundColor Yellow
        try {
            Remove-Item -Path $FilePath -Force -ErrorAction Stop
            Start-Sleep -Milliseconds 200
            Write-Host "    + Removed successfully" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "    - Failed to remove: $_" -ForegroundColor Red
            return $false
        }
    }
    Write-Host "    + File does not exist: $FilePath"
    return $true
}

function New-WordApplication {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet("COM", "Process")]
        $Mode = "COM"
    )

    try {
        Write-Host "    + Creating Word instance..." -ForegroundColor Yellow

        if ($Mode -eq "Process") {
            $interop = Initialize-ActiveObjectSupport
            $proc = Start-Process -FilePath "WINWORD" -PassThru
            $wordPID = $proc.Id

            $word = $null
            $timeout = 15 # Max timeout seconds
            $startTime = [System.DateTime]::Now

            # Deterministic loop checking if the process is active and responding to COM
            while ($null -eq $word -and ([System.DateTime]::Now - $startTime).TotalSeconds -lt $timeout) {
                if (-not (Get-Process -Id $wordPID -ErrorAction SilentlyContinue)) {
                    throw "Word process with PID $wordPID terminated unexpectedly."
                }

                try {
                    $word = $interop::GetActiveObject("Word.Application")
                } catch {
                    Start-Sleep -Milliseconds 100
                }
            }

            if ($null -eq $word) {
                throw "Timed out waiting for Word COM server to initialize."
            }

            $word.Visible = $true
        } else {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            
            # Obtain PID from recent Word process
            $wordProcess = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1
            $wordPID = if ($wordProcess) { $wordProcess.Id } else { 0 }
        }

        $word.DisplayAlerts = 0

        Write-Host "    + Word application with [PID: $wordPID] started" -ForegroundColor Green
        
        # Add PID as a NoteProperty to the word object
        $word | Add-Member -NotePropertyName "ProcessID" -NotePropertyValue $wordPID

        return $word
    } catch {
        Write-Host "    - Failed to create Word instance: $_" -ForegroundColor Red
        return $null
    }
}

function Close-WordApplication {
    param($WordApp)
    
    if (-not $WordApp) {
        Write-Host "    + No Word instance to close"
        return $true
    }

    $wordPID = $null

    try {
        $wordPID = $WordApp.ProcessID
        Write-Host "    + Closing Word instance..."
        $WordApp.Quit()
    } catch {
        Write-Host "    - Failed to call Quit() on Word: $_" -ForegroundColor Red
    } finally {
        # Explicitly release the COM wrapper
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($WordApp)) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApp) | Out-Null
        }
        
        # Force Garbage Collection twice for CoreCLR / PowerShell 7
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        if ($wordPID) {
            while (Get-Process -Id $wordPID -ErrorAction SilentlyContinue) {
                Start-Sleep -Milliseconds 100
            }
        }
    }

    Write-Host "    + Word application with [PID: $wordPID] closed" -ForegroundColor Green
    return $true
}

function New-Document {
    param($WordApp)
    
    try {
        # Write-Host "    + Creating new document..." -ForegroundColor Yellow
        $doc = $WordApp.Documents.Add()
        Write-Host "    + Document created: '$($doc.Name)'" -ForegroundColor Green
        return $doc
    } catch {
        Write-Host "    - Failed to create document: $_" -ForegroundColor Red
        return $null
    }
}

function Save-Document {
    param(
        $Document,
        $FilePath,
        [int]$WordPID = 0,
        [switch]$RemovePersonalInfo
    )
    
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $format = if ($extension -eq ".docm") { 13 } else { 15 }
    
    $dismissJob = $null
    try {
        if ($RemovePersonalInfo) {
            $Document.RemovePersonalInformation = $true
            if ($WordPID -gt 0) {
                $dismissJob = Start-DocumentInspectorDismissJob -WordPID $WordPID
            }
        }
        
        $Document.SaveAs2([ref][System.Object]$FilePath, [ref][System.Object]$format, [ref]$false)
        Write-Host "    + Document saved: $FilePath" -ForegroundColor Green
        
        if ($dismissJob) {
            Stop-DocumentInspectorJob -Job $dismissJob
        }
        return $true
    } catch {
        if ($dismissJob) {
            Stop-DocumentInspectorJob -Job $dismissJob
        }
        Write-Host "    - Failed to save document: $_" -ForegroundColor Red
        return $false
    }
}

function Start-DocumentInspectorDismissJob {
    param([int]$WordPID)
    
    if (-not $WordPID) { return $null }
    
    return Start-Job -ScriptBlock {
        param($targetPid)
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes

        $procCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty, 
            $targetPid
        )

        for ($i = 1; $i -le 30; $i++) {
            $topLevelWindows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
                [System.Windows.Automation.TreeScope]::Children, 
                $procCondition
            )

            $targetButton = $null
            foreach ($win in $topLevelWindows) {
                $targetButton = $win.FindFirst(
                    [System.Windows.Automation.TreeScope]::Descendants,
                    [System.Windows.Automation.AndCondition]::new(
                        [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::AutomationIdProperty, "1"),
                        [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ClassNameProperty, "NetUIButton")
                    )
                )
                if ($targetButton) { break }
            }

            if ($targetButton) {
                $invokePattern = $targetButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                if ($invokePattern) {
                    $invokePattern.Invoke()
                    return
                }
            }
            Start-Sleep -Milliseconds 300
        }
    } -ArgumentList $WordPID
}

function Stop-DocumentInspectorJob {
    param($Job)
    if ($Job) {
        Stop-Job $Job -ErrorAction SilentlyContinue
        Remove-Job $Job -ErrorAction SilentlyContinue
    }
}

function Close-Document {
    param($Document)
    
    if ($Document) {
        $docName = "Unknown"
        try {
            $docName = $Document.Name
            Write-Host "    + Closing document '$docName'..." -ForegroundColor Yellow
            
            # SaveChanges = 0 (wdDoNotSaveChanges) to prevent prompts during teardown
            $Document.Close([ref]0) 
            Write-Host "    + Document '$docName' closed successfully" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "    - Failed to close document '$docName': $_" -ForegroundColor Red
            return $false
        } finally {
            # Guarantee COM release regardless of whether .Close() succeeded or threw an error
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($Document)) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
            }
        }
    }
    
    Write-Host "    + No document to close"
    return $true
}

function Import-VBAComponents {
    param(
        $project,
        [array]$VBAList,
        $pythonExe,
        $pythonServerPath,
        [switch]$UpdatePythonPaths
    )
    
    if (-not $VBAList -or $VBAList.Count -eq 0) {
        Write-Host "    - No VBA modules provided to import." -ForegroundColor Red
        return $false
    }

    # Resolve all files into FileInfo objects
    $files = @()
    foreach ($item in $VBAList) {
        if (Test-Path $item) {
            $files += Get-Item $item
        } else {
            Write-Host "    - Module file not found: $item" -ForegroundColor Red
            return $false
        }
    }

    Write-Host "    + Found $($files.Count) VBA files to import" -ForegroundColor Green

    # PHASE 1: DELETE ALL EXISTING COMPONENTS ONCE
    $components = @($project.VBComponents)

    foreach ($component in $components) {
        $name = $component.Name
        
        # Skip document-level objects (Type 10 = vbext_ct_Document or named "ThisDocument")
        if ($component.Type -ne 10 -and $name -ne "ThisDocument") {
            $project.VBComponents.Remove($component)
            Write-Host "    - Deleted: $name" -ForegroundColor Red
        }
    }

    # PHASE 2: IMPORT ALL SPECIFIED FILES
    foreach ($file in $files) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        
        try {
            # For .bas files, ensure Attribute VB_Name exists
            if ($file.Extension -eq ".bas") {
                # Only update Python paths if the switch is provided AND paths are not null/empty
                if ($UpdatePythonPaths -and $file.Name -eq "aWordbotRibbonLLMFunctions.bas" -and $pythonExe -and $pythonServerPath) {
                    $raw = Get-Content $file.FullName -Raw -Encoding UTF8
                    if ($raw -match "\{\{PYTHON_SERVER_PATH\}\}" -or $raw -match "\{\{PYTHON_EXE_PATH\}\}") {
                        $updated = $raw -replace "\{\{PYTHON_EXE_PATH\}\}", $pythonExe
                        $updated = $updated -replace "\{\{PYTHON_SERVER_PATH\}\}", $pythonServerPath

                        if ($updated -notmatch "^Attribute VB_Name") {
                            $updated = "Attribute VB_Name = `"$name`"`r`n$updated"
                        }
                        
                        $tempFile = [System.IO.Path]::GetTempFileName()
                        [System.IO.File]::WriteAllText($tempFile, $updated, [System.Text.UTF8Encoding]::new($false))
                        $null = $project.VBComponents.Import($tempFile)
                        Remove-Item $tempFile -Force
                        Write-Host "    + Configured Python executable and server path in $($file.Name)" -ForegroundColor Green
                        continue
                    }
                }

                $firstLine = Get-Content $file.FullName -First 1
                if ($firstLine -notmatch "^Attribute VB_Name") {
                    $content = Get-Content $file.FullName -Raw -Encoding UTF8
                    $newContent = "Attribute VB_Name = `"$name`"`r`n$content"
                    
                    # Write modified content to temporary file and import
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    [System.IO.File]::WriteAllText($tempFile, $newContent, [System.Text.UTF8Encoding]::new($false))
                    $null = $project.VBComponents.Import($tempFile)
                    Remove-Item $tempFile -Force
                    Write-Host "    + Imported: $($file.Name)" -ForegroundColor Magenta
                } else {
                    $null = $project.VBComponents.Import($file.FullName)
                    Write-Host "    + Imported: $($file.Name)" -ForegroundColor Magenta
                }
            } else {
                # .cls and .frm files - import directly
                $null = $project.VBComponents.Import($file.FullName)
                Write-Host "    + Imported: $($file.Name)" -ForegroundColor Magenta
            }
            
        } catch {
            Write-Host "    - Failed: $($file.Name) - $_" -ForegroundColor Red
            return $false
        }
    }
    
    Write-Host "    + All VBA components imported successfully" -ForegroundColor Green
    return $true
}

function Add-RibbonToTemplate {
    param(
        $FilePath,
        $RibbonXmlPath
    )
    
    try {
        
        if (-not (Test-Path $RibbonXmlPath)) {
            Write-Host "    - Ribbon XML file not found: $RibbonXmlPath" -ForegroundColor Red
            return $false
        } else {
            $ribbonFileName = [System.IO.Path]::GetFileName($RibbonXmlPath)
            Write-Host "    + Adding Ribbon XML file: $ribbonFileName" -ForegroundColor Magenta
        }
        
        $xmlContent = Get-Content $RibbonXmlPath -Raw
        $package = [System.IO.Packaging.Package]::Open($FilePath, 'Open', 'ReadWrite')
        $partUri = New-Object System.Uri("/customUI/customUI14.xml", [System.UriKind]::Relative)
        
        if (-not $package.PartExists($partUri)) {
            $part = $package.CreatePart($partUri, "application/xml")
            $encoding = New-Object System.Text.UTF8Encoding($false)
            $writer = New-Object System.IO.StreamWriter($part.GetStream(), $encoding)
            $writer.Write($xmlContent)
            $writer.Close()
            $package.CreateRelationship($partUri, [System.IO.Packaging.TargetMode]::Internal, "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility") | Out-Null
            Write-Host "    + Ribbon added successfully" -ForegroundColor Green
        } else {
            Write-Host "    + Ribbon part already exists" -ForegroundColor Yellow
        }
        
        $package.Close()
        return $true
        
    } catch {
        Write-Host "    - Failed to add ribbon: $_" -ForegroundColor Red
        return $false
    }
}

# Define the GetActiveObject method once when module loads
$script:interop = $null

function Initialize-ActiveObjectSupport {
    if ("Interop.Interop" -as [type]) {
        $script:interop = [Interop.Interop]
        return $script:interop
    }

    $methodDefinition = @'
using System;
using System.Runtime.InteropServices;

namespace Interop
{
    public static class Interop
    {
        [DllImport("ole32")]
        private static extern int CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid lpclsid);

        [DllImport("oleaut32")]
        private static extern int GetActiveObject([MarshalAs(UnmanagedType.LPStruct)] Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public static object GetActiveObject(string progId)
        {
            if (progId == null)
                throw new ArgumentNullException("progId");

            Guid clsid;
            var hr = CLSIDFromProgIDEx(progId, out clsid);
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
                return null;
            }

            object obj;
            hr = GetActiveObject(clsid, IntPtr.Zero, out obj);
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
                return null;
            }
            return obj;
        }
    }
}
'@

    $script:interop = Add-Type -TypeDefinition $methodDefinition -PassThru
    return $script:interop
}

function Get-ActiveWordInstance {
    try {
        $interop = Initialize-ActiveObjectSupport
        $WordApp = $interop::GetActiveObject("Word.Application")
        
        # Get the PID of the Word process
        $wordProcess = Get-Process -Name "WINWORD" | Sort-Object StartTime -Descending | Select-Object -First 1
        if ($wordProcess) {
            $wordPID = $wordProcess.Id
            Write-Host "Connected to Word (PID: $wordPID)" -ForegroundColor Green
            # Add PID as a NoteProperty to the word object
            $WordApp | Add-Member -NotePropertyName "ProcessID" -NotePropertyValue $wordPID -Force
        }
        
        return $WordApp
    } catch {
        Write-Host "    - No active Word document found." -ForegroundColor Red
        return $null
    }
}

function Get-WordInfo {
    param(
        $WordApp,
        [switch]$ReturnObjects
    )
    
    $result = @{
        ProcessID = $WordApp.ProcessID
        DocumentCount = 0
        Documents = @()
        ActiveDocument = $null
        AttachedTemplate = $null
        VBAProject = $null
        VBAProtected = $false
        LoadedTemplates = @()
    }
    
    try {
        Write-Host ""
        
        # Show PID
        if ($WordApp.ProcessID) {
            Write-Host "Word Info for PID ($($WordApp.ProcessID))" -ForegroundColor Cyan
        } else {
            Write-Host "Word Info:" -ForegroundColor Cyan
        }
        
        # List all open documents
        $docCount = $WordApp.Documents.Count
        $result.DocumentCount = $docCount
        Write-Host "Documents open: $docCount" -ForegroundColor Gray
        
        if ($docCount -gt 0) {
            Write-Host "`nOpen Document Names:" -ForegroundColor Cyan
            $i = 1
            foreach ($doc in $WordApp.Documents) {
                $result.Documents += $doc.Name
                Write-Host "    $i. $($doc.Name)" -ForegroundColor Gray
                $i++
            }
            
            # Active document
            $activeDoc = $WordApp.ActiveDocument
            $result.ActiveDocument = $activeDoc
            Write-Host ""
            Write-Host "    Active Document: $($activeDoc.Name)" -ForegroundColor Yellow
            
            # Attached template
            try {
                $template = $activeDoc.AttachedTemplate
                if ($template) {
                    $result.AttachedTemplate = $template
                    Write-Host "    Attached Template:  $($template.Name)" -ForegroundColor Yellow
                    Write-Host "    Template Path:      $($template.Path)"
                }
            } catch {
                Write-Host "Attached Template: Not accessible" -ForegroundColor Red
            }
            
            # VBA project
            try {
                $project = $activeDoc.VBProject
                $result.VBAProject = $project
                $result.VBAProtected = ($project.Protection -ne 0)
                Write-Host "    VBA Project: $($project.Name)" -ForegroundColor Yellow
                Write-Host "    VBA Protected: $($project.Protection -ne 0)"
            } catch {
                Write-Host "    VBA Project: Not accessible" -ForegroundColor Red
            }
        }
        
        # Show loaded templates
        Write-Host ""
        Write-Host "Loaded Templates:" -ForegroundColor Cyan
        $templateCount = 0
        foreach ($t in $WordApp.Templates) {
            $templateCount++
            $result.LoadedTemplates += $t
            Write-Host "    $templateCount. $($t.Name)" -ForegroundColor Gray
        }
        if ($templateCount -eq 0) {
            Write-Host "No templates loaded" -ForegroundColor Gray
        }
        
        if ($ReturnObjects) {
            return $result
        }
        return $true
        
    } catch {
        Write-Host "    - Failed to get Word info: $_" -ForegroundColor Red
        return $false
    }
}

function Test-LoadedTemplate {
    param(
        [Parameter(Mandatory)]$WordAppInfo,
        [Parameter(Mandatory)]$TemplateName
    )
    
    $extension = [System.IO.Path]::GetExtension($TemplateName)
    if ($extension -ne ".dotm") {
        return $null
    }
    
    # Check if template is already loaded
    $loadedTemplate = $WordAppInfo.LoadedTemplates | Where-Object { $_.Name -eq $TemplateName } | Select-Object -First 1
    
    if (-not $loadedTemplate) {
        Write-Host "No templates in word startup with same name as '$TemplateName'. Continuing..." -ForegroundColor Magenta
        return $null
    }
    
    # Template is loaded - ask user what to do
    Write-Host ""
    Write-Host "IMPORTANT!!: Your project template '$TemplateName' is currently loaded in Word." -ForegroundColor Yellow
    Write-Host "    It has the project:   '$($loadedTemplate.VBProject.Name)'" -ForegroundColor Yellow
    Write-Host "    You can either choose:"
    Write-Host "        y- Edit the loaded template directly (macros will be saved to template)" -ForegroundColor Magenta
    Write-Host "        n- Close Word, remove from STARTUP, and test with a .docm" -ForegroundColor Magenta
    Write-Host ""
    
    $choice = Read-Host "Edit the loaded template directly? Choice (y/n)"
    
    if ($choice -eq 'y') {
        Write-Host "The chosen one: '$TemplateName'" -ForegroundColor Green            
        $targetProject = $loadedTemplate.VBProject
        Write-Host "Project name:   '$($targetProject.Name)'" -ForegroundColor Green
        return $targetProject
    } else {
        Write-Host "Exiting. Please close Word, remove the '$TemplateName' from STARTUP, and try again." -ForegroundColor Red
        exit 1
    }
}

function New-BlankDocument {
    param(
        [Parameter(Mandatory)]$WordApp,
        [Parameter(Mandatory)]$WordAppInfo
    )
    $WordApp.Visible = $true
    $targetDocument = $WordApp.Documents.Add()
    $targetProject = $targetDocument.VBProject
    
    if ($WordApp.ActiveWindow) {
        $WordApp.ActiveWindow.Visible = $true
        $WordApp.ActiveWindow.WindowState = 1 # wdWindowStateMaximize
    }
    $WordApp.Activate()
    
    # Update WordAppInfo
    $WordAppInfo.DocumentCount = 1
    $WordAppInfo.ActiveDocument = $targetDocument
    $WordAppInfo.Documents = @($targetDocument.Name)
    $WordAppInfo.VBAProject = $targetProject
    
    Write-Host "    Created new blank document." -ForegroundColor Green
    Write-Host "    Project name: '$($targetProject.Name)'" -ForegroundColor Green
    Write-Host "    INFO: When saving, save as .docm or .dotm to keep macros!" -ForegroundColor Yellow
    
    return @{
        WordApp = $WordApp
        WordAppInfo = $WordAppInfo
        TargetProject = $targetProject
        IsNewDocument = $true # Flag to track that we created new document
    }
}

function Get-VBACompileError {
    param(
        $WordApp,
        $Project,
        [switch]$Verbose
    )
    
    # 1. Make BE explicitly visible to ensure UI Automation and VBE commands work
    $WordApp.VBE.MainWindow.Visible = $true
    
    # 2. Set active project
    $Project.VBE.ActiveVBProject = $Project
    
    # 3. FORCE VBE to open a code pane so the parser initializes
    try {
        if ($Project.VBComponents.Count -gt 0) {
            # Opening the first standard module forces the VBE command bar to refresh its state
            $firstComponent = $Project.VBComponents | Where-Object { $_.Type -eq 1 } | Select-Object -First 1
            if ($firstComponent) {
                $null = $firstComponent.CodePane.Show()
            }
        }
    } catch {}
    
    Start-Sleep -Milliseconds 300

    # 4. Query the Compile button
    if ($Verbose) { Write-Host "[VERBOSE] Querying VBA compile button..." -ForegroundColor Cyan }
    $compileButton = $WordApp.VBE.CommandBars.FindControl([int]1, [int]578)

    if ($null -ne $compileButton) {
        # Execute compile
        try {
            if ($Verbose) { Write-Host "[VERBOSE] Executing VBA compilation..." -ForegroundColor Cyan }
            $compileButton.Execute()
        } catch {
            if ($Verbose) { Write-Host "[VERBOSE] Execute failed: $_" -ForegroundColor Red }
        }
    }
    
    Start-Sleep -Milliseconds 250

    # 5. Check if compilation succeeded!
    # If no modal dialog pops up and the compile button is now disabled, it compiled cleanly.
    if ($null -ne $compileButton -and -not $compileButton.Enabled) {
        if ($Verbose) { Write-Host "[VERBOSE] VBA compilation succeeded cleanly!" -ForegroundColor Green }
        return @{ 
            Success          = $true
            Error            = $null
            Module           = $null
            ModuleExtension  = $null
            FullModuleName   = $null
            ErrorLine        = $null
            ErrorLineContent = $null
        }
    }

    # 6. If we reached here, compilation failed -> process error dialog
    if ($Verbose) { Write-Host "[VERBOSE] Compilation failed - waiting for error dialog..." -ForegroundColor Red }
    
    $messageBox = $null
    $maxRetries = 30
    $retryCount = 0
    $moduleName = $null
    $moduleExtension = $null
    $errorLine = $null
    $errorLineContent = $null
    $errorMsg = $null
    
    try {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
    } catch {
        if ($Verbose) { Write-Host "[VERBOSE] UI Automation not available" -ForegroundColor Red }
        return @{ 
            Success          = $false
            Error            = "UI Automation not available"
            Module           = $null
            ModuleExtension  = $null
            FullModuleName   = $null
            ErrorLine        = $null
            ErrorLineContent = $null
        }
    }
    
    while ($retryCount -lt $maxRetries -and $null -eq $messageBox) {
        Start-Sleep -Milliseconds 150
        
        try {
            $condition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ClassNameProperty, "#32770"
            )
            
            $messageBox = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
                [System.Windows.Automation.TreeScope]::Descendants, $condition
            )
            
            if ($null -ne $messageBox) {
                $title = $messageBox.Current.Name
                if ($title -like "Microsoft Visual Basic*") {
                    break
                } else {
                    $messageBox = $null
                }
            }
            
            $retryCount++
        } catch {
            $retryCount++
        }
    }
    
    # Extract error message text from dialog
    if ($null -ne $messageBox) {
        try {
            $textCondition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::AutomationIdProperty, "65535"
            )
            $textElement = $messageBox.FindFirst(
                [System.Windows.Automation.TreeScope]::Descendants, $textCondition
            )
            
            if ($null -ne $textElement) {
                $errorMsg = $textElement.Current.Name
            } else {
                $allElements = $messageBox.FindAll([System.Windows.Automation.TreeScope]::Descendants, 
                    [System.Windows.Automation.Condition]::TrueCondition)
                foreach ($element in $allElements) {
                    $name = $element.Current.Name
                    if ($name -and $name.Trim().Length -gt 0 -and $name -notlike "*OK*" -and $name -notlike "*Cancel*") {
                        $errorMsg = $name
                        break
                    }
                }
            }
        } catch {}
        
        # Close error dialog via Win32
        try {
            if (-not ("User32CompileApi" -as [type])) {
                Add-Type @"
using System;
using System.Runtime.InteropServices;
public class User32CompileApi {
[DllImport("user32.dll", SetLastError = true)]
public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

[DllImport("user32.dll", SetLastError = true)]
public static extern uint SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

public const uint BM_CLICK = 0x00F5;
}
"@ -ErrorAction SilentlyContinue
            }
                
            $hwnd = [IntPtr]$messageBox.Current.NativeWindowHandle
            $okHwnd = [User32CompileApi]::FindWindowEx($hwnd, [IntPtr]::Zero, "Button", "OK")
            
            if ($okHwnd -ne [IntPtr]::Zero) {
                [User32CompileApi]::SendMessage($okHwnd, [User32CompileApi]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
                if ($Verbose) { Write-Host "[VERBOSE] Error dialog closed" -ForegroundColor Green }
            }
        } catch {
            if ($Verbose) { Write-Host "[VERBOSE] Could not close dialog: $($_.Exception.Message)" -ForegroundColor Yellow }
        }
    }
    
    # Retrieve active module and line coordinates highlighted by VBE compiler
    Start-Sleep -Milliseconds 150
    
    if ($Verbose) { Write-Host "[VERBOSE] Getting error line from VBE..." -ForegroundColor Cyan }
    
    try {
        if ($WordApp.VBE.ActiveCodePane) {
            $codePane = $WordApp.VBE.ActiveCodePane
            $codeModule = $codePane.CodeModule
            
            $moduleName = $codeModule.Name
            if ($Verbose) { Write-Host "[VERBOSE] Active module: $moduleName" -ForegroundColor Green }
            
            # Retrieve module extension
            try {
                $vbaProject = $WordApp.VBE.ActiveVBProject
                foreach ($component in $vbaProject.VBComponents) {
                    if ($component.Name -eq $moduleName) {
                        $moduleType = $component.Type
                        switch ($moduleType) {
                            1 { $moduleExtension = ".bas" }
                            2 { $moduleExtension = ".cls" }
                            3 { $moduleExtension = ".frm" }
                            100 { $moduleExtension = ".doc" }
                            default { $moduleExtension = ".unknown" }
                        }
                        if ($Verbose) { Write-Host "[VERBOSE] Module extension: $moduleExtension" -ForegroundColor Cyan }
                        break
                    }
                }
            } catch {}
            
            # Extract line number via VBE GetSelection ref variables
            try {
                $sl = 0; $sc = 0; $el = 0; $ec = 0
                $codePane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
                
                if ($sl -gt 0) {
                    $errorLine = $sl
                    $errorLineContent = $codeModule.Lines($errorLine, 1).Trim()
                    if ($Verbose) { Write-Host "[VERBOSE] Error line from VBE selection: $errorLine" -ForegroundColor Green }
                    if ($Verbose) { Write-Host "[VERBOSE] Line content: $errorLineContent" -ForegroundColor Yellow }
                }
            } catch {
                if ($Verbose) { Write-Host "[VERBOSE] Could not get VBE selection: $($_.Exception.Message)" -ForegroundColor Yellow }
            }
        }
    } catch {
        if ($Verbose) { Write-Host "[VERBOSE] Could not get module info: $($_.Exception.Message)" -ForegroundColor Red }
    }
    
    $sanitizedError = if ($errorMsg) { 
        ($errorMsg -replace '[\r\n]+', ' ' -replace '\s+', ' ').Trim() 
    } else { 
        "Unknown compile error" 
    }

    $result = @{ 
        Success          = $false
        Error            = $sanitizedError
        Module           = $moduleName
        ModuleExtension  = $moduleExtension
        FullModuleName   = if ($moduleName -and $moduleExtension) { "$moduleName$moduleExtension" } else { $moduleName }
        ErrorLine        = $errorLine
        ErrorLineContent = $errorLineContent
    }
    
    if ($Verbose) {
        Write-Host "[VERBOSE] Error: $($result.Error)" -ForegroundColor Cyan
        Write-Host "[VERBOSE] Module: $($result.FullModuleName)" -ForegroundColor Green
        Write-Host "[VERBOSE] Line: $($result.ErrorLine)" -ForegroundColor Cyan
        if ($errorLineContent) {
            Write-Host "[VERBOSE] Line content: $errorLineContent" -ForegroundColor Yellow
        }
    }

    return $result
}

function Get-VBAListForEdition {
    param (
        [string]$VBAFolder,
        [string]$Edition
    )
    
    # Base Markdown modules required by all editions
    $modules = @(
        "Markdown_Main.bas", "MarkdownCleanup.bas", "MarkdownCodeblock.bas",
        "MarkdownHeadersStyle.bas", "MarkdownLatex.bas", "MarkdownLinks.bas",
        "MarkdownListBullet.bas", "MarkdownTable.bas", "MarkdownTextStyle.bas",
        "MarkdownUnicode.bas", "aWordbotCommonFunctions.bas","HeadingNumbers.bas"
    )

    if ($Edition -eq "LLM" -or $Edition -eq "Research") {
        $modules += @("aPythonTask.bas", "aWordbotRibbonLLMFunctions.bas","aWordbotJsonParser.bas", "aWordbot_Main.bas","aWordbotRibbonKeytip.bas")
    }

    if ($Edition -eq "Research") {
        $modules += @("aWordbot_Research.bas", "aWordbot_ResearchCitation.bas")
    }

    return $modules | ForEach-Object { Join-Path $VBAFolder $_ }
}

function Invoke-ScriptCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $WordApp,

        [Parameter(Mandatory = $false)]
        $Document
    )
    
    Write-Host "`nCleaning up resources..." -ForegroundColor Yellow
    
    if ($Document) { 
        Close-Document -Document $Document | Out-Null
    }
    
    if ($WordApp) { 
        Close-WordApplication -WordApp $WordApp | Out-Null
    }
}

function Test-UserConfirmation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [switch]$Force
    )

    if ($Force) {
        Write-Host "$Message (y/n): y [Forced]" -ForegroundColor Cyan
        return $true
    }

    Write-Host ""
    Write-Host $Message -ForegroundColor Yellow
    Write-Host "    You can either choose:"
    Write-Host "        y- Proceed (macros will be updated/overwritten)" -ForegroundColor Magenta
    Write-Host "        n- Do not proceed and stop" -ForegroundColor Magenta
    Write-Host ""

    $choice = Read-Host "Choice (y/n)"
    if ($choice.Trim().ToLower() -eq 'y') {
        return $true
    }

    Write-Host "Stopping script" -ForegroundColor Red
    return $false
}

function Get-OpenTargetDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $WordApp,
        [Parameter(Mandatory = $true)]
        [string]$FullFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedProjectName,
        [Parameter(Mandatory = $true)]
        [string]$FileName,
        [switch]$Force
    )

    $doc = $null
    $isNew = $false
    $fileExists = Test-Path $FullFilePath

    # 1. Match by exact VBA project name across active open documents/templates
    foreach ($openDoc in $WordApp.Documents) {
        try {
            if ($openDoc.VBProject.Name -eq $ExpectedProjectName) {
                $doc = $openDoc
                Write-Host "    + Found open live project: $($doc.Name) (Project: $ExpectedProjectName)" -ForegroundColor Green
                
                # Bring Word to focus so the user sees the active document
                $WordApp.Visible = $true
                if ($WordApp.ActiveWindow) {
                    $WordApp.ActiveWindow.WindowState = 1 # Maximize
                }
                $WordApp.Activate()

                $confirmed = Test-UserConfirmation -Message "IMPORTANT!!: The project '$ExpectedProjectName' is currently active in Word (Document: '$($doc.Name)')." -Force:$Force
                if (-not $confirmed) {
                    return @{
                        Document = $null
                        IsNewDocument = $null
                        ClearVars = $true
                    }
                }
                break
            }
        } catch {}
    }

    # 2. Match by full file path if not already caught by project name
    if (-not $doc) {
        foreach ($openDoc in $WordApp.Documents) {
            if ($openDoc.FullName -eq $FullFilePath) {
                $doc = $openDoc
                Write-Host "    + Document already open by path: $FullFilePath" -ForegroundColor Gray
                break
            }
        }
    }
    
    # 3. Fallback to opening from disk or creating a blank document
    if (-not $doc) {
        if ($fileExists) {
            Write-Host "    + Opening existing document: $FullFilePath" -ForegroundColor Yellow
            $doc = $WordApp.Documents.Open($FullFilePath)
            
            # Make visible and focus window BEFORE prompting user
            $WordApp.Visible = $true
            if ($WordApp.ActiveWindow) {
                $WordApp.ActiveWindow.WindowState = 1 # Maximize
            }
            $WordApp.Activate()

            $confirmed = Test-UserConfirmation -Message "IMPORTANT!!: The file '$FullFilePath' already exists on disk." -Force:$Force
            if (-not $confirmed) {
                return @{
                    Document = $null
                    IsNewDocument = $null
                    ClearVars = $true
                }
            }
        } else {
            Write-Host "`nNo existing files found or loaded with name '$FileName'" -ForegroundColor Cyan
            Write-Host "    + Creating new blank document..." -ForegroundColor Yellow
            
            $WordAppInfo = Get-WordInfo -WordApp $WordApp -ReturnObjects
            $result = New-BlankDocument -WordApp $WordApp -WordAppInfo $WordAppInfo
            
            $doc = if ($result.ActiveDocument) { $result.ActiveDocument } else { $WordApp.ActiveDocument }
            $isNew = $true
            
            if ($result.TargetProject.Name -eq "Project") {
                $result.TargetProject.Name = $ExpectedProjectName
                Write-Host "    Project name renamed to '$ExpectedProjectName'" -ForegroundColor Magenta
            }
        }
    }

    return @{
        Document = $doc
        IsNewDocument = $isNew
        ClearVars = $false
    }
}