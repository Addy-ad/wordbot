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
            ($line -split '[=<>]')[0].Trim()
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
    }

    # Default fallback values
    if (-not $fileName) { $fileName = "Wordbot" }
    if (-not $extension) { $extension = ".dotm" }
    if (-not $outputPath) { $outputPath = $wordStartupPath }
    
    $fullFileName = $fileName + $extension

    $vbaRel = if ($Config.VBAFolderRelativePath) { $Config.VBAFolderRelativePath } else { "Resources\WordVBAmacros" }
    $ribbonRel = if ($Config.RibbonXmlRelativePath) { $Config.RibbonXmlRelativePath } else { "Resources\WordbotRibbon.xml" }
    $pythonRel = if ($Config.PythonServerLocation) { $Config.PythonServerLocation } else { "..\Python\main.py" }

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
        [switch]$Ask
    )
    
    if (Test-Path $FilePath) {
        if ($Ask) {
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
            Write-Host "    - File may be open or locked. Close any programs using it and try again." -ForegroundColor Yellow
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

            # Poll the COM server until Word is fully ready
            $word = $null
            $timeout = 10 # seconds
            $startTime = [System.DateTime]::Now

            while ($null -eq $word -and ([System.DateTime]::Now - $startTime).TotalSeconds -lt $timeout) {
                try {
                    $word = $interop::GetActiveObject("Word.Application")
                } catch {
                    Start-Sleep -Milliseconds 250
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
    
    if ($WordApp) {
        try {
            $wordPID = $WordApp.ProcessID  # Get the stored PID
            Write-Host "    + Closing Word instance..."
            $WordApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApp) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Write-Host "    + Word application with [PID: $wordPID] closed" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "    - Failed to close Word instance: $_" -ForegroundColor Red
            return $false
        }
    }
    Write-Host "    + No Word instance to close"
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
        $FilePath
    )
    
    # Determine format from file extension
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $format = if ($extension -eq ".docm") { 13 } else { 15 }
    
    try {
        # Write-Host "    + Saving document..." -ForegroundColor Yellow
        $Document.SaveAs2([ref][System.Object]$FilePath, [ref][System.Object]$format, [ref]$false)
        Write-Host "    + Document saved: $FilePath" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "    - Failed to save document: $_" -ForegroundColor Red
        return $false
    }
}

function Close-Document {
    param($Document)
    
    if ($Document) {
        try {
            $docName = $Document.Name
            Write-Host "    + Closing document..."
            $Document.Close()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
            Write-Host "    + Document '$docName' closed successfully" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "    - Failed to close document: $_" -ForegroundColor Red
            return $false
        }
    }
    Write-Host "    + No document to close"
    return $true
}

function Import-VBAComponents {
    param(
        $project,
        $VBAFolder,
        $pythonExe
    )
    
    if (-not (Test-Path $VBAFolder)) {
        Write-Host "    - VBA folder not found: $VBAFolder" -ForegroundColor Red
        return $false
    }
    
    # Import VBA components
    $files = Get-ChildItem -Path "$VBAFolder\*.bas", "$VBAFolder\*.cls", "$VBAFolder\*.frm"
    Write-Host "    + Found $($files.Count) VBA files" -ForegroundColor Green
    
    if ($files.Count -eq 0) {
        Write-Host "    - No VBA files found in: $VBAFolder" -ForegroundColor Red
        return $false
    }
    
	# PHASE 1: DELETE ALL EXISTING COMPONENTS ONCE
	# Get a static snapshot array of all components
	$components = @($project.VBComponents)

	foreach ($component in $components) {
		$name = $component.Name
		
		# Skip document-level objects (Type 10 = vbext_ct_Document or named "ThisDocument")
		if ($component.Type -ne 10 -and $name -ne "ThisDocument") {
			$project.VBComponents.Remove($component)
			Write-Host "    - Deleted: $name" -ForegroundColor Red
		}
	}

	# PHASE 2: IMPORT ALL FILES
	foreach ($file in $files) {
		$name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
		
		try {
			# For .bas files, ensure Attribute VB_Name exists
			if ($file.Extension -eq ".bas") {
				if ($file.Name -eq "aWordbotRibbonBtnFunctions.bas") {
					$raw = Get-Content $file.FullName -Raw -Encoding UTF8
					if ($raw -match "\{\{PYTHON_SERVER_PATH\}\}" -or $raw -match "\{\{PYTHON_EXE_PATH\}\}") {
						$updated = $raw -replace "\{\{PYTHON_EXE_PATH\}\}", $pythonExe
						$updated = $updated -replace "\{\{PYTHON_SERVER_PATH\}\}", $paths.pythonServerPath

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
            Write-Host "Open Document Names:" -ForegroundColor Yellow
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

function Get-TargetProject {
    param(
        [Parameter(Mandatory)]$WordApp,
        [Parameter(Mandatory)]$WordAppInfo,
        [string]$TemplateName,      # Name of template (if .dotm provided)
        [string]$File              # Full path to .docm or .dotm
    )

    # Set default template name if not provided
    if ($TemplateName -eq "" -or $null -eq $TemplateName) {
        $TemplateName = "Wordbot.dotm"
    }

    # CASE 1: Check if template is loaded
    $targetProject = Test-LoadedTemplate -WordAppInfo $WordAppInfo -TemplateName $TemplateName

    Write-Host ""
    if ($targetProject) {
        # Template loaded but no document open
        if ($WordAppInfo.DocumentCount -eq 0) {
            Write-Host "No documents open. Create a new blank document?" -ForegroundColor Yellow
            $choice = Read-Host "Choice (y/n)"
            
            if ($choice -eq 'y') {
                $result = New-BlankDocument -WordApp $WordApp -WordAppInfo $WordAppInfo
                return @{
                    WordApp = $result.WordApp
                    WordAppInfo = $result.WordAppInfo
                    TargetProject = $targetProject  # Keep template project
                }
            } else {
                Write-Host "Exiting without document." -ForegroundColor Red
                exit 1
            }
        } else {
            # Template loaded and documents exist
            return @{
                WordApp = $WordApp
                WordAppInfo = $WordAppInfo
                TargetProject = $targetProject
            }
        }
    }

    # CASE 2: This is the case when the template is not loaded and no document is open
    if ($WordAppInfo.DocumentCount -eq 0) {
        Write-Host "No documents open. Create a new blank document?" -ForegroundColor Yellow
        $choice = Read-Host "Choice (y/n)"
        
        if ($choice -eq 'y') {
            $result = New-BlankDocument -WordApp $WordApp -WordAppInfo $WordAppInfo
            if ($result.TargetProject.Name -eq "Project") {
                Write-Host "Project name renamed" -ForegroundColor Magenta
                $result.TargetProject.Name = ([System.IO.Path]::GetFileNameWithoutExtension($TemplateName)+"Project")
            }

            return @{
                WordApp = $result.WordApp
                WordAppInfo = $result.WordAppInfo
                TargetProject = $result.TargetProject  # Use the new document's project
            }
        } else {
            Write-Host "Exiting without document." -ForegroundColor Red
            exit 1
        }
    }

    # CASE 3: Check if user provided a file in the list of open documents of not, use existing active document
    if (-not [string]::IsNullOrEmpty($File)) {
        $fileNameWithExtension = [System.IO.Path]::GetFileName($File)
        $fileNameOnly = [System.IO.Path]::GetFileNameWithoutExtension($File)
        
        $targetDocument = $WordApp.Documents | Where-Object { 
            $_.Name -eq $fileNameWithExtension -or $_.FullName -eq $File 
        } | Select-Object -First 1
        
        if ($targetDocument) {
            $targetProject = $targetDocument.VBProject
            Write-Host "Found file in word:             '$fileNameWithExtension'" -ForegroundColor Green
            Write-Host "File path of the found file:    '$($targetDocument.FullName)'" -ForegroundColor Gray
            Write-Host "Project name:                   '$($targetProject.Name)'" -ForegroundColor Green
            Write-Host "Matches with user file:         '$File'" -ForegroundColor Gray

            return @{
                WordApp = $WordApp
                WordAppInfo = $WordAppInfo
                TargetProject = $targetProject
            }
        } else {
            Write-Host "The provided file '$fileNameWithExtension' is not open in Word." -ForegroundColor Red
            Write-Host ""

            # Show all open documents with numbers
            Write-Host "But other Open Documents exists:" -ForegroundColor Yellow
            $docIndex = 1
            foreach ($doc in $WordApp.Documents) {
                Write-Host "    $docIndex. $($doc.Name)" -ForegroundColor Gray
                $docIndex++
            }
            Write-Host "    0. Create a new blank document" -ForegroundColor Gray
            Write-Host ""

            $choice = Read-Host "Select a document by number to work on that document"

            if ($choice -eq '0') {
                $result = New-BlankDocument -WordApp $WordApp -WordAppInfo $WordAppInfo
                return @{
                    WordApp = $result.WordApp
                    WordAppInfo = $result.WordAppInfo
                    TargetProject = $result.TargetProject
                }
            } else {
                $selectedIndex = [int]$choice - 1
                if ($selectedIndex -ge 0 -and $selectedIndex -lt $WordApp.Documents.Count) {
                    $targetDocument = $WordApp.Documents.Item($selectedIndex + 1)
                    $targetProject = $targetDocument.VBProject
                    if ($targetProject.Name -eq "Project") {
                        Write-Host "Project name renamed" -ForegroundColor Magenta
                        $targetProject.Name = $fileNameOnly + "Project"
                    }
                    Write-Host "Selected: '$($targetDocument.Name)'" -ForegroundColor Green
                    Write-Host "Project name: '$($targetProject.Name)'" -ForegroundColor Green
                    return @{
                        WordApp = $WordApp
                        WordAppInfo = $WordAppInfo
                        TargetProject = $targetProject
                    }
                } else {
                    Write-Host "Invalid selection. Exiting." -ForegroundColor Red
                    exit 1
                }
            }
        }
    } else {
        Write-Host "Using existing active document: '$($WordAppInfo.ActiveDocument.Name)'" -ForegroundColor Green
        $targetProject = $WordAppInfo.ActiveDocument.VBProject
        if ($targetProject.Name -eq "Project") {
            Write-Host "Project name renamed" -ForegroundColor Magenta
            $targetProject.Name = ([System.IO.Path]::GetFileNameWithoutExtension($WordAppInfo.ActiveDocument.Name) + "Project")
        }
        Write-Host "Project name: '$($targetProject.Name)'" -ForegroundColor Green
        return @{
            WordApp = $WordApp
            WordAppInfo = $WordAppInfo
            TargetProject = $targetProject
        }
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
    
    $WordApp.ActiveWindow.Visible = $true
    $WordApp.ActiveWindow.WindowState = 1
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
        [switch]$Verbose
    )
    
    if ($Verbose) { Write-Host "[VERBOSE] Checking VBA compilation..." -ForegroundColor Cyan }
    
    # Get the compile button
    $compileButton = $WordApp.VBE.CommandBars.FindControl([Microsoft.Office.Core.MsoControlType]::msoControlButton, 578)
    
    if (-not $compileButton) {
        if ($Verbose) { Write-Host "[VERBOSE] Compile button not found" -ForegroundColor Red }
        return @{ Success = $true; Error = $null; Module = $null; ModuleExtension = $null; ErrorLine = $null; ErrorLineContent = $null }
    }
    
    if (-not $compileButton.Enabled) {
        if ($Verbose) { Write-Host "[VERBOSE] No compilation needed (button disabled)" -ForegroundColor Green }
        return @{ Success = $true; Error = $null; Module = $null; ModuleExtension = $null; ErrorLine = $null; ErrorLineContent = $null }
    }
    
    # Execute compile
    if ($Verbose) { Write-Host "[VERBOSE] Compiling VBA project..." -ForegroundColor Cyan }
    $compileButton.Execute()
    
    # Wait for dialog using UI Automation
    $messageBox = $null
    $maxRetries = 10
    $retryCount = 0
    $moduleName = $null
    $moduleExtension = $null
    $errorLine = $null
    $errorLineContent = $null
    
    try {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
    } catch {
        if ($Verbose) { Write-Host "[VERBOSE] UI Automation not available" -ForegroundColor Red }
        return @{ Success = $false; Error = "UI Automation not available"; Module = $null; ModuleExtension = $null; ErrorLine = $null; ErrorLineContent = $null }
    }
    
    while ($retryCount -lt $maxRetries -and $null -eq $messageBox) {
        Start-Sleep -Milliseconds 300
        
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
    
    # Get fresh reference to compile button
    $compileButton = $WordApp.VBE.CommandBars.FindControl([Microsoft.Office.Core.MsoControlType]::msoControlButton, 578)
    
    if ($compileButton.Enabled) {
        if ($Verbose) { Write-Host "[VERBOSE] Compilation failed" -ForegroundColor Red }
        $errorMsg = $null
        
        # ============================================================
        # Get module info from VBE
        # ============================================================
        if ($Verbose) { Write-Host "[VERBOSE] Getting module information from VBE..." -ForegroundColor Cyan }
        
        try {
            if ($WordApp.VBE.ActiveCodePane) {
                $codePane = $WordApp.VBE.ActiveCodePane
                $codeModule = $codePane.CodeModule
                
                $moduleName = $codeModule.Name
                if ($Verbose) { Write-Host "[VERBOSE] Active module: $moduleName" -ForegroundColor Green }
                
                # Get module extension
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
                
                # ============================================================
                # Find the ACTUAL error line
                # ============================================================
                try {
                    # Method 1: Get selection (what's highlighted)
                    $selection = $codePane.GetSelection(1, 1, 1, 1)
                    if ($selection) {
                        $errorLine = $selection.TopLine
                        if ($Verbose) { Write-Host "[VERBOSE] Selection line: $errorLine" -ForegroundColor Cyan }
                    }
                } catch {
                    if ($Verbose) { Write-Host "[VERBOSE] Could not get selection: $($_.Exception.Message)" -ForegroundColor Yellow }
                }
                
                # Method 2: Check TopLine for problematic patterns
                if (-not $errorLine) {
                    try {
                        $topLine = $codePane.TopLine
                        if ($topLine -gt 0) {
                            $lineToCheck = $codeModule.Lines($topLine, 1)
                            $problematicPatterns = @(
                                "*ParentNode As node*",
                                "*Children As Collection*",
                                "*ParentNode As*",
                                "*Public *"
                            )
                            foreach ($pattern in $problematicPatterns) {
                                if ($lineToCheck -like $pattern -and $topLine -gt 5) {
                                    $errorLine = $topLine
                                    $errorLineContent = $lineToCheck
                                    if ($Verbose) { Write-Host "[VERBOSE] Found error line from TopLine: $errorLine" -ForegroundColor Cyan }
                                    break
                                }
                            }
                        }
                    } catch {}
                }
                
                # Method 3: Search entire module for problematic patterns
                if (-not $errorLine) {
                    try {
                        $totalLines = $codeModule.CountOfLines
                        $problematicPatterns = @(
                            "*ParentNode As node*",
                            "*Children As Collection*",
                            "*ParentNode As*",
                            "*Public *"
                        )
                        
                        for ($i = 1; $i -le [Math]::Min($totalLines, 50); $i++) {
                            $line = $codeModule.Lines($i, 1)
                            foreach ($pattern in $problematicPatterns) {
                                if ($line -like $pattern -and $i -gt 5) {
                                    $errorLine = $i
                                    $errorLineContent = $line
                                    if ($Verbose) { Write-Host "[VERBOSE] Found pattern '$pattern' at line: $errorLine" -ForegroundColor Cyan }
                                    break
                                }
                            }
                            if ($errorLine) { break }
                        }
                    } catch {}
                }
                
                if ($Verbose -and $errorLine) {
                    Write-Host "[VERBOSE] Error line content: $errorLineContent" -ForegroundColor Yellow
                }
            }
        } catch {
            if ($Verbose) { Write-Host "[VERBOSE] Could not get module info: $($_.Exception.Message)" -ForegroundColor Red }
        }
        
        # ============================================================
        # Get error text from dialog
        # ============================================================
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
            
            # Close dialog
            try {
                Add-Type @"
using System;
using System.Runtime.InteropServices;
public class User32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    public const uint BM_CLICK = 0x00F5;
}
"@ -ErrorAction SilentlyContinue
                
                $hwnd = [IntPtr]$messageBox.Current.NativeWindowHandle
                $okHwnd = [User32]::FindWindowEx($hwnd, [IntPtr]::Zero, "Button", "OK")
                
                if ($okHwnd -ne [IntPtr]::Zero) {
                    [User32]::SendMessage($okHwnd, [User32]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
                }
            } catch {}
        }
        
        $result = @{ 
            Success = $false
            Error = if ($errorMsg) { $errorMsg.Trim() } else { "Unknown compile error" }
            Module = $moduleName
            ModuleExtension = $moduleExtension
            FullModuleName = if ($moduleName -and $moduleExtension) { "$moduleName$moduleExtension" } else { $moduleName }
            ErrorLine = $errorLine
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
        
    } else {
        if ($Verbose) { Write-Host "[VERBOSE] Compilation successful" -ForegroundColor Green }
        return @{ Success = $true; Error = $null; Module = $null; ModuleExtension = $null; ErrorLine = $null; ErrorLineContent = $null }
    }
}