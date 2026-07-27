' docLiveUpdateMacros Launcher

Set shell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the current folder where this VBS is located  
currentFolder = fso.GetParentFolderName(WScript.ScriptFullName)

' Full path to the PowerShell script
psScript = currentFolder & "\docLiveUpdateMacros.ps1"

' Verify the PowerShell script exists
If Not fso.FileExists(psScript) Then
    MsgBox "PowerShell script not found: " & psScript, vbCritical, "Error"
    WScript.Quit 1
End If

' Run PowerShell script as administrator
shell.ShellExecute "powershell.exe", "-NoProfile -ExecutionPolicy Bypass -Command ""& '" & psScript & "'; Write-Host ''; Read-Host 'Press Enter to exit'""", "", "open", 1