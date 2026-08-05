Public Function PythonTask(ByVal endpoint As String, ByVal jsonBody As String) As String
    Dim url As String
    Dim req As Object
    Dim dotCounter As Long
    Dim baseMessage As String

    url = "http://localhost:3670" & endpoint

    ' Check if running on Mac
    #If Mac Then
        ' Mac doesn't have MSXML2, use a wrapper for curl using AppleScript to call curl
        PythonTask = MacPythonTaskCurl(url, jsonBody)
        Exit Function
    #Else
        Set req = CreateObject("MSXML2.XMLHTTP")
        On Error GoTo NetworkError
    #End If
    
    ' Set up dynamic message variables based on endpoint
    If endpoint = "/process" Then
        baseMessage = "LLM processing in background"
    ElseIf endpoint = "/research" Then
        baseMessage = "Searching Zotero"
    Else
        baseMessage = "Processing request"
    End If
    
    stopStatusRefresh = False
    dotCounter = 0
    
    With req
        .Open "POST", url, True
        .setRequestHeader "Content-Type", "application/json"
        .send jsonBody
        
        ' Dynamic Polling Loop: Updates the UI immediately!
        Do While .readyState < 4
            ' Generate a clean, animated loading indicator (e.g., "Searching Zotero...")
            dotCounter = (dotCounter Mod 5) + 1
            Application.StatusBar = baseMessage & String(dotCounter, ".")
            
            DoEvents
            Pause 0.5 ' Slightly longer pause to make the dots visibly dance smoothly
        Loop
        
        ' Handle Response Status
        If .Status = 200 Then
            stopStatusRefresh = True
            If endpoint = "/process" Then
                Application.StatusBar = "Obtained LLM response!"
                ' Safe to use OnTime here because the heavy processing is over
                Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
            ElseIf endpoint = "/research" Then
                Application.StatusBar = "Obtained Zotseek response."
            End If
            PythonTask = .responseText
        ElseIf .Status = 569 Then
            stopStatusRefresh = True
            If endpoint = "/process" Then
                MsgBox "Error from LLM server: " & .Status, vbCritical
            ElseIf endpoint = "/research" Then
                MsgBox "Error from Zotero server: " & .responseText, vbCritical, "Zotero Error"
            Else
                MsgBox "Error from server: " & .responseText, vbCritical
            End If
            PythonTask = ""
        ElseIf .Status = 503 Then
            stopStatusRefresh = True
            If endpoint = "/process" Then
                MsgBox "The LLM Backend is not running!", vbCritical, "Server Offline"
            ElseIf endpoint = "/research" Then
                MsgBox .responseText, vbCritical, "Server Offline"
            Else
                MsgBox "Server is offline.", vbCritical, "Server Offline"
            End If
            PythonTask = ""
        Else
            GoTo NetworkError
        End If
    End With
    
    Set req = Nothing
    Exit Function
    
NetworkError:
    stopStatusRefresh = True
    MsgBox "Could not connect to the Python server.", vbCritical
    Application.StatusBar = ""
    PythonTask = ""
End Function

Private Sub ClearStatusBar()
    Application.StatusBar = ""
End Sub

Private Sub Pause(seconds As Double)
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + seconds
        DoEvents
    Loop
End Sub

' Mac-specific function using curl via AppleScriptTask
#If Mac Then
Private Function MacPythonTaskCurl(ByVal url As String, ByVal jsonBody As String) As String
    ' Escape JSON for shell - handle special characters properly
    Dim escapedJson As String
    escapedJson = Replace(jsonBody, "\", "\\")
    escapedJson = Replace(escapedJson, "'", "'\''")
    escapedJson = Replace(escapedJson, vbCrLf, "\n")
    escapedJson = Replace(escapedJson, vbLf, "\n")
    
    Dim curlCommand As String
    Dim result As String
    
    ' Build curl shell command string directly without do shell script wrapper
    curlCommand = "curl -s -X POST " & url & " -H 'Content-Type: application/json' -d '" & escapedJson & "'"
    
    On Error Resume Next
    ' Execute AppleScript via AppleScriptTask helper for Office 2016+ compatibility
    result = AppleScriptTask("WordbotCurl.scpt", "doShellCurl", curlCommand)
    
    If Err.Number <> 0 Then
        ' Error occurred
        MacPythonTaskCurl = ""
        MsgBox "Curl request failed: " & Err.Description, vbCritical
        Err.Clear
    Else
        ' Check if result contains error indicators
        If InStr(result, "Could not connect") > 0 Or InStr(result, "Failed to connect") > 0 Then
            MacPythonTaskCurl = ""
            MsgBox "Could not connect to Python server at " & url, vbCritical
        Else
            MacPythonTaskCurl = result
            Application.StatusBar = "Request completed"
        End If
    End If
    On Error GoTo 0
End Function
#End If
