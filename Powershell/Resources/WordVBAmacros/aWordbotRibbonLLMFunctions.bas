Public Sub AutoExec()
    ' Bind LLM keys
    BindUnbindLLM 1
    
    ' Safely trigger Research keybindings if that module exists in this edition
    On Error Resume Next
    Application.Run "BindUnbindResearch", 1
    On Error GoTo 0
End Sub

Public Sub AutoExit()
    ' Unbind LLM keys
    BindUnbindLLM 0
    
    ' Safely trigger Research key unbinding if that module exists in this edition
    On Error Resume Next
    Application.Run "BindUnbindResearch", 0
    On Error GoTo 0
End Sub

Public Sub BindUnbindLLM(ByVal Action As Long)
    Dim bindings As Variant
    Dim i As Long
    Dim keyCode As Long
    Dim macroName As String
    Dim kb As KeyBinding
    Dim mod1 As Long, mod2 As Long
    
    #If Mac Then
        bindings = Array( _
            Array(BuildKeyCode(wdKeyControl, wdKeyOption, wdKeyR), "Ribbon_Run"), _
            Array(BuildKeyCode(wdKeyControl, wdKeyOption, wdKeyE), "Ribbon_Expand"), _
            Array(BuildKeyCode(wdKeyControl, wdKeyOption, wdKeyS), "Ribbon_Summarize"), _
            Array(BuildKeyCode(wdKeyControl, wdKeyOption, wdKeyF), "Ribbon_FixGrammar"), _
            Array(BuildKeyCode(wdKeyControl, wdKeyOption, wdKeyT), "Ribbon_Translate"), _
            Array(BuildKeyCode(wdKeyControl, wdKeyOption, wdKeyD), "Ribbon_Define"), _
            Array(BuildKeyCode(wdKeyControl, wdKeyOption, wdKeyC), "Ribbon_Custom") _
        )
    #Else
        bindings = Array( _
            Array(BuildKeyCode(wdKeyAlt, wdKeyR), "Ribbon_Run"), _
            Array(BuildKeyCode(wdKeyAlt, wdKeyE), "Ribbon_Expand"), _
            Array(BuildKeyCode(wdKeyAlt, wdKeyS), "Ribbon_Summarize"), _
            Array(BuildKeyCode(wdKeyAlt, wdKeyF), "Ribbon_FixGrammar"), _
            Array(BuildKeyCode(wdKeyAlt, wdKeyT), "Ribbon_Translate"), _
            Array(BuildKeyCode(wdKeyAlt, wdKeyD), "Ribbon_Define"), _
            Array(BuildKeyCode(wdKeyAlt, wdKeyC), "Ribbon_Custom") _
        )
    #End If
    
    On Error Resume Next
    
    If Action = 1 Then
        CustomizationContext = ThisDocument
    End If
    
    For i = LBound(bindings) To UBound(bindings)
        keyCode = CLng(bindings(i)(0))
        macroName = CStr(bindings(i)(1))
        
        Set kb = FindKey(keyCode)
        kb.Clear
        
        If Action = 1 Then
            KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                            Command:=macroName, _
                            KeyCode:=keyCode
        End If
    Next i
    
    ThisDocument.Saved = True
    On Error GoTo 0
End Sub

' ============================================
' RIBBON CALLBACKS (for button clicks)
' ============================================

Public Sub RibbonProxy_StartServer(control As IRibbonControl)
    Ribbon_StartServer
End Sub

Public Sub RibbonProxy_AICommand(control As IRibbonControl)
    Select Case control.ID
        Case "btnRun"
            Ribbon_Run
        Case "btnExpand"
            Ribbon_Expand
        Case "btnSummarize"
            Ribbon_Summarize
        Case "btnFixGrammar"
            Ribbon_FixGrammar
        Case "btnTranslate"
            Ribbon_Translate
        Case "btnDefine"
            Ribbon_Define
        Case "btnCustom"
            Ribbon_Custom
    End Select
End Sub

' ============================================
' CORE RIBBON BUTTON FUNCTIONS
' ============================================

Public Sub Ribbon_StartServer()
    Dim pythonPath As String
    Dim scriptPath As String
    
    pythonPath = "{{PYTHON_EXE_PATH}}"
    scriptPath = "{{PYTHON_SERVER_PATH}}"
    
    #If Mac Then
        ' Replace backslashes with POSIX forward slashes for macOS paths
        pythonPath = Replace(pythonPath, "\", "/")
        scriptPath = Replace(scriptPath, "\", "/")
        
        ' Build raw shell command to run Python server in the background
        Dim cmd As String
        cmd = "'" & Replace(pythonPath, "'", "'\''") & "' '" & Replace(scriptPath, "'", "'\''") & "' > /dev/null 2>&1 &"
        
        On Error Resume Next
        ' Execute AppleScript via AppleScriptTask helper for Office 2016+ compatibility
        AppleScriptTask "WordbotCurl.scpt", "doShellCurl", cmd
        On Error GoTo 0
    #Else
        ' Windows execution via Standard Shell
        Shell pythonPath & " """ & scriptPath & """", vbNormalFocus
    #End If
End Sub

Public Sub Ribbon_Run()
    Dim txt As String
    txt = GetValidatedText("run")
    If txt <> "" Then wordbot "Execute the user's request while following the system prompt", txt
End Sub

Public Sub Ribbon_Expand()
    Dim txt As String
    txt = GetValidatedText("expand")
    If txt <> "" Then wordbot "Expand the selected text with relevant details, examples, and elaboration. Add depth and clarity. Do not add headings, introductory or concluding remarks. Begin directly with the expanded content. Use lists, tables, or equations if they help explain the content.", txt
End Sub

Public Sub Ribbon_Summarize()
    Dim txt As String
    txt = GetValidatedText("summarize")
    If txt <> "" Then wordbot "Summarize the selected text concisely, preserving key information. Begin directly with the summary. No headings, no lists.", txt
End Sub

Public Sub Ribbon_FixGrammar()
    Dim txt As String
    txt = GetValidatedText("fix grammar")
    If txt <> "" Then wordbot "Fix grammar and spelling errors only. Preserve all markdown formatting, LaTeX equations, and original structure. Do not rewrite, rephrase, or add explanations. Return only the corrected text.", txt
End Sub

Public Sub Ribbon_Translate()
    Dim txt As String
    txt = GetValidatedText("translate")
    If txt <> "" Then wordbot "Translate the selected text to English. Preserve original meaning and tone. Do not add phrases like 'Here is the translation'. Begin directly with the translation.", txt
End Sub

Public Sub Ribbon_Define()
    Dim txt As String
    txt = GetValidatedText("define")
    If txt <> "" Then wordbot "Explain the selected text. For a single word: **word** - definition, then brief explanation and interesting fact. For a phrase or sentence: provide a clear breakdown of what it means, simple examples, and an interesting fact if possible. Keep it concise. No extra formatting.", txt
End Sub

Public Sub Ribbon_Custom()
    Dim prompt As String
    Dim txt As String
    
    txt = Selection.text
    
    ' 1. Use a temporary Variant to catch the InputBox result
    Dim tempInput As Variant
    tempInput = InputBox("What should the AI do? Enter custom instruction or just press OK", "Custom AI Task")
    
    ' 2. Check if the user pressed Cancel (StrPtr is 0 for Cancel)
    If StrPtr(tempInput) = 0 Then
        Exit Sub
    End If
    
    ' 3. Now safely assign the string
    prompt = CStr(tempInput)
    
    ' 4. Process the logic
    If Len(prompt) > 0 Then
        wordbot prompt, txt
    Else
        wordbot "answer", txt
    End If
End Sub