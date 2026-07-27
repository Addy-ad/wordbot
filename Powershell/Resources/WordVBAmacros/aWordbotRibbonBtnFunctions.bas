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
        
        ' Escape double quotes for MacScript / AppleScript execution
        Dim cmd As String
        cmd = "do shell script """ & Replace(pythonPath, """", "\""") & " " & Replace(scriptPath, """", "\""") & " > /dev/null 2>&1 &"""
        
        On Error Resume Next
        MacScript cmd
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

Public Sub Ribbon_wordbotResearch()
    Dim txt As String
    txt = GetValidatedText("Research")
    If txt <> "" Then wordbotResearch txt
End Sub

Public Sub Ribbon_FormatMarkdown()
    Dim targetRng As Range
    Set targetRng = GetValidatedText("format markdown", "range")
    
    If Not targetRng Is Nothing Then
        FormatMarkdownAll targetRng
    End If
End Sub

Public Sub Ribbon_ConvertSelectionToFields()
    Dim targetRng As Range
    Set targetRng = GetValidatedText("convert citations", "range")
    
    If Not targetRng Is Nothing Then
        ConvertSelectionToFields_callback targetRng
    End If
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

Public Sub Ribbon_HeadingNumber()
    Dim objUndo As UndoRecord
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Apply Heading Numbers")

    Dim lt As ListTemplate
    Dim i As Integer, j As Integer
    Dim formatStr As String

    ' 1. Get or Create the Template
    ' Change: Don't just look for it; ensure it is set up correctly every time.
    On Error Resume Next
    Set lt = ActiveDocument.ListTemplates("HeadingNumberTemplate")
    On Error GoTo 0
    
    If lt Is Nothing Then
        Set lt = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True, Name:="HeadingNumberTemplate")
    End If

    ' 2. Forced Re-Linking
    ' Even if it existed, we re-apply the link to the style in case a
    ' "Reset Style" command wiped the association.
    For i = 1 To 9
        ' ActiveDocument.Styles("Heading " & i).ParagraphFormat.Reset
        formatStr = "%1"
        For j = 2 To i
            formatStr = formatStr & ".%" & j
        Next j
        
        With lt.ListLevels(i)
            .NumberFormat = formatStr
            .NumberStyle = wdListNumberStyleArabic
            ' This is the critical line that gets wiped by "Reset"
            .LinkedStyle = ActiveDocument.Styles("Heading " & i).NameLocal
            .Alignment = wdListLevelAlignLeft
            .TextPosition = InchesToPoints(0) ' Where text wraps to
            .NumberPosition = InchesToPoints(0)  ' Where number sits
            .TrailingCharacter = wdTrailingSpace
        End With
        
        With ActiveDocument.Styles("Heading " & i).ParagraphFormat
            .LeftIndent = InchesToPoints(0)
            .FirstLineIndent = InchesToPoints(0)
        End With
    Next i

    objUndo.EndCustomRecord
End Sub

' ============================================
' Helper functions
' ============================================

Public Function GetValidatedText(Optional actionVerb As String = "process", Optional returnType As String = "string") As Variant
    ' 1. Check if selection is valid initial selection
    If Selection.Type = wdSelectionIP Or Len(Trim(Selection.text)) <= 1 Then
        MsgBox "Please select some text first to " & actionVerb & "!", vbExclamation, "No Selection"
        
        If LCase(returnType) = "range" Then
            Set GetValidatedText = Nothing
        Else
            GetValidatedText = ""
        End If
        Exit Function
    End If
    
    ' 2. Handle Range Return Type
    If LCase(returnType) = "range" Then
        Set GetValidatedText = Selection.Range.Duplicate
        Exit Function
    End If
    
    ' 3. Handle String Return Type (Validate after stripping fields)
    Dim stripped As String
    stripped = StripFields(Trim(Selection.text))
    
    If Len(stripped) <= 1 Then
        MsgBox "Selection contains only fields, no plain text to " & actionVerb & "!", vbExclamation, "No Valid Text"
        GetValidatedText = ""
        Exit Function
    End If
    
    GetValidatedText = stripped
End Function

Public Function StripFields(ByVal txt As String) As String
    ' If you need to strip based on the actual selection's fields:
    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    
    If rng.Fields.Count > 0 Then
        Dim fld As Field
        For Each fld In rng.Fields
            txt = Replace(txt, fld.Result.text, "")
        Next fld
    End If
    
    StripFields = txt
End Function