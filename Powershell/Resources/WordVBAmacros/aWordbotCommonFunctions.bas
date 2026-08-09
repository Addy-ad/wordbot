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