Public Sub MarkdownFinalCleanup(targetRange As Range)
    ' STEP 1: Remove horizontal rules
    With targetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .text = "^13---^13"
        .Replacement.text = "^13"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' STEP 2: Handle remaining rules with spaces
    With targetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .text = "[^13]@---[^13]@"
        .Replacement.text = "^13"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' STEP 3: Clean up multiple newlines
    With targetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^13^13+"
        .Replacement.text = "^13"
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    ' STEP 4: Remove leading newlines safely with DoEvents and length checks
    Dim safetyCounter As Long
    safetyCounter = 0
    
    Do While Len(targetRange.text) > 0 And (Left(targetRange.text, 1) = vbCr Or Left(targetRange.text, 1) = vbLf)
        DoEvents
        safetyCounter = safetyCounter + 1
        If safetyCounter > 100 Then Exit Do
        
        targetRange.Characters(1).Delete
    Loop
End Sub