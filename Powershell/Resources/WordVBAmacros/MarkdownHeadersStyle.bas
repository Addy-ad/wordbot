Public Sub ApplyHeaderStyle(targetRange As Range, marker As String, builtInStyle As Long)
    Dim rng As Range
    Dim searchLimit As Long
    Dim styleRange As Range
    
    searchLimit = targetRange.End
    
    Set rng = targetRange.Duplicate
    With rng.Find
        .ClearFormatting
        .text = marker & " "
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            If rng.Start >= searchLimit Then Exit Do
            
            ' --- MODIFIED GUARD ---
            ' We check if the paragraph is inside a code block AND if it's in a table.
            ' If it's a code block, we ignore this match completely.
            If Not IsInsideCodeBlock(rng.Paragraphs(1).Range) Then
                If Not rng.Information(wdWithInTable) Then
                    
                    ' Clear only the markdown prefix text sequence safely
                    rng.Delete
                    
                    ' Reference the clean paragraph structure
                    Set styleRange = rng.Paragraphs(1).Range
                    
                    ' Safeguard style application
                    styleRange.style = ActiveDocument.Styles(builtInStyle)
                End If
            End If
            
            ' Shift pointer beyond the modified paragraph boundary
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub
