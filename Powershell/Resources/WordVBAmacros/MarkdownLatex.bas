Public Sub ConvertLaTeXToWordMath(targetRange As Range)
    ' Process block equations ($$ ... $$) first
    ProcessMathNative targetRange, "$$", True
    
    ' Process inline equations ($ ... $) next
    ProcessMathNative targetRange, "$", False
End Sub

Private Sub ProcessMathNative(targetRange As Range, delimiter As String, isBlock As Boolean)
    Dim searchRng As Range
    Dim eqRng As Range
    Dim startPos As Long, endPos As Long
    Dim rawText As String, cleanText As String
    Dim mathObj As OMath
    
    ' Duplicate the range so we don't modify the user's original selection bounds
    Set searchRng = targetRange.Duplicate
    
    With searchRng.Find
        .ClearFormatting
        .text = delimiter
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        
        Do While .Execute()
            ' We found an opening delimiter. Mark the spot.
            startPos = searchRng.Start
            searchRng.Collapse wdCollapseEnd
            
            ' Now hunt for the closing delimiter
            If .Execute() Then
                endPos = searchRng.End
                
                ' Create a precise range spanning from start to end delimiters
                Set eqRng = targetRange.Document.Range(startPos, endPos)
                rawText = eqRng.text
                
                ' Strip the delimiters based on block or inline length
                If isBlock Then
                    cleanText = Mid(rawText, 3, Len(rawText) - 4)
                Else
                    cleanText = Mid(rawText, 2, Len(rawText) - 2)
                End If
                
                ' SAFETY GUARD: Inline equations should not span across paragraphs.
                ' If it does, the AI likely missed a closing $ sign. Skip it.
                If Not isBlock And InStr(cleanText, vbCr) > 0 Then
                    searchRng.Start = startPos + 1
                    If searchRng.Start < targetRange.End Then searchRng.End = targetRange.End
                    GoTo NextMatch
                End If
                
                ' Scrub hidden newlines and junk (vbCr = Paragraph, vbLf = Line Feed, Chr(11) = Soft Return)
                cleanText = Trim(cleanText)
                Do While Left(cleanText, 1) = vbCr Or Left(cleanText, 1) = vbLf Or Left(cleanText, 1) = Chr(11)
                    cleanText = Trim(Mid(cleanText, 2))
                Loop
                Do While Right(cleanText, 1) = vbCr Or Right(cleanText, 1) = vbLf Or Right(cleanText, 1) = Chr(11)
                    cleanText = Trim(Left(cleanText, Len(cleanText) - 1))
                Loop
                
                ' Only process if there is actual math to convert
                If Len(cleanText) > 0 Then
                    eqRng.text = cleanText
                    eqRng.OMaths.Add eqRng
                    eqRng.OMaths(1).BuildUp
                    
                    If isBlock Then
                        eqRng.Paragraphs(1).Alignment = wdAlignParagraphCenter
                    End If
                End If
                
                ' CRITICAL RE-ANCHOR: Move search range to immediately after the new equation
                ' This prevents infinite loops and ensures we don't scan the math we just built
                searchRng.Start = eqRng.End
                
                ' Ensure we haven't exceeded the document boundaries
                If searchRng.Start <= targetRange.End Then
                    searchRng.End = targetRange.End
                Else
                    Exit Do
                End If
            Else
                ' Found an opening delimiter but no closing one. Reached the end.
                Exit Do
            End If
NextMatch:
        Loop
    End With
End Sub
