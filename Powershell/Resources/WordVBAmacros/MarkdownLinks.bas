Public Sub FormatMarkdownLinks(targetRange As Range)
    Dim linkRange As Range
    Dim displayText As String
    Dim urlAddress As String
    Dim startPos As Long, midPos As Long, endPos As Long
    Dim safetyCounter As Long
    
    Set linkRange = targetRange.Duplicate
    safetyCounter = 0
    
    Do
        DoEvents ' Yields control to Word/Windows on every single link evaluation
        
        safetyCounter = safetyCounter + 1
        If safetyCounter > 500 Then
            MsgBox "Link formatting loop exceeded safety limit and was stopped.", vbExclamation
            Exit Do
        End If
        
        With linkRange.Find
            .ClearFormatting
            .text = "\[[!\]]@\]\([!\(]@\)"
            .MatchWildcards = True
            .Forward = True
            .Wrap = wdFindStop
            
            If Not .Execute Then Exit Do
        End With
        
        ' Check if match is inside code block
        If IsInsideCodeBlock(linkRange) Then
            ' Step pointer forward by 1 character so .Find moves past the current match
            linkRange.Start = linkRange.Start + 1
            If linkRange.Start >= targetRange.End Then Exit Do
            linkRange.End = targetRange.End
            GoTo ContinueLoop
        End If
        
        startPos = InStr(linkRange.text, "[")
        midPos = InStr(linkRange.text, "](")
        endPos = InStrRev(linkRange.text, ")")
        
        If startPos > 0 And midPos > startPos And endPos > midPos Then
            displayText = Mid(linkRange.text, startPos + 1, midPos - startPos - 1)
            urlAddress = Mid(linkRange.text, midPos + 2, endPos - midPos - 2)
            
            linkRange.text = displayText
            ActiveDocument.Hyperlinks.Add Anchor:=linkRange, Address:=urlAddress, TextToDisplay:=displayText
        End If
        
        ' Collapse and advance search pointer
        linkRange.Collapse Direction:=wdCollapseEnd
        If linkRange.Start >= targetRange.End Then Exit Do
        linkRange.End = targetRange.End
        
ContinueLoop:
    Loop While linkRange.Start < targetRange.End
End Sub