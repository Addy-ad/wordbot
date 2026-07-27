Public Sub ConvertMarkdownLists(targetRange As Range)
    Dim p As Paragraph
    Dim pText As String, cleanText As String
    Dim dotPos As Integer
    Dim rReplace As Range
    Dim i As Long
    
    ' Loop through all paragraphs in the target range
    For i = 1 To targetRange.Paragraphs.count
        Set p = targetRange.Paragraphs(i)
        
        ' 1. Guard Clause: Skip if it's a code block (has shading)
        If p.Shading.BackgroundPatternColor = wdColorGray10 Then
            ' Skip this paragraph, continue to the next one
            GoTo ContinueLoop
        End If
        
        pText = p.Range.text
        
        ' Remove trailing paragraph mark for comparison
        If Right(pText, 1) = vbCr Then pText = Left(pText, Len(pText) - 1)
        
        ' Clean text
        pText = Replace(pText, ChrW(160), " ")
        pText = Replace(pText, ChrW(183), "-")
        pText = Replace(pText, ChrW(149), "-")
        pText = Replace(pText, ChrW(8226), "-")
        pText = Replace(pText, vbTab, " ")
        pText = LTrim(pText)
        
        If Len(pText) > 0 Then
            ' BULLET CHECK: Starts with "-" or "*" and a space
            If (Left(pText, 1) = "-" Or Left(pText, 1) = "*") And Mid(pText, 2, 1) = " " Then
                cleanText = LTrim(Mid(pText, 3))
                
                Set rReplace = p.Range
                rReplace.End = rReplace.End - 1 ' Keep the vbCr intact
                rReplace.text = cleanText
                
                p.Range.ListFormat.RemoveNumbers
                p.Range.ListFormat.ApplyBulletDefault (wdStyleListBullet)
                
            ' Check for Number
            ElseIf pText Like "#.*" Or pText Like "##.*" Or pText Like "###.*" Then
                dotPos = InStr(pText, ".")
                If dotPos > 0 And Mid(pText, dotPos + 1, 1) = " " Then
                    cleanText = LTrim(Mid(pText, dotPos + 1))
                    
                    Set rReplace = p.Range
                    rReplace.End = rReplace.End - 1
                    rReplace.text = cleanText
                    
                    p.Range.ListFormat.RemoveNumbers
                    p.Range.ListFormat.ApplyNumberDefault (wdNumberListNum)
                End If
            End If
        End If

ContinueLoop:
    Next i
End Sub
