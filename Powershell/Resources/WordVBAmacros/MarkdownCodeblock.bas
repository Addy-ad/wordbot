Public Sub ConvertCodeBlocks(targetRange As Range)
    Dim originalText As String
    Dim lines As Variant
    Dim resultLines As String
    Dim i As Long
    Dim inCodeBlock As Boolean
    Dim codeContent As String
    Dim para As Paragraph
    Dim codeRanges As Collection
    Dim rng As Range
    
    ' Turn off screen updating
    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo Cleanup
    
    ' Get the text
    originalText = targetRange.text
    
    ' Split into lines
    lines = Split(originalText, vbCr)
    inCodeBlock = False
    codeContent = ""
    resultLines = ""
    
    ' Process line by line
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        Dim trimmedLine As String
        
        line = lines(i)
        trimmedLine = Trim(line)
        
        If Not inCodeBlock And Left(trimmedLine, 3) = "```" Then
            ' Start of code block
            inCodeBlock = True
            codeContent = ""
            ' Don't add the opening line to output
        ElseIf inCodeBlock And trimmedLine = "```" Then
            ' End of code block
            inCodeBlock = False
            
            ' Add the code block with a unique marker
            If Len(codeContent) > 0 Then
                ' Replace newlines in code with a placeholder
                resultLines = resultLines & "%%CODEBLOCK%%" & codeContent & "%%ENDCODEBLOCK%%"
            End If
            ' Add a newline after the code block
            resultLines = resultLines & vbCr
        ElseIf inCodeBlock Then
            ' Inside code block
            codeContent = codeContent & line & vbCr
        Else
            ' Regular text
            resultLines = resultLines & line & vbCr
        End If
    Next i
    
    ' Handle unclosed code block
    If inCodeBlock And Len(codeContent) > 0 Then
        resultLines = resultLines & "%%CODEBLOCK%%" & codeContent & "%%ENDCODEBLOCK%%" & vbCr
    End If
    
    ' Replace the content
    targetRange.text = resultLines
    
    ' Now find and format code blocks
    Dim searchRange As Range
    Set searchRange = targetRange.Duplicate
    
    With searchRange.Find
        .ClearFormatting
        .text = "%%CODEBLOCK%%*%%ENDCODEBLOCK%%"
        .MatchWildcards = True
        .Wrap = wdFindStop
        
        While .Execute
            Dim codeBlockRange As Range
            Set codeBlockRange = searchRange.Duplicate
            
            ' Extract the code
            Dim codeText As String
            codeText = codeBlockRange.text
            codeText = Replace(codeText, "%%CODEBLOCK%%", "")
            codeText = Replace(codeText, "%%ENDCODEBLOCK%%", "")
            
            ' Replace with code
            codeBlockRange.text = codeText
            
            ' Format the code
            Dim formatRange As Range
            Set formatRange = codeBlockRange.Duplicate
            
            ' Expand to the paragraph
            formatRange.Expand Unit:=wdParagraph
            
            ' Apply formatting
            With formatRange.Font
                .Name = "Consolas"
                .Size = 10
            End With
            
            ' Apply shading and borders
            For Each para In formatRange.Paragraphs
                With para
                    .Shading.BackgroundPatternColor = wdColorGray10
                    .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    .LineSpacingRule = wdLineSpaceSingle
                    .LeftIndent = InchesToPoints(0.1)
                    .RightIndent = InchesToPoints(0.1)
                    .SpaceBefore = 3
                    .SpaceAfter = 3
                End With
            Next para
            
            ' Move to next
            searchRange.Collapse wdCollapseEnd
            searchRange.End = targetRange.End
        Wend
    End With

Cleanup:
    Application.ScreenUpdating = oldScreenUpdating
    Application.StatusBar = "Code blocks conversion complete"
End Sub