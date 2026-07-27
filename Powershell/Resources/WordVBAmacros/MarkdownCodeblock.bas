Public Sub ConvertCodeBlocks(targetRange As Range)
    Dim foundRange As Range
    Dim codeContent As String
    Dim langPos As Long
    Dim endOfFirstLine As Long
    Dim para As Paragraph
    Dim searchPatterns(1) As String
    Dim i As Integer
    
    ' Match triple backticks with optional language specifier
    searchPatterns(0) = "```[!^13]@^13(*[!`]@)```"
    searchPatterns(1) = "```^13(*[!`]@)```"
    
    For i = 0 To 1
        Set foundRange = targetRange.Duplicate
        
        With foundRange.Find
            .ClearFormatting
            .text = searchPatterns(i)
            .MatchWildcards = True
            .Forward = True
            .Wrap = wdFindStop
            
            While .Execute
                ' Extract the code content (everything between the backticks)
                codeContent = foundRange.text
                
                ' Remove the opening triple backticks and language specifier
                ' Find the end of the first line (after the language specifier)
                endOfFirstLine = InStr(codeContent, vbCr)
                If endOfFirstLine = 0 Then endOfFirstLine = InStr(codeContent, vbLf)
                
                If endOfFirstLine > 0 Then
                    ' Remove the first line (```python or ```)
                    codeContent = Mid(codeContent, endOfFirstLine + 1)
                Else
                    ' Fallback: just remove the first 3 characters
                    codeContent = Mid(codeContent, 4)
                End If
                
                ' Remove the trailing triple backticks
                If Right(codeContent, 3) = "```" Then
                    codeContent = Left(codeContent, Len(codeContent) - 3)
                End If
                
                ' Clean up: remove any leading/trailing blank lines
                codeContent = Trim(codeContent)
                
                ' Replace the found range with the cleaned code
                foundRange.text = codeContent
                
                ' Apply code formatting
                With foundRange.Font
                    .Name = "Consolas"
                    .Size = 10
                End With
                
                ' Apply shading and borders to each paragraph in the code block
                For Each para In foundRange.Paragraphs
                    With para
                        ' Add a light gray background
                        .Shading.BackgroundPatternColor = wdColorGray10
                        ' Add a thin border
                        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                        ' Single line spacing for code
                        .LineSpacingRule = wdLineSpaceSingle
                        ' Add some padding
                        .LeftIndent = InchesToPoints(0.1)
                        .RightIndent = InchesToPoints(0.1)
                        .SpaceBefore = 3
                        .SpaceAfter = 3
                    End With
                Next para
                
                ' Collapse to continue searching
                foundRange.Collapse wdCollapseEnd
                foundRange.End = targetRange.End
            Wend
        End With
    Next i
End Sub
