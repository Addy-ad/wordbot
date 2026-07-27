Public Sub ApplyTextStyle(targetRange As Range)
    Application.ScreenUpdating = False
    
    ' SAFE TEXT TOKENS: Plain text placeholders that do not collide with Word's wildcard engine
    Const MASK_BOLD_ITALIC_OPEN As String = "||MBI||"
    Const MASK_BOLD_ITALIC_CLOSE As String = "||/MBI||"
    Const MASK_BOLD_OPEN As String = "||MB||"
    Const MASK_BOLD_CLOSE As String = "||/MB||"
    Const MASK_ITALIC_OPEN As String = "||MI||"
    Const MASK_ITALIC_CLOSE As String = "||/MI||"
    Const MASK_CODE_OPEN As String = "||MC||"
    Const MASK_CODE_CLOSE As String = "||/MC||"
    
    ' Pass 1: Parse and format everything using masks (order matters - inner to outer)
    ApplyMarkdownFormat targetRange, "\`(*)\`", False, False, "Consolas", MASK_CODE_OPEN, MASK_CODE_CLOSE, 1
    ApplyMarkdownFormat targetRange, "\*(*)\*", False, True, "", MASK_ITALIC_OPEN, MASK_ITALIC_CLOSE, 1
    ApplyMarkdownFormat targetRange, "\*\*(*)\*\*", True, False, "", MASK_BOLD_OPEN, MASK_BOLD_CLOSE, 2
    ApplyMarkdownFormat targetRange, "\*\*\*(*)\*\*\*", True, True, "", MASK_BOLD_ITALIC_OPEN, MASK_BOLD_ITALIC_CLOSE, 3
    
    ' Pass 2: Clean up the safe structural tokens
    RemoveMarkdownMasks targetRange
    
    Application.ScreenUpdating = True
End Sub

Public Sub ApplyMarkdownFormat(targetRange As Range, pattern As String, isBold As Boolean, isItalic As Boolean, fontName As String, maskOpen As String, maskClose As String, delimiterLengthInput As Long)
    
    Dim rngSearch As Range
    Dim matchRange As Range
    Dim innerRange As Range
    Dim found As Boolean
    Dim delimiterLength As Long
    Dim firstInnerChar As String
    Dim lastInnerChar As String
    Dim delimRange As Range
    Dim absoluteSafetyPointer As Long
    
    Set rngSearch = targetRange.Duplicate
    
    ' Fix: Stripped away the brittle string matching lookup.
    ' We assign the delimiter size strictly based on the reliable input parameter.
    If delimiterLengthInput > 0 Then
        delimiterLength = delimiterLengthInput
    Else
        delimiterLength = 1
    End If
    
    With rngSearch.Find
        .ClearFormatting
        .text = pattern
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        
        found = .Execute
        Do While found
            
            Set matchRange = rngSearch.Duplicate
            
            ' Crash Guard 1: Instantly terminate loop if selection structures invert or break bounds
            If matchRange.Start < targetRange.Start Or matchRange.Start >= targetRange.End Or matchRange.Start >= matchRange.End Then
                Exit Do
            End If
            
            If Not IsInsideCodeBlock(matchRange) Then
                
                ' Flanking Spacing Validation Framework
                If matchRange.Characters.count > (delimiterLength * 2) Then
                    firstInnerChar = matchRange.Characters(delimiterLength + 1).text
                    lastInnerChar = matchRange.Characters(matchRange.Characters.count - delimiterLength).text
                    
                    If firstInnerChar = " " Or lastInnerChar = " " Or InStr(matchRange.text, vbCr) > 0 Then
                        rngSearch.Start = matchRange.Start + 1
                        GoTo NextIteration
                    End If
                Else
                    rngSearch.Start = matchRange.End
                    GoTo NextIteration
                End If
                
                Set innerRange = matchRange.Duplicate
                innerRange.Start = matchRange.Start + delimiterLength
                innerRange.End = matchRange.End - delimiterLength
                
                If innerRange.Start < innerRange.End And Len(innerRange.text) > 0 Then
                    
                    absoluteSafetyPointer = matchRange.Start
                    
                    With innerRange.Font
                        If isBold Then .Bold = True
                        If isItalic Then .Italic = True
                        If fontName <> "" Then .Name = fontName
                    End With
                    
                    If fontName = "Consolas" Then
                        On Error Resume Next
                        innerRange.Paragraphs(1).LineSpacingRule = wdLineSpaceSingle
                        On Error GoTo 0
                    End If
                    
                    ' Execute marker substitutions safely
                    Set delimRange = matchRange.Duplicate
                    delimRange.End = matchRange.Start + delimiterLength
                    delimRange.text = maskOpen
                    
                    innerRange.Start = matchRange.Start + Len(maskOpen)
                    innerRange.End = matchRange.End - delimiterLength
                    
                    Set delimRange = innerRange.Duplicate
                    delimRange.Start = innerRange.End
                    delimRange.End = innerRange.End + delimiterLength
                    delimRange.text = maskClose
                    
                    ' Crash Guard 2: Force scan track forward beyond modification shifts
                    If delimRange.End <= absoluteSafetyPointer Then
                        rngSearch.Start = absoluteSafetyPointer + Len(maskOpen) + Len(maskClose) + 1
                    Else
                        rngSearch.Start = delimRange.End
                    End If
                Else
                    rngSearch.Start = matchRange.End
                End If
            Else
                rngSearch.Start = matchRange.End
            End If
            
NextIteration:
            ' Crash Guard 3: Stop loop before exceeding execution boundaries
            If rngSearch.Start >= targetRange.End Then Exit Do
            
            rngSearch.End = targetRange.End
            found = .Execute
        Loop
    End With
End Sub

Public Sub RemoveMarkdownMasks(targetRange As Range)
    Dim tokens As Variant
    Dim token As Variant
    Dim searchRange As Range
    Dim found As Boolean
    
    tokens = Array("||MBI||", "||/MBI||", "||MB||", "||/MB||", "||MI||", "||/MI||", "||MC||", "||/MC||")
    
    On Error Resume Next
    
    For Each token In tokens
        ' Reset search range to the beginning of targetRange for each token
        Set searchRange = targetRange.Duplicate
        searchRange.Collapse Direction:=wdCollapseStart
        
        With searchRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = token
            .Replacement.text = ""
            .MatchWildcards = False
            .MatchCase = False
            .Forward = True
            .Wrap = wdFindStop  ' STOP at end of range - DO NOT continue to document
            .Format = False
            
            ' Find and replace within the targetRange only
            Do While .Execute
                ' Check if we're still within targetRange
                If searchRange.Start >= targetRange.Start And searchRange.End <= targetRange.End Then
                    ' Replace the found token with nothing
                    searchRange.text = ""
                    
                    ' Adjust searchRange to continue from the same position
                    searchRange.Collapse Direction:=wdCollapseEnd
                Else
                    Exit Do
                End If
                
                ' Safety check to prevent infinite loop
                If searchRange.Start >= targetRange.End Then Exit Do
            Loop
        End With
    Next token
    
    On Error GoTo 0
End Sub

Private Function HasInvalidSpacing(rng As Range, delimiterLength As Long) As Boolean
    Dim fullText As String
    Dim firstInnerChar As String
    Dim lastInnerChar As String
    
    If rng.Characters.count <= delimiterLength * 2 Then
        HasInvalidSpacing = True
        Exit Function
    End If
    
    fullText = rng.text
    firstInnerChar = Mid(fullText, delimiterLength + 1, 1)
    lastInnerChar = Mid(fullText, Len(fullText) - delimiterLength, 1)
    
    ' Check for spaces or line breaks adjacent to delimiters
    HasInvalidSpacing = (firstInnerChar = " " Or lastInnerChar = " " Or InStr(fullText, vbCr) > 0)
End Function
