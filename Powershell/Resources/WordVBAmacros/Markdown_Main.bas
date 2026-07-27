Public Sub FormatMarkdownAll(targetRange As Range)
    Dim objUndo As UndoRecord
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Format Markdown")
    On Error GoTo Cleanup
    
    ' Store the original range boundaries
    Dim originalStart As Long
    Dim originalEnd As Long
    Dim doc As Document
    Set doc = targetRange.Document

    originalStart = targetRange.Start
    originalEnd = targetRange.End
    
    ' Process each step with fresh range
    Dim currentRange As Range
    
    Application.StatusBar = "Step 1/10: Converting code blocks..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ConvertCodeBlocks currentRange
    
    Application.StatusBar = "Step 2/10: Converting markdown tables..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ConvertMarkdownTables currentRange
    
    Application.StatusBar = "Step 3/10: Converting markdown lists..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ConvertMarkdownLists currentRange
    
    Application.StatusBar = "Step 4/10: Converting LaTeX to Word math..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ConvertLaTeXToWordMath currentRange
    
    Application.StatusBar = "Step 5/10: Applying H4 headings (####)..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ApplyHeaderStyle currentRange, "####", wdStyleHeading3
    
    Application.StatusBar = "Step 6/10: Applying H3 headings (###)..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ApplyHeaderStyle currentRange, "###", wdStyleHeading2
    
    Application.StatusBar = "Step 7/10: Applying H2 headings (##)..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ApplyHeaderStyle currentRange, "##", wdStyleHeading1
    
    Application.StatusBar = "Step 8/10: Applying H1 headings (#)..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ApplyHeaderStyle currentRange, "#", wdStyleTitle
    
    Application.StatusBar = "Step 9/10: Applying bold, italic, and code formatting..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then ApplyTextStyle currentRange
    
    Application.StatusBar = "Step 10/10: Converting links and final cleanup..."
    DoEvents
    Set currentRange = GetSafeRange(originalStart, originalEnd, doc)
    If Not currentRange Is Nothing Then FormatMarkdownLinks currentRange
    If Not currentRange Is Nothing Then MarkdownFinalCleanup currentRange
    
    Application.StatusBar = "Markdown formatting complete!"
    
Cleanup:
    objUndo.EndCustomRecord
    If Err.Number <> 0 Then
        MsgBox "Error at step: " & Err.Description
        Application.StatusBar = "Error at step " & Err.Description
    End If
End Sub

Private Function GetSafeRange(startPos As Long, endPos As Long, doc As Document) As Range
    On Error GoTo ErrorHandler
    
    ' Ensure boundaries are valid
    Dim docEnd As Long
    docEnd = doc.Range.End
    
    ' Clamp positions
    If startPos < 0 Then startPos = 0
    If endPos < 0 Then endPos = 0
    If startPos > docEnd Then startPos = docEnd
    If endPos > docEnd Then endPos = docEnd
    
    ' Swap if start > end
    If startPos > endPos Then
        Dim temp As Long
        temp = startPos
        startPos = endPos
        endPos = temp
    End If
    
    ' Return range only if valid
    If startPos < endPos Then
        Set GetSafeRange = doc.Range(startPos, endPos)
    Else
        Set GetSafeRange = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set GetSafeRange = Nothing
End Function

Public Function IsInsideCodeBlock(rng As Range) As Boolean
    ' Check the shading (The Atomic Guard)
    If rng.Shading.BackgroundPatternColor = wdColorGray10 Then
        IsInsideCodeBlock = True
        Exit Function
    End If
    
    ' Check font as fallback
    If rng.Font.Name = "Consolas" Then
        IsInsideCodeBlock = True
        Exit Function
    End If
    
    IsInsideCodeBlock = False
End Function
