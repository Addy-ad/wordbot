Public Sub FormatMarkdownAll(targetRange As Range)
    Dim objUndo As UndoRecord
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Format Markdown")
    On Error GoTo Cleanup
    
    If targetRange Is Nothing Then Exit Sub
    If targetRange.Start >= targetRange.End Then Exit Sub
    
    ' Create a in-memory duplicate range to hold execution scope
    Dim memRange As Range
    Set memRange = targetRange.Duplicate
    
    ' Execute all steps directly on the isolated memory range
    Application.StatusBar = "Step 1/10: Converting code blocks..."
    DoEvents
    ConvertCodeBlocks memRange
    
    Application.StatusBar = "Step 2/10: Converting markdown tables..."
    DoEvents
    ConvertMarkdownTables memRange
    
    Application.StatusBar = "Step 3/10: Converting markdown lists..."
    DoEvents
    ConvertMarkdownLists memRange
    
    Application.StatusBar = "Step 4/10: Converting LaTeX to Word math..."
    DoEvents
    ConvertLaTeXToWordMath memRange
    
    Application.StatusBar = "Step 5/10: Applying H4 headings (####)..."
    DoEvents
    ApplyHeaderStyle memRange, "####", wdStyleHeading3
    
    Application.StatusBar = "Step 6/10: Applying H3 headings (###)..."
    DoEvents
    ApplyHeaderStyle memRange, "###", wdStyleHeading2
    
    Application.StatusBar = "Step 7/10: Applying H2 headings (##)..."
    DoEvents
    ApplyHeaderStyle memRange, "##", wdStyleHeading1
    
    Application.StatusBar = "Step 8/10: Applying H1 headings (#)..."
    DoEvents
    ApplyHeaderStyle memRange, "#", wdStyleTitle
    
    Application.StatusBar = "Step 9/10: Applying bold, italic, and code formatting..."
    DoEvents
    ApplyTextStyle memRange
    
    Application.StatusBar = "Step 10/10: Converting links and final cleanup..."
    DoEvents
    FormatMarkdownLinks memRange
    MarkdownFinalCleanup memRange
    
    ' Re-select strictly the updated in-memory range
    memRange.Select
    
    Application.StatusBar = "Markdown formatting complete!"

Cleanup:
    objUndo.EndCustomRecord
    If Err.Number <> 0 Then
        MsgBox "Error during formatting: " & Err.Description
    End If
End Sub
