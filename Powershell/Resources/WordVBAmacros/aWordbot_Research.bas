Public Sub wordbotResearch(selectedText As String)
    Dim txt As String
    Dim body As String
    Dim response As String
    Dim objUndo As UndoRecord
    
    ' Configuration Variables
    Dim topK As String
    Dim mode As String
    Dim minSimilarity As String
    
    topK = "4"
    mode = "hybrid"
    minSimilarity = "0.3"
    
    ' 2. Construct JSON body
    body = "{""text"": """ & EscapeJSON(selectedText) & """, " & _
           """topK"": """ & topK & """, " & _
           """mode"": """ & mode & """, " & _
           """minSimilarity"": """ & minSimilarity & """}"
    
    ' 3. Send request
    response = PythonTask("/research", body)
    
    ' 4. Insert response at cursor position
    If response <> "" Then
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.TypeParagraph
        
        Set objUndo = Application.UndoRecord
        objUndo.StartCustomRecord ("1. Raw Research Response")
        ' Define the start point BEFORE typing
        Dim startPos As Long
        startPos = Selection.End
        
        ' ACTUALLY TYPE THE TEXT
        Selection.TypeText response
        
        ' CRITICAL: Capture the end position AFTER typing
        Dim endPos As Long
        endPos = Selection.End
        
        objUndo.EndCustomRecord
        
        ' Define the range that now contains the response
        Dim docRange As Range
        Set docRange = ActiveDocument.Range(startPos, endPos)
        
        ' OPTIONAL: Select the range so the user can see what was processed
        docRange.Select
        
        ' STEP 1: Apply Markdown Formatting FIRST (on plain text)
        ' FormatMarkdownAll has its own undo record
        Application.ScreenUpdating = False
        FormatMarkdownAll docRange
        Application.ScreenUpdating = True
        
        ' Process only the newly inserted text
        Ribbon_ConvertSelectionToFields
        
        Application.StatusBar = "Research completed."

    End If
End Sub
