Public Sub wordbot(taskType As String, selectedText As String)
    Dim body As String
    Dim objUndo As UndoRecord
    Dim targetRange As Range
    Dim formattedRange As Range
    Dim responseText As String
    Dim startRange As Long
    Dim rawJSON As String
    Dim response As String
    
    Set objUndo = Application.UndoRecord
    body = "{""text"": """ & EscapeJSON(selectedText) & """, ""task"": """ & EscapeJSON(taskType) & """}"
    
    ' Store the current cursor position BEFORE sending
    Set targetRange = Selection.Range.Duplicate
    
    ' Send request
    response = PythonTask("/process", body)
    If response = "" Then Exit Sub
    
    rawJSON = response
    
    ' CHECKPOINT 1: Paste Raw Server JSON String
    objUndo.StartCustomRecord ("1. Raw LLM Response")
    
    targetRange.Collapse Direction:=wdCollapseEnd
    targetRange.Select
    Selection.TypeText vbCr & vbCr
    startRange = Selection.End
    
    Set formattedRange = ActiveDocument.Range(startRange, startRange)
    formattedRange.text = rawJSON
    formattedRange.End = startRange + Len(rawJSON)
    
    objUndo.EndCustomRecord
    
    ' CHECKPOINT 2: Extract Content from JSON & Overwrite Range
    objUndo.StartCustomRecord ("2. JSON to Markdown")
    
    responseText = ExtractContentFromJSON(formattedRange.text, "result")
    formattedRange.text = responseText
    formattedRange.End = startRange + Len(responseText)
    
    objUndo.EndCustomRecord
    
    ' CHECKPOINT 3: Run Layout & Markdown Structural Styling
    ' FormatMarkdownAll has its own undo record
    Application.ScreenUpdating = False
    FormatMarkdownAll formattedRange
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Done!"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub