Public Sub ConvertMarkdownTables(targetRange As Range)
    Dim para As Paragraph
    Dim inTable As Boolean
    Dim tableLines As Collection
    Dim lineRange As Range
    Dim lineText As String
    Dim searchRange As Range
    
    Set searchRange = targetRange.Duplicate
    inTable = False
    
    On Error GoTo TableError
    
    For Each para In searchRange.Paragraphs
    
        ' Guard codeblock
        If IsInsideCodeBlock(para.Range) Then
            ' If we were in a table, close it before skipping the code block line
            If inTable Then
                inTable = False
                If Not tableLines Is Nothing Then
                    If tableLines.count >= 2 Then Call CreateTableFromCollection(tableLines, lineRange)
                    Set tableLines = Nothing
                End If
            End If
            GoTo NextParagraph
        End If
        
        ' Get the paragraph text without the trailing paragraph mark
        lineText = Trim(para.Range.text)
        If Right(lineText, 1) = vbCr Then
            lineText = Left(lineText, Len(lineText) - 1)
        End If
        
        ' Skip empty lines
        If Len(lineText) = 0 Then
            If inTable Then
                ' Empty line ends the table
                inTable = False
                If Not tableLines Is Nothing Then
                    If tableLines.count >= 2 Then
                        Call CreateTableFromCollection(tableLines, lineRange)
                    End If
                    Set tableLines = Nothing
                End If
            End If
            GoTo NextParagraph
        End If
        
        ' Check if this looks like a markdown table row
        If IsMarkdownTableRow(lineText) Then
            ' Clean the line (remove bold/italic markers that might interfere)
            lineText = CleanTableLine(lineText)
            
            If Not inTable Then
                ' Start of a new table
                inTable = True
                Set tableLines = New Collection
                Set lineRange = para.Range.Duplicate
                lineRange.Collapse wdCollapseStart
            End If
            
            ' Add this line to our table collection
            On Error Resume Next
            tableLines.Add lineText
            On Error GoTo TableError
        Else
            If inTable Then
                ' End of table - process it
                inTable = False
                If Not tableLines Is Nothing Then
                    If tableLines.count >= 2 Then
                        Call CreateTableFromCollection(tableLines, lineRange)
                    End If
                    Set tableLines = Nothing
                End If
            End If
        End If
        
NextParagraph:
    Next para
    
    ' Handle table at the end of document
    If inTable And Not tableLines Is Nothing Then
        If tableLines.count >= 2 Then
            Call CreateTableFromCollection(tableLines, lineRange)
        End If
        Set tableLines = Nothing
    End If
    
    Exit Sub
    
TableError:
    ' If there's an error, just skip this table and continue
    On Error Resume Next
    inTable = False
    Set tableLines = Nothing
    Resume Next
End Sub

Private Function IsMarkdownTableRow(lineText As String) As Boolean
    ' A valid markdown table row MUST:
    ' 1. Contain at least one pipe character
    ' 2. Have pipes in a table-like pattern
    
    Dim trimmed As String
    Dim pipeCount As Long
    Dim firstPipePos As Long
    Dim lastPipePos As Long
    
    trimmed = Trim(lineText)
    
    ' Must contain a pipe
    If InStr(trimmed, "|") = 0 Then
        IsMarkdownTableRow = False
        Exit Function
    End If
    
    ' Count pipes
    pipeCount = Len(trimmed) - Len(Replace(trimmed, "|", ""))
    
    ' Table rows typically have at least 2 pipes (for 2+ columns)
    If pipeCount < 2 Then
        IsMarkdownTableRow = False
        Exit Function
    End If
    
    ' Find first and last pipe positions
    firstPipePos = InStr(trimmed, "|")
    lastPipePos = InStrRev(trimmed, "|")
    
    ' Check if it's a separator row (contains dashes and colons)
    If InStr(trimmed, "---") > 0 Or InStr(trimmed, ":--") > 0 Or InStr(trimmed, "--:") > 0 Then
        ' It's a separator row - valid for tables
        IsMarkdownTableRow = True
        Exit Function
    End If
    
    ' For data rows, pipe should be at or near the beginning (within first 3 chars)
    ' OR the line should have a pipe structure like "| cell | cell |"
    If firstPipePos <= 3 Or (Left(trimmed, 1) = "|") Then
        IsMarkdownTableRow = True
    Else
        IsMarkdownTableRow = False
    End If
End Function

Private Sub CreateTableFromCollection(tableLines As Collection, insertionPoint As Range)
    On Error GoTo CreateError
    
    If tableLines Is Nothing Then Exit Sub
    If tableLines.count < 2 Then Exit Sub
    
    Dim dataRows() As String
    Dim colCount As Long
    Dim i As Long, j As Long
    Dim rowIndex As Long
    Dim lineText As String
    Dim parts() As String
    Dim tbl As Table
    Dim startRange As Range
    Dim endRange As Range
    
    ' First pass: separate data rows from separator rows
    ReDim dataRows(0)
    rowIndex = 0
    colCount = 0
    
    For i = 1 To tableLines.count
        lineText = Trim(tableLines(i))
        
        If IsSeparatorRowFast(lineText) Then
            ' This is the separator row - use it to determine column count
            colCount = CountColumnsFast(lineText)
        Else
            ' This is a data row
            ReDim Preserve dataRows(rowIndex)
            dataRows(rowIndex) = lineText
            rowIndex = rowIndex + 1
        End If
    Next i
    
    ' If we didn't get column count from separator, try first data row
    If colCount < 2 And rowIndex > 0 Then
        colCount = CountColumnsFast(dataRows(0))
    End If
    
    ' Validate we have a valid table
    If colCount < 2 Or rowIndex < 1 Then
        Exit Sub
    End If
    
    ' Get the range that contains the entire table
    Set startRange = insertionPoint.Duplicate
    startRange.Collapse wdCollapseStart
    
    ' Find the end of the table (last paragraph in the collection)
    Set endRange = startRange.Duplicate
    For i = 1 To tableLines.count
        On Error Resume Next
        endRange.MoveEnd wdParagraph, 1
        On Error GoTo CreateError
    Next i
    
    ' Delete the markdown table text
    On Error Resume Next
    startRange.End = endRange.End
    startRange.Delete
    On Error GoTo CreateError
    
    ' Create the table at the insertion point
    Set tbl = ActiveDocument.Tables.Add(Range:=startRange, NumRows:=rowIndex, NumColumns:=colCount)
    
    ' Fill the table with error handling for each cell
    For i = 0 To rowIndex - 1
        ' Parse the row - split by pipe
        parts = Split(dataRows(i), "|")
        
        For j = 0 To colCount - 1
            Dim cellContent As String
            On Error Resume Next
            If j + 1 <= UBound(parts) Then
                cellContent = Trim(parts(j + 1))
                ' Clean any remaining markdown from cell content
                cellContent = CleanTableCell(cellContent)
            Else
                cellContent = ""
            End If
            
            ' Safely set cell content
            If i + 1 <= tbl.Rows.count And j + 1 <= tbl.Columns.count Then
                tbl.Cell(i + 1, j + 1).Range.text = cellContent
            End If
            On Error GoTo CreateError
        Next j
    Next i
    
    ' Format the table with error handling
    On Error Resume Next
    With tbl
        .AutoFormat Format:=wdTableFormatGrid1, ApplyBorders:=True, ApplyFont:=True, ApplyColor:=False
        If rowIndex >= 1 Then
            .Rows(1).HeadingFormat = True
            .Rows(1).Range.Bold = True
        End If
        ' Add spacing after the table
        .Range.Select
        'Selection.Collapse Direction:=wdCollapseEnd
        'Selection.TypeText vbCrLf
        ' Add spacing after the table using Range instead of Selection
        Dim tableRange As Range
        Set tableRange = .Range
        tableRange.Collapse Direction:=wdCollapseEnd
        tableRange.InsertAfter vbCrLf
    End With
    On Error GoTo CreateError
    
    Exit Sub
    
CreateError:
    ' If table creation fails, just exit gracefully
    On Error Resume Next
    Debug.Print "Table creation error: " & Err.Description & " at line: " & Erl
    ' Don't leave any partial table
    If Not tbl Is Nothing Then
        tbl.Delete
    End If
End Sub

Private Function CleanTableLine(line As String) As String
    Dim result As String
    result = line
    
    ' Remove bold markers but keep content
    result = Replace(result, "**", "")
    
    ' Remove italic markers
    result = Replace(result, "*", "")
    
    ' Remove inline code markers
    result = Replace(result, "`", "")
    
    CleanTableLine = result
End Function

Private Function IsSeparatorRowFast(lineText As String) As Boolean
    Dim trimmed As String
    Dim temp As String
    Dim i As Long
    Dim ch As String
    
    trimmed = Trim(lineText)
    
    ' Must contain pipe and dash
    If InStr(trimmed, "|") = 0 Or InStr(trimmed, "-") = 0 Then
        IsSeparatorRowFast = False
        Exit Function
    End If
    
    ' Remove spaces and pipes for checking
    temp = Replace(trimmed, " ", "")
    temp = Replace(temp, "|", "")
    
    ' If nothing left after removing pipes, it's just pipes (invalid separator)
    If Len(temp) = 0 Then
        IsSeparatorRowFast = False
        Exit Function
    End If
    
    ' Check if remaining characters are only dashes and colons
    For i = 1 To Len(temp)
        ch = Mid(temp, i, 1)
        If ch <> "-" And ch <> ":" Then
            IsSeparatorRowFast = False
            Exit Function
        End If
    Next i
    
    IsSeparatorRowFast = (Len(temp) > 0)
End Function

Private Function CountColumnsFast(lineText As String) As Long
    Dim trimmed As String
    Dim pipeCount As Long
    
    trimmed = Trim(lineText)
    pipeCount = Len(trimmed) - Len(Replace(trimmed, "|", ""))
    
    If pipeCount >= 2 Then
        CountColumnsFast = pipeCount - 1
    Else
        CountColumnsFast = 0
    End If
End Function

Private Function CleanTableCell(text As String) As String
    Static regEx As Object
    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.pattern = "\[([^\]]*)\]\([^\)]*\)"
    End If
    Dim result As String
    result = Replace(Replace(Replace(text, "**", ""), "*", ""), "`", "")
    On Error Resume Next
    result = regEx.Replace(result, "$1")
    On Error GoTo 0
    CleanTableCell = Trim(result)
End Function
