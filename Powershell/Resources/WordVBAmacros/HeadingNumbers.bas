Public Sub RibbonProxy_HeadingNumber(control As IRibbonControl)
    Ribbon_HeadingNumber
End Sub

Public Sub Ribbon_HeadingNumber()
    Dim objUndo As UndoRecord
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Apply Heading Numbers")

    Dim lt As ListTemplate
    Dim i As Integer, j As Integer
    Dim formatStr As String

    ' 1. Get or Create the Template
    ' Change: Don't just look for it; ensure it is set up correctly every time.
    On Error Resume Next
    Set lt = ActiveDocument.ListTemplates("HeadingNumberTemplate")
    On Error GoTo 0
    
    If lt Is Nothing Then
        Set lt = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True, Name:="HeadingNumberTemplate")
    End If

    ' 2. Forced Re-Linking
    ' Even if it existed, we re-apply the link to the style in case a
    ' "Reset Style" command wiped the association.
    For i = 1 To 9
        ' ActiveDocument.Styles("Heading " & i).ParagraphFormat.Reset
        formatStr = "%1"
        For j = 2 To i
            formatStr = formatStr & ".%" & j
        Next j
        
        With lt.ListLevels(i)
            .NumberFormat = formatStr
            .NumberStyle = wdListNumberStyleArabic
            ' This is the critical line that gets wiped by "Reset"
            .LinkedStyle = ActiveDocument.Styles("Heading " & i).NameLocal
            .Alignment = wdListLevelAlignLeft
            .TextPosition = InchesToPoints(0) ' Where text wraps to
            .NumberPosition = InchesToPoints(0)  ' Where number sits
            .TrailingCharacter = wdTrailingSpace
        End With
        
        With ActiveDocument.Styles("Heading " & i).ParagraphFormat
            .LeftIndent = InchesToPoints(0)
            .FirstLineIndent = InchesToPoints(0)
        End With
    Next i

    objUndo.EndCustomRecord
End Sub