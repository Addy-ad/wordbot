Public Sub DecodeUnicodeEscapes(targetRange As Range)
    Dim text As String
    Dim result As String
    Dim i As Long
    Dim charCode As Long
    
    text = targetRange.text
    result = text
    i = 1
    
    Do While i <= Len(result) - 5
        If Mid(result, i, 2) = "\u" Then
            On Error Resume Next
            charCode = CLng("&H" & Mid(result, i + 2, 4))
            On Error GoTo 0
            
            If charCode > 0 And charCode <= 65535 Then
                result = Left(result, i - 1) & ChrW(charCode) & Mid(result, i + 6)
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    
    If result <> text Then
        targetRange.text = result
    End If
End Sub
