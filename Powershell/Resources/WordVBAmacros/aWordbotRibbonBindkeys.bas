' Runs ONCE automatically when Word launches
Public Sub AutoExec()
    ' Bind shortcuts once for the entire Word session
    BindUnbindWrapper 1
End Sub

' Runs ONCE automatically when Word closes
Public Sub AutoExit()
    ' Clean up shortcuts on exit
    BindUnbindWrapper 0
End Sub

Public Sub BindUnbindWrapper(ByVal Action As Long)
    Dim dictBindings As New Dictionary
    Dim key As Variant
    Dim kb As KeyBinding
    Dim bindingArray As Variant
    
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyR), Array("Alt+R", "Ribbon_Run")
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyE), Array("Alt+E", "Ribbon_Expand")
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyS), Array("Alt+S", "Ribbon_Summarize")
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyF), Array("Alt+F", "Ribbon_FixGrammar")
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyT), Array("Alt+T", "Ribbon_Translate")
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyD), Array("Alt+D", "Ribbon_Define")
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyC), Array("Alt+C", "Ribbon_Custom")
    dictBindings.Add BuildKeyCode(wdKeyAlt, wdKeyZ), Array("Alt+Z", "Ribbon_wordbotResearch")
    
    On Error Resume Next
    
    If Action = 1 Then 
        CustomizationContext = ThisDocument
    End If
    
    For Each key In dictBindings.Keys
        Set kb = FindKey(CLng(key))
        
        ' ' --- 1. ALWAYS INSPECT / CLEAR FIRST ---
        ' If kb.Command <> "" Then
        '     ' Check if assigned to something OTHER than our WordBot command
        '     If kb.Command <> wbCommands(i) Then
        '         Select Case kb.KeyCategory
        '             Case wdKeyCategoryMacro: categoryName = "Macro"
        '             Case wdKeyCategoryCommand: categoryName = "Built-in Word Command"
        '             Case wdKeyCategoryFont: categoryName = "Font"
        '             Case wdKeyCategoryAutoText: categoryName = "AutoText"
        '             Case wdKeyCategoryStyle: categoryName = "Style"
        '             Case wdKeyCategorySymbol: categoryName = "Symbol"
        '             Case Else: categoryName = "Custom Category (" & kb.KeyCategory & ")"
        '         End Select
                
        '         MsgBox keyNames(i) & " is currently assigned to:" & vbCrLf & _
        '                "• Type: " & categoryName & vbCrLf & _
        '                "• Target: " & kb.Command & vbCrLf & vbCrLf & _
        '                "Clearing existing binding.", _
        '                vbInformation, "Keybinding Conflict"
        '     End If
        ' End If
        
        ' Clear existing binding regardless of mode
        kb.Clear
        
        If Action = 1 Then
            bindingArray = dictBindings(key)
            KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                            Command:=bindingArray(LBound(bindingArray) + 1), _
                            KeyCode:=CLng(key)
        End If
    Next key
    
    ThisDocument.Saved = True
    On Error GoTo 0
End Sub