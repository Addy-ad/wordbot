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
    Dim bindings As Variant
    Dim i As Long
    Dim keyCode As Long
    Dim macroName As String
    Dim kb As KeyBinding
    Dim categoryName As String
    
    ' Define all keyboard shortcuts: [keyCode, macroName]
    bindings = Array( _
        Array(BuildKeyCode(wdKeyAlt, wdKeyR), "Ribbon_Run"), _
        Array(BuildKeyCode(wdKeyAlt, wdKeyE), "Ribbon_Expand"), _
        Array(BuildKeyCode(wdKeyAlt, wdKeyS), "Ribbon_Summarize"), _
        Array(BuildKeyCode(wdKeyAlt, wdKeyF), "Ribbon_FixGrammar"), _
        Array(BuildKeyCode(wdKeyAlt, wdKeyT), "Ribbon_Translate"), _
        Array(BuildKeyCode(wdKeyAlt, wdKeyD), "Ribbon_Define"), _
        Array(BuildKeyCode(wdKeyAlt, wdKeyC), "Ribbon_Custom"), _
        Array(BuildKeyCode(wdKeyAlt, wdKeyZ), "Ribbon_wordbotResearch") _
    )
    
    On Error Resume Next
    
    If Action = 1 Then
        CustomizationContext = ThisDocument
    End If
    
    For i = LBound(bindings) To UBound(bindings)
        keyCode = CLng(bindings(i)(0))
        macroName = CStr(bindings(i)(1))
        
        Set kb = FindKey(keyCode)
        
        ' ' Check for conflicts - if key is assigned to something else, warn user
        ' If kb.Command <> "" Then
        '     If kb.Command <> macroName Then
        '         Select Case kb.KeyCategory
        '             Case wdKeyCategoryMacro: categoryName = "Macro"
        '             Case wdKeyCategoryCommand: categoryName = "Built-in Word Command"
        '             Case wdKeyCategoryFont: categoryName = "Font"
        '             Case wdKeyCategoryAutoText: categoryName = "AutoText"
        '             Case wdKeyCategoryStyle: categoryName = "Style"
        '             Case wdKeyCategorySymbol: categoryName = "Symbol"
        '             Case Else: categoryName = "Custom Category (" & kb.KeyCategory & ")"
        '         End Select
                
        '         MsgBox "Conflict detected for " & GetKeyName(keyCode) & ":" & vbCrLf & _
        '                "Currently assigned to: " & kb.Command & " (" & categoryName & ")" & vbCrLf & vbCrLf & _
        '                "Clearing existing binding.", _
        '                vbInformation, "Keybinding Conflict"
        '     End If
        ' End If
        
        ' Clear existing binding
        kb.Clear
        
        ' Bind if Action = 1
        If Action = 1 Then
            KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                            Command:=macroName, _
                            KeyCode:=keyCode
        End If
    Next i
    
    ThisDocument.Saved = True
    On Error GoTo 0
End Sub

' Helper function to get readable key names
Private Function GetKeyName(ByVal keyCode As Long) As String
    Select Case keyCode
        Case BuildKeyCode(wdKeyAlt, wdKeyR): GetKeyName = "Alt+R"
        Case BuildKeyCode(wdKeyAlt, wdKeyE): GetKeyName = "Alt+E"
        Case BuildKeyCode(wdKeyAlt, wdKeyS): GetKeyName = "Alt+S"
        Case BuildKeyCode(wdKeyAlt, wdKeyF): GetKeyName = "Alt+F"
        Case BuildKeyCode(wdKeyAlt, wdKeyT): GetKeyName = "Alt+T"
        Case BuildKeyCode(wdKeyAlt, wdKeyD): GetKeyName = "Alt+D"
        Case BuildKeyCode(wdKeyAlt, wdKeyC): GetKeyName = "Alt+C"
        Case BuildKeyCode(wdKeyAlt, wdKeyZ): GetKeyName = "Alt+Z"
        Case Else: GetKeyName = "Key " & keyCode
    End Select
End Function