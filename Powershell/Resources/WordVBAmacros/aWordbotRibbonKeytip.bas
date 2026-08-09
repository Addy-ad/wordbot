' ============================================
' RIBBON SCREENTIP CALLBACK
' ============================================

Public Sub GetScreentip_AICommand(control As IRibbonControl, ByRef returnedVal)
    Dim isMac As Boolean
    Dim shortcut As String
    
    #If Mac Then
        isMac = True
    #Else
        isMac = False
    #End If
    
    Select Case control.ID
        Case "btnRun"
            shortcut = IIf(isMac, "Ctrl + Option + R", "Alt + R")
            returnedVal = "Run (" & shortcut & ")"
            
        Case "btnExpand"
            shortcut = IIf(isMac, "Ctrl + Option + E", "Alt + E")
            returnedVal = "Expand (" & shortcut & ")"
            
        Case "btnSummarize"
            shortcut = IIf(isMac, "Ctrl + Option + S", "Alt + S")
            returnedVal = "Summarize (" & shortcut & ")"
            
        Case "btnFixGrammar"
            shortcut = IIf(isMac, "Ctrl + Option + F", "Alt + F")
            returnedVal = "Fix Grammar (" & shortcut & ")"
            
        Case "btnTranslate"
            shortcut = IIf(isMac, "Ctrl + Option + T", "Alt + T")
            returnedVal = "Translate (" & shortcut & ")"
            
        Case "btnDefine"
            shortcut = IIf(isMac, "Ctrl + Option + D", "Alt + D")
            returnedVal = "Define (" & shortcut & ")"
            
        Case "btnCustom"
            shortcut = IIf(isMac, "Ctrl + Option + C", "Alt + C")
            returnedVal = "Custom instruction (" & shortcut & ")"
            
        Case "btnResearch"
            shortcut = IIf(isMac, "Ctrl + Option + Z", "Alt + Z")
            returnedVal = "Zotero Research (" & shortcut & ")"
            
        Case Else
            returnedVal = control.Context
    End Select
End Sub