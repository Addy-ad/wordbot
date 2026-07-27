' ============================================
' RIBBON CALLBACKS (for button clicks)
' ============================================

Public Sub RibbonProxy_AICommand(control As IRibbonControl)
    Select Case control.ID
        Case "btnRun"
            Ribbon_Run
        Case "btnExpand"
            Ribbon_Expand
        Case "btnSummarize"
            Ribbon_Summarize
        Case "btnFixGrammar"
            Ribbon_FixGrammar
        Case "btnTranslate"
            Ribbon_Translate
        Case "btnDefine"
            Ribbon_Define
        Case "btnCustom"
            Ribbon_Custom
    End Select
End Sub

Public Sub RibbonProxy_StartServer(control As IRibbonControl)
    Ribbon_StartServer
End Sub

Public Sub RibbonProxy_HeadingNumber(control As IRibbonControl)
    Ribbon_HeadingNumber
End Sub

Public Sub RibbonProxy_FormatMarkdown(control As IRibbonControl)
    Ribbon_FormatMarkdown
End Sub

Public Sub RibbonProxy_Research(control As IRibbonControl)
    Ribbon_wordbotResearch
End Sub

Public Sub RibbonProxy_ConvertCitations(control As IRibbonControl)
    Ribbon_ConvertSelectionToFields
End Sub

Public Sub RibbonProxy_UpdateCitations(control As IRibbonControl)
    UpdateAllReferences
End Sub