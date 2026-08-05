Public Sub ConvertSelectionToFields_callback(ByRef docRange As Range)

    If docRange.Start = docRange.End Then
        MsgBox "The selected range is empty.", vbExclamation
        Exit Sub
    End If
    
    Dim fld As Field
    Dim rng As Range
    Dim doc As Document
    Dim objUndo As UndoRecord
    
    Set doc = docRange.Document
    
    If docRange Is Nothing Then
        MsgBox "No valid range provided."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False

    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord ("Convert References to Fields")
    
    Dim rangeStart As Long, rangeEnd As Long
    rangeStart = docRange.Start
    rangeEnd = docRange.End
    
    ' ---------- Collect all matches ----------
    Dim matches As Collection
    Set matches = New Collection
    
    Set rng = doc.Range(rangeStart, rangeEnd)
    
    With rng.Find
        .Text = "zotero://[A-Za-z0-9]{8}"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Replacement.Text = ""
        
        Do While .Execute
            If rng.Start >= rangeStart And rng.End <= rangeEnd Then
                ' Store start, end, and the full matched text
                matches.Add Array(rng.Text, rng.Start, rng.End)
            Else
                Exit Do
            End If
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    If matches.Count = 0 Then
        MsgBox "No Zotero placeholders found in the selection.", vbInformation
        Exit Sub
    End If
    
    ' ---------- Group matches that are separated only by whitespace ----------
    Dim groups As Collection
    Set groups = New Collection
    
    Dim i As Integer
    i = 1
    Do While i <= matches.Count
        Dim groupMembers As Collection
        Set groupMembers = New Collection
        groupMembers.Add i
        
        ' Look ahead to see if the next match is adjacent (whitespace only between)
        Do While i < matches.Count
            Dim nextIdx As Integer
            nextIdx = i + 1
            Dim matchData As Variant, nextMatchData As Variant
            matchData = matches(i)
            nextMatchData = matches(nextIdx)
            
            ' Range between the current match's end and the next match's start
            Dim rngBetween As Range
            Set rngBetween = doc.Range(matchData(2), nextMatchData(1)) ' end of current, start of next
            Dim betweenText As String
            betweenText = rngBetween.Text
            
            If IsOnlyWhitespace(betweenText) Then
                groupMembers.Add nextIdx
                i = nextIdx   ' move forward
            Else
                Exit Do
            End If
        Loop
        
        groups.Add groupMembers
        i = i + 1
    Loop
    
    ' ---------- Process groups from right to left ----------
    Dim g As Integer
    For g = groups.Count To 1 Step -1
        Dim group As Collection
        Set group = groups(g)
        
        ' Get first and last match indices in this group
        Dim firstIdx As Integer, lastIdx As Integer
        firstIdx = group(1)
        lastIdx = group(group.Count)
        
        ' Retrieve their start and end positions
        Dim firstMatch As Variant, lastMatch As Variant
        firstMatch = matches(firstIdx)
        lastMatch = matches(lastIdx)
        Dim groupStart As Long, groupEnd As Long
        groupStart = firstMatch(1)
        groupEnd = lastMatch(2)
        
        ' Collect all citation keys in forward order
        Dim keys As Collection
        Set keys = New Collection
        Dim k As Integer
        For k = 1 To group.Count
            Dim idx As Integer
            idx = group(k)
            Dim matchItem As Variant
            matchItem = matches(idx)
            Dim fullText As String
            fullText = matchItem(0)
            Dim key As String
            key = Mid(fullText, 10, 8)   ' extract 8 chars after "zotero://"
            keys.Add key
        Next k
        
        ' ---------- Prepare insertion point ----------
        Set rng = doc.Range(groupStart, groupEnd)
        
        ' Clean up any old fields or bookmarks (though they shouldn't exist)
        Dim existingFld As Field
        For Each existingFld In rng.Fields
            existingFld.Delete
        Next existingFld
        
        Dim bm As Bookmark
        For Each bm In rng.Bookmarks
            bm.Delete
        Next bm
        
        ' ---------- Delete the entire group range ----------
        rng.Delete   ' now rng is collapsed at groupStart

        ' ---------- Check and add leading space AFTER deletion ----------
        If rng.Start > 0 Then
            Dim charBefore As Range
            Set charBefore = doc.Range(rng.Start - 1, rng.Start)
            Dim leftChar As String
            leftChar = charBefore.Text
            ' Only add space if preceding char is NOT whitespace
            If leftChar <> " " And leftChar <> vbTab And leftChar <> vbCr And leftChar <> vbLf Then
                rng.InsertBefore " "
                rng.Collapse wdCollapseEnd
            End If
        End If
        
        ' ---------- Build combined JSON with multiple citationItems ----------
        Dim jsonPayload As String
        jsonPayload = BuildCombinedJSON(keys)
        
        ' ---------- Insert the Zotero field ----------
        Set fld = doc.Fields.Add(Range:=rng, Type:=wdFieldEmpty, PreserveFormatting:=False)
        fld.Code.Text = "ADDIN ZOTERO_ITEM CSL_CITATION " & jsonPayload
        fld.Result.Text = "[CITATION]"
        
        ' ---------- Handle OMath (if field lands inside equation) ----------
        If fld.Result.OMaths.Count > 0 Then
            Dim parentMath As OMath
            Dim outRng As Range
            
            Set parentMath = fld.Result.OMaths(1)
            fld.Delete   ' remove the field that landed inside math
            
            Set outRng = parentMath.Range.Duplicate
            outRng.Collapse wdCollapseEnd
            
            ' Move past all nested OMath bounds
            Do While outRng.OMaths.Count > 0 And outRng.End < doc.Range.End
                outRng.Collapse wdCollapseEnd
                outRng.MoveEnd wdCharacter, 1
            Loop
            
            ' Add a leading space if needed (same logic as before)
            If outRng.Start > docRange.Start Then
                Dim charBeforeMath As Range
                Set charBeforeMath = doc.Range(outRng.Start - 1, outRng.Start)
                If charBeforeMath.Text <> " " And charBeforeMath.Text <> vbCr And charBeforeMath.Text <> vbLf Then
                    outRng.InsertBefore " "
                    outRng.Collapse wdCollapseEnd
                End If
            End If
            
            ' Re-create the field cleanly outside the math region
            Set fld = doc.Fields.Add(Range:=outRng, Type:=wdFieldEmpty, PreserveFormatting:=False)
            fld.Code.Text = "ADDIN ZOTERO_ITEM CSL_CITATION " & jsonPayload
            fld.Result.Text = "[CITATION]"
        End If
    Next g
    
    objUndo.EndCustomRecord
    Application.ScreenUpdating = True
    
    ' Trigger Zotero to refresh
    UpdateAllReferences
    
    ' Force the active window view to hide field codes
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.View.ShowFieldCodes = True
        ActiveWindow.View.ShowFieldCodes = False
    End If
    
End Sub

' ------------------------------------------------------------
' Helper: Check if a string contains only whitespace characters
' ------------------------------------------------------------
Private Function IsOnlyWhitespace(ByVal text As String) As Boolean
    Dim ch As String
    Dim i As Integer
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        Select Case ch
            Case " ", vbTab, vbCr, vbLf, vbVerticalTab, vbFormFeed
                ' continue
            Case Else
                IsOnlyWhitespace = False
                Exit Function
        End Select
    Next i
    IsOnlyWhitespace = True
End Function

' ------------------------------------------------------------
' Helper: Build JSON payload with multiple citation items
' ------------------------------------------------------------
Private Function BuildCombinedJSON(ByRef keys As Collection) As String
    Dim randomID As String
    randomID = LCase(Right("0000000" & Hex(Int(Rnd * 16777215)), 8))
    
    Dim json As String
    json = "{" & _
        """citationID"": """ & randomID & """," & _
        """properties"": {""unsorted"": false, ""noteIndex"": 0, ""formattedCitation"": ""[CITATION]"", ""plainCitation"": ""[CITATION]""}," & _
        """citationItems"": ["
    
    Dim firstItem As Boolean
    firstItem = True
    Dim key As Variant
    For Each key In keys
        If Not firstItem Then
            json = json & ","
        End If
        json = json & "{""uris"": [""http://zotero.org/users/local/0/items/" & key & """]}"
        firstItem = False
    Next key
    
    json = json & "]," & _
        """schema"": ""https://github.com/citation-style-language/schema/raw/master/csl-citation.json""" & _
        "}"
    
    BuildCombinedJSON = json
End Function

Public Sub UpdateAllReferences()
    #If Mac Then
        MsgBox "Fields converted successfully! Please click 'Refresh' on the Zotero ribbon tab to update citations and bibliography.", _
               vbInformation, "Zotero Refresh Required"
    #Else
        On Error Resume Next
        ' Calls ZoteroRefresh located in the Zotero template module
        Application.Run "Zotero.ZoteroCommand", "refresh", False
            
        If Err.Number <> 0 Then
            MsgBox "Could not trigger Zotero. Please ensure the Zotero Word add-in is active.", vbExclamation
        End If
        On Error GoTo 0
    #End If

End Sub
