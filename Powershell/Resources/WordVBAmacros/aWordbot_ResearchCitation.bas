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
    
    Dim matches As Collection
    Set matches = New Collection
    
    Set rng = doc.Range(rangeStart, rangeEnd)
    
    With rng.Find
        .text = "zotero://[A-Za-z0-9]{8}"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Replacement.text = ""
        
        Do While .Execute
            If rng.Start >= rangeStart And rng.End <= rangeEnd Then
                matches.Add Array(rng.text, rng.Start, rng.End)
            Else
                Exit Do
            End If
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Process matches backwards
    Dim i As Integer
    For i = matches.Count To 1 Step -1
        Dim matchData As Variant
        matchData = matches(i)
        
        Dim fullText As String, itemKey As String
        Dim keyStart As Long
        
        fullText = matchData(0)
        ' Key is the 8 characters after "zotero://"
        itemKey = Mid(fullText, 10, 8)
        
        Set rng = doc.Range(matchData(1), matchData(2))
        
        ' Clean up any old corrupted fields or bookmarks in this spot
        Dim existingFld As Field
        For Each existingFld In rng.Fields
            existingFld.Delete
        Next existingFld
        
        Dim bm As Bookmark
        For Each bm In rng.Bookmarks
            bm.Delete
        Next bm
        
        ' Delete the matched LLM text completely
        rng.Delete
        
        ' Build Zotero JSON with [CITATION] placeholders to bypass edit detection
        Dim jsonPayload As String, randomID As String
        randomID = LCase(Right("0000000" & Hex(Int(Rnd * 16777215)), 8))
        
        jsonPayload = "{" & _
            """citationID"": """ & randomID & """," & _
            """properties"": {""unsorted"": false, ""noteIndex"": 0, ""formattedCitation"": ""[CITATION]"", ""plainCitation"": ""[CITATION]""}," & _
            """citationItems"": [{" & _
                """uris"": [""http://zotero.org/users/local/0/items/" & itemKey & """]" & _
            "}]," & _
            """schema"": ""https://github.com/citation-style-language/schema/raw/master/csl-citation.json""" & _
        "}"
        
        ' Create an empty field to strictly control the field code and boundary
        Set fld = doc.Fields.Add(Range:=rng, Type:=wdFieldEmpty, PreserveFormatting:=False)
        
        ' Write the exact AddIn string and match the visible result to the JSON
        fld.Code.text = "ADDIN ZOTERO_ITEM CSL_CITATION " & jsonPayload
        fld.result.text = "[CITATION]"
        
        ' This parts is to check if the new Field landed inside an OMath object ---
        If fld.result.OMaths.Count > 0 Then
            Dim parentMath As OMath
            Dim outRng As Range
            
            Set parentMath = fld.result.OMaths(1)
            
            fld.Delete
            
            Set outRng = parentMath.Range.Duplicate
            outRng.Collapse wdCollapseEnd
            
            ' Ensure we move past all nested OMath bounds
            Do While outRng.OMaths.Count > 0 And outRng.End < doc.Range.End
                outRng.Collapse wdCollapseEnd
                outRng.MoveEnd wdCharacter, 1
            Loop
            
            ' Add a space separator if adjacent to text
            If outRng.Start > docRange.Start Then
                Dim charBefore As Range
                Set charBefore = doc.Range(outRng.Start - 1, outRng.Start)
                If charBefore.text <> " " And charBefore.text <> vbCr And charBefore.text <> vbLf Then
                    outRng.InsertBefore " "
                    outRng.Collapse wdCollapseEnd
                End If
            End If
            
            ' Re-create the field cleanly using clean jsonPayload outside math region
            Set fld = doc.Fields.Add(Range:=outRng, Type:=wdFieldEmpty, PreserveFormatting:=False)
            fld.Code.text = "ADDIN ZOTERO_ITEM CSL_CITATION " & jsonPayload
            fld.result.text = "[CITATION]"
        End If
    Next i
    
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
