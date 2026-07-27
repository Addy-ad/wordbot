' Helper function to escape special characters for JSON
Public Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function

' Extract the assistant's reply from the JSON response
Public Function ExtractContentFromJSON(jsonStr As String, searchKey_term As String) As String
    Dim startPos As Long, endPos As Long
    Dim searchKey As String
    
    ' Assign the key used by your Python Flask server
    searchKey = """" & searchKey_term & """:"
    startPos = InStr(jsonStr, searchKey)
    
    If startPos = 0 Then
        ExtractContentFromJSON = "Could not parse response."
        Exit Function
    End If
    
    ' Find the opening quote after the key
    startPos = InStr(startPos + Len(searchKey), jsonStr, """")
    If startPos = 0 Then
        ExtractContentFromJSON = "Could not find opening quote."
        Exit Function
    End If
    
    startPos = startPos + 1 ' Move exactly one character inside the quotes

    ' Find the closing quote that is not escaped
    endPos = startPos
    Do While endPos <= Len(jsonStr)
        endPos = InStr(endPos, jsonStr, """")
        If endPos = 0 Then Exit Do
        If Mid(jsonStr, endPos - 1, 1) <> "\" Then Exit Do
        endPos = endPos + 1
    Loop

    If endPos = 0 Then
        ExtractContentFromJSON = "Could not find closing quote."
        Exit Function
    End If

    ExtractContentFromJSON = Mid(jsonStr, startPos, endPos - startPos)

    ' ======================================================================
    ' THE SOLUTION SEQUENCE: SAFE ESCAPE ISOLATION MAPPING
    ' ======================================================================

    ' STEP 1: Turn double backslashes into a custom mathematical marker.
    ' This transforms "\\nabla" into "[[LATEX_BACKSLASH]]nabla", completely
    ' moving your math commands out of danger.
    ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\\", "[[LATEX_BACKSLASH]]")

    ' STEP 2: Now that your LaTeX is protected, process standard structural JSON escapes.
    ' This will safely catch true paragraph changes or quotes sent by the server.
    ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\""", """")
    ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\n", vbCrLf)
    ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\t", vbTab)

    ' STEP 3: Re-hydrate your math tokens back into single LaTeX backslashes.
    ExtractContentFromJSON = Replace(ExtractContentFromJSON, "[[LATEX_BACKSLASH]]", "\")
End Function

