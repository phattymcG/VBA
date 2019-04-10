Sub iterateSection()

'iterations sections of the full draft

Dim section As Integer
Dim numberFound As Integer
Dim start As Single, finish As Single, totalTime As Single

start = Timer

For section = 1 To 1
    'extractGlossaryTerms ActiveDocument.Sections(section).Range
    'extractExternalRefs ActiveDocument.Sections(section).Range
    advancedFind ActiveDocument.Sections(section).Range, _
        numberFound
Next section

finish = Timer
totalTime = finish - start

MsgBox CStr(numberFound) + " instances found and fixed in " _
        + CStr(totalTime) + " seconds"

End Sub
Sub extractGlossaryTerms(section As Range)

'TO MAKE INDEPENDENT:
'Dim section As Range
'Set section = ActiveDocument.Content

Dim wrd As Range
Dim wb As Workbook, ws As Worksheet
Dim wsNumber As Integer
Dim wbName As String
Dim extractedTerm As String

Dim styleType As String
Dim headingLevel As String
Dim check As String

wbName = "Glossary terms.xlsx"
wsNumber = 1
extractedTerm = "Glossary terms"

styleType = "Glossary Char"
headingLevel = "2"

Set wb = setWorkbook(wb, wbName)
Set ws = wb.Worksheets(wsNumber)
ws.Range("A1").Value = extractedTerm


For Each wrd In section.Words

'the "check" term functions as a lock on checking words
'if one of the checking subroutines finds a match, it sets
'the "check" term to avoid checking for multiple criteria on the
'same word, as meeting just one criterion is sufficient to
'to trigger an export of a word (or multi-word term)
'I think this saves cycles compared to iterating the entire document
'for each type of check subroutine
'Seems like there should be a more elegant solution to this logic,
'but maybe not

If check = "" Then
    checkForStyle wrd, styleType, _
        headingLevel, ws, check
ElseIf check = "style" Then
    checkForStyle wrd, styleType, _
        headingLevel, ws, check
End If

If check = "" Then
    checkForColor ws, wrd, extractedItem, check
ElseIf check = "color" Then
    checkForColor ws, wrd, extractedItem, check
End If

Next wrd

End Sub


Sub extractExternalRefs(section As Range)

'extracts external refs to a spreadsheet at with heading level 4
'association

'TO MAKE INDEPENDENT:
'Dim section As Range
'Set section = ActiveDocument.Content

Dim wrd As Range
Dim wb As Workbook, ws As Worksheet
Dim wsNumber As Integer
Dim wbName As String
Dim extractedTerm As String

wbName = "External References.xlsx"
wsNumber = 1
extractedTerm = "External References"

Set wb = setWorkbook(wbName)
Set ws = wb.Worksheets(wsNumber)
ws.Range("A1").Value = extractedTerm

For Each wrd In section.Words

checkForStyle wrd, "External Reference Char", _
                    "4", ws

Next wrd

End Sub


Sub checkForColor(ws As Worksheet, wrd As Range, _
                    check As String)

Dim export As Boolean
Static extractedItem As String

If wrd.style = "Normal" Then
If wrd.text <> vbCr Then
If wrd.Font.color = wdColorBlue Then
    extractedItem = extractedItem + wrd.text
    check = "color"
Else
    If extractedItem <> "" Then export = True
End If

If export Then
    exportextractedItem ws, extractedItem, 2
    extractedItem = ""
    check = ""
    export = False
End If

End If
End If

End Sub

Sub checkForStyle(wrd As Range, _
                    styleType As String, _
                    headingLevel As String, _
                    ws As Worksheet, _
                    Optional check As String)

'checks for specified heading level and a specified style, and
' exports an entire phrase when it's been identified

Dim wrdStyle As style
Dim wrdStyleName As String
Dim lastWrdStyle As String
Dim export As Boolean
Static headingText As String
Static extractedItem As String

Set wrdStyle = wrd.style
wrdStyleName = CStr(wrdStyle)

expandHeadingLevel headingLevel

If Not wrdStyle Is Nothing Then
Select Case wrdStyleName
Case styleType
    extractedItem = extractedItem + wrd.text
    check = "style"
Case headingLevel
    If Not extractedItem = "" Then
        extractedItem = extractedItem + wrd.text
    Else
        extractedItem = CStr(wrd.ListFormat.ListString) _
        + " " + wrd.text
    End If
    check = "style"
    headingText = extractedItem
Case Else
    'exports the phrase once it hits a word with another style
    If extractedItem <> "" Then export = True
End Select

If export Then
    exportextractedItem ws, extractedItem, headingText, _
                        lastWrdStyle
    'reset the static variables
    extractedItem = ""
    check = ""
    export = False
End If

lastWrdStyle = wrdStyleName

End If

End Sub
Sub exportextractedItem(ws As Worksheet, _
                        extractedItem As String, _
                        headingText As String, _
                        Optional lastWrdStyle As String)

Static lastRowFound As Boolean
Static firstBlankRow

'On Error GoTo handler
'Set wb = Workbooks("Glossary terms.xlsx")
'On Error GoTo 0

If Not lastRowFound Then
    If ws.Range("A2").text <> "" Then
        firstBlankRow = ws.Range("A1", ws.Range("A1").End(xlDown)) _
                    .Rows.Count + 1
    Else
        firstBlankRow = 2
    End If
lastRowFound = True
Else
    'add an extra blank line between headings
    If lastWrdStyle = "Heading 2" Then firstBlankRow _
                                        = firstBlankRow + 1
End If

ws.Cells(firstBlankRow, 1).Value = headingText

If extractedItem <> headingText Then
    ws.Cells(firstBlankRow, 2).Value = extractedItem
End If

firstBlankRow = firstBlankRow + 1

'Exit Sub

'handler:
'MsgBox "Please open the glossary terms.xlsx file"
'End

End Sub

Function setWorkbook(wbName As String) _
                    As Workbook

Dim wb As Workbook
Dim appExcel As Excel.Application

Set appExcel = GetObject(, "Excel.Application")
appExcel.Visible = True
Set setWorkbook = appExcel.Workbooks(wbName)

End Function
Function expandHeadingLevel(headingLevel As String)

Select Case headingLevel
    Case "4"
        headingLevel = "Heading 4"
    Case "3"
        headingLevel = "Heading 3"
    Case "2"
        headingLevel = "Heading 2"
    Case "1"
        headingLevel = "Heading 1"
End Select

End Function

Function loadBuffer(section As Range, buffer() As Range, _
                    bufferLength As Long)

'loads a passed array with specified number of words (as Ranges)

Dim wrd As Integer

If section.Words.Count >= bufferLength Then
    For wrd = 1 To bufferLength
        Set buffer(wrd) = section.Words(wrd)
    Next wrd
Else
    End
End If

End Function
