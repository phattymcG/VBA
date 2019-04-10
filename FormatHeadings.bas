Sub formatHeadings(docRange As Range)

' applies the correct heading style to lines of text that are
' manually numbered in the style x.x(.x.x)
' processed text is then manually copied into a template with
' the correct formatting for the heading styles

Dim headingLevel As Integer
Dim heading As String

'check for lines that start with a number
'---don't need to check for correctly applied headings,
'   because it's not possible to manually screw up the
'   number without some serious option clicking

heading = Trim(docRange.Text)

'see if raw text starts with a number
If Left(heading, 1) Like "#" Then
        
    headingLevel = determineHeadingLevel(Left(heading, _
                InStr(heading, " ") - 1), docRange)
    
    'remove manual numbers
    'have to do this before setting style or style will be reset
    'to normal (!)
    docRange.Text = Right(heading, Len(heading) _
                    - InStr(heading, " "))

    'set the heading style
    setHeading docRange, headingLevel
    
'see if line has manually applied outline numbering
ElseIf Not docRange.ListFormat.List Is Nothing Then
    If docRange.ListFormat.listType = wdListOutlineNumbering Then
    
    headingLevel = docRange.ListFormat.ListLevelNumber
    
    'set the heading style
    setHeading docRange, headingLevel
    
    'remove outline numbers
    'unlike above, have to do this AFTER setting style, or
    'numbering comes back!
    docRange.ListFormat.RemoveNumbers
    
    End If
'DEBUG
'Else
'MsgBox "No numbering (or heading style numbering) on " _
    + heading + "!"

End If

End Sub

Function determineHeadingLevel(numberString As String, docRange As Range)

'determines heading level if a number was found
'at the beginning of a line


'Dim compStringHdg1 As String
Dim compStringHdg2 As String
Dim compStringHdg3 As String
Dim compStringHdg4 As String

'compStringHdg1 = "#"
compStringHdg2 = "#*.#*"
compStringHdg3 = "#*.#*.#*"
compStringHdg4 = "#*.#*.#*.#*"

'OR, string comp method

If numberString Like compStringHdg4 Then
determineHeadingLevel = 4
ElseIf numberString Like compStringHdg3 Then
determineHeadingLevel = 3
ElseIf numberString Like compStringHdg2 Then
determineHeadingLevel = 2
End If

End Function

Sub setHeading(docRange As Range, headingLevel As Integer)

'sets heading style if a number was found
'at the beginning of a line
                
Select Case headingLevel
    Case 4
        docRange.Style = wdStyleHeading4
    Case 3
        docRange.Style = wdStyleHeading3
    Case 2
        docRange.Style = wdStyleHeading2
End Select

End Sub

Sub checkListNumbering(docRange As Range)

'checks to see if any type of list numbering other than
'outline numbering or bullets exists

Dim listType As String
Dim diffNumbering As Boolean

If Not docRange.ListFormat.List Is Nothing Then
    If Not docRange.ListFormat.listType = wdListOutlineNumbering Then
    If Not docRange.ListFormat.listType = wdListBullet Then
    listType = docRange.ListFormat.listType
    MsgBox "list type " + listType + " on " _
        + docRange.Text + "!"
    diffNumbering = True
    End If
    End If
End If



'type number conversions:
'wdListBullet 2 Bulleted list.
'wdListListNumOnly 1 ListNum fields that can be used in the body of a paragraph.
'wdListMixedNumbering 5 Mixed numeric list.
'wdListNoNumbering 0 List with no bullets, numbering, or outlining.
'wdListOutlineNumbering 4 Outlined list.
'wdListPictureBullet 6 Picture bulleted list.
'wdListSimpleNumbering 3 Simple numeric list.


End Sub

Sub checkManualNumbering(docRange As Range)

If Left(docRange.Text, 1) Like "#" Then
    
    MsgBox "manual text numbering on " + docRange.Text + "!"
    
End If
End Sub
