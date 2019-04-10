Sub getURLs()

'exports list of URLs and display text from a Word file
'to an Excel file

'REQUIRES Microsoft Excel 15.0 Object Library reference (see
'Tools - References to enable)

'custom URL extraction from PDF in which each display
'word in a hyperlink had a separate copy of the hyperlink
'attached to it (duplicate hyperlinks for each .TextToDisplay
'word)

Dim docCurrent As Document
Dim excelApp As Object
Dim ssNew As Workbook
Dim oLink As Hyperlink
Dim lastLinkAddress As String
Dim rng As Range
Dim ctr As Integer

Application.ScreenUpdating = False

Set docCurrent = ActiveDocument

'On Error Resume Next
'neither of the below lines should create a new instance
'  of Excel, but they do on the current machine
Set excelApp = excel.Application
'Set excelApp = GetObject(, "Excel.Application")
'If Err.Number <> 0 Then
'    MsgBox "Please open Excel and try again"
'    End
'End If
'On Error GoTo 0

Set ssNew = Workbooks.Add
'ssNew.SaveAs docCurrent.Path & "\hyperlinks in current KC2.xlsx"
'Set ssNew = Workbooks.Open(docCurrent.Path & "\hyperlinks in current KC.xlsx")
ssNew.Application.Visible = True
'excelApp.Application.Visible = True
'excelApp.Parent.Windows(1).Visible = True
'ssNew.Parent.Windows(1).Visible = True

ctr = 2

For parag = 1 To docCurrent.Paragraphs.Count

    If parag Mod 2 = 1 Then
        ssNew.Worksheets(1).Cells(ctr, 1).Value = _
        docCurrent.Paragraphs(parag).Range.Text
    Else
        ssNew.Worksheets(1).Cells(ctr, 2).Value = _
        docCurrent.Paragraphs(parag).Range.Text
        ctr = ctr + 1
    End If
    
Next

ssNew.Save

Application.ScreenUpdating = True
'Application.ScreenRefresh

End Sub
