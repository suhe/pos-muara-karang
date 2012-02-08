Attribute VB_Name = "modExcel"
Option Explicit

Public xlApp                    As New Excel.Application
Public xlBook                   As Excel.Workbook
Public xlSheet                  As Excel.Worksheet


Public Function SaveAsExcel(rsErr As ADODB.Recordset, sFileName As String, _
            sSheet As String, sOpen As String)
Dim fd          As Field
Dim CellCnt     As Integer
Dim i           As Integer

On Error GoTo Err_Handler

Screen.MousePointer = vbHourglass

Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

'Get the field names
CellCnt = 1
xlSheet.Name = sSheet
For Each fd In rsErr.Fields 'Add the headers..
       xlSheet.Cells(1, CellCnt).Value = fd.Name
       xlSheet.Cells(1, CellCnt).Interior.ColorIndex = 33
       xlSheet.Cells(1, CellCnt).Font.Bold = True
       xlSheet.Cells(1, CellCnt).BorderAround xlContinuous
       CellCnt = CellCnt + 1
Next

'Rewind the recordset so we can get at the actual records...
rsErr.MoveFirst
i = 2
Do While Not rsErr.EOF()
     CellCnt = 1
     For Each fd In rsErr.Fields
        xlSheet.Cells(i, CellCnt).Value = rsErr.Fields(fd.Name).Value
        CellCnt = CellCnt + 1
     Next
     rsErr.MoveNext
     i = i + 1
 Loop

'AutoFit all columns
CellCnt = 1
For Each fd In rsErr.Fields
    xlSheet.Columns(CellCnt).AutoFit
    CellCnt = CellCnt + 1
Next
     
xlSheet.SaveAs sFileName ' Save the Worksheet.
xlBook.Close ' Close the Workbook
xlApp.Quit ' Close Microsoft Excel with the Quit method.

If sOpen = "YES" Then ' Open Excel Workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(sFileName)
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Application.Visible = True
Else
    Set xlApp = Nothing ' Release the Excel objects.
    Set xlBook = Nothing
    Set xlSheet = Nothing
End If

Err_Handler:
    If err = 0 Then
        Screen.MousePointer = vbDefault
    Else
        MsgBox "An error has occurred! " & vbCrLf & vbCrLf & err & ":" & Error & " ", vbExclamation
        Screen.MousePointer = vbDefault
    End If
End Function




