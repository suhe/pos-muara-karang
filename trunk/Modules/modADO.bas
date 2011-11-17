Attribute VB_Name = "modADO"
Option Explicit

Public Function OpenDB() As Boolean
    Dim isOpen      As Boolean
    Dim ANS         As VbMsgBoxResult
    isOpen = False
    On Error GoTo err
        Do Until isOpen = True
                CN.CursorLocation = adUseClient
                'CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False;Jet OLEDB:Database Password=philiprj"
                CN.Open "DSN=pos_db"
            isOpen = True
        Loop
        OpenDB = isOpen
    Exit Function
err:
    ANS = MsgBox("Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical + vbRetryCancel)
    If ANS = vbCancel Then
        OpenDB = False
    ElseIf ANS = vbRetry Then
        Resume
    End If
End Function

Public Sub CloseDB()
    CN.Close
    Set CN = Nothing
End Sub

Public Function getIndex(ByVal srcPK As String, ByVal srcTable As String) As Long
    On Error GoTo err
    Dim rs As New Recordset
    Dim RI As Long
    
    rs.CursorLocation = adUseClient
    sql = "SELECT * FROM " & srcTable & " ORDER BY " & srcPK & " DESC "
    'MsgBox sql
    rs.Open sql, CN, adOpenStatic, adLockOptimistic
    If (rs.RecordCount > 0) Then
        RI = rs.Fields(0)
        RI = Val(RI) + 1
    Else
        RI = 1
    End If
    getIndex = RI
    srcPK = ""
    srcTable = ""
    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getIndex = 1: Resume Next
End Function

'Function used to get the sum  of fields
Public Function getSumOfFields(ByVal sTable As String, ByVal sField As String, ByRef sCN As ADODB.Connection, Optional inclField As String, Optional sCondition As String, Optional wCondition As String) As Double
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If wCondition <> "" Then wCondition = " WHERE " & wCondition & ""
    If sCondition <> "" Then sCondition = " GROUP BY " & inclField & " "
    If inclField <> "" Then inclField = "," & inclField
    sql = "SELECT Sum(" & sField & ") AS fTotal FROM " & sTable & wCondition & sCondition
    MsgBox sql
    rs.Open sql, sCN, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            getSumOfFields = getSumOfFields + rs.Fields("fTotal")
            rs.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If
    
    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function

'Function used to get the sum  of fields
Public Function getSUMTotal(ByVal xsql As String, ByVal xtotal As String, ByRef sCN As ADODB.Connection) As Double
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open xsql, sCN, adOpenStatic, adLockOptimistic
    'MsgBox rs.Fields("total")
    If rs.RecordCount > 0 Then
            getSUMTotal = rs.Fields("" & xtotal & "")
    Else
        getSUMTotal = 0
    End If
    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSUMTotal = 0: Resume Next
End Function


'Procedure used to generate DSN
Public Sub GenerateDSN()
Open App.path & "\rptCN.dsn" For Output As #1
    Print #1, "[ODBC]"
    Print #1, "DRIVER=Microsoft Access Driver (*.mdb)"
    Print #1, "UID=admin"
    Print #1, "UserCommitSync=Yes"
    Print #1, "Threads=3"
    Print #1, "SafeTransactions=0"
    Print #1, "PageTimeout=5"
    Print #1, "MaxScanRows=8"
    Print #1, "MaxBufferSize=2048"
    Print #1, "FIL=MS Access"
    Print #1, "DriverId=25"
    Print #1, "DefaultDir=" & App.path
    Print #1, "DBQ=" & App.path & "\Db.mdb"
Close #1
End Sub

'Procedure used to remove DSN
Public Sub RemoveDSN()
On Error Resume Next
Kill App.path & "\rptCN.dsn"
End Sub
