Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub Main()
    Dim sUser As String
    Dim sComputer As String
    Dim lpBuff As String * 1024

    'Get the Login User Name
    GetUserName lpBuff, Len(lpBuff)
    sUser = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    lpBuff = ""
    
    'Get the Computer Name
    GetComputerName lpBuff, Len(lpBuff)
    sComputer = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    lpBuff = ""

    MsgBox "Login User: " & sUser & vbCrLf & _
           "Computer Name: " & sComputer
    
    End
End Sub

