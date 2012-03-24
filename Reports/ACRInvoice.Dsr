VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ACRInvoice 
   Caption         =   "Invoice"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   Icon            =   "ACRInvoice.dsx":0000
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   21643
   _ExtentY        =   15028
   SectionData     =   "ACRInvoice.dsx":038A
End
Attribute VB_Name = "ACRInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_Activate()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fffffft"
End Sub

Private Sub ActiveReport_Deactivate()
    MDIMainMenu.HideTBButton "", True
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        'Case vbKeyF1: MDIMainMenu.tbMenu.Button.Key = "Shortcut"
        Case vbKeyF2: CommandPass "New"
        Case vbKeyF3: CommandPass "Edit"
        Case vbKeyF4: CommandPass "Search"
        Case vbKeyF5: CommandPass "Delete"
        Case vbKeyF6: CommandPass "Refresh"
        Case vbKeyF8: CommandPass "Close"
    End Select
End Sub

Private Sub ActiveReport_ReportEnd()
    MDIMainMenu.RemToWin Me.Caption
    Call Invoice_lunas
End Sub

Private Sub ActiveReport_ReportStart()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fffffft"
    MDIMainMenu.AddToWin Me.Caption, Name
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "Close"
            Unload Me
            Call cetak_Invoice
    End Select
    Exit Sub
    'Trap the error
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
End Sub





