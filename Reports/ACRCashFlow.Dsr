VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ACRCashFlow 
   Caption         =   "Cash Flow Report"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ACRCashFlow.dsx":0000
End
Attribute VB_Name = "ACRCashFlow"
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
End Sub

Private Sub ActiveReport_ReportStart()
    MDIMainMenu.AddToWin Me.Caption, Name
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "Close"
            Unload Me
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

