VERSION 5.00
Begin VB.Form frmPurchaseAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase No Fak :"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmPurchaseAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   405
      ScaleHeight     =   30
      ScaleWidth      =   12015
      TabIndex        =   3
      Top             =   3465
      Width           =   12015
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   120
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Payment"
      Text            =   "0"
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmPurchaseAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public State                As FormState
Public PK                   As Integer
Public srcText              As TextBox
Dim HaveAction              As Boolean
Dim rs                      As New Recordset




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(1), True) = True Then Exit Sub
        sql = "UPDATE tbl_beli b "
        sql = sql + "SET "
        sql = sql + " bayar=" & txtEntry(1).Text & " "
        sql = sql + " WHERE b.id_beli=" & tbl.TABLE_NO_FAK
        CN.Execute sql
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtEntry(1).Text = tbl.TABLE_PAYMENT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPurchase.RefreshRecords
    Set frmPurchaseAE = Nothing
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
     If (Index = 1) Then
        NumberOnly KeyAscii
     End If
End Sub
