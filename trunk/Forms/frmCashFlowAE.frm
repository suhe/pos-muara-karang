VERSION 5.00
Begin VB.Form frmCashFlowAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Flow"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   Icon            =   "frmCashFlowAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text1 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "frmCashFlowAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode
Public srcTextAdd           As TextBox 'Used in pop-up mode -> Display the customer address
Public srcTextCP            As TextBox 'Used in pop-up mode -> Display the customer contact person
Public srcTextDisc          As Object  'Used in pop-up mode -> Display the customer Discount (can be combo or textbox)
Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    With rs
         Text1.Text = .Fields("tgl_cash")
         text2.Text = .Fields("cash")
    End With
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    text2.SetFocus
End Sub

Private Sub cmdSave_Click()
    If toNumber(text2.Text) < 1000 Then MsgBox "Field value must be 1000 to larger Number only.", vbExclamation: Exit Sub
    If text2.Text = "" Then MsgBox "Field to Cash Return value No Empty And Number only..", vbExclamation:  Exit Sub
    If is_high(text2.Text, tbl.TABLE_TOTAL, True) = True Then Exit Sub
    If State = adStateAddMode Or State = adStatePopupMode Then
    Else
        'Dim transfer As Double
        Dim laba, transfer, total As Double
        
        With tbl
            laba = Val(.TABLE_LABA_BERSIH)
            transfer = Val(text2.Text)
            total = laba - transfer
                .TABLE_TRANSFER = Format(text2.Text, "")
                .TABLE_KAS_SISA = Format(total, "")
        End With
                    
        sql = "UPDATE tbl_cash "
        sql = sql + "SET "
        sql = sql + " cash=" & text2.Text & " "
        sql = sql + " WHERE id=" & PK & ""
        CN.Execute sql
        Call cetak_Transfer
        frmCashFlow.RefreshRecords
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_cash WHERE id = " & PK, CN, adOpenStatic, adLockOptimistic
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        Text1.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmCashFlow.RefreshRecords
        ElseIf State = adStatePopupMode Then
        End If
    End If
    Set frmCashFlowAE = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        Call cmdSave_Click
    End If

End Sub
