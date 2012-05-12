VERSION 5.00
Begin VB.Form frmSupplierAE 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSupplierAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1350
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Negara"
      Top             =   2025
      Width           =   2370
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "Kota"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Contact Person"
      Top             =   1275
      Width           =   5175
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1350
      MaxLength       =   200
      TabIndex        =   2
      Tag             =   "Telepon"
      Top             =   900
      Width           =   1935
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Alamat"
      Top             =   525
      Width           =   5175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   120
      Width           =   5205
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -150
      ScaleHeight     =   30
      ScaleWidth      =   12015
      TabIndex        =   14
      Top             =   3225
      Width           =   12015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "No.Account"
      Height          =   240
      Index           =   5
      Left            =   -75
      TabIndex        =   13
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Nm.Account"
      Height          =   240
      Index           =   4
      Left            =   -75
      TabIndex        =   12
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact Person"
      Height          =   240
      Index           =   3
      Left            =   -75
      TabIndex        =   11
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Telepon"
      Height          =   240
      Index           =   2
      Left            =   -75
      TabIndex        =   10
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Alamat"
      Height          =   240
      Index           =   1
      Left            =   675
      TabIndex        =   9
      Top             =   525
      Width           =   615
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   8
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "frmSupplierAE"
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

Private Sub DisplayForEditing()
    On Error GoTo err
    With rs
        txtEntry(0).Text = .Fields("nm_supplier")
        txtEntry(1).Text = .Fields("almt_supplier")
        txtEntry(2).Text = .Fields("tlp_supplier")
        txtEntry(3).Text = .Fields("cp_supplier")
        txtEntry(4).Text = .Fields("kota_supplier")
        txtEntry(5).Text = .Fields("negara_supplier")
    End With
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    On Error Resume Next
    clearText Me
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    If is_empty(txtEntry(3), True) = True Then Exit Sub
    If is_empty(txtEntry(4), True) = True Then Exit Sub
    If is_empty(txtEntry(5), True) = True Then Exit Sub
    
    If State = adStateAddMode Or State = adStatePopupMode Then
       With rs
        .AddNew
        .Fields("tgl_input") = Now
        .Fields("id_pengguna") = CurrUser.USER_PK
        .Fields("nm_supplier") = txtEntry(0).Text
        .Fields("almt_supplier") = txtEntry(1).Text
        .Fields("tlp_supplier") = txtEntry(2).Text
        .Fields("cp_supplier") = txtEntry(3).Text
        .Fields("kota_supplier") = txtEntry(4).Text
        .Fields("negara_supplier") = txtEntry(5).Text
        .Update
      End With
   Else
        sql = "UPDATE tbl_supplier "
        sql = sql + "SET "
        sql = sql + " nm_supplier='" & txtEntry(0).Text & "', "
        sql = sql + " almt_supplier='" & txtEntry(1).Text & "', "
        sql = sql + " tlp_supplier='" & txtEntry(2).Text & "', "
        sql = sql + " cp_supplier='" & txtEntry(3).Text & "', "
        sql = sql + " kota_supplier='" & txtEntry(4).Text & "', "
        sql = sql + " negara_supplier='" & txtEntry(5).Text & "' "
        sql = sql + " WHERE id_supplier=" & PK
        CN.Execute sql
    End If
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New Record Has Been Successfully Saved.", vbInformation
        If MsgBox("Do You Want To Add Another New Record?", vbQuestion + vbYesNo) = vbYes Then
            frmSupplier.RefreshRecords
            ResetFields
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        'POP-UP MODE HERE
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_supplier WHERE id_supplier =" & PK, CN, adOpenStatic, adLockOptimistic
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        frmSupplier.RefreshRecords
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmSupplier.RefreshRecords
        ElseIf State = adStatePopupMode Then
            
        End If
    End If
    Set frmSupplierAE = Nothing
End Sub

