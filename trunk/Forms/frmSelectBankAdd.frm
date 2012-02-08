VERSION 5.00
Begin VB.Form frmGroupAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Information"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectBankAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "Telepon"
      Text            =   "10000000"
      Top             =   2040
      Width           =   1635
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   150
      TabIndex        =   11
      Tag             =   "Kota"
      Text            =   "0"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   9
      Tag             =   "Cabang"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      MaxLength       =   150
      TabIndex        =   2
      Tag             =   "Kota"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      MaxLength       =   150
      TabIndex        =   1
      Tag             =   "Alamat"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      MaxLength       =   150
      TabIndex        =   0
      Tag             =   "Cabang"
      Top             =   510
      Width           =   2655
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -225
      ScaleHeight     =   30
      ScaleWidth      =   12015
      TabIndex        =   8
      Top             =   2025
      Width           =   12015
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Default Plafon"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1050
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Hari"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   1560
      Width           =   330
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Jangka Waktu"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Kode"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Kota"
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   1260
      Width           =   1050
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Alamat"
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   885
      Width           =   1020
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Cabang"
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   510
      Width           =   1050
   End
End
Attribute VB_Name = "frmGroupAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ADD_STATE        As Boolean
Public PK         As Long

Private Sub Command1_Click()
    If is_empty(text1) = True Then Exit Sub
    If is_empty(Text2) = True Then Exit Sub
    If is_empty(text3) = True Then Exit Sub
    If is_empty(Text4) = True Then Exit Sub
    If is_empty(Text5) = True Then Exit Sub
    If is_empty(text6) = True Then Exit Sub
    
    If ADD_STATE = False Then
        If isRecordExist("tbl_cabang", "id_cabang", PK) = False Then
            MsgBox "This Group is no longer exist in the record. Click ok to reload the records!", vbExclamation, "Unable To Edit"
            frmGroup.reload_rec
            Unload Me
            Exit Sub
        End If
    End If
    
    On Error GoTo err
    With frmGroup.rsGroup
        If ADD_STATE = True Then
            .AddNew
            .Fields("kd_cabang") = Text4.Text
            .Fields("nm_cabang") = text1.Text
            .Fields("almt_cabang") = Text2.Text
            .Fields("kota_cabang") = text3.Text
            .Fields("jw_waktu") = Text5.Text
            .Fields("plafon_default") = text6.Text
            .Fields("tgl_input") = Now
            .Fields("id_pengguna") = CurrUser.USER_PK
            .Update
        Else
            sql = "UPDATE tbl_cabang "
            sql = sql + " SET "
            sql = sql + " nm_cabang='" & text1.Text & "', "
            sql = sql + " almt_cabang='" & Text2.Text & "', "
            sql = sql + " jw_waktu='" & Text5.Text & "', "
            sql = sql + " plafon_default='" & text6.Text & "', "
            sql = sql + " kota_cabang='" & text3.Text & "' "
            sql = sql + " WHERE id_cabang=" & PK
            CN.Execute sql
        End If
    End With
    frmGroup.reload_rec
    Unload Me
    Exit Sub
err:
        prompt_err err, Me.Name, "Command1_Click"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If ADD_STATE = True Then
       Caption = "Add New"
       Text4.Enabled = True
    Else
        Caption = "Edit Existing"
        customMove frmGroup.rsGroup, False, PK, "id_cabang"
        With frmGroup.rsGroup
            Text4.Text = .Fields("kd_cabang")
            text1.Text = .Fields("nm_cabang")
            Text2.Text = .Fields("almt_cabang")
            text3.Text = .Fields("kota_cabang")
            Text5.Text = .Fields("jw_waktu")
            text6.Text = .Fields("plafon_default")
            Text4.Enabled = False
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGroupAE = Nothing
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub
