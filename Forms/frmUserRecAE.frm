VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUserRecAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserRecAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4185
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbkelamin 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   2415
   End
   Begin VB.OptionButton OPUser 
      Caption         =   "User"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2160
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton OpManager 
      Caption         =   "Manager"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.OptionButton OpAdminstrator 
      Caption         =   "Administrator"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Username"
      Top             =   150
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1320
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Password"
      Top             =   525
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -150
      ScaleHeight     =   30
      ScaleWidth      =   12015
      TabIndex        =   4
      Top             =   2355
      Width           =   12015
   End
   Begin MSDataListLib.DataCombo dcGroup 
      Height          =   360
      Left            =   1320
      TabIndex        =   12
      Tag             =   "Group"
      Top             =   2520
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Group"
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Gender"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   1080
      Picture         =   "frmUserRecAE.frx":038A
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2475
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   525
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmUserRecAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    With rs
        txtEntry(0).Text = .Fields("nm_pengguna")
        txtEntry(1).Text = .Fields("password")
        cmbkelamin.Text = .Fields("jk_pengguna")
        If (.Fields("level") = "Administrator") Then
            OpAdminstrator.Value = True
        ElseIf (.Fields("level") = "Manager") Then
            OpManager.Value = True
        Else
            OPUser.Value = True
        End If
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
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(cmbkelamin, True) = True Then Exit Sub
    If OpAdminstrator.Value = True Then
    Else
        If is_empty(dcGroup, True) = True Then Exit Sub
    End If
    
    Dim level As String
    If (OpAdminstrator.Value = True) Then
            level = "Administrator"
    ElseIf (OpManager.Value = True) Then
            level = "Manager"
    Else
            level = "User"
    End If
    
    If State = adStateAddMode Then
        rs.AddNew
        rs.Fields("tgl_input") = Now
        rs.Fields("id_admin") = CurrUser.USER_PK
        rs.Fields("nm_pengguna") = txtEntry(0).Text
        rs.Fields("password") = txtEntry(1).Text
        rs.Fields("jk_pengguna") = cmbkelamin.Text
        rs.Fields("level") = level
        If OpAdminstrator.Value = False Then
            rs.Fields("user_cabang") = dcGroup.BoundText
        End If
        rs.Update
    Else
        sql = "UPDATE tbl_pengguna "
        sql = sql + "SET "
        sql = sql + " nm_pengguna='" & txtEntry(0).Text & "', "
        sql = sql + " password='" & txtEntry(1).Text & "', "
        If OpAdminstrator.Value = False Then
            sql = sql + " user_cabang=" & dcGroup.BoundText & ", "
        End If
        sql = sql + " jk_pengguna='" & cmbkelamin.Text & "', "
        sql = sql + " level='" & level & "' "
        sql = sql + " WHERE id=" & PK
        CN.Execute sql
    End If
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_pengguna WHERE id = " & PK, CN, adOpenStatic, adLockOptimistic
    If State = adStateAddMode Then
        Caption = "Create New Entry"
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    cmbkelamin.AddItem "Pria"
    cmbkelamin.AddItem "Wanita"
    bind_dc "SELECT * FROM tbl_cabang", "nm_cabang", dcGroup, "id_cabang"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmUserRec.CommandPass 5
        End If
    End If
    
    Set frmUserRecAE = Nothing
End Sub

Private Sub OpAdminstrator_Click()
    If OpAdminstrator.Value = True Then
        dcGroup.Enabled = False
    Else
        dcGroup.Enabled = True
    End If
End Sub

Private Sub OpManager_Click()
    If OpAdminstrator.Value = True Then
        dcGroup.Enabled = False
    Else
        dcGroup.Enabled = True
    End If
End Sub

Private Sub OPUser_Click()
    If OpAdminstrator.Value = True Then
        dcGroup.Enabled = False
    Else
        dcGroup.Enabled = True
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub
