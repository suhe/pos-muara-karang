VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDepartementAE 
   BorderStyle     =   0  'None
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   Icon            =   "frmDepartementAE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDepartement 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   200
      TabIndex        =   24
      Top             =   600
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "u"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   960
      MaxLength       =   100
      TabIndex        =   18
      Tag             =   "VN"
      Text            =   "0"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   960
      MaxLength       =   200
      TabIndex        =   6
      Tag             =   "RN"
      Text            =   "0"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   960
      MaxLength       =   100
      TabIndex        =   7
      Tag             =   "PN"
      Text            =   "0"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "Name"
      Text            =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.ComboBox cmbKode 
      Height          =   315
      ItemData        =   "frmDepartementAE.frx":038A
      Left            =   960
      List            =   "frmDepartementAE.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "Kode"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   960
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "An"
      Text            =   "0"
      Top             =   1755
      Width           =   2055
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   960
      MaxLength       =   200
      TabIndex        =   4
      Tag             =   "BN"
      Text            =   "0"
      Top             =   1380
      Width           =   2055
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Nama"
      Top             =   1005
      Width           =   5175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "Name"
      Text            =   "00"
      Top             =   240
      Width           =   405
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   12015
      TabIndex        =   10
      Top             =   3345
      Width           =   12015
   End
   Begin MSDataListLib.DataCombo dcDepartement 
      Height          =   360
      Left            =   960
      TabIndex        =   21
      Tag             =   "Parent Departement"
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check For Enabled Parent"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Parent"
      Height          =   240
      Index           =   8
      Left            =   0
      TabIndex        =   20
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "VN"
      Height          =   240
      Index           =   7
      Left            =   -600
      TabIndex        =   19
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   240
      Index           =   6
      Left            =   1320
      TabIndex        =   17
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "RN"
      Height          =   240
      Index           =   5
      Left            =   -600
      TabIndex        =   16
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "PN"
      Height          =   240
      Index           =   4
      Left            =   -600
      TabIndex        =   15
      Top             =   2880
      Width           =   1485
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "AN"
      Height          =   240
      Index           =   3
      Left            =   -525
      TabIndex        =   14
      Top             =   1755
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "BN"
      Height          =   240
      Index           =   2
      Left            =   -525
      TabIndex        =   13
      Top             =   1380
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Nama"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   12
      Top             =   1005
      Width           =   615
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Kode "
      Height          =   240
      Index           =   0
      Left            =   45
      TabIndex        =   11
      Top             =   270
      Width           =   915
   End
End
Attribute VB_Name = "frmDepartementAE"
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
        If (Left(.Fields("kd_departement"), 1) = 2) Then
            txtEntry(4).Visible = True
        End If
        cmbKode.Enabled = False
        cmbKode.Text = Left(.Fields("kd_departement"), 1)
        dcDepartement.BoundText = .Fields("parent_id")
        txtEntry(0).Text = Mid(.Fields("kd_departement"), 2, 2)
        txtEntry(0).Enabled = False
        txtEntry(4).Text = Mid(.Fields("kd_departement"), 4, 1)
        txtEntry(4).Enabled = False
        txtEntry(1).Text = .Fields("nm_departement")
        txtEntry(2).Text = Val(.Fields("bn"))
        txtEntry(3).Text = Val(.Fields("an"))
        txtEntry(5).Text = Val(.Fields("rn"))
        txtEntry(6).Text = Val(.Fields("pn"))
        txtEntry(7).Text = Val(.Fields("vn"))
    End With
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub Check1_Click()
    Dim rsdep As New Recordset
    Dim rsdepParent As New Recordset
    If Check1.Value = 1 Then
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        cmbKode.Enabled = False
        txtEntry(0).Enabled = False
        If rsdep.State = 1 Then rsdep.Close
        rsdep.CursorLocation = adUseClient
        sql = "SELECT * FROM tbl_departement WHERE parent_id=0 AND kd_departement LIKE '%" & cmbKode.Text & txtEntry(0).Text & "%' ORDER BY kd_departement ASC "
        rsdep.Open sql, CN, adOpenStatic, adLockOptimistic
        If rsdep.RecordCount > 0 Then
            txtDepartement.Text = rsdep.Fields("nm_departement")
            tbl.TABLE_ID_DEPT = rsdep.Fields("id_departement")
            tbl.TABLE_NM_DEPT = rsdep.Fields("nm_departement")
            txtEntry(1).Text = rsdep.Fields("nm_departement")
            
            If rsdepParent.State = 1 Then rsdepParent.Close
            rsdepParent.CursorLocation = adUseClient
            sql = "SELECT * FROM tbl_departement WHERE parent_id=" & rsdep.Fields("id_departement") & " ORDER BY kd_departement DESC "
            rsdepParent.Open sql, CN, adOpenStatic, adLockOptimistic
            If rsdepParent.RecordCount > 0 Then
                txtEntry(4).Text = Val(Right(rsdepParent.Fields("kd_departement"), 1)) + 1
            Else
                txtEntry(4).Text = 1
            End If
        Else
            txtDepartement.Text = "Root Level"
            tbl.TABLE_ID_DEPT = "0"
            tbl.TABLE_NM_DEPT = "0"
        End If
        'bind_dc "SELECT * FROM tbl_departement WHERE parent_id=0 AND kd_departement LIKE '%" & cmbKode.Text & txtEntry(0).Text & "%' ORDER BY kd_departement ASC ", "nm_departement", dcDepartement, "id_departement"
    Else
       cmdSave.Enabled = False
       cmdCancel.Enabled = False
       cmbKode.Enabled = True
       txtEntry(0).Enabled = True
       dcDepartement.Text = ""
       'bind_dc "SELECT * FROM tbl_departement WHERE parent_id=1000 AND kd_departement LIKE '%" & cmbKode.Text & txtEntry(0).Text & "%' ORDER BY kd_departement ASC ", "nm_departement", dcDepartement, "id_departement"
    End If
End Sub

Private Sub cmbKode_Change()
    On Error Resume Next
    If cmbKode.Text <> 1 Then
        cmbKode.Text = 1
    End If
End Sub

Private Sub cmbKode_Click()
    If (cmbKode.Text = 1) Then
        txtEntry(2).Text = 0
        txtEntry(2).Enabled = False
        txtEntry(3).Enabled = False
        txtEntry(3).Text = 0
        txtEntry(4).Visible = False
        dcDepartement.Enabled = False
        Check1.Visible = False
        Label1.Visible = False
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
    Else
        txtEntry(2).Enabled = True
        txtEntry(3).Enabled = True
        txtEntry(4).Visible = True
        dcDepartement.Enabled = True
        Check1.Visible = True
        Label1.Visible = True
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
    End If
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
    If is_empty(txtEntry(5), True) = True Then Exit Sub
    If is_empty(txtEntry(6), True) = True Then Exit Sub
    If is_empty(txtEntry(7), True) = True Then Exit Sub
    'If is_zero(txtEntry(6), True) = True Then Exit Sub
    If is_empty(cmbKode, True) = True Then Exit Sub
    If txtEntry(4).Visible = True Then
        If is_empty(txtEntry(4), True) = True Then Exit Sub
        'If is_empty(txtDepartement, True) = True Then Exit Sub
    End If
    
    Dim kode As String
    If (txtEntry(4).Visible = True) Then
        kode = txtEntry(4).Text
    Else
        kode = ""
    End If
    
    Dim total As Byte
    
    If State = adStateAddMode Or State = adStatePopupMode Then
       
       total = getRecordCount("id_departement", "tbl_departement", "WHERE kd_departement ='" & cmbKode.Text & txtEntry(0).Text & kode & "' ")
        If (total > 0) Then
            MsgBox "This Kode Is Already Exists In Database Please Change this Kode !", vbCritical + vbInformation
            Exit Sub
        End If
    
       With rs
        .AddNew
        .Fields("tgl_input") = Now
        .Fields("id_pengguna") = CurrUser.USER_PK
        .Fields("kd_departement") = cmbKode.Text & txtEntry(0).Text & kode
        .Fields("nm_departement") = txtEntry(1).Text
        If (dcDepartement.Text <> "") Then
            .Fields("parent_id") = dcDepartement.BoundText
        End If
        .Fields("bn") = txtEntry(2).Text
        .Fields("an") = txtEntry(3).Text
        .Fields("rn") = txtEntry(5).Text
        .Fields("pn") = txtEntry(6).Text
        .Fields("vn") = txtEntry(7).Text
        .Fields("group_departement") = cmbKode.Text
        .Update
      End With
   Else
        sql = "UPDATE tbl_departement "
        sql = sql + "SET "
        sql = sql + " nm_departement='" & txtEntry(1).Text & "', "
        sql = sql + " bn=" & Val(txtEntry(2).Text) & ", "
        sql = sql + " an=" & Val(txtEntry(3).Text) & ", "
        sql = sql + " rn=" & Val(txtEntry(5).Text) & ", "
        sql = sql + " vn=" & Val(txtEntry(7).Text) & ", "
        sql = sql + " pn=" & Val(txtEntry(6).Text) & " "
        sql = sql + " WHERE id_departement=" & PK
        CN.Execute sql
    End If
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New Record Has Been Successfully Saved.", vbInformation
        If MsgBox("Do You Want To Add Another New Record?", vbQuestion + vbYesNo) = vbYes Then
            frmDepartement.RefreshRecords
            ResetFields
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_departement WHERE id_departement =" & PK, CN, adOpenStatic, adLockOptimistic
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        frmDepartement.RefreshRecords
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        If (tbl.TABLE_KD_DEPT = 1) Then
            txtEntry(2).Enabled = False
            txtEntry(3).Enabled = False
        Else
            txtEntry(2).Enabled = True
            txtEntry(3).Enabled = True
        End If
    End If
    cmbKode.AddItem "1"
    cmbKode.AddItem "2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmDepartement.RefreshRecords
        ElseIf State = adStatePopupMode Then
            'srcText.Text = rs![nm_kreditor]
            'srcText.Tag = rs![id_kreditor]
        End If
    End If
    Set frmDepartementAE = Nothing
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
     If (Index = 0) Or (Index > 1) Then
        NumberOnly KeyAscii
     End If
End Sub
