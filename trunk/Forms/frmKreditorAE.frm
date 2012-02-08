VERSION 5.00
Begin VB.Form frmKreditorAE 
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   Icon            =   "frmKreditorAE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   15
      Tag             =   "Telepon"
      Text            =   "10000000"
      Top             =   2400
      Width           =   2235
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   12015
      TabIndex        =   8
      Top             =   3195
      Width           =   12015
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "ID"
      Top             =   120
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   600
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   480
      Width           =   5415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1560
      MaxLength       =   200
      TabIndex        =   1
      Tag             =   "Address"
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "City / Town"
      Top             =   1245
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Contact Person"
      Top             =   1620
      Width           =   5415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Telepon"
      Top             =   1995
      Width           =   2355
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Plafon"
      Height          =   240
      Index           =   6
      Left            =   0
      TabIndex        =   16
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "ID"
      Height          =   240
      Index           =   0
      Left            =   525
      TabIndex        =   14
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   240
      Index           =   1
      Left            =   825
      TabIndex        =   13
      Top             =   495
      Width           =   615
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   12
      Top             =   870
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "City/Town"
      Height          =   240
      Index           =   3
      Left            =   75
      TabIndex        =   11
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact Person"
      Height          =   240
      Index           =   4
      Left            =   75
      TabIndex        =   10
      Top             =   1620
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Telepon"
      Height          =   240
      Index           =   5
      Left            =   75
      TabIndex        =   9
      Top             =   1995
      Width           =   1365
   End
End
Attribute VB_Name = "frmKreditorAE"
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
        txtEntry(0).Text = .Fields("id_kreditor")
        txtEntry(1).Text = .Fields("nm_kreditor")
        txtEntry(2).Text = .Fields("tlp_kreditor")
        txtEntry(3).Text = .Fields("kota_kreditor")
        txtEntry(4).Text = .Fields("cp_kreditor")
        txtEntry(5).Text = .Fields("tlp_kreditor")
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
    txtEntry(1).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    If is_empty(txtEntry(3), True) = True Then Exit Sub
    If is_empty(txtEntry(4), True) = True Then Exit Sub
    If is_empty(txtEntry(5), True) = True Then Exit Sub
    If is_empty(txtEntry(6), True) = True Then Exit Sub
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("nm_kreditor") = txtEntry(1).Text
        rs.Fields("almt_kreditor") = txtEntry(2).Text
        rs.Fields("kota_kreditor") = txtEntry(3).Text
        rs.Fields("cp_kreditor") = txtEntry(4).Text
        rs.Fields("tlp_kreditor") = txtEntry(5).Text
        rs.Fields("tgl_input") = Now
        rs.Fields("id_pengguna") = CurrUser.USER_PK
        rs.Update
    Else
        sql = "UPDATE tbl_kreditor "
        sql = sql + "SET "
        sql = sql + " nm_kreditor='" & txtEntry(1).Text & "', "
        sql = sql + " almt_kreditor='" & txtEntry(2).Text & "', "
        sql = sql + " kota_kreditor='" & txtEntry(3).Text & "', "
        sql = sql + " cp_kreditor='" & txtEntry(4).Text & "', "
        sql = sql + " tlp_kreditor='" & txtEntry(5).Text & "' "
        sql = sql + " WHERE id_kreditor=" & PK
        CN.Execute sql
    End If
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            frmKreditor.RefreshRecords
            ResetFields
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        'POP-UP MODE HERE
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        frmKreditor.RefreshRecords
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_kreditor WHERE id_kreditor = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        frmKreditor.RefreshRecords
        txtEntry(1).SetFocus
        If (CurrBiz.BUSINNES_PLAFON) Then
            txtEntry(6).Text = CurrBiz.BUSINNES_PLAFON
        End If
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmKreditor.RefreshRecords
        ElseIf State = adStatePopupMode Then
            'srcText.Text = rs![Name]
            'srcText.Tag = rs![PK]
        End If
    End If
    rs.Close
    Set rs = Nothing
    Set frmKreditorAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 6 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub
