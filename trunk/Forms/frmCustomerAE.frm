VERSION 5.00
Begin VB.Form frmPasienAE 
   BorderStyle     =   0  'None
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomerAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRegion 
      Height          =   315
      ItemData        =   "frmCustomerAE.frx":0A02
      Left            =   1440
      List            =   "frmCustomerAE.frx":0A04
      TabIndex        =   28
      Tag             =   "Region"
      Top             =   1320
      Width           =   2130
   End
   Begin VB.ComboBox cmbBulan 
      Height          =   315
      ItemData        =   "frmCustomerAE.frx":0A06
      Left            =   2160
      List            =   "frmCustomerAE.frx":0A08
      TabIndex        =   7
      Tag             =   "Month"
      Text            =   "cmbBulan"
      Top             =   3120
      Width           =   690
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   4
      Tag             =   "Tlp"
      Top             =   2400
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "Place"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.ComboBox cmbTahun 
      Height          =   315
      ItemData        =   "frmCustomerAE.frx":0A0A
      Left            =   2880
      List            =   "frmCustomerAE.frx":0A0C
      TabIndex        =   8
      Tag             =   "Year"
      Text            =   "cmbTahun"
      Top             =   3120
      Width           =   1050
   End
   Begin VB.ComboBox cmbHari 
      Height          =   315
      ItemData        =   "frmCustomerAE.frx":0A0E
      Left            =   1440
      List            =   "frmCustomerAE.frx":0A10
      TabIndex        =   6
      Tag             =   "Day"
      Text            =   "cmbHari"
      Top             =   3120
      Width           =   690
   End
   Begin VB.CheckBox Check3 
      Caption         =   "MEMBER TUUNECA "
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "MEMBER HOUSE OF TAAJ JAKARTA"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "UMUM"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -3000
      ScaleHeight     =   30
      ScaleWidth      =   11565
      TabIndex        =   23
      Top             =   3345
      Width           =   11565
   End
   Begin VB.ComboBox cmdGender 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Gender"
      Top             =   2760
      Width           =   1290
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "Relation"
      Top             =   1680
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "HP"
      Top             =   2040
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   200
      TabIndex        =   1
      Tag             =   "Address"
      Top             =   900
      Width           =   5655
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   525
      Width           =   5655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   27
      Tag             =   "ID"
      Top             =   150
      Width           =   1965
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Kode"
      Default         =   -1  'True
      Height          =   315
      Left            =   840
      TabIndex        =   26
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   14
      Left            =   2880
      TabIndex        =   34
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   13
      Left            =   3960
      TabIndex        =   33
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   16
      Left            =   3960
      TabIndex        =   32
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   12
      Left            =   3480
      TabIndex        =   31
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   11
      Left            =   7080
      TabIndex        =   30
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   10
      Left            =   7080
      TabIndex        =   29
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Tlp"
      Height          =   240
      Index           =   4
      Left            =   -120
      TabIndex        =   25
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Place"
      Height          =   240
      Index           =   5
      Left            =   0
      TabIndex        =   24
      Top             =   4680
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Gender"
      Height          =   240
      Index           =   9
      Left            =   120
      TabIndex        =   22
      Top             =   2790
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   240
      Index           =   8
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "HP"
      Height          =   240
      Index           =   7
      Left            =   -75
      TabIndex        =   20
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Relation"
      Height          =   240
      Index           =   6
      Left            =   -75
      TabIndex        =   19
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "City/Town"
      Height          =   240
      Index           =   3
      Left            =   -75
      TabIndex        =   18
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   2
      Left            =   -75
      TabIndex        =   17
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   240
      Index           =   1
      Left            =   675
      TabIndex        =   16
      Top             =   525
      Width           =   615
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   " ID"
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   15
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "frmPasienAE"
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
Dim rsKode                  As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    With rs
        txtEntry(0).Text = .Fields("kd_pasien")
        txtEntry(1).Text = .Fields("nm_pasien")
        txtEntry(2).Text = .Fields("alamat")
        txtEntry(4).Text = .Fields("relasi")
        txtEntry(5).Text = .Fields("no_hp")
        txtEntry(6).Text = .Fields("no_tlp")
        cmbRegion.Text = .Fields("kota")
        cmdGender.Text = .Fields("jk_pasien")
        cmbTahun.Text = Year(.Fields("tgl_lahir"))
        cmbBulan.Text = Month(.Fields("tgl_lahir"))
        cmbHari.Text = Day(.Fields("tgl_lahir"))
    End With
    txtEntry(0).Enabled = False
    txtEntry(1).Enabled = False
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    txtEntry(1).SetFocus
End Sub


Private Sub cmdGender_Change()
    On Error Resume Next
    If cmdGender.Text <> "Pria" Then
        cmdGender.Text = "Pria"
    End If
End Sub

Private Sub cmdSave_Click()
    'On Error Resume Next
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    If is_empty(cmbRegion, True) = True Then Exit Sub
    If is_empty(txtEntry(4), True) = True Then Exit Sub
    'If is_empty(txtEntry(5), True) = True Then Exit Sub
    'If is_empty(txtEntry(6), True) = True Then Exit Sub
    If is_empty(cmdGender, True) = True Then Exit Sub
    If is_empty(cmbHari, True) = True Then Exit Sub
    If is_empty(cmbBulan, True) = True Then Exit Sub
    If is_empty(cmbTahun, True) = True Then Exit Sub
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        With rs
            .AddNew
            .Fields("tgl_input") = Now
            .Fields("id_pengguna") = CurrUser.USER_PK
            .Fields("kd_pasien") = txtEntry(0).Text
            .Fields("nm_pasien") = txtEntry(1).Text
            .Fields("pk_pasien") = Mid(txtEntry(0).Text, 2, 5)
            .Fields("pl_pasien") = UCase(Left(txtEntry(1).Text, 1))
            .Fields("alamat") = txtEntry(2).Text
            .Fields("kota") = cmbRegion.Text
            .Fields("relasi") = txtEntry(4).Text
            If txtEntry(5).Text <> "" Then
                .Fields("no_hp") = Trim(txtEntry(5).Text)
            Else
                .Fields("no_hp") = "0"
            End If
            
            '.Fields("no_hp") = txtEntry(5).Text
            If txtEntry(6).Text <> "" Then
                .Fields("no_tlp") = Trim(txtEntry(6).Text)
            Else
                .Fields("no_tlp") = "0"
            End If
            '.Fields("no_tlp") = txtEntry(6).Text
            .Fields("tgl_lahir") = cmbTahun.Text & "-" & cmbBulan.Text & "-" & cmbHari.Text
            .Fields("jk_pasien") = cmdGender.Text
            .Update
        End With
    Else
            Dim tlp, hp As String
            
            If txtEntry(5).Text <> "" Then
                hp = txtEntry(5).Text
            Else
                hp = "0"
            End If
            
            
            If txtEntry(6).Text <> "" Then
                tlp = txtEntry(6).Text
            Else
                tlp = "0"
            End If
            
            sql = "UPDATE tbl_pasien "
            sql = sql + "SET "
            sql = sql + " nm_pasien='" & txtEntry(1).Text & "', "
            sql = sql + " alamat='" & txtEntry(2).Text & "', "
            sql = sql + " kota='" & cmbRegion.Text & "', "
            sql = sql + " relasi='" & txtEntry(4).Text & "', "
            sql = sql + " no_hp='" & Trim(hp) & "', "
            sql = sql + " no_tlp='" & Trim(tlp) & "', "
            sql = sql + " tgl_lahir='" & cmbTahun.Text & "-" & cmbBulan.Text & "-" & cmbHari.Text & "', "
            sql = sql + " jk_pasien='" & cmdGender.Text & "' "
            sql = sql + " WHERE id_pasien=" & PK
            CN.Execute sql
    End If
    
    HaveAction = True
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            frmPasien.RefreshRecords
            ResetFields
            txtEntry(1).SetFocus
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub


Private Function getPK(ByVal srcTable As String) As String
    On Error GoTo err
    Dim rsKode As New Recordset
    Dim RI As Long
    Dim pl As String
    rsKode.CursorLocation = adUseClient
    rsKode.Open "SELECT * FROM tbl_pasien WHERE pl_pasien = '" & srcTable & "' ORDER BY ABS(pk_pasien) DESC ", CN, adOpenStatic, adLockOptimistic
    If (rsKode.RecordCount > 0) Then
        pl = rsKode.Fields("pl_pasien")
        RI = Val(rsKode.Fields("pk_pasien")) + 1
    Else
        pl = srcTable
        RI = 1
    End If
    getPK = pl & RI
    Set rsKode = Nothing
    Exit Function
err:
        If err.Number = 94 Then getPK = 1: Resume Next
End Function


Private Sub Form_Load()
    On Error Resume Next
    cmdGender.AddItem "Pria"
    cmdGender.AddItem "Wanita"
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_pasien WHERE id_pasien = " & PK, CN, adOpenStatic, adLockOptimistic
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    Dim d, mm As Byte
    Dim Y As Integer
    
    For d = 1 To 31
        If (d <= 9) Then
            d = "0" & d
        End If
        cmbHari.AddItem d
    Next d
    
    For mm = 1 To 12
        cmbBulan.AddItem mm
    Next mm
    
    For Y = 1900 To Year(Now)
        cmbTahun.AddItem Y
    Next Y
    'cmbHari.Text = Day(Now)
    'cmbBulan.Text = Month(Now)
    'cmbTahun.Text = Year(Now)
    
    With cmbRegion
        .AddItem "Muara Karang"
        .AddItem "Tanjung Periok"
        .AddItem "Pluit"
        .AddItem "Pantai Mutiara"
        .AddItem "Muara Angke"
        .AddItem "pengasinan"
        .AddItem "Pantai Indah Kapuk"
        .AddItem "Lainnya"
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmPasien.RefreshRecords
        ElseIf State = adStatePopupMode Then
            'srcText.Text = txtEntry(0).Text
            'srcText.Tag = PK
            On Error Resume Next
        End If
        'MDIMainMenu.UpdateInfoMsg
    End If
    Set frmPasienAE = Nothing
End Sub

Private Sub txtEntry_Change(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        If State = adStateAddMode Or State = adStatePopupMode Then
            txtEntry(0).Text = getPK(UCase(Left(txtEntry(1).Text, 1)))
        End If
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 7 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If Index = 1 Then
        Call txtEntry_Change(1)
    End If
    
    If Index = 5 Or Index = 6 Then
        KeyAscii = isNumber(KeyAscii)
    End If
    
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 7 Then cmdSave.Default = True
    On Error Resume Next
    If Index = 1 Then
        Call txtEntry_Change(1)
    End If
End Sub
