VERSION 5.00
Begin VB.Form frmCashierCNewPasien 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   405
      Index           =   3
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   35
      Tag             =   "Tlp"
      Top             =   3120
      Width           =   690
   End
   Begin VB.ComboBox cmbRegion 
      Height          =   315
      ItemData        =   "frmCashierCNewPasien.frx":0000
      Left            =   1560
      List            =   "frmCashierCNewPasien.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Region"
      Top             =   1200
      Width           =   2130
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   13
      Tag             =   "ID"
      Top             =   120
      Width           =   1965
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1590
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   495
      Width           =   5415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1590
      MaxLength       =   200
      TabIndex        =   1
      Tag             =   "Address"
      Top             =   870
      Width           =   5415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1590
      MaxLength       =   20
      TabIndex        =   4
      Tag             =   "HP"
      Top             =   2010
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1590
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Relation"
      Top             =   1650
      Width           =   2490
   End
   Begin VB.ComboBox cmdGender 
      Height          =   315
      ItemData        =   "frmCashierCNewPasien.frx":0004
      Left            =   1545
      List            =   "frmCashierCNewPasien.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "Gender"
      Top             =   2760
      Width           =   1290
   End
   Begin VB.ComboBox cmbHari 
      Height          =   315
      ItemData        =   "frmCashierCNewPasien.frx":0020
      Left            =   1560
      List            =   "frmCashierCNewPasien.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Day"
      Top             =   3210
      Width           =   690
   End
   Begin VB.ComboBox cmbTahun 
      Height          =   315
      ItemData        =   "frmCashierCNewPasien.frx":0024
      Left            =   3000
      List            =   "frmCashierCNewPasien.frx":0026
      TabIndex        =   9
      Tag             =   "Year"
      Text            =   "cmbTahun"
      Top             =   3210
      Width           =   1050
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "Place"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "Tlp"
      Top             =   2370
      Width           =   2490
   End
   Begin VB.ComboBox cmbBulan 
      Height          =   315
      ItemData        =   "frmCashierCNewPasien.frx":0028
      Left            =   2280
      List            =   "frmCashierCNewPasien.frx":002A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Month"
      Top             =   3210
      Width           =   690
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Age :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   20
      Left            =   4800
      TabIndex        =   34
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   19
      Left            =   3720
      TabIndex        =   33
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   18
      Left            =   2880
      TabIndex        =   32
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   17
      Left            =   4200
      TabIndex        =   31
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   12
      Left            =   4080
      TabIndex        =   30
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   11
      Left            =   3720
      TabIndex        =   29
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   10
      Left            =   7080
      TabIndex        =   28
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "*"
      Height          =   240
      Index           =   16
      Left            =   7080
      TabIndex        =   27
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   " ID"
      Height          =   240
      Index           =   0
      Left            =   615
      TabIndex        =   26
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   240
      Index           =   1
      Left            =   915
      TabIndex        =   25
      Top             =   495
      Width           =   615
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   2
      Left            =   165
      TabIndex        =   24
      Top             =   870
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Region"
      Height          =   240
      Index           =   3
      Left            =   165
      TabIndex        =   23
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Relation"
      Height          =   240
      Index           =   6
      Left            =   165
      TabIndex        =   22
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "HP"
      Height          =   240
      Index           =   7
      Left            =   165
      TabIndex        =   21
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   240
      Index           =   8
      Left            =   360
      TabIndex        =   20
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Gender"
      Height          =   240
      Index           =   9
      Left            =   360
      TabIndex        =   19
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "DD"
      Height          =   240
      Index           =   13
      Left            =   1560
      TabIndex        =   18
      Top             =   3570
      Width           =   495
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "MM"
      Height          =   240
      Index           =   14
      Left            =   2280
      TabIndex        =   17
      Top             =   3570
      Width           =   375
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "YYYY"
      Height          =   240
      Index           =   15
      Left            =   3240
      TabIndex        =   16
      Top             =   3570
      Width           =   375
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Place"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   3810
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Tlp"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2370
      Width           =   1365
   End
End
Attribute VB_Name = "frmCashierCNewPasien"
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
        txtEntry(3).Text = .Fields("kota")
        txtEntry(4).Text = .Fields("relasi")
        txtEntry(5).Text = .Fields("no_hp")
        txtEntry(6).Text = .Fields("no_tlp")
        txtEntry(7).Text = .Fields("tmpt_lahir")
        cmbTahun.Text = Year(.Fields("tgl_lahir"))
        cmbBulan.Text = Month(.Fields("tgl_lahir"))
        cmbHari.Text = Day(.Fields("tgl_lahir"))
        cmdGender.Text = .Fields("jk_pasien")
    End With
    txtEntry(0).Enabled = False
    txtEntry(1).Enabled = False
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmbTahun_Click()
    On Error Resume Next
    Dim age As Byte
    If cmbTahun.Text <> "" Then
        age = Val(Year(Date)) - Val(cmbTahun.Text)
        txtEntry(3).Text = age
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    txtEntry(1).SetFocus
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
                .Fields("no_hp") = txtEntry(5).Text
            Else
                .Fields("no_hp") = "0"
            End If
            
            If txtEntry(6).Text <> "" Then
                .Fields("no_tlp") = txtEntry(6).Text
            Else
                .Fields("no_tlp") = "0"
            End If
            
            .Fields("tgl_lahir") = cmbTahun.Text & "-" & cmbBulan.Text & "-" & cmbHari.Text
            .Fields("jk_pasien") = cmdGender.Text
            .Update
        End With
    End If
    frmCashierCNew.lblCodeCust.Caption = txtEntry(0).Text
    frmCashierCNew.lblNamaCust.Caption = txtEntry(1).Text
    With frmCashier
        .lblKdPasien.Caption = txtEntry(0).Text
        .lblNmPasien.Caption = txtEntry(1).Text
        .lblTlpPasien.Caption = txtEntry(6).Text
        .lblAlmtPasien.Caption = txtEntry(2).Text
        .lblRelasi.Caption = Val(Year(Date)) - Val(cmbTahun.Text)
    End With
    Unload Me
End Sub


Private Function getPK(ByVal srcTable As String) As String
    On Error GoTo err
    Dim rsKode As New Recordset
    Dim RI As Long
    Dim pl As String
    rsKode.CursorLocation = adUseClient
    rsKode.Open "SELECT TRIM(pl_pasien) as pl_pasien,TRIM(pk_pasien) as pk_pasien FROM tbl_pasien WHERE TRIM(pl_pasien) = '" & srcTable & "' ORDER BY ABS(pk_pasien) DESC ", CN, adOpenStatic, adLockOptimistic
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
    Set frmCashierCNewPasien = Nothing
End Sub

Private Sub txtEntry_Change(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        If State = adStateAddMode Or State = adStatePopupMode Then
            txtEntry(0).Text = getPK(UCase(Left(txtEntry(1).Text, 1)))
        End If
    End If
    
    If Index = 3 Then
        Dim age As Integer
        If txtEntry(3).Text <> "" Then
            age = Val(Year(Date)) - Val(txtEntry(3).Text)
            cmbTahun.Text = age
        Else
            age = Val(Year(Date)) - Val(txtEntry(3).Text)
            cmbTahun.Text = age
        End If
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 7 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 5 Or Index = 6 Then
        KeyAscii = isNumber(KeyAscii)
    End If
    
    If Index = 3 Then
        Dim age As Integer
        If txtEntry(3).Text <> "" Then
            age = Val(Year(Date)) - Val(txtEntry(3).Text)
            cmbTahun.Text = age
        End If
    End If
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 7 Then cmdSave.Default = True
End Sub
