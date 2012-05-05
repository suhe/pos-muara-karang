VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCashierCNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Faktur"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "frmCashierCNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar (F6)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   315
      Left            =   4800
      TabIndex        =   19
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Cetak Faktur (F3)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Simpan (F2)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblNmKreditor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...................................................."
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
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Width           =   2340
      End
      Begin VB.Label lblCodeKreditor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         TabIndex        =   14
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         TabIndex        =   13
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   3960
         Width           =   765
      End
   End
   Begin VB.Frame fraCustomer 
      Caption         =   "Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5535
      Begin VB.CommandButton cmdOld 
         Caption         =   "Old"
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcDepartement 
         Height          =   360
         Left            =   720
         TabIndex        =   23
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dept."
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
         TabIndex        =   22
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Left            =   360
         TabIndex        =   9
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Kode "
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
         TabIndex        =   8
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         TabIndex        =   7
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblCodeCust 
         AutoSize        =   -1  'True
         Caption         =   "..."
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
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblNamaCust 
         AutoSize        =   -1  'True
         Caption         =   "..."
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
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   135
      End
   End
   Begin VB.Frame fraFaktur 
      Caption         =   "Faktur Jual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtFak 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Number"
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
         TabIndex        =   2
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ditanggung Kreditor"
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
      Left            =   480
      TabIndex        =   21
      Top             =   3120
      Width           =   1710
   End
End
Attribute VB_Name = "frmCashierCNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK      As Long
Dim rs As New Recordset

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
End Sub

Private Sub cmdNew_Click()
    frmPasienAE.State = adStateAddMode
    frmPasienAE.show vbModal
End Sub

Private Sub cmdOld_Click()
    frmCashierCustomer.show vbModal
End Sub

Private Sub cmdPrint_Click()
    Dim Lines As Integer, Y As Long, OutStr As String
    'On Error Resume Nex
    On Error GoTo opps
  '---- set jenis kertas jika paper=0 tidak ada yang akan dicetak.
        'If Paper <> "0" Then
        '    Printer.Font.Size = 10
        '    Printer.Font.Name = "control"
        '    Printer.Print Paper
        'End If
   '--- set font dan size, font
     
     With Printer
     If (lblCodeCust.Caption <> "...") Then
        '.Cls
        .Font.Name = "Times New Roman"
        .Font.Size = 8
        .CurrentY = .CurrentY + 200 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " No Faktur"; Spc(6); ":"; Spc(5); " " & txtFak.Text & " "; Tab(40); " Tanggal "; Spc(3); ":"; Spc(2); ""; Format(Now, "DD/MM/YYYY"); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Pasien"; Spc(3); ":"; Spc(5); " " & lblCodeCust.Caption & " "
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Pasien "; Spc(2); ":"; Spc(5); "" & lblNamaCust.Caption & "";
        Printer.Print ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Dept. "; Spc(3); ":"; Spc(5); "" & dcDepartement.Text & "";
        Printer.Print ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Kreditor "; Spc(0); ":"; Spc(5); "" & lblNmKreditor.Caption & "";
        .EndDoc
        Else
            MsgBox "No Data"
        End If
     End With
opps:
    'MB_Options = vbCritical
    'MsgBox "Printer Error.", MB_Options, "Error Message"
End Sub

Private Sub cmdProcess_Click()
    If is_empty(txtFak, True) = True Then Exit Sub
    If is_empty(dcDepartement, True) = True Then Exit Sub
    Call cash
    'MDIMainMenu.UpdateInfoMsg
    cmdPrint.SetFocus
End Sub

Private Sub cash()
    Dim i As Integer
    Dim subtotal As Double
    Dim intResponse As Integer
    With rs
        .AddNew
        .Fields("no_jual") = Trim(txtFak.Text)
        .Fields("tgl_jual") = Now
        .Fields("id_departement") = Trim(dcDepartement.BoundText)
        .Fields("kd_pasien") = Trim(lblCodeCust.Caption)
        .Fields("id_cabang") = CurrBiz.BUSINNES_GROUP
        If (lblCodeKreditor.Caption <> "...") Then
            .Fields("id_kreditor") = lblCodeKreditor.Caption
        End If
        .Fields("tgl_input") = Now
        .Fields("tgl_akhir") = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
        .Fields("id_pengguna") = CurrUser.USER_PK
        .Update
        End With
        rs.Close
        tbl.TABLE_NO_FAK = Trim(txtFak.Text)
        tbl.TABLE_KD_PASIEN = Trim(lblCodeCust.Caption)
        tbl.TABLE_NM_PASIEN = Trim(lblNamaCust.Caption)
        tbl.TABLE_NM_DEPT = Trim(dcDepartement.Text)
        tbl.TABLE_ID_KREDITUR = Trim(lblCodeKreditor.Caption)
        tbl.TABLE_NM_KREDITUR = lblNmKreditor.Caption
        tbl.TABLE_TANGGAL = Format(Now, "DD-MM-YYYY")
        cmdPrint.Enabled = True
        cmdProcess.Enabled = False
        fraCustomer.Enabled = False
        Frame1.Enabled = False
        intResponse = MsgBox("You Successfull Purchasing !", vbYes + vbInformation, "Warning")
End Sub

Private Sub Command1_Click()
    frmCashierKreditor.show vbModal
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub dcDepartement_Click(Area As Integer)
    If (dcDepartement.Text <> "") Then
        If (lblCodeCust.Caption = "...") Then: MsgBox "Lengkapi Kode pasien !", vbCritical + vbInformation: Exit Sub
        If (lblNamaCust.Caption = "...") Then: MsgBox "Lengkapi Nama pasien !", vbCritical + vbInformation: Exit Sub
            If (lblCodeCust.Caption <> "...") And (lblNamaCust.Caption <> "") Then
                cmdProcess.Enabled = True
            End If
        Else
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call GeneratePK
     rs.Open "SELECT * FROM tbl_jual WHERE no_jual=" & PK, CN, adOpenStatic, adLockOptimistic
     bind_dc "SELECT * FROM tbl_departement", "nm_departement", dcDepartement, "id_departement"
End Sub

Private Sub GeneratePK()
    On Error Resume Next
    PK = getIndex("id_jual", "tbl_jual")
    txtFak.Text = "K-" & PK
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set frmCashierCNew = Nothing
End Sub
