VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearchFaktur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmSearchFaktur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5760
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4440
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Condition "
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.ComboBox cmbBulan2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSearchFaktur.frx":038A
         Left            =   5040
         List            =   "frmSearchFaktur.frx":03B2
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "Month"
         Top             =   720
         Width           =   690
      End
      Begin VB.ComboBox cmbHari2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSearchFaktur.frx":03E6
         Left            =   4320
         List            =   "frmSearchFaktur.frx":03E8
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "Day"
         Top             =   720
         Width           =   690
      End
      Begin VB.ComboBox cmbTahun2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSearchFaktur.frx":03EA
         Left            =   5760
         List            =   "frmSearchFaktur.frx":03EC
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Year"
         Top             =   720
         Width           =   1050
      End
      Begin VB.ComboBox cmbHari 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSearchFaktur.frx":03EE
         Left            =   1680
         List            =   "frmSearchFaktur.frx":03F0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Day"
         Top             =   720
         Width           =   690
      End
      Begin VB.ComboBox cmbTahun 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSearchFaktur.frx":03F2
         Left            =   3120
         List            =   "frmSearchFaktur.frx":03F4
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Year"
         Top             =   720
         Width           =   1050
      End
      Begin VB.ComboBox cmbBulan 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSearchFaktur.frx":03F6
         Left            =   2400
         List            =   "frmSearchFaktur.frx":041E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Month"
         Top             =   720
         Width           =   690
      End
      Begin VB.CheckBox chTanggal 
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Tampilkan Hanya Hutang Kreditor"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tampilkan Semua Record"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   255
         Left            =   6240
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtFilter 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Tag             =   "Insert Name"
         ToolTipText     =   "Ketik Atau Cari dengan Tombol"
         Top             =   360
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   59572225
         CurrentDate     =   38207
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblNama 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   390
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSearchFaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim d, d2, mm, mm2 As Byte
Dim names, tgl, flag As String
Dim Y, y2 As Integer

Public srcColumnHeaders As ColumnHeaders 'Source column headers
Public srcNoOfCol As Long
Public srcform As Form 'Source form

Private Sub chTanggal_Click()
    If chTanggal.Value = 1 Then
        cmbHari.Enabled = True
        cmbBulan.Enabled = True
        cmbTahun.Enabled = True
        cmbHari2.Enabled = True
        cmbBulan2.Enabled = True
        cmbTahun2.Enabled = True
    Else
        cmbHari.Enabled = False
        cmbBulan.Enabled = False
        cmbTahun.Enabled = False
        cmbHari2.Enabled = False
        cmbBulan2.Enabled = False
        cmbTahun2.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim varx  As String
    If is_empty(txtFilter(0), True) = True Then Exit Sub
    'If cmbOperation(0).Text <> "Tanggal" Then If txtFilter(0).Text = "" Then txtFilter(0).SetFocus: Exit Sub
    On Error GoTo err
    If srcform.Name = "frmSales" Then
         names = "k.nm_kreditor"
         tgl = "j.tgl_jual"
         flag = "j.flag_kreditor"
         tbl.TABLE_SEARCH = 1
    ElseIf srcform.Name = "frmPurchase" Then
         names = "s.nm_supplier"
         tgl = "b.tgl_beli"
         flag = "b.flag_supplier"
         tbl.TABLE_SEARCH2 = 1
    ElseIf srcform.Name = "frmKomisi" Then
         names = "d.nm_departement"
         tgl = "j.tgl_jual"
         flag = "j.flag_debitor"
         tbl.TABLE_SEARCH3 = 1
    Else
        MsgBox "Invalid Operation", vbCritical + vbInformation
        Unload Me
    End If
    
    Dim strFilter, strFilter2 As String
    
    If txtFilter(0).Text <> "" Then
        tbl.TABLE_SEARCH_KREDITOR = txtFilter(0).Text
        tbl.TABLE_SEARCH_SUPPLIER = txtFilter(0).Text
        tbl.TABLE_SEARCH_DEP = txtFilter(0).Text
        strFilter = "" & names & ""
        strFilter = strFilter & " LIKE '%" & txtFilter(0).Text & "%'"
    End If
        
    If chTanggal.Value = 1 Then
        tbl.TABLE_SEARCH_TANGGAL = " DATE_FORMAT(" & tgl & ",'%Y-%m-%d') >= '" & cmbTahun.Text & "-" & cmbBulan.Text & "-" & cmbHari.Text & " ' AND DATE_FORMAT(" & tgl & ",'%Y-%m-%d') <= '" & cmbTahun2.Text & "-" & cmbBulan2.Text & "-" & cmbHari2.Text & "'"
        tbl.TABLE_SEARCH_TANGGAL_2 = " DATE_FORMAT(" & tgl & ",'%Y-%m-%d') >= '" & cmbTahun.Text & "-" & cmbBulan.Text & "-" & cmbHari.Text & " ' AND DATE_FORMAT(" & tgl & ",'%Y-%m-%d') <= '" & cmbTahun2.Text & "-" & cmbBulan2.Text & "-" & cmbHari2.Text & "'"
        tbl.TABLE_SEARCH_TANGGAL_3 = " DATE_FORMAT(" & tgl & ",'%Y-%m-%d') >= '" & cmbTahun.Text & "-" & cmbBulan.Text & "-" & cmbHari.Text & " ' AND DATE_FORMAT(" & tgl & ",'%Y-%m-%d') <= '" & cmbTahun2.Text & "-" & cmbBulan2.Text & "-" & cmbHari2.Text & "'"
        If txtFilter(0).Text <> "" Then
            strFilter = strFilter & " AND "
        End If
        tbl.TABLE_TANGGAL_AWAL = cmbTahun.Text & "-" & cmbBulan.Text & "-" & cmbHari.Text
        tbl.TABLE_TANGGAL_AKHIR = cmbTahun2.Text & "-" & cmbBulan2.Text & "-" & cmbHari2.Text
        
        strFilter = strFilter & tbl.TABLE_SEARCH_TANGGAL
    End If
    
     If Option2.Value = True Then
        If txtFilter(0).Text <> "" Or chTanggal.Value = 1 Then
            strFilter = strFilter & " AND"
        End If
        tbl.TABLE_SEARCH_FLAG = flag & "=1 "
        tbl.TABLE_SEARCH_FLAG_2 = flag & "=1 "
        tbl.TABLE_SEARCH_FLAG_3 = flag & "=1 "
        
        If srcform.Name = "frmKomisi" Then
            strFilter = strFilter & " " & flag & "=1 AND j.flag_kreditor=0 "
        Else
            strFilter = strFilter & " " & flag & "=1 "
        End If
    End If
    'MsgBox strFilter
    srcform.FilterRecord strFilter
    strFilter = vbNullString
    Call Active
    Unload Me
    Exit Sub
err:
        If err.Number = -2147352571 Then
            MsgBox "Invalid search operation.", vbExclamation
            Unload Me
        ElseIf err.Number = 3001 Then
            Resume Next
        Else
            prompt_err err, "frmFilter", "cmdOk_Click"
        End If
End Sub

Private Sub Active()
     With MDIMainMenu
         .tbMenu.Buttons(5).Caption = "Lunasi"
         .tbMenu.Buttons(5).Image = 9
     End With
End Sub

Private Sub Command1_Click()
    If srcform.Name = "frmSales" Then
        frmSearchFakturKreditor.show vbModal
    ElseIf srcform.Name = "frmPurchase" Then
        frmSearchFakturSupplier.show vbModal
    ElseIf srcform.Name = "frmKomisi" Then
        frmSearchFakturDepartement.show vbModal
    Else
        MsgBox "Invalid Operation", vbCritical + vbInformation
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    For d = 1 To 31
        If (d <= 9) Then
            d = "0" & d
        End If
        cmbHari.AddItem d
    Next d
    
    'For mm = 1 To 12
    '    If (mm <= 9) Then
    '        mm = "0" & mm
    '    End If
    '    cmbBulan.AddItem mm
    'Next mm
    
    For Y = 1900 To Year(Now)
        cmbTahun.AddItem Y
    Next Y
    
    For d2 = 1 To 31
        If (d2 <= 9) Then
            d2 = "0" & d2
        End If
        cmbHari2.AddItem d2
    Next d2
    
    'For mm2 = 1 To 12 Step mm2 + 1
    '    If (Val(mm2) <= 9) Then
    '        cmbBulan2.AddItem "0" & mm2
    '    Else
    '        cmbBulan2.AddItem mm2
    '    End If
    'Next mm2
    
    For y2 = 1900 To Year(Now)
        cmbTahun2.AddItem y2
    Next y2
    
    cmbHari.Text = "01"
    cmbBulan.Text = "08"
    cmbTahun.Text = "2011"
    cmbHari2.Text = "31"
    cmbBulan2.Text = "08"
    cmbTahun2.Text = "2011"
    
    If srcform.Name = "frmSales" Then
        Option2.Caption = "Tampilkan Hanya Hutang Kreditor"
    ElseIf srcform.Name = "frmPurchase" Then
        Option2.Caption = "Tampilkan Hanya Hutang Supplier"
    ElseIf srcform.Name = "frmKomisi" Then
        Option2.Caption = "Tampilkan Hanya Hutang Debitor"
    Else
        MsgBox "Invalid Operation", vbCritical + vbInformation
        Unload Me
    End If
End Sub
