VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKomisiFaktur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Komisi Faktur"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9900
      TabIndex        =   5
      Top             =   6000
      Width           =   9900
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   0
         Width           =   4150
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   60
            Visible         =   0   'False
            Width           =   2535
         End
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9900
      TabIndex        =   4
      Top             =   5985
      Width           =   9900
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9900
      TabIndex        =   3
      Top             =   5970
      Width           =   9900
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   9615
      Begin VB.CommandButton Command1 
         Caption         =   "&Lunasi* (F2)"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "&Keluar (F6)"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   5115
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   9022
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Jual"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No Jual"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tgl Jual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "F.Debitor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "P.Kd Pasien"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Nm Pasien"
         Object.Width           =   3995
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tlp Pasien"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Relasi"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ID Kreditor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Nm Kreditor"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "KD Departement"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Nm Departement"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Nm Cabang"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Payment"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Bayar"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Komisi"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Nm Pengguna"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmKomisiFaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsKomisiFaktur As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser

'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    SQLParser.RestoreStatement
    SQLParser.wCondition = srcCondition
    ReloadRecords SQLParser.SQLStatement
End Sub

Public Sub RefreshRecords()
    SQLParser.RestoreStatement
    ReloadRecords SQLParser.SQLStatement
End Sub

Public Sub ReloadRecords(ByVal srcSQL As String)
    On Error GoTo err
    With rsKomisiFaktur
        If .State = adStateOpen Then .Close
        .Open srcSQL
    End With
    RecordPage.Refresh
    FillList 1
    Exit Sub
err:
        If err.Number = -2147217913 Then
            srcSQL = Replace(srcSQL, "'", "", , , vbTextCompare)
            Resume
        ElseIf err.Number = -2147217900 Then
            MsgBox "Invalid search operation.", vbExclamation
            SQLParser.RestoreStatement
            srcSQL = SQLParser.SQLStatement
            Resume
        Else
            prompt_err err, Name, "ReloadRecords"
        End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdStruck_Click()
        Unload Me
        Call printKomisi
End Sub

Private Sub Command1_Click()
    Dim intResponse As String
    Dim i, total As Integer
    intResponse = MsgBox("Saya bersedia bertanggung jawab dengan mengklik tombol Yes!", vbYesNo + vbInformation, "Warning")
    If intResponse = vbYes Then
          If (lvList.ListItems.Count > 0) Then
            Unload Me
            Call printKomisi
          End If
    Else
            MsgBox "No Item Selected !", vbCritical + vbQuestion
    End If
        Call Deactive
        Unload Me
        On Error Resume Next
        frmKomisi.RefreshRecords
End Sub

Private Sub Deactive()
     With MDIMainMenu
         .tbMenu.Buttons(5).Caption = "Search"
         .tbMenu.Buttons(5).Image = 3
     End With
     tbl.TABLE_SEARCH3 = 0
     tbl.TABLE_SEARCH_FLAG_3 = ""
     tbl.TABLE_SEARCH_DEP = ""
     tbl.TABLE_SEARCH_TANGGAL_3 = ""
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF2: Call Command1_Click
        Case vbKeyF8: Call cmdKeluar_Click
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim subQty As Long
    Dim subtotal, subsubtotal As Double
    'Set the graphics for the controls
    With MDIMainMenu
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With
    
    Dim sql2 As String
    
    sql2 = "j.flag_debitor = 1 AND j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= curdate() "
    If (tbl.TABLE_SEARCH_DEP <> "") Then
        sql2 = sql2 & " AND d.nm_departement LIKE '%" & tbl.TABLE_SEARCH_DEP & "%'"
    End If
    
    If (tbl.TABLE_SEARCH_TANGGAL_3 <> "") Then
        sql2 = sql2 & " AND " & tbl.TABLE_SEARCH_TANGGAL_3
    End If
     
    With SQLParser
            .Fields = "j.id_jual,j.no_jual,j.tgl_jual,j.flag_debitor,j.kd_pasien,p.nm_pasien,p.no_tlp,p.relasi,j.id_kreditor,k.nm_kreditor,d.kd_departement,d.nm_departement,c.nm_cabang,j.type,j.payment,j.bayar,j.komisi,pp.nm_pengguna "
            .Tables = "tbl_jual j INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor INNER JOIN tbl_departement d ON d.id_departement=j.id_departement INNER JOIN tbl_cabang c ON c.id_cabang=j.id_cabang INNER JOIN tbl_pengguna pp ON pp.id=j.id_pengguna"
            .wCondition = sql2
            .SortOrder = "j.id_jual DESC"
            .SaveStatement
    End With
    
    If rsKomisiFaktur.State = 1 Then rsKomisiFaktur.Close
    rsKomisiFaktur.CursorLocation = adUseClient
    rsKomisiFaktur.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsKomisiFaktur, 10000000
        FillList 1
    End With
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsKomisiFaktur, RecordPage.PageStart, RecordPage.PageEnd, 22, 2, False, True, , , , "id_jual")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    lvList_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSalesFaktur = Nothing
End Sub



Private Sub lvList_Click()
    On Error GoTo err
    lblCurrentRecord.Caption = "Selected Record: " & RightSplitUF(lvList.SelectedItem.Tag)
    Exit Sub
err:
        lblCurrentRecord.Caption = "Selected Record: NONE"
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Sort the listview
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lvList.SortOrder = 0
    Else
        lvList.SortOrder = Abs(lvList.SortOrder - 1)
    End If
    lvList.SortKey = ColumnHeader.Index - 1
    
    lvList.Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

