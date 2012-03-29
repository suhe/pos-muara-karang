VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSales 
   Caption         =   "Sales"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   Icon            =   "frmSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   7455
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7455
      TabIndex        =   2
      Top             =   3855
      Width           =   7455
      Begin VB.ComboBox cbSort 
         Height          =   315
         ItemData        =   "frmSales.frx":038A
         Left            =   4800
         List            =   "frmSales.frx":0394
         TabIndex        =   14
         Text            =   "No.Faktur"
         Top             =   30
         Width           =   1335
      End
      Begin VB.ComboBox cbShow 
         Height          =   315
         ItemData        =   "frmSales.frx":03B4
         Left            =   3600
         List            =   "frmSales.frx":03B6
         TabIndex        =   13
         Text            =   "30"
         Top             =   30
         Width           =   735
      End
      Begin VB.ComboBox cbSortType 
         Height          =   315
         ItemData        =   "frmSales.frx":03B8
         Left            =   6120
         List            =   "frmSales.frx":03C2
         TabIndex        =   12
         Text            =   "DESC"
         Top             =   30
         Width           =   855
      End
      Begin VB.Label lbltotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Record : 0"
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sort"
         Height          =   195
         Left            =   4440
         TabIndex        =   16
         Top             =   60
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Show"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   60
         Width           =   390
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   3
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
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   3840
      Width           =   7455
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   3825
      Width           =   7455
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3435
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
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
      NumItems        =   22
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Jual"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No Fak"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tgl Jual"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tgl Lns Kd"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "F.KR"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "F.DE"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tgl Ln Dep"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "P.Kd Pasien"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Nm Pasien"
         Object.Width           =   3995
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Tlp Pasien"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Relasi"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ID Kreditor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Nm Kreditor"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "KD Departement"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Nm Departement"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Payment"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Piutang"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Sisa"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Komisi"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Laba Bersih"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Nm Pengguna"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.ComboBox cbYear2 
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cbMonth2 
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cbDay2 
      Height          =   315
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cbYear1 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cbMonth1 
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cbDay1 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   -240
      Top             =   0
      Width           =   7515
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsSales As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser

'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    SQLParser.RestoreStatement
    SQLParser.wCondition = srcCondition
    ReloadRecords SQLParser.SQLStatement
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    Select Case srcPerformWhat
        Case "Search"
            If tbl.TABLE_SEARCH = 1 Then
                With frmSalesFaktur
                    .show vbModal
                End With
            Else
                With frmSearchFaktur
                    Set .srcform = Me
                    Set .srcColumnHeaders = lvList.ColumnHeaders
                    .show vbModal
                End With
            End If
        Case "Refresh"
             tbl.TABLE_ID_KREDITUR = ""
             tbl.TABLE_TANGGAL_AWAL = ""
             tbl.TABLE_TANGGAL_AKHIR = ""
            Call Deactive
            RefreshRecords
        Case "Print"
            If lvList.ListItems.Count > 0 Then
                 RefreshRecords
                 Call printSalesSummary
              Else
                 MsgBox "Data Is empty", vbOKOnly + vbCritical, "Warning"
            End If
        Case "Close"
            Unload Me
    End Select
End Sub

Public Sub RefreshRecords()
    SQLParser.RestoreStatement
    ReloadRecords SQLParser.SQLStatement
End Sub

Public Sub ReloadRecords(ByVal srcSQL As String)
    On Error GoTo err
    With rsSales
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

Private Sub btnFirst_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillList 1
End Sub

Private Sub btnLast_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillList RecordPage.PAGE_TOTAL
End Sub

Private Sub btnNext_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillList RecordPage.PAGE_NEXT
End Sub

Private Sub btnPrev_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillList RecordPage.PAGE_PREVIOUS
End Sub

Private Sub btnRecOp_Click()
    frmCustomerRecOp.show vbModal
End Sub

Private Sub cbShow_Change()
    cbShow.Text = "30"
End Sub

Private Sub cbShow_Click()
    Call Form_Load
End Sub

Private Sub cbSort_Change()
    cbSort.Text = "No.Faktur"
End Sub

Private Sub cbSort_Click()
    Call Form_Load
End Sub

Private Sub cbSortType_Change()
    cbSortType.Text = "ASC"
End Sub

Private Sub cbSortType_Click()
    Call Form_Load
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fftfttt"
    CurrBiz.BUSINNES_SALE = 1
End Sub

Private Sub Form_Deactivate()
    MDIMainMenu.HideTBButton "", True
    CurrBiz.BUSINNES_SALE = 0
    Call Deactive
End Sub

Private Sub Deactive()
     With MDIMainMenu
         .tbMenu.Buttons(5).Caption = "Search"
         .tbMenu.Buttons(5).Image = 3
     End With
     tbl.TABLE_SEARCH = 0
     tbl.TABLE_SEARCH_FLAG = ""
     tbl.TABLE_SEARCH_KREDITOR = ""
     tbl.TABLE_SEARCH_TANGGAL = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF1: CommandPass "New"
        Case vbKeyF2: CommandPass "Edit"
        Case vbKeyF3: CommandPass "Search"
        Case vbKeyF4: CommandPass "Delete"
        Case vbKeyF5: CommandPass "Refresh"
        Case vbKeyF6: CommandPass "Print"
        Case vbKeyF8: CommandPass "Close"
    End Select
End Sub

Private Sub Form_Load()
    Dim sort As String
    'Set the graphics for the controls
    Call LoadShow(cbShow)
    With MDIMainMenu
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
        .AddToWin Me.Caption, Name
    End With
    
    Select Case cbSort.Text
        Case "No.Faktur": sort = " LEFT(j.no_jual,3) " & cbSortType.Text & ", ABS(MID(j.no_jual,4,7))" & cbSortType.Text
        Case "Nama Departemen": sort = "d.nm_departement " & cbSortType.Text
    End Select
    
    With tbl
        .TABLE_TANGGAL_AWAL = ""
        .TABLE_TANGGAL_AKHIR = ""
    End With
    
    With SQLParser
            .Fields = "j.id_jual,j.no_jual,DATE_FORMAT(j.tgl_jual,'%Y-%m-%d %H:%i:%s'),j.tgl_bayar,j.flag_kreditor,j.flag_debitor,j.tgl_komisi,j.kd_pasien,p.nm_pasien,p.no_tlp,p.relasi,j.id_kreditor,k.nm_kreditor,d.kd_departement,d.nm_departement,j.type,j.payment,j.piutang,(j.bayar-j.piutang) as sisa,j.komisi,((j.bayar-j.piutang)-j.komisi) as laba,pp.nm_pengguna "
            .Tables = "tbl_jual j INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor INNER JOIN tbl_departement d ON d.id_departement=j.id_departement INNER JOIN tbl_pengguna pp ON pp.id=j.id_pengguna"
            .SortOrder = sort & " LIMIT " & cbShow.Text
            .SaveStatement
    End With
    
    If rsSales.State = 1 Then rsSales.Close
    rsSales.CursorLocation = adUseClient
    rsSales.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsSales, 1000000
        FillList 1
    End With
    lbltotal.Caption = "Total Record : " & lvList.ListItems.Count
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsSales, RecordPage.PageStart, RecordPage.PageEnd, 25, 2, False, True, , , , "id_jual")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    'lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    lvList_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        shpBar.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth
        lvList.Height = (Me.ScaleHeight - (Picture1.Height)) - lvList.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainMenu.RemToWin Me.Caption
    MDIMainMenu.HideTBButton "", True
    CurrBiz.BUSINNES_SALE = 0
    Call Deactive
    Set frmSales = Nothing
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

Private Sub Browsesrcform()
    With frmSalesDetails
                Set .srcform = Me
                .Caption = lvList.SelectedItem.Text
                .show vbModal
    End With
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

