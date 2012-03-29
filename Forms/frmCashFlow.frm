VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCashFlow 
   Caption         =   "Cash Flow"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   Icon            =   "frmCashFlow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   9240
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9240
      TabIndex        =   3
      Top             =   4680
      Width           =   9240
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9240
      TabIndex        =   2
      Top             =   4695
      Width           =   9240
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9240
      TabIndex        =   0
      Top             =   4710
      Width           =   9240
      Begin VB.ComboBox cbSortType 
         Height          =   315
         ItemData        =   "frmCashFlow.frx":038A
         Left            =   6120
         List            =   "frmCashFlow.frx":0394
         TabIndex        =   8
         Text            =   "ASC"
         Top             =   30
         Width           =   855
      End
      Begin VB.ComboBox cbShow 
         Height          =   315
         ItemData        =   "frmCashFlow.frx":03A3
         Left            =   3600
         List            =   "frmCashFlow.frx":03A5
         TabIndex        =   7
         Text            =   "30"
         Top             =   30
         Width           =   735
      End
      Begin VB.ComboBox cbSort 
         Height          =   315
         ItemData        =   "frmCashFlow.frx":03A7
         Left            =   4800
         List            =   "frmCashFlow.frx":03B1
         TabIndex        =   6
         Text            =   "Tanggal"
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Show"
         Height          =   195
         Left            =   3120
         TabIndex        =   11
         Top             =   50
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sort"
         Height          =   195
         Left            =   4440
         TabIndex        =   10
         Top             =   50
         Width           =   300
      End
      Begin VB.Label lbltotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Record : 0"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4035
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   7117
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tgl Cash"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Kas Kemarin"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Lunas Piutang"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Laba Berjalan"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Lunas Hutang"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Beli Lunas"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Komis Dep."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Laba "
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Transf."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Kas Total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Pasien"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Flow"
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
      Left            =   75
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmCashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsCashFlow As New Recordset
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
        Case "Edit"
          If (lvList.SelectedItem.SubItems(10) > 0) Then
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_cash", "id", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    With tbl
                        .TABLE_TGL_CASH = lvList.SelectedItem.SubItems(1)
                        .TABLE_LABA_BERSIH = Format(lvList.SelectedItem.SubItems(10), "")
                    End With
                    'MsgBox Format(lvList.SelectedItem.SubItems(1), "DD/MM/YYYY") & " AND " & Format(Date, "DD/MM/YYYY")
                    If (Format(lvList.SelectedItem.SubItems(1), "DD/MM/YYYY") = Format(Date, "DD/MM/YYYY")) Then
                        With frmCashFlowAE
                            .State = adStateEditMode
                            .PK = CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                            tbl.TABLE_TOTAL = Format(lvList.SelectedItem.SubItems(10), "")
                            .show vbModal
                        End With
                    Else
                        MsgBox "The Day Without Transfer Is Must Be Same Day , Thank You", vbCritical + vbInformation
                    End If
                End If
            End If
           Else
            MsgBox "Tidak Laba Bersih Untuk Tanggal Ini ", vbCritical + vbInformation
           End If
        Case "Search"
            tbl.TABLE_TANGGAL = "Date"
            With frmSearch
                Set .srcform = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .show vbModal
            End With
        Case "Refresh"
            RefreshRecords
        Case "Print"
           On Error Resume Next
            If (lvList.ListItems.Count > 0) Then
                'MDIMainMenu.HideTBButton "", True
                Call printCashFlowdetails
                'Unload Me
            Else
                MsgBox "No Data View In the List"
            End If
        Case "Close"
            Unload Me
    End Select
End Sub

Public Sub RefreshRecords()
    SQLParser.RestoreStatement
    ReloadRecords SQLParser.SQLStatement
End Sub

'Procedure for reloadingrecords
Public Sub ReloadRecords(ByVal srcSQL As String)
    '-In this case I used SQL because it is faster than Filter function of VB
    '-when hundling millions of records.
    On Error GoTo err
    With rsCashFlow
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

Private Sub Active()
    With MDIMainMenu
        .tbMenu.Buttons(3).Caption = "New Cash"
        .tbMenu.Buttons(3).Image = 1
    End With
End Sub

Private Sub Deactive()
    With MDIMainMenu
         .tbMenu.Buttons(3).Caption = "New"
         .tbMenu.Buttons(3).Image = 1
         '.mnuRACN.Caption = "Create New"
    End With
End Sub

Private Sub cbShow_Change()
    cbShow.Text = "30"
End Sub

Private Sub cbShow_Click()
    Call Form_Load
End Sub

Private Sub cbSort_Change()
    cbSort.Text = "Tanggal"
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
    Active
End Sub

Private Sub Form_Deactivate()
    MDIMainMenu.HideTBButton "", True
    Deactive
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF1: CommandPass "New"
        Case vbKeyF2: CommandPass "Edit"
        Case vbKeyF3: CommandPass "Search"
        Case vbKeyF4: CommandPass "Delete"
        Case vbKeyF5: CommandPass "Refresh"
        Case vbKeyF8: CommandPass "Close"
    End Select
End Sub

Private Sub Form_Load()
    Call LoadShow(cbShow)
    'Set the graphics for the controls
    With MDIMainMenu
        'For listview
        .AddToWin Me.Caption, Name
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With
    
    'sort
    Select Case cbSort.Text
        Case "Tanggal": sort = "DATE_FORMAT(tgl_cash,'%Y-%m-%d') " & cbSortType.Text
        Case "Nama Pasien": sort = "p.nm_pasien " & cbSortType.Text
    End Select
    
    With SQLParser
        .Fields = "id,DATE_FORMAT(tgl_cash,'%Y-%m-%d'),FORMAT(money_cash,0),FORMAT(jual_sebelumnya,0),FORMAT(jual,0),FORMAT(beli_sebelumnya,0),FORMAT(beli,0),FORMAT(komisi,0),FORMAT(laba,0),FORMAT(cash,0),FORMAT((kas_total+(retur)),0) as kas_total,pasien"
        .Tables = "vw_cash_flow"
        .SortOrder = sort & " LIMIT " & cbShow.Text
        .SaveStatement
    End With
    
    If rsCashFlow.State = 1 Then rsCashFlow.Close
    Set rsCashFlow = New ADODB.Recordset
    rsCashFlow.CursorLocation = adUseClient
    rsCashFlow.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsCashFlow, 1000000
        FillList 1
    End With
    lbltotal.Caption = "Total Record : " & lvList.ListItems.Count
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsCashFlow, RecordPage.PageStart, RecordPage.PageEnd, 20, 2, False, True, , , , "id")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    lvList_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        shpBar.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth
        lvList.Height = (Me.ScaleHeight - Picture1.Height) - lvList.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainMenu.RemToWin Me.Caption
    MDIMainMenu.HideTBButton "", True
    Set frmCashFlow = Nothing
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

Private Sub lvList_DblClick()
    CommandPass "Edit"
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub
