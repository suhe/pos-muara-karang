VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPurchase 
   Caption         =   "List Purchase"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   Icon            =   "frmPurchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   9930
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9930
      TabIndex        =   9
      Top             =   4305
      Width           =   9930
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9930
      TabIndex        =   8
      Top             =   4320
      Width           =   9930
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9930
      TabIndex        =   0
      Top             =   4335
      Width           =   9930
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Next 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Last 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Previous 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "First 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   6
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
         TabIndex        =   7
         Top             =   60
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3915
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   6906
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Beli"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No Faktur"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tgl Beli"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tgl Bayar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Flag"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ID Supplier"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nm Supplier"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Payment"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Bayar"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Hutang"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Sisa"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Nm Pengguna"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.ComboBox cbDay1 
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox cbMonth1 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox cbYear1 
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox cbDay2 
      Height          =   315
      Left            =   3720
      TabIndex        =   15
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox cbMonth2 
      Height          =   315
      Left            =   4440
      TabIndex        =   16
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox cbYear2 
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase"
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
      TabIndex        =   11
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
      Width           =   9795
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsPurchase As New Recordset
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
        Case "New"
        If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_beli", "id_beli", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    If (lvList.SelectedItem.SubItems(11) < 0) Then
                        With frmPurchaseAE
                            tbl.TABLE_NO_FAK = lvList.SelectedItem.Text
                            tbl.TABLE_PAYMENT = lvList.SelectedItem.SubItems(10)
                            .State = adStateEditMode
                            .PK = CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                            .show vbModal
                        End With
                    Else
                        MsgBox "Your No Debt For No.Faktur : " & lvList.SelectedItem.SubItems(1) & " !", vbCritical + vbInformation
                    End If
                End If
            End If
        Case "Search"
            If tbl.TABLE_SEARCH2 = 1 Then
                With frmPurchaseFaktur
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
            Call Deactive
            RefreshRecords
        Case "Print"
            If lvList.ListItems.Count > 0 Then
                 'Call printPurchaseSummary
                 frmPurchasePrint.show vbModal
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
    With rsPurchase
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

Private Sub Active()
     'With MDIMainMenu
     '    .tbMenu.Buttons(3).Caption = "Lunasi"
     '    .tbMenu.Buttons(3).Image = 9
     'End With
End Sub

Private Sub Deactive()
     With MDIMainMenu
         .tbMenu.Buttons(5).Caption = "Search"
         .tbMenu.Buttons(5).Image = 3
     End With
     tbl.TABLE_SEARCH2 = 0
     tbl.TABLE_SEARCH_FLAG_2 = ""
     tbl.TABLE_SEARCH_SUPPLIER = ""
     tbl.TABLE_SEARCH_TANGGAL_2 = ""
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fftfttt"
    'Call Active
End Sub

Private Sub Form_Deactivate()
    MDIMainMenu.HideTBButton "", True
    Call Deactive
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
    MDIMainMenu.AddToWin Me.Caption, Name
    'Set the graphics for the controls
    With MDIMainMenu
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    
        btnFirst.Picture = .i16x16.ListImages(3).Picture
        btnPrev.Picture = .i16x16.ListImages(4).Picture
        btnNext.Picture = .i16x16.ListImages(5).Picture
        btnLast.Picture = .i16x16.ListImages(6).Picture
        
        btnFirst.DisabledPicture = .i16x16g.ListImages(3).Picture
        btnPrev.DisabledPicture = .i16x16g.ListImages(4).Picture
        btnNext.DisabledPicture = .i16x16g.ListImages(5).Picture
        btnLast.DisabledPicture = .i16x16g.ListImages(6).Picture
    End With
    
    With tbl
        .TABLE_TANGGAL_AWAL = ""
        .TABLE_TANGGAL_AKHIR = ""
    End With
    
    With SQLParser
        .Fields = "b.id_beli,b.no_beli,DATE_FORMAT(b.tgl_input,'%Y-%m-%d'),b.tgl_bayar,b.flag_supplier,b.id_supplier,s.nm_supplier,b.type,b.payment,b.bayar,b.hutang,(b.bayar-b.hutang) as sisa,p.nm_pengguna"
        .Tables = "tbl_beli b JOIN tbl_supplier s ON s.id_supplier=b.id_supplier LEFT JOIN tbl_pengguna p ON p.id=b.id_pengguna "
        .SortOrder = "b.id_beli DESC "
        .SaveStatement
    End With
    
    If rsPurchase.State = 1 Then rsPurchase.Close
    rsPurchase.CursorLocation = adUseClient
    rsPurchase.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsPurchase, 10000000
        FillList 1
    End With
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsPurchase, RecordPage.PageStart, RecordPage.PageEnd, 14, 2, False, True, , , , "id_beli")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
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
    Call Deactive
    Set frmPurchase = Nothing
End Sub

Private Sub SetNavigation()
    With RecordPage
        If .PAGE_TOTAL = 1 Then
            btnFirst.Enabled = False
            btnPrev.Enabled = False
            btnNext.Enabled = False
            btnLast.Enabled = False
        ElseIf .PAGE_CURRENT = 1 Then
            btnFirst.Enabled = False
            btnPrev.Enabled = False
            btnNext.Enabled = True
            btnLast.Enabled = True
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            btnFirst.Enabled = True
            btnPrev.Enabled = True
            btnNext.Enabled = False
            btnLast.Enabled = False
        Else
            btnFirst.Enabled = True
            btnPrev.Enabled = True
            btnNext.Enabled = True
            btnLast.Enabled = True
        End If
    End With
End Sub

Private Sub lblGrand_Click()

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
    'With frmPurchaseDetails
    '     Set .srcform = frmPurchase
    '         .show vbModal
    'End With
End Sub

Private Sub lvList_DblClick()
    'With lvList
    '    tbl.TABLE_NO_FAK = .SelectedItem.SubItems(1)
    '    tbl.TABLE_TANGGAL = .SelectedItem.SubItems(2)
    '    tbl.TABLE_FLAG_SUPPLIER = .SelectedItem.SubItems(4)
    '    tbl.TABLE_ID_SUPPLIER = .SelectedItem.SubItems(5)
    '    tbl.TABLE_NM_SUPPLIER = .SelectedItem.SubItems(6)
    '    tbl.TABLE_TYPE = .SelectedItem.SubItems(7)
    '    tbl.TABLE_PAY_TYPE = .SelectedItem.SubItems(8)
    '    tbl.TABLE_TOTAL = Format(.SelectedItem.SubItems(10), "")
    'End With
    'Call Browsesrcform
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    Call Browsesrcform
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub
