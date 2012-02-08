VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStock 
   Caption         =   "Stock Of Product"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   8475
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   8475
      TabIndex        =   2
      Top             =   3960
      Width           =   8475
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5040
         ScaleHeight     =   345
         ScaleWidth      =   4275
         TabIndex        =   3
         Top             =   0
         Width           =   4275
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   1320
            TabIndex        =   8
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   9
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
      ScaleWidth      =   8475
      TabIndex        =   1
      Top             =   3945
      Width           =   8475
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   3930
      Width           =   8475
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3675
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   6482
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bar"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Category Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Disc(%)"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Price Sell"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Price Buy"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Profit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Gudang"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Buy"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Sale"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Check + "
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Check - "
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Stock"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "GrandTotal"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   7515
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
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsStocks As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser

Private Sub Active()
    'With MDIMainMenu
    '    .tbMenu.Buttons(3).Caption = "Adjust"
    '    .tbMenu.Buttons(3).Image = 9
        '.mnuRACN.Caption = "Adjust Stock"
    'End With
End Sub

Private Sub Deactive()
    'With MDIMainMenu
    '    .tbMenu.Buttons(3).Caption = "New"
    '    .tbMenu.Buttons(3).Image = 1
        '.mnuRACN.Caption = "Create New"
    'End With
End Sub

'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    SQLParser.RestoreStatement
    SQLParser.wCondition = srcCondition
    ReloadRecords SQLParser.SQLStatement
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "New"
            'If lvList.ListItems.Count > 0 Then
            '    If isRecordExist("Categories", "PK", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
             '       MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
             '       RefreshRecords
             '       Exit Sub
             '   Else
                    With frmStockChecking
                        .Caption = lvList.SelectedItem.Text
                        '.PK = CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                        .show vbModal
                    End With
                'End If
            'End If
        Case "Search"
            With frmSearch
                .srcNoOfCol = 8
                Set .srcform = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                frmSearch.cmbFields.AddItem "MDIMainMenu Stock"
                .show vbModal
            End With
        Case "Refresh"
            RefreshRecords
        Case "Print"
             'Dim i As Integer
            'For i = 1 To 4
                Call printStock
            'Next i
            'ACRStock.show

          '  GenerateDSN
          '  With MAIN.CR
               ' .Reset: MAIN.InitCrys
                
           '     .ReportFileName = App.Path & "\Reports\rptStocksInformation.rpt"
           '     .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"
            
           '     .WindowTitle = "Stocks Information List"
        
           '     .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
           '     .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
           '     .ParameterFields(2) = "prmTitle;STOCKS INFORMATION LIST;True"
                    
           '     .PageZoom 100
           '     .Action = 1
          '  End With
          '  RemoveDSN
        Case "Close"
            Unload Me
    End Select
    Exit Sub
    'Trap the error
err:
        prompt_err err, Name, "CommandPass"
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
    With rsStocks
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

Private Sub cmdStatusFile_Click()
    On Error GoTo Err_Handler
    CommonDialog1.CancelError = True
    txbStatusFile.Text = ""
    CommonDialog1.FileName = Replace(lblTitle.Caption, " ", "") & "_" & cboList & "_" & Format(Now(), "ddmm") & ".xls"
    CommonDialog1.Filter = "Microsoft Excel .xls (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then txbStatusFile.Text = CommonDialog1.FileName
    Exit Sub
Err_Handler:
    If err = 32755 Then
        txbStatusFile.Text = ""
    Else
        MsgBox "An error has occurred! " & vbCrLf & vbCrLf & err & ": " & Error & " ", vbExclamation
    End If
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
    
    With SQLParser
       ' .Fields = "PK,ProductCode,ProductName,CategoryName,Discount,Price_buy,Price_Sell,StockP,Purchase_qty,Sales_qty,stock"
        .Fields = "*"
        .Tables = "QR_STOCKS_TOTAL"
        .SortOrder = "PK ASC,PK ASC"
        .SaveStatement
    End With
    
    rsStocks.CursorLocation = adUseClient
    rsStocks.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsStocks, 10000
        FillList 1
    End With
    
    rsStocks.Close
    
    Dim subtotal As Double
    Dim stok As Double
    
    subtotal = 0
    stok = 0
    'For i = 1 To lvList.ListItems.Count
        'stok = stok + lvList.ListItems(i).SubItems(13)
        'subtotal = subtotal + lvList.ListItems(i).SubItems(14)
    'Next i
   stok = getSumOfFields("QR_STOCKS_TOTAL", "Stock", CN)
   subtotal = getSumOfFields("QR_STOCKS_TOTAL", "GrandTotal", CN)
   Label1.Caption = "Stock : " & stok & " GrandTotal : " & Format(subtotal, "##,###0.00")
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsStocks, RecordPage.PageStart, RecordPage.PageEnd, 16, 2, False, True, , , , "PK")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation
    'Display the page information
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    'Display the selected record
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
        lvList.Refresh
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainMenu.RemToWin Me.Caption
    MDIMainMenu.HideTBButton "", True
    Set frmStock = Nothing
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
    CommandPass "New"
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then PopupMenu MAIN.mnuRecA
End Sub

