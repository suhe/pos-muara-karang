VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPurchasingSupplier 
   Caption         =   "Supplier Purchasing"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   Icon            =   "frmPurchasingSupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   19080
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   19080
      TabIndex        =   13
      Top             =   10710
      Width           =   19080
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   19080
      TabIndex        =   12
      Top             =   10725
      Width           =   19080
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   19080
      TabIndex        =   4
      Top             =   10740
      Width           =   19080
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   5
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Next 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Last 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Previous 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   6
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
            TabIndex        =   10
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
         TabIndex        =   11
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.Frame freSearch 
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
      TabIndex        =   3
      Top             =   4320
      Width           =   7215
      Begin VB.TextBox txtSrchStr 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmPurchasingSupplier.frx":038A
         Left            =   240
         List            =   "frmPurchasingSupplier.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image imgSearch 
         Height          =   480
         Left            =   6600
         Picture         =   "frmPurchasingSupplier.frx":038E
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4035
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   7260
      _ExtentX        =   12806
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Supplier"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nm Supplier"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Almt Supplier"
         Object.Width           =   4710
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kota"
         Object.Width           =   3863
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tlp Supplier"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact Person"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Contact Person"
         Object.Width           =   3836
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Records"
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
      Left            =   360
      TabIndex        =   14
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   0
      Width           =   7275
   End
End
Attribute VB_Name = "frmPurchasingSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsPurchaseSupplier As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser
Dim str As String

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
    With rsPurchaseSupplier
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    MDIMainMenu.AddToWin Me.Caption, Name
    With cboFilter
        .AddItem "Code"
        .AddItem "Name"
        .Text = "Name"
    End With
    
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
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsPurchaseSupplier, RecordPage.PageStart, RecordPage.PageEnd, 8, 2, False, True, , , , "id_supplier")
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
        freSearch.Width = Me.ScaleWidth
        txtSrchStr.Width = freSearch.Width - (txtSrchStr.Left + imgSearch.Width)
        lvList.Height = (Me.ScaleHeight - Picture1.Height) - lvList.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPurchasingSupplier = Nothing
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
    With tbl
        .TABLE_ID_SUPPLIER = lvList.SelectedItem.Text
        .TABLE_NM_SUPPLIER = lvList.SelectedItem.SubItems(1)
        .TABLE_TLP_SUPPLIER = lvList.SelectedItem.SubItems(4)
    End With
    frmPurchasing.lblCodeCust.Caption = lvList.SelectedItem.Text
    frmPurchasing.lblNamaCust.Caption = lvList.SelectedItem.SubItems(1)
    Unload Me
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmPurchasing.lblCodeCust.Caption = lvList.SelectedItem.Text
        frmPurchasing.lblNamaCust.Caption = lvList.SelectedItem.SubItems(1)
        Unload Me
     End If
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub txtSrchStr_Change()
    On Error Resume Next
    If cboFilter.Text = "Code" Then
        str = "id_supplier"
    Else
        str = "nm_supplier"
    End If
    
    If txtSrchStr.Text <> "" Then
    
        With SQLParser
            .Fields = "id_supplier,nm_supplier,almt_supplier,kota_supplier,tlp_supplier,cp_supplier"
            .Tables = " tbl_supplier"
            .wCondition = str & " LIKE '%" & txtSrchStr.Text & "%'"
            .SortOrder = " id_supplier ASC,nm_supplier ASC LIMIT 20 "
            .SaveStatement
        End With
        
        'SQLParser =
        If rsPurchaseSupplier.State = 1 Then rsPurchaseSupplier.Close
        rsPurchaseSupplier.CursorLocation = adUseClient
        rsPurchaseSupplier.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
        
        With RecordPage
            .Start rsPurchaseSupplier, 20
            FillList 1
        End With
        'rsPurchaseSupplier.Close
    End If
End Sub

Private Sub txtSrchStr_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        lvList.SetFocus
    End If
End Sub
