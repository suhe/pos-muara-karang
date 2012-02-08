VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReturPurchase 
   Caption         =   "List Retur Purchase"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmReturPurchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   10095
   Begin MSComctlLib.ListView lvList 
      Height          =   6795
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   11986
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
         Text            =   "ID Beli"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "B.No Beli"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tgl Beli"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tgl Retur"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID Supplier"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Nama Supplier"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Kd Obat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Nm Obat"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Jumlah"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Retur"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Sisa"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Kerugian"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.ComboBox cbDay1 
      Height          =   315
      Left            =   480
      TabIndex        =   17
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox cbMonth1 
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox cbYear1 
      Height          =   315
      Left            =   1920
      TabIndex        =   15
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cbDay2 
      Height          =   315
      Left            =   3360
      TabIndex        =   14
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox cbMonth2 
      Height          =   315
      Left            =   4080
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox cbYear2 
      Height          =   315
      Left            =   4800
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10095
      TabIndex        =   2
      Top             =   7215
      Width           =   10095
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   3
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "First 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Previous 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Last 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Next 250"
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
            TabIndex        =   8
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
      ScaleWidth      =   10095
      TabIndex        =   1
      Top             =   7200
      Width           =   10095
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   10095
      TabIndex        =   0
      Top             =   7185
      Width           =   10095
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Retur Purchase"
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
      Width           =   10035
   End
End
Attribute VB_Name = "frmReturPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsReturPurchase As New Recordset
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
            On Error Resume Next
            frmPurchaseReturAE.show vbModal
        Case "Search"
            With frmSearch
                Set .srcform = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .show vbModal
            End With
        Case "Refresh"
            RefreshRecords
        Case "Print"
            If lvList.ListItems.Count Then
                Call printRetur
            Else
                MsgBox "Dont Have A Report For this Form", vbOKOnly + vbCritical, "Warning"
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
    With rsReturPurchase
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
         .tbMenu.Buttons(3).Caption = "Retur"
         .tbMenu.Buttons(3).Image = 9
     End With
End Sub

Private Sub Deactive()
     With MDIMainMenu
         .tbMenu.Buttons(3).Caption = "New"
         .tbMenu.Buttons(3).Image = 1
     End With
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "tftfttt"
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
    Dim subtotal As Double
    MDIMainMenu.AddToWin Me.Caption, Name
    With MDIMainMenu
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
        .Fields = " b.id_beli,b.no_beli,b.tgl_beli,d.tgl_retur,b.id_supplier,s.nm_supplier,o.kd_obat,o.nm_obat,d.jumlah,d.retur,(d.jumlah-d.retur) AS sisa,(d.retur*d.harga_beli) AS rugi"
        .Tables = " tbl_beli_details d JOIN tbl_beli b ON b.no_beli=d.no_beli JOIN tbl_obat o ON o.id_obat=d.id_obat JOIN tbl_supplier s ON s.id_supplier=b.id_supplier "
        .wCondition = " d.retur > 0"
        .SaveStatement
    End With
    
    If rsReturPurchase.State = 1 Then rsReturPurchase.Close
    rsReturPurchase.CursorLocation = adUseClient
    rsReturPurchase.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsReturPurchase, 10000000
        FillList 1
    End With

End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsReturPurchase, RecordPage.PageStart, RecordPage.PageEnd, 12, 2, False, True, , , , "id_beli")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    lvList_Click
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
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
    Set frmReturPurchase = Nothing
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
    'With frmSalesProductDetails
      '          Set .srcform = frmReturPurchase
     '           .cmdNew.Visible = False
       '         .show vbModal
    'End With
End Sub

Private Sub lvList_DblClick()
    'Call Browsesrcform
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    'Call Browsesrcform
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub
