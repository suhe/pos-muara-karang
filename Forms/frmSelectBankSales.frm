VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelectBankSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Sales"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   14250
      TabIndex        =   14
      Top             =   7845
      Width           =   14250
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   14250
      TabIndex        =   13
      Top             =   7860
      Width           =   14250
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14250
      TabIndex        =   5
      Top             =   7875
      Width           =   14250
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   60
            Width           =   2535
         End
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14220
      TabIndex        =   0
      Top             =   7440
      Width           =   14250
      Begin VB.CommandButton cmdProses 
         Caption         =   "&Process"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   277282817
         CurrentDate     =   40493
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   277282817
         CurrentDate     =   40493
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   2040
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblGrand 
         BackStyle       =   0  'Transparent
         Caption         =   "GrandTotal"
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
         Left            =   4920
         TabIndex        =   2
         Top             =   120
         Width           =   7095
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   7155
      Left            =   0
      TabIndex        =   15
      Top             =   240
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   12621
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fak No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date Added"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer Name"
         Object.Width           =   3995
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Exp ID"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Expedition Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Exp Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "SubTotal"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "GrandTotal"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.ComboBox cbDay1 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cbMonth1 
      Height          =   315
      Left            =   840
      TabIndex        =   18
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cbYear1 
      Height          =   315
      Left            =   1560
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cbDay2 
      Height          =   315
      Left            =   3000
      TabIndex        =   20
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cbMonth2 
      Height          =   315
      Left            =   3720
      TabIndex        =   21
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cbYear2 
      Height          =   315
      Left            =   4440
      TabIndex        =   22
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Sales"
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
      Left            =   240
      TabIndex        =   16
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
      Width           =   14235
   End
End
Attribute VB_Name = "frmSelectBankSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsSalesBank As New Recordset
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
            With frmSearch
                Set .srcform = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .show vbModal
            End With
        Case "Refresh"
            RefreshRecords
        Case "Print"
            MsgBox "Dont Have A Report For this Form", vbOKOnly + vbCritical, "Warning"
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
    With rsSalesBank
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

Private Sub cmdProses_Click()
    On Error Resume Next
    
    'If (cbDay1.Text = "") Or (cbDay2.Text = "") Or (cbMonth1.Text = "") Or (cbMonth2.Text = "") Or (cbYear1.Text = "") Or (cbYear2.Text = "") Then
    '    MsgBox "Complete The Selection", vbOKCancel + vbCritical, "Warning Error !"
    'Else
        With SQLParser
            .Fields = "*"
            .Tables = "QR_SALES_BANKS"
            .wCondition = "(((QR_SALES_BANKS.DateAdded)>='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "' And (QR_SALES_BANKS.DateAdded)<='" & Format(DTPicker2.Value, "mm/dd/yyyy") & "') AND ((QR_SALES_BANKS.BankCode)=" & CurrBiz.BUSINNES_BANK & "))"
            '.wCondition = " DateAdded>=#" & cbMonth1.Text & "/" & cbDay1.Text & "/" & cbYear1.Text & "# " & _
                          "AND DateAdded<=#" & cbMonth2.Text & "/" & cbDay2.Text & "/" & cbYear2.Text & "# "
            .SortOrder = "DateAdded DESC"
            .SaveStatement
        End With
    
        rsSalesBank.CursorLocation = adUseClient
        rsSalesBank.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
        
        With RecordPage
            .Start rsSalesBank, 75
            FillList 1
        End With
        rsSalesBank.Close
        
    subAmount = 0
    subtotal = 0
    grandtotal = 0
    
    For i = 1 To lvList.ListItems.Count
        subAmount = subAmount + lvList.ListItems(i).SubItems(6)
        subtotal = subtotal + lvList.ListItems(i).SubItems(7)
        grandtotal = grandtotal + lvList.ListItems(i).SubItems(8)
    Next i
   lblGrand.Caption = "Exp.Amount : " & Format(subAmount, "##,###0.00") & " SubTotal : " & Format(subtotal, "##,###0.00") & " GrandTotal : " & Format(grandtotal, "##,###0.00")
    'End If
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fftftft"
End Sub

Private Sub Form_Deactivate()
    MDIMainMenu.HideTBButton "", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    CurrBiz.BUSINNES_BANK = frmSelectBank.ListView1.SelectedItem.SubItems(1)
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
        .Fields = "*"
        .Tables = "QR_SALES_BANKS"
        .wCondition = "BANKCODE=" & CurrBiz.BUSINNES_BANK
        .SaveStatement
    End With

    rsSalesBank.CursorLocation = adUseClient
    rsSalesBank.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsSalesBank, 75
        FillList 1
    End With
    rsSalesBank.Close
    'Set rsSalesBank = Nothing
    With cbDay1
        For i = 1 To 31
            cbDay1.AddItem i
        Next i
    End With
    
    With cbDay2
        For i = 1 To 31
            cbDay2.AddItem i
        Next i
    End With
    
    With cbMonth1
        For i = 1 To 12
            cbMonth1.AddItem i
        Next i
    End With
    
    With cbMonth2
        For i = 1 To 12
            cbMonth2.AddItem i
        Next i
    End With
    
    With cbYear1
        For i = 2010 To 2022
            cbYear1.AddItem i
        Next i
    End With
    
    With cbYear2
        For i = 2010 To 2022
            cbYear2.AddItem i
        Next i
    End With
    
    Dim subtotal As Double
    Dim subAmount As Double
    Dim grandtotal As Double
    
    subAmount = 0
    subtotal = 0
    grandtotal = 0
    
    For i = 1 To lvList.ListItems.Count
        subAmount = subAmount + lvList.ListItems(i).SubItems(6)
        subtotal = subtotal + lvList.ListItems(i).SubItems(7)
        grandtotal = grandtotal + lvList.ListItems(i).SubItems(8)
    Next i
   lblGrand.Caption = "Exp.Amount : " & Format(subAmount, "##,###0.00") & " SubTotal : " & Format(subtotal, "##,###0.00") & " GrandTotal : " & Format(grandtotal, "##,###0.00")
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsSalesBank, RecordPage.PageStart, RecordPage.PageEnd, 9, 2, False, True, , , , "PK")
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
        lvList.Height = (Me.ScaleHeight - (Picture1.Height + Picture3.Height)) - lvList.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainMenu.RemToWin Me.Caption
    MDIMainMenu.HideTBButton "", True
    Set frmSales = Nothing
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

Private Sub Browsesrcform()
    With frmSalesProductDetails
                'Set .tabledata = ""
                Set .srcform = Me
                .Caption = lvList.SelectedItem.Text
                .show vbModal
    End With
End Sub

Private Sub lvList_DblClick()
    On Error Resume Next
    Call Browsesrcform
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Call Browsesrcform
    End If
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


