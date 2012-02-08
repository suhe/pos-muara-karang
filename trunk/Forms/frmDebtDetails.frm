VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDebtDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hutang Kredior"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmDebtDetails.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   1815
      Index           =   8
      Left            =   14895
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   24
      Tag             =   "Remarks"
      Top             =   -120
      Width           =   3105
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   16635
      TabIndex        =   23
      Top             =   3105
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   15195
      TabIndex        =   22
      Top             =   3105
      Width           =   1335
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   9720
      ScaleHeight     =   30
      ScaleWidth      =   12015
      TabIndex        =   21
      Top             =   2955
      Width           =   12015
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9450
      TabIndex        =   12
      Top             =   6585
      Width           =   9450
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   13
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "First 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Previous 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Last 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   14
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
            TabIndex        =   18
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
         TabIndex        =   19
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
      ScaleWidth      =   9450
      TabIndex        =   11
      Top             =   6570
      Width           =   9450
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9450
      TabIndex        =   10
      Top             =   6555
      Width           =   9450
   End
   Begin VB.Frame fraFaktur 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   7
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   37
         Top             =   960
         Width           =   2490
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   35
         Top             =   960
         Width           =   2970
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   3
         Left            =   5520
         MaxLength       =   100
         TabIndex        =   33
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   2
         Left            =   5520
         MaxLength       =   200
         TabIndex        =   31
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   29
         Tag             =   "Name"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact No"
         Height          =   240
         Index           =   7
         Left            =   4560
         TabIndex        =   36
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Person"
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "City/Town"
         Height          =   240
         Index           =   3
         Left            =   4440
         TabIndex        =   32
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   240
         Index           =   2
         Left            =   4800
         TabIndex        =   30
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "ID"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdInvoice 
      Caption         =   "&Cetak Tagihan (F2)"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      ToolTipText     =   "Tagihan dicetak Apabila Pembayaran Hutang"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar (F8)"
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
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
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   9135
      Begin VB.TextBox txtDibayar 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmDebtDetails.frx":038A
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtBayar 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmDebtDetails.frx":038E
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtPiutang 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmDebtDetails.frx":039C
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Dibayar"
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
         Left            =   3120
         TabIndex        =   6
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Total Bayar"
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
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Piutang"
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
         Left            =   6240
         TabIndex        =   4
         Top             =   120
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4275
      Left            =   0
      TabIndex        =   39
      Top             =   2160
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   7541
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
      NumItems        =   16
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
         Text            =   "P.Kd Pasien"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nm Pasien"
         Object.Width           =   3995
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ID Kreditor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nm Kreditor"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "KD Departement"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Nm Departement"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Nm Cabang"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Payment"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Bayar"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Piutang"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Sisa"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Nm Pengguna"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Faktur HUtang "
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
      TabIndex        =   38
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Memo"
      Height          =   240
      Index           =   8
      Left            =   13845
      TabIndex        =   25
      Top             =   -120
      Width           =   990
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "..........."
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
      TabIndex        =   20
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   1920
      Width           =   9435
   End
End
Attribute VB_Name = "frmDebtDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CURR_COL As Integer
Dim rsSales As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode
Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

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
            If lvList.ListItems.Count > 0 Then
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

Private Sub cmdInvoice_Click()
    Unload Me
    On Error Resume Next
    If (lvList.ListItems.Count > 0) Then
        Call printInvoice
    Else
        MsgBox "No Data View In the List"
    End If
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF2: Call cmdInvoice_Click
        Case vbKeyF8: Call cmdKeluar_Click
    End Select
End Sub

Private Sub Form_Load()
    DisplayForEditing
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
            .Fields = "j.id_jual,j.no_jual,j.tgl_jual,j.kd_pasien,p.nm_pasien,j.id_kreditor,k.nm_kreditor,d.kd_departement,d.nm_departement,c.nm_cabang,j.type,j.payment,j.bayar,j.piutang,(j.bayar-j.piutang) as sisa,pp.nm_pengguna "
            .Tables = "tbl_jual j JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor JOIN tbl_departement d ON d.id_departement=j.id_departement LEFT JOIN tbl_cabang c ON c.id_cabang=j.id_cabang LEFT JOIN tbl_pengguna pp ON pp.id=j.id_pengguna"
            .SortOrder = "j.id_jual DESC"
            .wCondition = "j.id_kreditor = " & tbl.TABLE_ID_KREDITUR & " AND j.piutang > 0 "
            .SaveStatement
    End With
    
    If rsSales.State = 1 Then rsSales.Close
    rsSales.CursorLocation = adUseClient
    rsSales.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsSales, 10000000
        FillList 1
    End With
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsSales, RecordPage.PageStart, RecordPage.PageEnd, 16, 2, False, True, , , , "id_jual")
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
    Set frmDebtDetails = Nothing
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
    With frmSalesDetails
                Set .srcform = Me
                .Caption = lvList.SelectedItem.Text
                .show vbModal
    End With
End Sub

Private Sub lvList_DblClick()
    On Error Resume Next
    With lvList
        tbl.TABLE_NO_FAK = .SelectedItem.SubItems(1)
        tbl.TABLE_TANGGAL = .SelectedItem.SubItems(2)
        tbl.TABLE_KD_PASIEN = .SelectedItem.SubItems(3)
        tbl.TABLE_NM_PASIEN = .SelectedItem.SubItems(4)
        tbl.TABLE_ID_KREDITUR = .SelectedItem.SubItems(5)
        tbl.TABLE_NM_KREDITUR = .SelectedItem.SubItems(6)
        tbl.TABLE_KD_DEPT = .SelectedItem.SubItems(7)
        tbl.TABLE_NM_DEPT = .SelectedItem.SubItems(8)
        tbl.TABLE_TYPE = .SelectedItem.SubItems(9)
        tbl.TABLE_PAY_TYPE = .SelectedItem.SubItems(10)
        tbl.TABLE_TOTAL = .SelectedItem.SubItems(11) + .SelectedItem.SubItems(12)
        tbl.TABLE_MONEY = .SelectedItem.SubItems(11)
        tbl.TABLE_CBACK = .SelectedItem.SubItems(12)
    End With
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

Private Sub DisplayForEditing()
    On Error GoTo err
    With tbl
        txtEntry(0).Text = .TABLE_ID_KREDITUR
        txtEntry(1).Text = .TABLE_NM_KREDITUR
        txtEntry(2).Text = .TABLE_ALMT_KREDITUR
        txtEntry(3).Text = .TABLE_KOTA_KREDITUR
        txtEntry(6).Text = .TABLE_CP_KREDITOR
        txtEntry(7).Text = .TABLE_TLP_KREDITOR
    End With
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

