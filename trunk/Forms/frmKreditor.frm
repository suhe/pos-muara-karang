VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKreditor 
   Caption         =   "Kreditor"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmKreditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   7695
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7695
      TabIndex        =   2
      Top             =   5460
      Width           =   7695
      Begin VB.ComboBox cbShow 
         Height          =   315
         Left            =   3780
         TabIndex        =   8
         Text            =   "30"
         Top             =   30
         Width           =   735
      End
      Begin VB.ComboBox cbSort 
         Height          =   315
         ItemData        =   "frmKreditor.frx":038A
         Left            =   4920
         List            =   "frmKreditor.frx":0394
         TabIndex        =   7
         Text            =   "ID Kreditor"
         Top             =   30
         Width           =   1695
      End
      Begin VB.ComboBox cbSortType 
         Height          =   315
         ItemData        =   "frmKreditor.frx":03B4
         Left            =   6600
         List            =   "frmKreditor.frx":03BE
         TabIndex        =   6
         Text            =   "ASC"
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Show"
         Height          =   195
         Left            =   3360
         TabIndex        =   11
         Top             =   60
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sort"
         Height          =   195
         Left            =   4560
         TabIndex        =   10
         Top             =   60
         Width           =   300
      End
      Begin VB.Label lbltotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Record : 0"
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   60
         Width           =   1395
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
      ScaleWidth      =   7695
      TabIndex        =   1
      Top             =   5445
      Width           =   7695
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7695
      TabIndex        =   0
      Top             =   5430
      Width           =   7695
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3435
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   6059
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
         Text            =   "ID_Kreditor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nm Kreditor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Almt Kreditor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kota Kreditor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tlp Kreditor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CP Kreditor"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Plafon"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tgl Input"
         Object.Width           =   3572
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Nm Pengguna"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Kreditor Records"
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
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   -120
      Top             =   0
      Width           =   7755
   End
End
Attribute VB_Name = "frmKreditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsKreditor As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser

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
            frmKreditorAE.State = adStateAddMode
            frmKreditorAE.show vbModal
        Case "Edit"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_kreditor", "id_kreditor", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    With frmKreditorAE
                        .State = adStateEditMode
                        .PK = CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                        .show vbModal
                    End With
                End If
            End If
        Case "Search"
            With frmSearch
                Set .srcform = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .show vbModal
            End With
        Case "Delete"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_kreditor", "id_kreditor", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    Dim ANS As Integer
                    ANS = MsgBox("Are you sure you want to delete the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation.", vbCritical + vbYesNo, "Confirm Record Delete")
                    Me.MousePointer = vbHourglass
                    If ANS = vbYes Then
                        If isRecordExist("tbl_jual", "id_kreditor", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                            DelRecwSQL "tbl_kreditor", "id_kreditor", "", True, CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                            RefreshRecords
                            MDIMainMenu.UpdateInfoMsg
                            MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
                        Else
                            MsgBox "Record not been deleted , this is record in the transaction table !.", vbInformation, "Confirm"
                        End If
                    End If
                    ANS = 0
                    Me.MousePointer = vbDefault
                End If
            Else
                MsgBox "No record to delete.", vbExclamation
            End If
        Case "Refresh"
            RefreshRecords
        Case "Close"
            Unload Me
    End Select
    Exit Sub
    
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
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
    With rsKreditor
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
    frmSupplierRecOp.show vbModal
End Sub

Private Sub cbShow_Change()
    cbShow.Text = "30"
End Sub

Private Sub cbShow_Click()
    Call Form_Load
End Sub

Private Sub cbSort_Change()
    cbSort.Text = "ID Kreditor"
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
    'HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "", True
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "tttttft"
End Sub

Private Sub Form_Deactivate()
    MDIMainMenu.HideTBButton "", True
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
    Call LoadShow(cbShow)
    'MDIMainMenu
    With MDIMainMenu
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
        .AddToWin Me.Caption, Name
    End With
    
    'sort berdsarkan
    Select Case cbSort.Text
        Case "ID Kreditor": sort = "ABS(k.id_kreditor) " & cbSortType.Text
        Case "Nama Kreditor": sort = "k.nm_kreditor " & cbSortType.Text
    End Select
    
    With SQLParser
        .Fields = "k.id_kreditor,k.nm_kreditor,k.almt_kreditor,k.kota_kreditor,k.tlp_kreditor,cp_kreditor,k.plafon,k.tgl_input,p.nm_pengguna"
        .Tables = "tbl_kreditor k INNER JOIN tbl_pengguna p ON p.id=k.id_pengguna"
        .SortOrder = sort & " LIMIT " & cbShow.Text
        .SaveStatement
    End With
    
    If rsKreditor.State = 1 Then rsKreditor.Close
    rsKreditor.CursorLocation = adUseClient
    rsKreditor.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsKreditor, 10000000
        FillList 1
    End With
    lbltotal.Caption = "Total Record : " & lvList.ListItems.Count
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsKreditor, RecordPage.PageStart, RecordPage.PageEnd, 14, 2, False, True, , , , "id_kreditor")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    'Display the page information
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
    Set rsKreditor = Nothing
    Set frmKreditor = Nothing
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

