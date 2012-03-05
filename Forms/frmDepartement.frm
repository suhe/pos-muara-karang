VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDepartement 
   Caption         =   "Departement"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmDepartement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   7695
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7695
      TabIndex        =   2
      Top             =   5280
      Width           =   7695
      Begin VB.ComboBox cbShow 
         Height          =   315
         ItemData        =   "frmDepartement.frx":038A
         Left            =   3780
         List            =   "frmDepartement.frx":038C
         TabIndex        =   8
         Text            =   "30"
         Top             =   30
         Width           =   735
      End
      Begin VB.ComboBox cbSort 
         Height          =   315
         ItemData        =   "frmDepartement.frx":038E
         Left            =   4920
         List            =   "frmDepartement.frx":0398
         TabIndex        =   7
         Text            =   "Kd Departemen"
         Top             =   30
         Width           =   1695
      End
      Begin VB.ComboBox cbSortType 
         Height          =   315
         ItemData        =   "frmDepartement.frx":03BC
         Left            =   6600
         List            =   "frmDepartement.frx":03C6
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
         Left            =   1680
         TabIndex        =   9
         Top             =   60
         Width           =   1515
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
      Top             =   5265
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
      Top             =   5250
      Width           =   7695
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4875
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8599
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID_Departement"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kd Departement"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nm Departement"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Parent Departement"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "BN"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "AN"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "VN"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "RN(%)"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "PN"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Tgl Input"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Nm Pengguna"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Departement Records"
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
Attribute VB_Name = "frmDepartement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDepartement As New Recordset
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
            frmDepartementAE.State = adStateAddMode
            frmDepartementAE.show vbModal
        Case "Edit"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_departement", "id_departement", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    With frmDepartementAE
                        .State = adStateEditMode
                        .PK = CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                        tbl.TABLE_KD_DEPT = Left(lvList.SelectedItem.SubItems(1), 1)
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
                If isRecordExist("tbl_departement", "id_departement", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    Dim ANS As Integer
                    ANS = MsgBox("Are you sure you want to delete the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation.", vbCritical + vbYesNo, "Confirm Record Delete")
                    Me.MousePointer = vbHourglass
                    If ANS = vbYes Then
                        If isRecordExist("tbl_jual", "id_departement", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                            DelRecwSQL "tbl_departement", "id_departement", "", True, CLng(LeftSplitUF(lvList.SelectedItem.Tag))
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
    On Error GoTo err
    With rsDepartement
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
    cbSort.Text = "Kd Departemen"
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
    'Set the graphics for the controls
    With MDIMainMenu
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
        .AddToWin Me.Caption, Name
    End With
    
    'sort berdsarkan
    Select Case cbSort.Text
        Case "Kd Departemen": sort = "d.kd_departement " & cbSortType.Text
        Case "Nama Departemen": sort = "d.nm_departement " & cbSortType.Text
    End Select
    
    With SQLParser
        .Fields = "d.id_departement,d.kd_departement,d.nm_departement,d2.nm_departement,d.bn,d.an,d.vn,d.rn,d.pn,DATE_FORMAT(d.tgl_input,'%Y-%m-%d'),p.nm_pengguna"
        .Tables = "tbl_departement d LEFT JOIN tbl_departement d2 ON d2.id_departement=d.parent_id LEFT JOIN tbl_pengguna p ON p.id=d.id_pengguna"
        .wCondition = "d.id_departement<>0"
        .SortOrder = sort & " LIMIT " & cbShow.Text
        .SaveStatement
    End With
        
    If rsDepartement.State = 1 Then rsDepartement.Close
    rsDepartement.CursorLocation = adUseClient
    rsDepartement.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsDepartement, 10000000
        FillList 1
    End With
    lbltotal.Caption = "Total Record : " & lvList.ListItems.Count
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsDepartement, RecordPage.PageStart, RecordPage.PageEnd, 13, 2, False, True, , , , "id_departement")
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
    Set rsDepartement = Nothing
    Set frmDepartement = Nothing
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
    CommandPass "Edit"
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub
