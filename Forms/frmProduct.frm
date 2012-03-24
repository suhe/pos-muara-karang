VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduct 
   Caption         =   "Medicine List"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   14865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProduct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   14865
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   14865
      TabIndex        =   4
      Top             =   6540
      Width           =   14865
      Begin VB.ComboBox cbSortType 
         Height          =   315
         ItemData        =   "frmProduct.frx":0A02
         Left            =   6960
         List            =   "frmProduct.frx":0A0C
         TabIndex        =   12
         Text            =   "ASC"
         Top             =   50
         Width           =   855
      End
      Begin VB.ComboBox cbSort 
         Height          =   315
         ItemData        =   "frmProduct.frx":0A1B
         Left            =   5520
         List            =   "frmProduct.frx":0A2B
         TabIndex        =   10
         Text            =   "Kode Obat"
         Top             =   50
         Width           =   1335
      End
      Begin VB.ComboBox cbShow 
         Height          =   315
         ItemData        =   "frmProduct.frx":0A5B
         Left            =   4140
         List            =   "frmProduct.frx":0A5D
         TabIndex        =   8
         Text            =   "30"
         Top             =   30
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   10200
         ScaleHeight     =   345
         ScaleWidth      =   6315
         TabIndex        =   5
         Top             =   0
         Width           =   6315
      End
      Begin VB.Label lbltotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Record : 0"
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   60
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sort"
         Height          =   195
         Left            =   5040
         TabIndex        =   9
         Top             =   60
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Show"
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   60
         Width           =   390
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   14865
      TabIndex        =   1
      Top             =   6945
      Width           =   14865
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   14865
      TabIndex        =   0
      Top             =   6960
      Width           =   14865
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6165
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Obat"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kd Obat"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nm Obat"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nm Ilmiah"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nm Kategori"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kemasan"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Harga Jual"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Harga Beli"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Profit"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Gudang"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Beli"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Jual"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Retur"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Sisa Stok"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Stok Min"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Tgl Input"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Nm Pengguna"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine List"
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
      TabIndex        =   2
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
      Width           =   8955
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CURR_COL As Integer
Dim rsproduct As New Recordset
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
            frmProductAE.State = adStateAddMode
            frmProductAE.show vbModal
        Case "Edit"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_obat", "id_obat", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    With frmProductAE
                        .State = adStateEditMode
                        .PK = CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                        .show vbModal
                    End With
                End If
            End If
        Case "Search"
            With frmSearch
                .srcNoOfCol = 4
                Set .srcform = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .show vbModal
            End With
        Case "Delete"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_obat", "id_obat", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    Dim ANS As Integer
                    ANS = MsgBox("Are you sure you want to delete the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation.", vbCritical + vbYesNo, "Confirm Record Delete")
                    Me.MousePointer = vbHourglass
                    If ANS = vbYes Then
                        If isRecordExist("tbl_jual_details", "id_obat", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                            DelRecwSQL "tbl_obat", "id_obat", "", True, CLng(LeftSplitUF(lvList.SelectedItem.Tag))
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
        Case "Print"
            On Error Resume Next
            If (lvList.ListItems.Count > 0) Then
                Call printStock
            Else
                MsgBox "No Data View In the List"
            End If
        Case "Close"
            Unload Me
    End Select
    Exit Sub
    'Trap the error
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
    With rsproduct
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

Private Sub cbShow_Change()
    cbShow.Text = 0
End Sub

Private Sub cbShow_Click()
    Call Form_Load
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub cbSort_Change()
    cbSort.Text = "Kode Obat"
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
    If CurrUser.USER_ISADMIN Then
        HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "ttttttt"
    Else
        HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "tftfttt"
    End If
    
    With MDIMainMenu
         .tbMenu.Buttons(3).Caption = "New"
         .tbMenu.Buttons(3).Image = 1
    End With
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
    With MDIMainMenu
        .AddToWin Me.Caption, Name
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With
    'sort berdsarkan
    Select Case cbSort.Text
        Case "Kode Obat": sort = "ABS(o.kd_obat) " & cbSortType.Text
        Case "Nama Obat": sort = "o.nm_obat " & cbSortType.Text
        Case "Penjualan": sort = "ABS(@jual) " & cbSortType.Text
        Case "Stok Sisa": sort = "ABS(@stok) " & cbSortType.Text
    End Select
    'Set the graphics for the controls
    sql = "o.id_obat,o.kd_obat ,o.nm_obat,o.nm_ilmiah,k.nm_kategori,o.kemasan,o.harga_jual,o.harga_beli,(o.harga_jual-o.harga_beli)as profit,o.stok,"
    sql = sql & "(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0)) AS beli,"
    sql = sql & "@jual:=(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0)) AS jual,"
    sql = sql & "(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0)) AS rugi,"
    sql = sql & "@stok:=((IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))-"
    sql = sql & "(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0))+ (o.stok) - "
    sql = sql & "(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))"
    sql = sql & " ) AS sisa,o.stok_min,"
    sql = sql & "DATE_FORMAT(o.tgl_input,'%Y-%m-%d'),p.nm_pengguna"
    
    With SQLParser
        .Fields = sql
        .Tables = " tbl_obat o INNER JOIN tbl_kategori k ON k.id_kategori =o.id_kategori INNER JOIN tbl_pengguna p ON p.id=o.id_pengguna "
        .SortOrder = sort & " LIMIT  " & cbShow.Text
        .SaveStatement
    End With
    
    If rsproduct.State = 1 Then rsproduct.Close
    rsproduct.CursorLocation = adUseClient
    rsproduct.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsproduct, 100000
        FillList 1
    End With
    lbltotal.Caption = "Total Record : " & lvList.ListItems.Count
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsproduct, RecordPage.PageStart, RecordPage.PageEnd, 17, 2, False, True, , , , "id_obat")
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
    Set frmProduct = Nothing
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

Private Sub lvList_DblClick()
    CommandPass "Edit"
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub
