VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchFakturDepartement 
   Caption         =   "Departement"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7290
      TabIndex        =   8
      Top             =   4995
      Width           =   7290
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7290
      TabIndex        =   7
      Top             =   5010
      Width           =   7290
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7290
      TabIndex        =   3
      Top             =   5025
      Width           =   7290
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   4
         Top             =   0
         Width           =   4150
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   5
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
         TabIndex        =   6
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
      Left            =   0
      TabIndex        =   0
      Top             =   4200
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
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmSearchFakturDepartement.frx":0000
         Left            =   240
         List            =   "frmSearchFakturDepartement.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image imgSearch 
         Height          =   480
         Left            =   6600
         Picture         =   "frmSearchFakturDepartement.frx":0004
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3915
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   7260
      _ExtentX        =   12806
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Departement"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kd Departement"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nm Departement"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "AN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "BN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "VN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "RN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "PN"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Name Of Departement over textbox"
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
      TabIndex        =   10
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
      Width           =   7275
   End
End
Attribute VB_Name = "frmSearchFakturDepartement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsSearchDepartement As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser

'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    SQLParser.RestoreStatement
    SQLParser.wCondition = srcCondition
    ReloadRecords SQLParser.SQLStatement
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
    With rsSearchDepartement
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

Private Sub Form_Load()
    On Error Resume Next
    Me.txtSrchStr.SetFocus
    MDIMainMenu.AddToWin Me.Caption, Name
    With cboFilter
        .AddItem "Name"
        .Text = "Name"
    End With
    'Set the graphics for the controls
    With MDIMainMenu
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsSearchDepartement, RecordPage.PageStart, RecordPage.PageEnd, 12, 2, False, True, , , , "id_departement")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
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
        'lvList.Height = (Me.ScaleHeight - Picture1.Height) - lvList.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSearchFakturDepartement = Nothing
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
    If (lvList.ListItems.Count > 0) Then
        tbl.TABLE_ID_DEPT = lvList.SelectedItem.Text
        tbl.TABLE_KD_DEPT = lvList.SelectedItem.SubItems(1)
        tbl.TABLE_NM_DEPT = lvList.SelectedItem.SubItems(2)
        With frmSearchFaktur
            .txtFilter(0).Text = tbl.TABLE_NM_DEPT
        End With
        Unload Me
    Else
        MsgBox "NO Data Selected !", vbCritical + vbInformation
    End If
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lvList_DblClick
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
    Dim str As String
    If (cboFilter.Text = "Code") Then
        str = "kd_departement"
    Else
        str = "nm_departement"
    End If
    
    If txtSrchStr.Text <> "" Then
        With SQLParser
            .Fields = "id_departement,kd_departement,nm_departement,an,bn,vn,rn,pn"
            .Tables = "tbl_departement"
            If (str = "kd_departement") Then
            .wCondition = str & " = " & txtSrchStr.Text & ""
            Else
            .wCondition = str & " LIKE '%" & txtSrchStr.Text & "%'"
            End If
            .SortOrder = "kd_departement ASC,bn ASC,an DESC LIMIT 20 "
            .SaveStatement
        End With
        
        If rsSearchDepartement.State = 1 Then rsSearchDepartement.Close
        Set rsSearchDepartement = New ADODB.Recordset
        rsSearchDepartement.CursorLocation = adUseClient
        rsSearchDepartement.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
        
        With RecordPage
            .Start rsSearchDepartement, 20
            FillList 1
        End With
    End If
End Sub

Private Sub txtSrchStr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        lvList.SetFocus
     End If
End Sub
