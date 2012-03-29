VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   120
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectBank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton sel5 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Refresh"
      Top             =   1635
      Width           =   315
   End
   Begin VB.CommandButton sel4 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete"
      Top             =   1320
      Width           =   315
   End
   Begin VB.CommandButton sel3 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit"
      Top             =   1005
      Width           =   315
   End
   Begin VB.CommandButton sel2 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "New"
      Top             =   690
      Width           =   315
   End
   Begin VB.CommandButton sel1 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Find"
      Top             =   375
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6960
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   315
      Left            =   8160
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4785
      Left            =   465
      TabIndex        =   2
      Top             =   375
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   8440
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
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
         Text            =   "Sort"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID Cabang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kd Cabang"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nm Cabang"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Almt Cabang"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kota Cabang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Jangka Waktu"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Group"
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
      Left            =   150
      TabIndex        =   8
      Top             =   75
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   75
      Top             =   75
      Width           =   9465
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public srcTextGroup     As TextBox
Public srcTextGAddress  As TextBox
Public rsGroup          As New Recordset
Public OPEN_COMMAND     As Integer

Private Sub Command1_Click()
    Call selectCurList
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If OPEN_COMMAND = 1 Then
        Me.Height = 6000
        Command1.Visible = False
        Command2.Visible = False
        
        lblTitle.Caption = Me.Caption
    End If
    '-Start setting up the graphics
    With MDIMainMenu
        sel1.Picture = .i16x16.ListImages(9).Picture
        sel2.Picture = .i16x16.ListImages(10).Picture
        sel3.Picture = .i16x16.ListImages(11).Picture
        sel4.Picture = .i16x16.ListImages(12).Picture
        sel5.Picture = .i16x16.ListImages(13).Picture
        
        Set ListView1.SmallIcons = .i16x16
        Set ListView1.Icons = .i16x16
    End With
    
    If rsGroup.State = 1 Then rsGroup.Close
    Set rsGroup = New ADODB.Recordset
    rsGroup.CursorLocation = adUseClient
    rsGroup.Open "SELECT * FROM tbl_cabang ORDER BY id_cabang ASC", CN, adOpenStatic, adLockOptimistic
    Call reload_rec
End Sub

Public Sub reload_rec()
    rsGroup.Filter = ""
    rsGroup.Requery
    FillListView ListView1, rsGroup, 8, 2, True, True, "id_cabang"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGroup = Nothing
End Sub

Private Sub selectCurList()
    If ListView1.ListItems.Count < 1 Then MsgBox "No record to select.", vbExclamation: Exit Sub
    On Error Resume Next
    srcTextGroup.Text = ListView1.SelectedItem.ListSubItems(1)
    srcTextGAddress = ListView1.SelectedItem.ListSubItems(2)
    If ListView1.SelectedItem.ListSubItems(3) <> "" Then srcTextGAddress.Text = srcTextGAddress.Text & "," & ListView1.SelectedItem.ListSubItems(3)
    If ListView1.SelectedItem.ListSubItems(4) <> "" Then srcTextGAddress.Text = srcTextGAddress.Text & "," & ListView1.SelectedItem.ListSubItems(4)
    If ListView1.SelectedItem.ListSubItems(5) <> "" Then srcTextGAddress.Text = srcTextGAddress.Text & "," & ListView1.SelectedItem.ListSubItems(5)
    Unload Me
End Sub

Private Sub sel1_Click()
    If ListView1.ListItems.Count < 1 Then MsgBox "No record to search.", vbExclamation: Exit Sub
    With frmFind
        Set .srcListView = ListView1
        .show vbModal
    End With
End Sub

Private Sub sel2_Click()
    With frmGroupAE
        .ADD_STATE = True
        .show vbModal
    End With
End Sub

Private Sub sel3_Click()
    If ListView1.ListItems.Count < 1 Then
        MsgBox "There is no record to edit.", vbInformation
        Exit Sub
    End If
    With frmGroupAE
        .ADD_STATE = False
        .PK = toNumber(ListView1.SelectedItem.Tag)
        .show vbModal
    End With
End Sub

Private Sub sel4_Click()
    On Error GoTo err
    With rsGroup
        '-Check if there is no record
        If .RecordCount < 1 Then MsgBox "No record to delete.", vbExclamation: Exit Sub
        '-Confirm deletion of record
        Dim ANS As Integer
        ANS = MsgBox("Are you sure you want to delete the selected record?", vbCritical + vbYesNo, "Confirm Record Delete")
        Me.MousePointer = vbHourglass
        If ANS = vbYes Then
            If isRecordExist("tbl_cabang", "id_cabang", toNumber(ListView1.SelectedItem.Tag)) = False Then
                MsgBox "This zip code is no longer exist in the record. Click ok to reload the records!", vbExclamation, "Unable To Edit"
                Me.MousePointer = vbDefault
                reload_rec
                Exit Sub
            End If
            '-Delete the record
            .AbsolutePosition = CInt(ListView1.SelectedItem)
            If isRecordExist("tbl_jual", "id_cabang", CLng(LeftSplitUF(ListView1.SelectedItem.Tag))) = False Then
                If isRecordExist("tbl_pengguna", "user_cabang", CLng(LeftSplitUF(ListView1.SelectedItem.Tag))) = False Then
                    .Delete
                    reload_rec
                    MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
                    MDIMainMenu.UpdateInfoMsg
                 End If
            Else
                MsgBox "Record not been deleted , this is record in the transaction OR FK Other table !.", vbInformation, "Confirm"
            End If
        End If
        ANS = 0
        Me.MousePointer = vbDefault
    End With
    Exit Sub
err:
        'prompt_err err, "frmSelectZipCode", "sel4_Click"
        Me.MousePointer = vbDefault
End Sub

Private Sub sel5_Click()
    Call reload_rec
End Sub
