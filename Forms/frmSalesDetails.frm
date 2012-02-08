VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalesDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Details"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "frmSalesDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Lunasi* (F2)"
      Height          =   495
      Left            =   4200
      TabIndex        =   46
      Top             =   2760
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
      TabIndex        =   41
      Top             =   6480
      Width           =   9135
      Begin VB.TextBox txtPiutang 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Text            =   "frmSalesDetails.frx":038A
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtBayar 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "frmSalesDetails.frx":0397
         Top             =   360
         Width           =   2775
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
         TabIndex        =   44
         Top             =   120
         Width           =   645
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
         Left            =   3360
         TabIndex        =   43
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kreditor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      TabIndex        =   33
      Top             =   1440
      Width           =   5175
      Begin VB.CheckBox chKreditor 
         Caption         =   "Lunas Kreditor"
         Height          =   255
         Left            =   2280
         TabIndex        =   48
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtKdKreditor 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Text            =   "frmSalesDetails.frx":03A5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtNmKreditor 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Text            =   "frmSalesDetails.frx":03A9
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         TabIndex        =   38
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Left            =   360
         TabIndex        =   37
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
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
         TabIndex        =   36
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Departement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   27
      Top             =   1440
      Width           =   3975
      Begin VB.CheckBox chDep 
         Caption         =   "Lunas Debitor"
         Height          =   255
         Left            =   2280
         TabIndex        =   47
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtKdDepartement 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "frmSalesDetails.frx":03C3
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtNmDepartement 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "frmSalesDetails.frx":03C8
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         TabIndex        =   32
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Left            =   360
         TabIndex        =   31
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
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
         TabIndex        =   30
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar (F6)"
      Height          =   495
      Left            =   7560
      TabIndex        =   26
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdStruck 
      Caption         =   "&Cetak Struk (F3)"
      Height          =   495
      Left            =   5880
      TabIndex        =   25
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame fraCustomer 
      Caption         =   "Pasien"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtNmPasien 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmSalesDetails.frx":03D6
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txtKdPasien 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "frmSalesDetails.frx":03E4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
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
         TabIndex        =   19
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Left            =   360
         TabIndex        =   17
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         TabIndex        =   16
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame fraFaktur 
      Caption         =   "Faktur Jual"
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
      TabIndex        =   12
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtTanggal 
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Text            =   "frmSalesDetails.frx":03E9
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTypeBayar 
         Height          =   375
         Left            =   2640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "frmSalesDetails.frx":03F6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtType 
         Height          =   375
         Left            =   600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Text            =   "frmSalesDetails.frx":03FF
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFak 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal."
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
         Left            =   2520
         TabIndex        =   39
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bayar"
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
         Left            =   2040
         TabIndex        =   23
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         TabIndex        =   21
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No."
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
         TabIndex        =   14
         Top             =   240
         Width           =   255
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
      ScaleWidth      =   9270
      TabIndex        =   9
      Top             =   7425
      Width           =   9270
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   9270
      TabIndex        =   8
      Top             =   7440
      Width           =   9270
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9270
      TabIndex        =   0
      Top             =   7455
      Width           =   9270
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Next 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Last 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Previous 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   2
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
            TabIndex        =   6
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
         TabIndex        =   7
         Top             =   60
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3075
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   5424
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Obat"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kd Obat"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nm Obat"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Harga Jual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Jumlah"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "SubTotal"
         Object.Width           =   3175
      EndProperty
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
      TabIndex        =   11
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmSalesDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURR_COL As Integer
Dim rsSalesDetails As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser As New clsSQLSelectParser
Public srcform As Form

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
    On Error GoTo err
    With rsSalesDetails
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

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdStruck_Click()
    On Error Resume Next
    Call cetak_Faktur2
End Sub

Private Sub Command1_Click()
    Dim dep, kreditor As Byte
    Dim payment As String
    Dim bayar, piutang As Double
    If chDep.Value = 1 Then
        dep = 0
    Else
        dep = 1
    End If
    
    If chKreditor.Value = 1 Then
        kreditor = 0
        bayar = tbl.TABLE_TOTAL
        piutang = 0
        payment = "Lunas"
    Else
        kreditor = 1
        bayar = 0
        piutang = tbl.TABLE_TOTAL
        payment = "Hutang"
    End If
    Dim intResponse As String
    intResponse = MsgBox("Saya bersedia bertanggung jawab dengan mengklik tombol Yes!", vbYesNo + vbInformation, "Warning")
    If intResponse = vbYes Then
        sql = "UPDATE tbl_jual "
        sql = sql + "SET "
        sql = sql + " payment='" & payment & "', "
        sql = sql + " flag_kreditor=" & kreditor & ", "
        sql = sql + " bayar=" & bayar & ", "
        sql = sql + " dibayar=" & bayar & ", "
        sql = sql + " piutang=" & piutang & ", "
        sql = sql + " flag_debitor=" & dep & " "
        sql = sql + " WHERE no_jual='" & tbl.TABLE_NO_FAK & "'"
        CN.Execute sql
        MsgBox sql
    Else
        Unload Me
    End If
   
    'On Error Resume Next
    frmSales.RefreshRecords
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
    frmPurchaseDetails.Caption = "No Fak Beli : " & tbl.TABLE_NO_FAK
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
        .Fields = "tbd.id_obat,o.kd_obat,o.nm_obat,tbd.harga_jual,tbd.jumlah,(tbd.jumlah* tbd.harga_jual) as subtotal"
        .Tables = "tbl_jual_details tbd JOIN tbl_obat o ON o.id_obat=tbd.id_obat"
        .wCondition = " tbd.no_jual='" & tbl.TABLE_NO_FAK & "'"
        .SaveStatement
    End With
    
    If rsSalesDetails.State = 1 Then rsSalesDetails.Close
    rsSalesDetails.CursorLocation = adUseClient
    rsSalesDetails.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPage
        .Start rsSalesDetails, 1000000
        FillList 1
    End With
    'rsSalesDetails .Close
    txtFak.Text = tbl.TABLE_NO_FAK
    txtTanggal.Text = tbl.TABLE_TANGGAL
    txtKdPasien.Text = tbl.TABLE_KD_PASIEN
    txtNmPasien.Text = tbl.TABLE_NM_PASIEN
    txtKdDepartement.Text = tbl.TABLE_KD_DEPT
    txtNmDepartement.Text = tbl.TABLE_NM_DEPT
    If (tbl.TABLE_NM_KREDITUR = "") Then
        txtKdKreditor.Text = "Tidak Ada"
        txtNmKreditor.Text = "Tidak Ada"
    Else
        txtKdKreditor.Text = tbl.TABLE_ID_KREDITUR
        txtNmKreditor.Text = tbl.TABLE_NM_KREDITUR
    End If
    
    If (tbl.TABLE_FLAG_DEPT = 1) Then
        chDep.Value = 0
    End If
    
    If (tbl.TABLE_FLAG_KREDITOR = 1) Then
        chKreditor.Value = 0
    End If
    
    txtType.Text = tbl.TABLE_TYPE
    txtTypeBayar.Text = tbl.TABLE_PAY_TYPE
    txtBayar.Text = tbl.TABLE_TOTAL
    txtPiutang.Text = tbl.TABLE_CBACK
    If (tbl.TABLE_PAY_TYPE <> "Hutang") Then
        'cmdInvoice.Enabled = False
    End If
    rsSalesDetails.Close
    Set rsSalesDetails = Nothing
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rsSalesDetails, RecordPage.PageStart, RecordPage.PageEnd, 9, 2, False, True, , , , "id_obat")
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
        lvList.Height = (Me.ScaleHeight - (Picture1.Height + Frame3.Height)) - lvList.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSalesDetails = Nothing
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

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

