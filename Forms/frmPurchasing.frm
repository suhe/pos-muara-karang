VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPurchasing 
   Caption         =   "Purchasing"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   Icon            =   "frmPurchasing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   12060
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel (F11)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Save (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame fraAmount 
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   3000
      Width           =   6855
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   4560
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
      Left            =   7080
      TabIndex        =   28
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmPurchasing.frx":038A
         Left            =   240
         List            =   "frmPurchasing.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
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
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Image imgSearch 
         Height          =   480
         Left            =   4320
         Picture         =   "frmPurchasing.frx":038E
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame fraProduct 
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   3855
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Harga Beli"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Stok "
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
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Obat"
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
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblBrand 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1320
         TabIndex        =   24
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1320
         TabIndex        =   23
         Top             =   1320
         Width           =   2385
      End
      Begin VB.Label Label8 
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
         TabIndex        =   22
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblstock 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   2385
      End
   End
   Begin VB.Frame fraFaktur 
      Caption         =   "Faktur Beli"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtFak 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Number"
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
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   12060
      TabIndex        =   13
      Top             =   8430
      Width           =   12060
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   14
         Top             =   0
         Width           =   4150
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
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
         TabIndex        =   16
         Top             =   60
         Width           =   1365
      End
   End
   Begin ComctlLib.ListView lstOrders 
      Height          =   3495
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Kd Obat"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nama Obat"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Kemasan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Harga Beli"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Jumlah"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sub Total"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3555
      Left            =   7080
      TabIndex        =   30
      Top             =   1080
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   6271
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kd Obat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nm Obat"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Kemasan"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Harga Beli"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Stok"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Stok Min"
         Object.Width           =   1940
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   5640
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchasing.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchasing.frx":1932
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchasing.frx":1C7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchasing.frx":43AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchasing.frx":5D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchasing.frx":8F75
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchasing.frx":984F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMoneyBack 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   33
      Text            =   "0"
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Frame fraSupplier 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cbtypePayment 
         Height          =   315
         ItemData        =   "frmPurchasing.frx":B1E1
         Left            =   1560
         List            =   "frmPurchasing.frx":B1EB
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cbpayment 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pembayaran"
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
         Top             =   1560
         Width           =   1080
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
         TabIndex        =   12
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Supplier ID"
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
         TabIndex        =   11
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cash/Transfer"
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
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblCodeCust 
         AutoSize        =   -1  'True
         Caption         =   "code"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblNamaCust 
         AutoSize        =   -1  'True
         Caption         =   "nama"
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
         TabIndex        =   8
         Top             =   720
         Width           =   2640
      End
   End
   Begin VB.Image imgQty 
      Height          =   480
      Left            =   5280
      Picture         =   "frmPurchasing.frx":B1FE
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "frmPurchasing.frx":BE42
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   7080
      Top             =   840
      Width           =   4755
   End
End
Attribute VB_Name = "frmPurchasing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CURR_COL   As Integer
Dim rspurchasing  As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser  As New clsSQLSelectParser
Dim rs         As New Recordset
Dim rsdetails  As New Recordset
Dim rskredit   As New Recordset
Public PK      As Long
Dim rsnew As Boolean

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF1: cmdNew_Click
        Case vbKeyF8: CommandPass "Close"
        Case vbKeyF10:
            If cmdNew.Enabled = False Then
                cmdProcess_Click
            Else
                MsgBox "Please New Transaction !", vbCritical + vbInformation
            End If
        Case vbKeyF11: cmdCancel_Click
    End Select
End Sub

Private Sub CONTROL(Active As Boolean)
    fraFaktur.Enabled = Active
    fraSupplier.Enabled = Active
    fraAmount.Enabled = Active
    fraProduct.Enabled = Active
    freSearch.Enabled = Active
    lstOrders.Enabled = Active
    cmdNew.Enabled = Not Active
    cmdProcess.Enabled = Active
    cmdCancel.Enabled = Active
End Sub

Private Sub clearText()
    On Error Resume Next
    lblTotal.Caption = 0
    txtFak.Text = ""
    txtSrchStr.Text = ""
    txtMoneyBack.Text = ""
    lvList.ListItems.Clear
    lblCodeCust.Caption = "......"
    lblNamaCust.Caption = "......"
    lblBrand.Caption = "---"
    lblstock.Caption = "---"
    lblPrice.Caption = "---"
    lvList.ListItems.Clear
    lstOrders.ListItems.Clear
End Sub

Private Sub GeneratePK()
    PK = getIndex("id_beli", "tbl_beli")
    txtFak.Text = "K" & tbl.TABLE_GROUP & PK
End Sub

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
    With rspurchasing
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

Private Sub cboFilter_LostFocus()
    On Error Resume Next
    txtSrchStr.SetFocus
End Sub

Private Sub cmdBrowse_Click()
    frmPurchasingSupplier.show vbModal
End Sub

Private Sub cmdCancel_Click()
    Call txtSrchStr_Change
    CONTROL False
    Call Form_Activate
    Call clearText
    txtFak.Text = ""
End Sub

Private Sub cmdNew_Click()
    CONTROL True
    Call clearText
    Call GeneratePK
    cmdRemove.Visible = True
    If (rsnew = True) Then
        frmPurchasingSupplier.show vbModal
    Else
        cmdBrowse.SetFocus
    End If
End Sub

Private Sub cash()
    Dim i As Integer
    Dim subtotal As Double
    Dim intResponse As Integer
        With rsdetails
            On Error Resume Next
            subtotal = 0
            For i = 1 To lstOrders.ListItems.Count
                .AddNew
                .Fields("no_beli") = Trim(txtFak.Text)
                .Fields("id_obat") = Trim(lstOrders.ListItems(i).Text)
                .Fields("harga_beli") = Format(lstOrders.ListItems(i).SubItems(4), "")
                .Fields("jumlah") = Format(lstOrders.ListItems(i).SubItems(5), "")
                .Fields("retur") = 0
                .Update
            Next i
        End With
        rsdetails.Close
        
        With tbl
            .TABLE_NO_FAK = txtFak.Text
            .TABLE_TANGGAL = Format(Date, "DD-MM-YYYY")
            .TABLE_TOTAL = Format(lblTotal.Caption, "")
        End With
        
        With rs
                .AddNew
                .Fields("no_beli") = Trim(txtFak.Text)
                .Fields("tgl_beli") = Now
                .Fields("tgl_akhir") = Now
                .Fields("id_supplier") = Trim(lblCodeCust.Caption)
                .Fields("tgl_input") = Now
                .Fields("id_pengguna") = Trim(CurrUser.USER_PK)
                .Fields("type") = "Cash"
                If (cbtypePayment.Text = "Lunas") Then
                    .Fields("payment") = "Lunas"
                    .Fields("flag_supplier") = 0
                    .Fields("tgl_bayar") = Format(Date, "YYYY-MM-DD")
                    .Fields("bayar") = Format(lblTotal.Caption, "")
                ElseIf (cbtypePayment.Text = "Hutang") Then
                    .Fields("payment") = "Hutang"
                    .Fields("flag_supplier") = 1
                    .Fields("hutang") = Format(lblTotal.Caption, "")
                    .Fields("tgl_bayar") = "-"
                Else
                    MsgBox "Invalid Payment Type", vbCritical + vbInformation
                    Exit Sub
                End If
                .Update
        End With
        rs.Close
End Sub

Private Sub cmdProcess_Click()
    If cbtypePayment.Text = "" Then MsgBox "Empty Type Of Payment", vbOKOnly + vbCritical: Exit Sub
    If lstOrders.ListItems.Count < 1 Then MsgBox "Empty Product", vbOKOnly + vbCritical: Exit Sub
    If lblTotal.Caption = 0 Then MsgBox "Please Insert Medicine ! ", vbOKOnly + vbCritical: Exit Sub
    If lblCodeCust.Caption = "......" Then
        MsgBox "Please Fill The Supplier Product! ", vbOKOnly + vbCritical, "Supplier Not Found"
        frmPurchasingSupplier.show vbModal
    Else
        Call cash
        On Error Resume Next
        Call cetak_FakturBeli
        lstOrders.ListItems.Clear
        Call Form_Load
        Call txtSrchStr_Change
    End If
    cmdRemove.Visible = False
    'MDIMainMenu.UpdateInfoMsg
End Sub

Private Sub Active()
    With MDIMainMenu
        .tbMenu.Buttons(9).Caption = "Close"
        .tbMenu.Buttons(9).Image = 7
    End With
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    If lstOrders.SelectedItem Is Nothing Then Exit Sub
    lstOrders.ListItems.Remove lstOrders.SelectedItem.Index
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fffffft"
    Call Active
End Sub

Public Sub counttotal()
    Dim i As Integer
    Dim subtotal As Double
    On Error Resume Next
    subtotal = 0
    For i = 0 To lstOrders.ListItems.Count
        subtotal = subtotal + lstOrders.ListItems(i).SubItems(6)
    Next i
    lblTotal.Caption = Format(subtotal, "##,###0.00")
End Sub

Private Sub Form_Deactivate()
    MDIMainMenu.HideTBButton "", True
End Sub
    

Private Sub Form_Load()
    On Error Resume Next
    rsnew = True
    MDIMainMenu.AddToWin Me.Caption, Name
    lstOrders.ListItems.Clear
    'Set the graphics for the controls
    With MDIMainMenu
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With
    
    With cboFilter
        .AddItem "Name"
        .AddItem "Code"
        .Text = "Name"
    End With
    
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM tbl_beli WHERE no_beli=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsdetails.State = 1 Then rsdetails.Close
    rsdetails.Open "SELECT * FROM tbl_beli_details WHERE no_beli=" & PK, CN, adOpenStatic, adLockOptimistic
    
    cboFilter.Text = "Name"
    lblCodeCust.Caption = "......"
    lblNamaCust.Caption = "......"
    CONTROL False
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rspurchasing, RecordPage.PageStart, RecordPage.PageEnd, 16, 2, False, True, , , , "id_obat")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        shpBar.Width = ScaleWidth
        lvList.Width = ScaleWidth - (lvList.Left + 100)
        lstOrders.Width = Me.ScaleWidth
        lstOrders.Height = (Me.ScaleHeight - Picture1.Height) - lvList.Top
        freSearch.Width = ScaleWidth - (freSearch.Left + 100)
        txtSrchStr.Width = freSearch.Width - (txtSrchStr.Left + imgSearch.Width)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainMenu.RemToWin Me.Caption
    MDIMainMenu.HideTBButton "", True
    Set frmPurchasing = Nothing
End Sub

Private Sub lvList_Click()
    If (lvList.ListItems.Count > 0) Then
        With lvList.SelectedItem
            lblBrand.Caption = .SubItems(1) & "(" & .SubItems(2) & ")"
            lblPrice.Caption = .SubItems(4)
            lblstock.Caption = .SubItems(5)
        End With
    End If
End Sub

Private Sub lvList_DblClick()
    On Error Resume Next
    Call callBrand
    frmPurchasingAE.show vbModal
End Sub

Private Sub callBrand()
    With frmPurchasingAE
        .lblBarCode.Caption = lvList.SelectedItem.Text
        .lblCode.Caption = lvList.SelectedItem.SubItems(1)
        .lblname.Caption = lvList.SelectedItem.SubItems(2)
        .lblKemasan.Caption = lvList.SelectedItem.SubItems(3)
        .lblPrice.Caption = lvList.SelectedItem.SubItems(4)
        .lblstock.Caption = lvList.SelectedItem.SubItems(5)
    End With
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Call callBrand
    frmPurchasingAE.txtQty.SetFocus
    frmPurchasingAE.show vbModal
End Sub

Private Sub m_save_transaction_Click()
    Call cmdProcess_Click
End Sub

Private Sub lvList_LostFocus()
    On Error Resume Next
    txtSrchStr.SetFocus
End Sub

Private Sub m_cancel_Click()
    On Error Resume Next
    Call cmdCancel_Click
End Sub

Private Sub m_new_Click()
    On Error Resume Next
    Call cmdNew_Click
End Sub

Private Sub m_save_Click()
    On Error Resume Next
    Call cmdProcess_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub InsertList()
   Dim itmX As ListItem
    With lstOrders.ListItems.Add
        .Text = "test"
        .SubItems(1) = "test"
    End With
   Set itmX = Nothing
End Sub

Private Sub txtSrchStr_Change()
    Dim str As String
    On Error Resume Next
    If cboFilter.Text = "Code" Then
        str = "kd_obat"
    Else
        str = "nm_obat"
    End If
     If txtSrchStr.Text <> "" Then
        sql = "o.id_obat,o.kd_obat ,o.nm_obat,o.kemasan,FORMAT(o.harga_beli,0),"
        sql = sql & "((IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))-"
        sql = sql & "(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0))+ (o.stok) - "
        sql = sql & "(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))"
        sql = sql & " ) AS sisa,o.stok_min"

        With SQLParser
            .Fields = sql
            .Tables = " tbl_obat o INNER JOIN tbl_kategori k ON k.id_kategori =o.id_kategori INNER JOIN tbl_pengguna p ON p.id=o.id_pengguna "
            .wCondition = str & " Like '%" & txtSrchStr.Text & "%'"
            .SortOrder = " o.id_obat ASC LIMIT 20"
            .SaveStatement
        End With
    
        If rspurchasing.State = 1 Then rspurchasing.Close
        Set rspurchasing = New ADODB.Recordset
        rspurchasing.CursorLocation = adUseClient
        rspurchasing.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
        
        With RecordPage
            .Start rspurchasing, 20
            FillList 1
        End With
        
        'rspurchasing.Close
        'Set rspurchasing = Nothing
        If rspurchasing.State = 1 Then rspurchasing.Close
        
    End If
End Sub

Private Sub txtSrchStr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        lvList.SetFocus
     End If
End Sub
