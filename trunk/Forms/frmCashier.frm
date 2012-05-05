VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCashier 
   Caption         =   "Cashier"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   12030
   Icon            =   "frmCashier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPasien 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   37
      Top             =   1080
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "&Pasien Baru"
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
         TabIndex        =   54
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTlpPasien 
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
         Left            =   840
         TabIndex        =   47
         Top             =   1560
         Width           =   2385
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Tlp"
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
         TabIndex        =   46
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lblRelasi 
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
         Left            =   840
         TabIndex        =   45
         Top             =   1200
         Width           =   2385
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
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
         TabIndex        =   44
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label17 
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
         TabIndex        =   43
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label16 
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
         TabIndex        =   42
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblKdPasien 
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
         Left            =   840
         TabIndex        =   41
         Top             =   120
         Width           =   2385
      End
      Begin VB.Label lblAlmtPasien 
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
         Left            =   840
         TabIndex        =   40
         Top             =   840
         Width           =   2865
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
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
         TabIndex        =   39
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label lblNmPasien 
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
         Left            =   840
         TabIndex        =   38
         Top             =   480
         Width           =   2385
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1935
      Left            =   4080
      TabIndex        =   35
      Top             =   0
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmCashier.frx":038A
         Left            =   1560
         List            =   "frmCashier.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Tag             =   "Jangka Waktu"
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   57
         Top             =   960
         Width           =   255
      End
      Begin VB.ComboBox cbPay 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCashier.frx":038E
         Left            =   1200
         List            =   "frmCashier.frx":0398
         TabIndex        =   52
         Text            =   "Lunas"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cbpayment 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCashier.frx":03AB
         Left            =   120
         List            =   "frmCashier.frx":03B5
         TabIndex        =   51
         Text            =   "Cash"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtKreditor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Tag             =   "Kreditor"
         Text            =   "frmCashier.frx":03C9
         Top             =   960
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dcDepartement 
         Height          =   360
         Left            =   120
         TabIndex        =   48
         Tag             =   "Departement"
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jangka Waktu"
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
         TabIndex        =   60
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hari"
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
         Left            =   2400
         TabIndex        =   59
         Top             =   1560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label18 
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
         TabIndex        =   36
         Top             =   3960
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdResep 
      Caption         =   "&Lunasi Resep (F12)"
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
      Left            =   4080
      TabIndex        =   32
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal (F5)"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame fraCashBack 
      Caption         =   "Money Back"
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
      Left            =   3600
      TabIndex        =   30
      Top             =   4080
      Width           =   3375
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraPayment 
      Caption         =   "Payment Money"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   4080
      Width           =   3375
      Begin VB.TextBox txtPayment 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Rp""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   2
         EndProperty
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
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   12030
      TabIndex        =   21
      Top             =   8310
      Width           =   12030
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
         Left            =   1680
         TabIndex        =   33
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
         TabIndex        =   22
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "First 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Previous 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Last 250"
            Top             =   10
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   23
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
            TabIndex        =   27
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
         TabIndex        =   28
         Top             =   60
         Width           =   1365
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
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtFak 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   480
         Width           =   3495
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
         TabIndex        =   20
         Top             =   240
         Width           =   660
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
      Height          =   1455
      Left            =   4080
      TabIndex        =   10
      Top             =   1920
      Width           =   2895
      Begin VB.Label lblstock 
         AutoSize        =   -1  'True
         Caption         =   "................................"
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
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   1920
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
         TabIndex        =   16
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "................................"
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
         Left            =   840
         TabIndex        =   15
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label lblBrand 
         AutoSize        =   -1  'True
         Caption         =   "................................"
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
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   1920
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Obat"
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
         TabIndex        =   13
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Stok"
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
         TabIndex        =   12
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Harga"
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
         Top             =   1080
         Width           =   510
      End
   End
   Begin VB.Frame fraSearch 
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
      TabIndex        =   6
      Top             =   0
      Width           =   4815
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
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtSrchStrPasien 
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
         Width           =   2055
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmCashier.frx":03CD
         Left            =   240
         List            =   "frmCashier.frx":03CF
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image imgSearch 
         Height          =   480
         Left            =   4320
         Picture         =   "frmCashier.frx":03D1
         Top             =   120
         Width           =   480
      End
   End
   Begin ComctlLib.ListView lstOrders 
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5106
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
         Size            =   11.25
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
         Text            =   "ID Obat"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Kd Obat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nm Obat"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Kemasan"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Harga Jual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Jumlah"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Dosis"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   120
      Top             =   7680
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
            Picture         =   "frmCashier.frx":0C9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashier.frx":1975
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashier.frx":1CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashier.frx":43ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashier.frx":5D7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashier.frx":8FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashier.frx":9892
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAMount 
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
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3360
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
         Height          =   555
         Left            =   2040
         TabIndex        =   9
         Top             =   120
         Width           =   4560
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Berikutnya (F1)"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Cetak (F2)"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4275
      Left            =   7080
      TabIndex        =   53
      Top             =   960
      Width           =   4785
      _ExtentX        =   8440
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kd Pasien"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nm Pasien"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "No.Tlp"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Alamat"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Umur"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView lvListObat 
      Height          =   4275
      Left            =   7080
      TabIndex        =   55
      Top             =   960
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Obat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kd Obat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Merk"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kemasan"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Harga Jual"
         Object.Width           =   2293
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
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Closed Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   34
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   7080
      Top             =   720
      Width           =   4755
   End
End
Attribute VB_Name = "frmCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CURR_COL   As Integer
Dim rscashier  As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser  As New clsSQLSelectParser
Dim rs         As New Recordset
Dim rsdetails  As New Recordset
Public PK      As Long

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

Private Sub CONTROL(Active As Boolean)
    fraFaktur.Enabled = Active
    dcDepartement.Enabled = Active
    fraAMount.Enabled = Active
    fraPayment.Enabled = Active
    fraCashBack.Enabled = Active
    fraProduct.Enabled = Active
    fraSearch.Enabled = Active
    lstOrders.Enabled = Active
    cmdNew.Enabled = Not Active
    cmdCancel.Enabled = Active
    cmdProcess.Enabled = Active
    txtPayment.Enabled = Active
    txtMoneyBack.Enabled = Active
    Command2.Enabled = Active
End Sub

Private Sub controlPasien(Active As Boolean)
    cmdResep.Enabled = Active
    fraProduct.Enabled = Active
    Command1.Enabled = Not Active
    Command2.Enabled = Active
    cbPay.Enabled = Active
    cbpayment.Enabled = Active
    lstOrders.Enabled = Active
    cmdRemove.Enabled = Active
    dcDepartement.Text = ""
End Sub

Private Sub GeneratePK()
    On Error Resume Next
    PK = getIndex("id_jual", "tbl_jual")
    txtFak.Text = "M" & tbl.TABLE_GROUP & PK
End Sub

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
    With rscashier
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

Private Sub cmdBrowse_Click()
    On Error Resume Next
    frmCashierCustomer.show vbModal
End Sub

Private Sub cmdCancel_Click()
    Set rsdetails = Nothing
    Call txtSrchStr_Change
    CONTROL False
    controlPasien True
    lvListObat.Visible = False
    lvList.Visible = True
    txtSrchStr.Visible = False
    txtSrchStrPasien.Visible = True
    Call Form_Activate
    lblStatus.Caption = "Cancel Transaction"
    Call clearText
    txtFak.Text = ""
End Sub

Private Sub cmdNew_Click()
    CONTROL True
    Command1.Visible = True
    Call clearText
    dcDepartement.Text = ""
    dcDepartement.Enabled = False
    Combo1.Visible = False
    Label1.Visible = False
    Label5.Visible = False
    txtKreditor.Text = "0"
    txtSrchStr.Text = ""
    Call GeneratePK
    controlPasien False
    With CurrBiz
        .BUSINNES_NEW = 1
        .BUSINNES_RECEPT = 0
    End With
    txtPayment.Enabled = False
    txtMoneyBack.Enabled = False
    lvListObat.ListItems.Clear
    lvListObat.Visible = False
    lvList.Visible = True
    txtSrchStrPasien.Visible = True
    txtSrchStr.Visible = False
End Sub

Private Sub clearText()
    lbltotal.Caption = 0
    txtFak.Text = ""
    txtSrchStrPasien.Text = ""
    txtPayment.Text = ""
    txtMoneyBack.Text = ""
    lstOrders.ListItems.Clear
    lvList.ListItems.Clear
    lblBrand.Caption = "---"
    lblStock.Caption = "---"
    lblPrice.Caption = "---"
    lblKdPasien.Caption = "..."
    lblNmPasien.Caption = "..."
    lblAlmtPasien.Caption = "..."
    lblRelasi.Caption = "..."
    lblTlpPasien.Caption = "..."
End Sub

Private Sub cmdPrint_Click()
    Dim intResponse As Integer
    If txtFak.Text = "" Then
        MsgBox "No Data Is Printed", vbOKOnly + vbCritical, "Warning"
    Else
        intResponse = MsgBox("Are you sure you want to Print!", vbYesNo + vbCritical, "Warning")
        If intResponse = vbYes Then
            
        End If
    End If
End Sub

Private Sub cmdProcess_Click()
    On Error Resume Next
    If CurrBiz.BUSINNES_NEW = 1 And CurrBiz.BUSINNES_RECEPT = 0 Then
        Call newPatient
    ElseIf CurrBiz.BUSINNES_RECEPT = 1 And CurrBiz.BUSINNES_NEW = 0 Then
        Call recipeMedicine
    End If
    lvListObat.ListItems.Clear
    lvList.ListItems.Clear
End Sub

Private Sub newPatient()
    Dim i As Integer
    Dim subtotal As Double
    Dim intResponse As String
    Dim details As Byte
    If is_empty(txtFak, True) = True Then Exit Sub
    If lblKdPasien.Caption = "..." Then MsgBox "Data Pasien Tidak Boleh Kosong !": Exit Sub
    details = getRecordCount("no_jual", "tbl_jual", "WHERE no_jual ='" & PK & "' AND kd_pasien='" & Trim(lblKdPasien.Caption) & "' ")
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM tbl_jual WHERE no_jual='" & PK & "' LIMIT 1 ", CN, adOpenStatic, adLockOptimistic
    If (rs.RecordCount > 0) Then
        MsgBox "Not A Data this Faktur !"
    Else
        With rs
            .AddNew
            .Fields("no_jual") = Trim(txtFak.Text)
            .Fields("tgl_jual") = Now
            .Fields("kd_pasien") = Trim(lblKdPasien.Caption)
            .Fields("id_cabang") = CurrBiz.BUSINNES_GROUP
            .Fields("tgl_input") = Now
            .Fields("tgl_akhir") = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
            .Fields("id_pengguna") = CurrUser.USER_PK
            .Update
        End With
            
            tbl.TABLE_NO_FAK = Trim(txtFak.Text)
            tbl.TABLE_KD_PASIEN = Trim(lblKdPasien.Caption)
            tbl.TABLE_NM_PASIEN = lblNmPasien.Caption
            tbl.TABLE_NM_DEPT = dcDepartement.Text
            tbl.TABLE_UMUR_PASIEN = lblRelasi.Caption
            tbl.TABLE_TLP_PASIEN = lblTlpPasien.Caption
            tbl.TABLE_TANGGAL = Format(Now, "DD-MM-YYYY")
            MsgBox "Thank You!", vbOKOnly + vbInformation, "Warning"
            Call cetak_Faktur
            controlPasien True
         Call Form_Load
         Combo1.Visible = False
         Label1.Visible = False
         Label5.Visible = False
    End If
End Sub

Private Sub recipeMedicine()
    If (txtPayment.Text <> "Credit") Then
        If txtPayment.Text = "" Then MsgBox "Empty Payment ", vbOKOnly + vbCritical: Exit Sub
        If txtMoneyBack.Text = "" Then MsgBox "Empty Cashback", vbOKOnly + vbCritical: Exit Sub
        If txtPayment.Text = 0 Then MsgBox "Please Insert Payment ! ", vbOKOnly + vbCritical: Exit Sub
    End If
    
    If (Combo1.Visible = True) Then
        If is_empty(Combo1, True) = True Then Exit Sub
    End If
    
    If lstOrders.ListItems.Count < 1 Then MsgBox "Please Insert Medicine To Cashier ! ", vbOKOnly + vbCritical: Exit Sub
    If dcDepartement.Text = "" Then MsgBox "Empty Departement ", vbOKOnly + vbCritical: Exit Sub
    cmdRemove.Visible = False
    Dim i As Integer
    Dim kode As String
    Dim payment, total, BN, AN, PN, RN, VN, Om, Kn As Double
    Dim intResponse As Integer
    payment = Format(txtPayment.Text, "")
    payment = Replace(payment, ".", ",")
    total = Format(lbltotal.Caption, "")
    If (txtPayment.Text <> "Credit") Then
        If (Val(total) > Val(payment)) Then MsgBox "Sorry Not Enought Money ,Please Insert Money! ", vbOKOnly + vbCritical: Exit Sub
    Else
    
    End If
    
        'Perhitungan Komisi
        Dim komisi, bayar, piutang As Double
        Dim strpay As String
        If txtKreditor.Text = 0 Then
            bayar = total
            piutang = 0
            strpay = "Lunas"
            tbl.TABLE_TYPE = "Cash"
        Else
            bayar = 0
            piutang = Format(lbltotal.Caption, "")
            strpay = "Piutang"
            tbl.TABLE_TYPE = "Credit"
        End If
               
        'cari departement
        Dim rsKomisi As New Recordset
        Set rsKomisi = New Recordset
        
        If rsKomisi.State = 1 Then rsKomisi.Close
        rsKomisi.Open "SELECT * FROM tbl_departement WHERE id_departement=" & dcDepartement.BoundText, CN, adOpenStatic, adLockReadOnly
        If (rsKomisi.RecordCount > 0) Then
            kode = rsKomisi.Fields("kd_departement")
            BN = rsKomisi.Fields("bn")
            AN = rsKomisi.Fields("an")
            PN = rsKomisi.Fields("pn")
            RN = (rsKomisi.Fields("rn")) / 100
            VN = rsKomisi.Fields("vn")
        Else
            BN = 0
            AN = 0
            PN = 0
            RN = 0
            VN = 0
        End If
        rsKomisi.Close
        Set rsKomisi = Nothing
        'perhitungan komisi
        Om = bayar + piutang
        If (Left(kode, 1) = 1) Then
        'Rumus Kn = ((Om - Vn) * Rn) + Pn
             Kn = ((Om - VN) * RN) + PN
        ElseIf (Left(kode, 1) = 2) Then
        '  Rumus Kn = ((Om - Vn) * Rn) + Pn
            If (Om > BN) And (Om < AN) Then
                Kn = ((Om - VN) * RN) + PN
            Else
                If rsKomisi.State = 1 Then rsKomisi.Close
                rsKomisi.Open "SELECT * FROM tbl_departement WHERE parent_id=" & dcDepartement.BoundText & " AND bn <= " & Om & " AND an >= " & Om & " ORDER BY bn ASC ", CN, adOpenStatic, adLockReadOnly
                If rsKomisi.RecordCount > 0 Then
                        BN = rsKomisi.Fields("bn")
                        AN = rsKomisi.Fields("an")
                        PN = rsKomisi.Fields("pn")
                        RN = (rsKomisi.Fields("rn")) / 100
                        VN = rsKomisi.Fields("vn")
                        Kn = ((Om - VN) * RN) + PN
                Else
                        BN = 0
                        AN = 0
                        PN = 0
                        RN = 0
                        VN = 0
                        Kn = ((Om - VN) * RN) + PN
                End If
            End If
        End If
        
        With rsdetails
            For i = 1 To lstOrders.ListItems.Count
                Dim details As Byte
                details = getRecordCount("no_jual", "tbl_jual_details", "WHERE no_jual ='" & Trim(txtFak.Text) & "' AND id_obat='" & Trim(lstOrders.ListItems(i).Text) & "' ")
                If (details > 0) Then
                Else
                    .AddNew
                    .Fields("no_jual") = txtFak.Text
                    .Fields("id_obat") = lstOrders.ListItems(i).Text
                    .Fields("harga_jual") = lstOrders.ListItems(i).SubItems(4)
                    .Fields("jumlah") = lstOrders.ListItems(i).SubItems(5)
                    .Fields("dosis") = lstOrders.ListItems(i).SubItems(6)
                    .Update
                End If
            Next i
        End With
        rsdetails.Close
        
        sql = "UPDATE tbl_jual "
        sql = sql + " SET "
        sql = sql + " payment='" & strpay & "', "
        sql = sql + " type='Cash', "
        sql = sql + " id_departement=" & Trim(dcDepartement.BoundText) & ", "
        sql = sql + " id_kreditor=" & Trim(txtKreditor.Text) & ", "
        sql = sql + " bayar=" & Format(bayar, "") & ", "
        sql = sql + " piutang=" & Format(piutang, "") & ", "
        
        'kreditor
        If ((txtKreditor.Text = 0) And (txtPayment.Text <> "Credit")) Then
            sql = sql + " flag_kreditor= 0,"
            sql = sql + " flag_debitor = 1,"
            sql = sql + " jw=0,"
            sql = sql + " dibayar=" & Format(txtPayment.Text, "") & ", "
            sql = sql + " tgl_bayar='" & Format(Date, "YYYY-MM-DD") & "',"
        Else
            sql = sql + " flag_kreditor=1,"
            sql = sql + " flag_debitor =1,"
            sql = sql + " jw=" & Combo1.Text & ","
            sql = sql + " dibayar=" & Format(Replace(bayar, ",", "."), "") & ", "
            sql = sql + " tgl_bayar='-' ,"
        End If
        
        sql = sql + " komisi=" & Replace(Kn, ",", ".") & " "
        sql = sql + " WHERE no_jual='" & Trim(txtFak.Text) & "' "
        sql = sql + " AND kd_pasien='" & Trim(lblKdPasien.Caption) & "'"
        
        CN.Execute sql
        Call txtSrchStr_Change
        CONTROL False
        Call Form_Activate
        lstOrders.Enabled = False
        lblStatus.Caption = "Saved Transaction !"
        MDIMainMenu.UpdateInfoMsg 'Display the business status
        tbl.TABLE_NO_FAK = Trim(txtFak.Text)
        tbl.TABLE_KD_PASIEN = Trim(lblKdPasien.Caption)
        tbl.TABLE_NM_PASIEN = lblNmPasien.Caption
        tbl.TABLE_NM_DEPT = dcDepartement.Text
        tbl.TABLE_RELASI = lblRelasi.Caption
        tbl.TABLE_TLP_PASIEN = lblTlpPasien.Caption
        tbl.TABLE_PAY_TYPE = txtPayment.Text
        tbl.TABLE_NM_DEPT = dcDepartement.Text
        tbl.TABLE_ID_KREDITUR = Trim(txtKreditor.Text)
        tbl.TABLE_TANGGAL = Format(Now, "DD-MM-YYYY")
        tbl.TABLE_TOTAL = lbltotal.Caption
        tbl.TABLE_KOMISI = komisi
        If txtPayment <> "Credit" Then
            tbl.TABLE_MONEY = txtPayment.Text
            tbl.TABLE_CBACK = txtMoneyBack.Text
        End If
        Set rsKomisi = Nothing
        frmCashierDesc.show vbModal
        Call Form_Load
        cmdResep.Enabled = True
        Combo1.Visible = False
        Label1.Visible = False
        Label5.Visible = False
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

Private Sub clearPayment()
    lbltotal.Caption = 0
    txtPayment.Text = ""
    txtMoneyBack.Text = ""
    txtKreditor.Text = 0
    lblKdPasien.Caption = "---"
    lblNmPasien.Caption = "---"
    lblRelasi.Caption = "---"
    lblTlpPasien.Caption = "---"
    txtSrchStr.Text = ""
    txtSrchStrPasien.Text = ""
    dcDepartement.Text = ""
    lstOrders.ListItems.Clear
    lvList.ListItems.Clear
    lvListObat.ListItems.Clear
End Sub

Private Sub cmdResep_Click()
    Dim strName As String
    Dim strVB As String
    Call clearPayment
    txtPayment.Enabled = True
    txtMoneyBack.Enabled = True
    Combo1.Visible = False
    Label1.Visible = False
    Label5.Visible = False
    With CurrBiz
        .BUSINNES_NEW = 0
        .BUSINNES_RECEPT = 1
    End With
    strName = InputBox("Masukkan Nomor Faktur : ", "No Faktur")
    Dim total, total_obat As Byte
    strName = Trim(strName)
    total = getRecordCount("id_jual", "tbl_jual", "WHERE no_jual ='" & strName & "' ")
    total_obat = getRecordCount("no_jual", "tbl_jual_details", "WHERE no_jual ='" & strName & "' ")
    If ((total > 0) And (total_obat < 1)) Then
        strVB = MsgBox("No Faktur Ditemukan Data ditampilkan !", vbOKCancel + vbInformation)
        If (strVB = vbOK) Then
            tbl.TABLE_NO_FAK = Trim(strName)
            Call Cashier
            CONTROL True
            lvListObat.Visible = True
            lvList.Visible = False
            txtSrchStr.Visible = True
            txtSrchStrPasien.Visible = False
            Command1.Visible = False
            cmdResep.Enabled = False
            rs.Close
            Set rs = Nothing
        End If
    Else:
        strVB = MsgBox("No Faktur Tidak Ditemukan Atau Obat Sudah ditebus,Silahkan Isi Kembali !", vbOKCancel + vbInformation)
        If (strVB = vbOK) Then
            Call cmdResep_Click
        End If
    End If
End Sub

Private Sub Cashier()
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM tbl_jual j INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien LEFT JOIN tbl_departement d ON d.id_departement=j.id_departement LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor WHERE j.no_jual='" & tbl.TABLE_NO_FAK & "' LIMIT 1 ", CN, adOpenStatic, adLockOptimistic
    If rsdetails.State = 1 Then rsdetails.Close
    rsdetails.Open "SELECT * FROM tbl_jual_details WHERE no_jual='" & tbl.TABLE_NO_FAK & "' LIMIT 1 ", CN, adOpenStatic, adLockOptimistic
    If (rs.RecordCount > 0) Then
        txtFak.Text = rs.Fields("no_jual")
        lblKdPasien.Caption = rs.Fields("kd_pasien")
        lblNmPasien.Caption = rs.Fields("nm_pasien")
        lblRelasi.Caption = rs.Fields("relasi")
        lblAlmtPasien.Caption = rs.Fields("alamat")
        lblTlpPasien.Caption = rs.Fields("no_tlp")
    End If
End Sub

Private Sub Command1_Click()
    frmCashierCNewPasien.show vbModal
End Sub

Private Sub Command2_Click()
    frmCashierKreditor.show vbModal
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
    If (lstOrders.ListItems.Count > 0) Then
        For i = 0 To lstOrders.ListItems.Count
            subtotal = subtotal + (lstOrders.ListItems(i).SubItems(4) * lstOrders.ListItems(i).SubItems(5))
        Next i
    End If
    lbltotal.Caption = Format(subtotal, "##,###0.00")
End Sub

Private Sub Form_Deactivate()
    MDIMainMenu.HideTBButton "", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF1: cmdNew_Click
        Case vbKeyF2: cmdProcess_Click
        Case vbKeyF12: cmdResep_Click
        Case vbKeyF5: cmdCancel_Click
        Case vbKeyF8: CommandPass "Close"
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    MDIMainMenu.AddToWin Me.Caption, Name
    lstOrders.ListItems.Clear
    'Set the graphics for the controls
    With MDIMainMenu
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
        Set lvListObat.SmallIcons = .i16x16
        Set lvListObat.Icons = .i16x16
    
        btnFirst.Picture = .i16x16.ListImages(3).Picture
        btnPrev.Picture = .i16x16.ListImages(4).Picture
        btnNext.Picture = .i16x16.ListImages(5).Picture
        btnLast.Picture = .i16x16.ListImages(6).Picture
        
        btnFirst.DisabledPicture = .i16x16g.ListImages(3).Picture
        btnPrev.DisabledPicture = .i16x16g.ListImages(4).Picture
        btnNext.DisabledPicture = .i16x16g.ListImages(5).Picture
        btnLast.DisabledPicture = .i16x16g.ListImages(6).Picture
    End With
    
    With cboFilter
        .AddItem "Code"
        .AddItem "Name"
    End With
    bind_dc "SELECT * FROM tbl_departement WHERE parent_id=0 ORDER BY kd_departement ASC", "nm_departement", dcDepartement, "id_departement"
    cboFilter.Text = "Name"
    CONTROL False
    cmdNew.SetFocus
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rscashier, RecordPage.PageStart, RecordPage.PageEnd, 16, 2, False, True, , , , "kd_pasien")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    'Display the page information
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    'Display the selected record
    'lvList_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        shpBar.Width = ScaleWidth
        lvList.Width = ScaleWidth - (lvList.Left + 100)
        lvListObat.Width = ScaleWidth - (lvListObat.Left + 100)
        lstOrders.Width = Me.ScaleWidth
        lstOrders.Height = (Me.ScaleHeight - Picture1.Height) - lvList.Top
        fraSearch.Width = ScaleWidth - (fraSearch.Left + 100)
        txtSrchStrPasien.Width = fraSearch.Width - (txtSrchStrPasien.Left)
        txtSrchStr.Width = fraSearch.Width - (txtSrchStr.Left)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainMenu.RemToWin Me.Caption
    MDIMainMenu.HideTBButton "", True
    Set frmCashier = Nothing
End Sub

Private Sub lstOrders_AfterLabelEdit(Cancel As Integer, NewString As String)
    Call counttotal
End Sub

Private Sub lvList_Click()
    If (lvList.ListItems.Count > 0) Then
        With lvList.SelectedItem
            lblKdPasien.Caption = .Text
            lblNmPasien.Caption = .SubItems(1)
            lblTlpPasien.Caption = .SubItems(2)
            lblAlmtPasien.Caption = .SubItems(3)
            lblRelasi.Caption = .SubItems(4)
        End With
    End If
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    Call lvList_Click
End Sub

Private Sub lvList_DblClick()
    On Error Resume Next
    Call lvList_Click
End Sub

Private Sub callBrand()
    With frmCashierAE
        .lblBarCode.Caption = lvListObat.SelectedItem.Text
        .lblCode.Caption = lvListObat.SelectedItem.SubItems(1)
        .lblname.Caption = lvListObat.SelectedItem.SubItems(2)
        .lblKemasan.Caption = lvListObat.SelectedItem.SubItems(3)
        .lblPrice.Caption = lvListObat.SelectedItem.SubItems(4)
        .txtHarga.Text = Format(lvListObat.SelectedItem.SubItems(4), "")
        
        If (lvListObat.SelectedItem.SubItems(1) = "999999") Then
            .txtHarga.Enabled = True
            .txtHarga.Text = ""
            .txtQty.Text = 1
            .txtQty.Enabled = False
        Else
            On Error Resume Next
            .txtHarga.Text = "0"
            .txtHarga.Enabled = False
            .txtQty.Text = 10
            .txtQty.Enabled = True
            .txtQty.SetFocus
        End If
        
        If (lvListObat.SelectedItem.SubItems(5) = 0) Then
            .lblStock.Caption = 0
        Else
            .lblStock.Caption = lvListObat.SelectedItem.SubItems(5)
        End If
    End With
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Call lvList_Click
    End If
End Sub

Private Sub lvList_LostFocus()
    On Error Resume Next
    txtSrchStrPasien.SetFocus
End Sub

Private Sub lvListObat_Click()
    If (lvListObat.ListItems.Count > 0) Then
        With lvListObat.SelectedItem
            lblBrand.Caption = .SubItems(2)
            lblPrice.Caption = .SubItems(4)
            lblStock.Caption = .SubItems(5)
        End With
    End If
End Sub

Private Sub lvListObat_DblClick()
    On Error Resume Next
        Call callBrand
        frmCashierAE.txtQty.SetFocus
        frmCashierAE.show vbModal
End Sub

Private Sub lvListObat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lvListObat_DblClick
    End If
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub


Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        Call counttotal
        Call txtPayment_Change
    End If
End Sub

Private Sub txtKreditor_Change()
    If txtKreditor.Text <> 0 Then
        cbPay.Enabled = True
        cbpayment.Enabled = False
        txtPayment.Text = "Credit"
        txtMoneyBack.Text = "Credit"
        txtMoneyBack.Enabled = False
        txtPayment.Enabled = False
        Label1.Visible = True
        Label5.Visible = True
        Combo1.Visible = True
        Dim rsjw As New Recordset
        Dim i As Byte
        If rsjw.State = 1 Then rsjw.Close
        sql = "SELECT jw_waktu FROM tbl_cabang WHERE id_cabang = " & CurrBiz.BUSINNES_GROUP
        rsjw.Open sql, CN, adOpenStatic, adLockReadOnly
        If (rsjw.RecordCount > 0) Then
            For i = 1 To rsjw.Fields("jw_waktu")
                Combo1.AddItem i
            Next i
        End If
        rsjw.Close
    Else
        cbPay.Enabled = False
        cbpayment.Enabled = False
        Label1.Visible = False
        Label5.Visible = False
        Combo1.Visible = False
        txtPayment.Text = ""
        txtMoneyBack.Text = "0"
        txtMoneyBack.Enabled = False
        txtPayment.Enabled = True
        
    End If
End Sub

Private Sub txtKreditor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCashierKreditor.show vbModal
    End If
End Sub

Private Sub txtPayment_Change()
    On Error Resume Next
    txtPayment.Text = Format(txtPayment.Text, "#,##0")
    Dim payment, total, cback As Double
    payment = Format(txtPayment.Text, "")
    total = Format(lbltotal.Caption, "")
    cback = payment - total
    txtMoneyBack.Text = Format(cback, "#,###0")
    txtPayment.SelStart = Len(txtPayment.Text)
End Sub


Private Sub txtPayment_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        Call cmdProcess_Click
    End If
End Sub

Private Sub txtPayment_Validate(Cancel As Boolean)
    toMoney (toNumber(txtPayment.Text))
End Sub

Private Sub txtSrchStr_Change()
    Dim str As String
    If cboFilter.Text = "Code" Then
        str = "o.kd_obat"
    Else
        str = "o.nm_obat"
    End If
    If txtSrchStr.Text <> "" Then
        sql = "o.id_obat,o.kd_obat ,o.nm_obat,o.kemasan,FORMAT(o.harga_jual,0),"
        sql = sql & "((IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))-"
        sql = sql & "(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0))+ (o.stok) - "
        sql = sql & "(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))"
        sql = sql & " ) AS sisa,o.stok_min"

        With SQLParser
            .Fields = sql
            .Tables = " tbl_obat o INNER JOIN tbl_kategori k ON k.id_kategori =o.id_kategori INNER JOIN tbl_pengguna p ON p.id=o.id_pengguna "
            .wCondition = str & " Like '%" & txtSrchStr.Text & "%'  "
            .SortOrder = " o.id_obat ASC LIMIT 15"
            .SaveStatement
        End With
        
        If rscashier.State = 1 Then rscashier.Close
        rscashier.CursorLocation = adUseClient
        rscashier.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
        
        With RecordPage
            .Start rscashier, 15
            FillList2 1
        End With
        
        rscashier.Close
        Set rscashier = Nothing
        
    End If
End Sub

Private Sub FillList2(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvListObat, rscashier, RecordPage.PageStart, RecordPage.PageEnd, 16, 2, False, True, , , , "id_obat")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    lblPageInfo.Caption = "Record " & RecordPage.PageInfo
    lvListObat_Click
End Sub

Private Sub txtSrchStr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        lvList.SetFocus
     End If
End Sub

Private Sub txtSrchStrPasien_Change()
    Dim str As String
    On Error Resume Next
    If cboFilter.Text = "Code" Then
        str = "p.kd_pasien"
    Else
        str = "p.nm_pasien"
    End If
    If txtSrchStrPasien.Text <> "" Then
        sql = "p.kd_pasien,p.nm_pasien,p.no_tlp,p.alamat,(YEAR(curdate())-YEAR(p.tgl_lahir)) as umur"
        With SQLParser
            .Fields = sql
            .Tables = " tbl_pasien p INNER JOIN tbl_pengguna pp ON pp.id=p.id_pengguna "
            .wCondition = str & " Like '%" & txtSrchStrPasien.Text & "%'"
            .SortOrder = " p.kd_pasien ASC LIMIT 100"
            .SaveStatement
        End With
        
        If rscashier.State = 1 Then rscashier.Close
        rscashier.CursorLocation = adUseClient
        rscashier.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
        
        With RecordPage
            .Start rscashier, 100
            FillList 1
        End With
        
        rscashier.Close
        Set rscashier = Nothing
        
    End If
End Sub
