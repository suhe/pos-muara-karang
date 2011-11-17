VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIMainMenu 
   BackColor       =   &H8000000C&
   Caption         =   "POS APPLICATION"
   ClientHeight    =   8430
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   15240
   Icon            =   "MDIMainMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSeparator 
      Align           =   4  'Align Right
      Height          =   7350
      Left            =   12495
      ScaleHeight     =   7290
      ScaleWidth      =   135
      TabIndex        =   5
      Top             =   780
      Width           =   195
      Begin VB.CommandButton cmdHie 
         Height          =   1335
         Left            =   0
         TabIndex        =   6
         Top             =   3120
         Width           =   135
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7350
      Left            =   12690
      ScaleHeight     =   7350
      ScaleWidth      =   2550
      TabIndex        =   1
      Top             =   780
      Width           =   2550
      Begin VB.Frame Frame1 
         Height          =   465
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Width           =   2250
         Begin VB.Image Image1 
            Height          =   240
            Left            =   75
            Picture         =   "MDIMainMenu.frx":5CC28
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Opened Forms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   3
            Top             =   195
            Width           =   1290
         End
      End
      Begin MSComctlLib.ListView lvWin 
         Height          =   4050
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MDIMainMenu.frx":5D62A
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Form Name"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   1560
         Picture         =   "MDIMainMenu.frx":5E304
         Top             =   5040
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   1920
         Picture         =   "MDIMainMenu.frx":5F04E
         Top             =   5040
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Timer tmrMemStatus 
      Interval        =   1000
      Left            =   4755
      Top             =   6315
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5955
      Top             =   5040
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   780
      Width           =   15240
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   2520
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":5FD98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":607AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":611BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":61556
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":618F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":61C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":62024
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":62A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":63448
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":63E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":6486C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":6527E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":65C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":666A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":66C3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16g 
      Left            =   2280
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":671DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":67774
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":67D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":680A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":68442
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":687DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ig24x24 
      Left            =   4080
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":68B76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   960
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":68DA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":6A735
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":6C0C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":6DA59
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":6F3EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":70D7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":7270F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":740A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":75A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":773C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":780A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":78983
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":7965F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":7A33B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":7B017
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":7BCF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainMenu.frx":7C9CF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   8130
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "MDIMainMenu.frx":7D2AB
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12850
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MDIMainMenu.frx":7D647
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11/4/2011"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "8:17 PM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1376
      ButtonWidth     =   1402
      ButtonHeight    =   1376
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Shortcuts"
            Key             =   "Shortcuts"
            Object.ToolTipText     =   "F10"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "F1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "Edit"
            Object.ToolTipText     =   "F2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "F3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "F4"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "F5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "F6"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
            Object.ToolTipText     =   "F8"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.PictureBox picFreeMem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   10440
         ScaleHeight     =   825
         ScaleWidth      =   4815
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AVAILABLE FREE MEMORY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   2070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   2520
         Y1              =   175
         Y2              =   175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         X1              =   825
         X2              =   825
         Y1              =   225
         Y2              =   525
      End
      Begin VB.Label lblVMem 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "                    "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   960
         TabIndex        =   11
         Top             =   405
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Virtual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   75
         TabIndex        =   10
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Physical"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":7D9E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":7E633
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":7F285
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":7FED7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":80B29
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   5160
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":8177B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":823CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":8301F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":83C71
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":848C3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   8400
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMainMenu.frx":85515
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu m_dashboard 
      Caption         =   "&Dashboard"
      Begin VB.Menu m_logout 
         Caption         =   "&Log Out"
         Shortcut        =   ^O
      End
      Begin VB.Menu mShorcut 
         Caption         =   "&Shortcut"
         Shortcut        =   {F1}
      End
      Begin VB.Menu m_Exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu m_Master 
      Caption         =   "&Master"
      Begin VB.Menu m_medicine 
         Caption         =   "&Medicine"
         Shortcut        =   ^M
      End
      Begin VB.Menu m_categories 
         Caption         =   "&Categories"
         Shortcut        =   ^C
      End
      Begin VB.Menu m_supplier 
         Caption         =   "&Supplier"
         Shortcut        =   ^S
      End
      Begin VB.Menu m_pasien 
         Caption         =   "&Pasien"
         Shortcut        =   ^P
      End
      Begin VB.Menu m_kreditor 
         Caption         =   "&Kreditor"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu m_transaction 
      Caption         =   "&Transaction"
      Begin VB.Menu m_cashier 
         Caption         =   "&Cashier"
         Shortcut        =   ^A
      End
      Begin VB.Menu m_purchasing 
         Caption         =   "&Purchasing"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu m_trans_view 
      Caption         =   "&Transaction Views"
      Begin VB.Menu m_Sales 
         Caption         =   "&Sales"
         Shortcut        =   ^J
      End
      Begin VB.Menu m_purchase 
         Caption         =   "&Purchase"
         Shortcut        =   ^Z
      End
      Begin VB.Menu m_CashFLow 
         Caption         =   "&Cash Flow"
      End
   End
   Begin VB.Menu m_setting 
      Caption         =   "&Setting"
      Begin VB.Menu m_info 
         Caption         =   "&Businnes Info"
         Shortcut        =   ^I
      End
      Begin VB.Menu m_user 
         Caption         =   "&User & Staff"
         Shortcut        =   ^U
      End
      Begin VB.Menu mGroup 
         Caption         =   "&Group Klinik"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "MDIMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MF_BYPOSITION = &H400
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Dim cursor_pos  As PointAPI
Public CloseMe  As Boolean
Dim resize_down     As Boolean
Dim show_mnu        As Boolean
Dim pos_num         As Integer
Dim Theme As Integer

Public Sub AddToWin(ByVal srcDName As String, ByVal srcFormName As String)
    On Error Resume Next
    Dim xItem As ListItem
    Set xItem = lvWin.ListItems.Add(, srcFormName, srcDName, 1, 1)
    xItem.ToolTipText = srcDName
    xItem.SubItems(1) = "***" & srcDName & "***"
    xItem.Selected = True
    Set xItem = Nothing
End Sub

Public Sub RemToWin(ByVal srcDName As String)
    On Error Resume Next
    search_in_listview lvWin, "***" & srcDName & "***"
    lvWin.ListItems.Remove (lvWin.SelectedItem.Index)
End Sub

Private Sub cmdHie_Click()
    show_mnu = Not show_mnu
    show_menu (show_mnu)
End Sub

Private Sub show_menu(ByVal show As Boolean)
    Dim img As Image
    If show = True Then
        Set img = Image2
    Else
        Set img = Image5
    End If
    'Set the style button graphics
    'With StyleButton2
     '   Set .PictureDown = img.Picture
      '  Set .PictureFocus = img.Picture
       ' Set .PictureHover = img.Picture
       ' Set .PictureUp = img.Picture
    'End With
    'Set picture visibility
    picLeft.Visible = show
    
    If show = True Then cmdHie.ToolTipText = "Hide": picSeparator.MousePointer = vbSizeWE Else picSeparator.MousePointer = vbArrow: cmdHie.ToolTipText = "Expand"
    
    Set img = Nothing
End Sub

Private Sub lvWin_Click()
    If lvWin.ListItems.Count < 1 Then Exit Sub
    Select Case lvWin.SelectedItem.Key
        Case "frmShortcuts": frmShortcuts.show: frmShortcuts.WindowState = vbMaximized: frmShortcuts.SetFocus
        Case "frmProduct": LoadForm frmProduct
        Case "frmProductList": LoadForm frmProductList
        Case "frmCashFlow": LoadForm frmCashFlow
        Case "frmCashier": LoadForm frmCashier
        Case "frmCategories": LoadForm frmCategories
        Case "frmPasien": LoadForm frmPasien
        Case "frmKomisi": LoadForm frmKomisi
        Case "frmPurchase": LoadForm frmPurchase
        Case "frmPurchasing": LoadForm frmPurchasing
        Case "frmSales": LoadForm frmSales
        Case "frmDebt": LoadForm frmDebt
        Case "frmReturPurchase": LoadForm frmReturPurchase
        Case "frmGroup": frmGroup.show vbModal
        Case "frmDepartement": LoadForm frmDepartement
        Case "frmSupplier": LoadForm frmSupplier
        Case "frmUserRec": frmUserRec.show vbModal
        Case "frmBusinessInfo": frmBusinessInfo.show vbModal
    End Select
End Sub

Private Sub m_CashFLow_Click()
    frmCashFlow.show
End Sub

Private Sub m_cashier_Click()
    frmCashier.show
End Sub

Private Sub m_categories_Click()
    LoadForm frmCategories
End Sub

Private Sub m_dashboard_Click()
    frmShortcuts.show
End Sub

Private Sub m_Exit_Click()
    End
End Sub

Private Sub m_info_Click()
    frmBusinessInfo.show vbModal
End Sub

Private Sub m_kreditor_Click()
    LoadForm frmKreditor
End Sub

Private Sub m_logout_Click()
    If MsgBox("Are you sure you want to log out?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    'SendMessage frmShortcuts.hwnd, WM_CLOSE, 0, 0
    UnloadChilds
    SendMessage frmShortcuts.hWnd, WM_ACTIVATE, 0, 0
    ClearInfoMsg
    StatusBar1.Panels(3).Text = ""
    StatusBar1.Panels(4).Text = ""
    CurrUser.USER_NAME = ""
    CurrUser.USER_PK = 0
    Unload frmShortcuts
    frmLogin.show vbModal: If CloseMe = True Then Unload Me: Exit Sub: Exit Sub
    DisplayUserInfo
    UpdateInfoMsg
End Sub

Private Sub m_medicine_Click()
    LoadForm frmProduct
End Sub

Private Sub m_pasien_Click()
    LoadForm frmPasien
End Sub

Private Sub m_purchase_Click()
    LoadForm frmPurchase
End Sub

Private Sub m_purchasing_Click()
    LoadForm frmPurchase
End Sub

Private Sub m_Sales_Click()
    LoadForm frmSales
End Sub

Private Sub m_supplier_Click()
    LoadForm frmSupplier
End Sub

Private Sub m_user_Click()
    frmUserRec.show vbModal
End Sub

Private Sub MDIForm_Load()
    Dim hndMenu As Long, ItemCount As Long, i As Integer, lngID As Long
    hndMenu = GetSystemMenu(Me.hWnd, 0)
    If hndMenu Then
        ItemCount = GetMenuItemCount(hndMenu)
        lngID = GetMenuItemID(hndMenu, ItemCount - 1)
        'MsgBox lngID
        'Disable the X button
        For i = 1 To 2
        Call RemoveMenu(hndMenu, ItemCount - i, MF_BYPOSITION)
    Next
    End If

    DBPath = "DSN=pos_db"
    'ClearInfoMsg
    HideTBButton "", True
    Me.show
    frmSplash.show vbModal
    If OpenDB = False Then CloseMe = True: Unload Me: Exit Sub
    frmLogin.show vbModal: If CloseMe = True Then Unload Me: Exit Sub: Exit Sub
    frmShortcuts.show
    'frmDateChecker.show vbModal
    If CurrUser.USER_ISCASHIER = True Then
        picLeft.Visible = False
        picSeparator.Visible = False
        show_mnu = Not show_mnu
        frmCashier.show
    End If
    Set lvWin.SmallIcons = i16x16
    Set lvWin.Icons = i16x16
    DisplayUserInfo
    lvWin.ListItems.Add(, "frmShortcuts", "@Shortcuts", 1, 1).Bold = True
    'UpdateInfoMsg 'Display the business status
    show_mnu = True
    show_menu (show_mnu)
    If CurrUser.USER_ISADMIN = False Then
        m_Master.Enabled = False
        m_transaction.Enabled = False
        m_trans_view.Enabled = False
        m_setting.Enabled = False
    End If
End Sub

Private Sub DisplayUserInfo()
    If CurrUser.USER_ISADMIN = True Then
        StatusBar1.Panels(4).Text = "Manager OR Administrator"
    Else
        StatusBar1.Panels(4).Text = "User"
    End If
    StatusBar1.Panels(3).Text = CurrUser.USER_NAME
    Dim rs As New Recordset
    rs.Open "SELECT * FROM tbl_business_info", CN, adOpenStatic, adLockReadOnly
    CurrBiz.BUSINNES_NAME = rs.Fields("bussines_name")
    CurrBiz.BUSINESS_ADDRESS = rs.Fields("bussines_address")
    CurrBiz.BUSINESS_CONTACT_INFO = rs.Fields("bussines_cp")
    CurrBiz.BUSINNES_CITY = rs.Fields("bussines_city")
    CurrBiz.BUSINNES_NOTE = rs.Fields("bussines_note")
    Set rs = Nothing
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    picFreeMem.Left = (Me.Width - picFreeMem.ScaleWidth) - 200
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set MDIMainMenu = Nothing
End Sub

Private Sub mShorcut_Click()
    LoadForm frmShortcuts
End Sub

Private Sub picSeparator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = True
        resize_down = True
    End If
End Sub

Private Sub picSeparator_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = False
        resize_down = False
    End If
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Shortcuts" Then
        frmShortcuts.show
        frmShortcuts.WindowState = vbMaximized
        frmShortcuts.SetFocus
    Else
        On Error Resume Next
        ActiveForm.CommandPass Button.Key
    End If
End Sub

Public Sub HideTBButton(ByVal srcPatern As String, Optional srcAllButton As Boolean)
    If srcAllButton = True Then srcPatern = "ttttttt"
    If Mid$(srcPatern, 1, 1) = "t" Then tbMenu.Buttons(3).Visible = False
    If Mid$(srcPatern, 2, 1) = "t" Then tbMenu.Buttons(4).Visible = False
    If Mid$(srcPatern, 3, 1) = "t" Then tbMenu.Buttons(5).Visible = False
    If Mid$(srcPatern, 4, 1) = "t" Then tbMenu.Buttons(6).Visible = False
    If Mid$(srcPatern, 5, 1) = "t" Then tbMenu.Buttons(7).Visible = False
    If Mid$(srcPatern, 6, 1) = "t" Then tbMenu.Buttons(8).Visible = False
    If Mid$(srcPatern, 7, 1) = "t" Then tbMenu.Buttons(9).Visible = False
    'If mnuRAC.Visible = False Then mnuRASep2.Visible = False
End Sub

Public Sub ShowTBButton(ByVal srcPatern As String, Optional srcAllButton As Boolean)
    'Highligh active form in opened form list
    If srcAllButton = True Then srcPatern = "ttttttt"
    If Mid$(srcPatern, 1, 1) = "t" Then tbMenu.Buttons(3).Visible = True
    If Mid$(srcPatern, 2, 1) = "t" Then tbMenu.Buttons(4).Visible = True
    If Mid$(srcPatern, 3, 1) = "t" Then tbMenu.Buttons(5).Visible = True
    If Mid$(srcPatern, 4, 1) = "t" Then tbMenu.Buttons(6).Visible = True
    If Mid$(srcPatern, 5, 1) = "t" Then tbMenu.Buttons(7).Visible = True
    If Mid$(srcPatern, 6, 1) = "t" Then tbMenu.Buttons(8).Visible = True
    If Mid$(srcPatern, 7, 1) = "t" Then tbMenu.Buttons(9).Visible = True
    'If mnuRAC.Visible = True Then mnuRASep2.Visible = True
End Sub

Public Sub UpdateInfoMsg()
    Dim strHTML As String
    Screen.MousePointer = vbHourglass
    ' Header html
    strHTML = "<html><body topmargin=9 leftmargin=0 bgcolor=#" & Hex$(80) & Hex$(80) & Hex$(80) & "><b>"
    
    ' Body html
    strHTML = strHTML & "<marquee direction=left scrolldelay=75>"
    
    '- For new caompany
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            ":::" & CurrBiz.BUSINNES_NAME & " ::: " & _
                        "</font>"
    
    '- For Pasien
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Jumlah Pasien = " & getRecordCount("id_pasien", "tbl_pasien", "WHERE tgl_lahir <>'' ") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
    '- For Kreditor
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Jumlah Kreditor = " & getRecordCount("id_kreditor", "tbl_kreditor", "WHERE id_kreditor <>'' ") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For Depp
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Jumlah Departement = " & getRecordCount("id_departement", "tbl_departement", "WHERE id_departement <>'' ") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For Cabang
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Jumlah Cabang = " & getRecordCount("id_cabang", "tbl_cabang", "WHERE id_cabang <>'' ") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For all Supplier
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(128) & Hex$(191) & Hex$(28) & ">" & _
                            "Jumlah Supplier = " & getRecordCount("id_supplier", "tbl_supplier") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                                    
    '- For no. Medicine
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Obat = " & getRecordCount("id_obat", "tbl_obat") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"

    '- For no. Cat Medicine
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Jumlah Kategori Obat = " & getRecordCount("id_kategori", "tbl_kategori") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                 
                 
    '- Laba Pendapatan
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Laba Pendapatan Lunas Hari ini = Rp." & toMoney(getSUMTotal("SELECT jual AS total FROM vw_cash_flow WHERE DAY(tgl_cash)=DAY(curdate())", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- Laba Pelunasan
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pelunasan Piutang Hari ini = Rp." & toMoney(getSUMTotal("SELECT jual_sebelumnya AS total FROM vw_cash_flow WHERE DAY(tgl_cash)=DAY(curdate())", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- Laba Total Pendapatan
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Laba Pendapatan Hari ini = Rp." & toMoney(getSUMTotal("SELECT jual+jual_sebelumnya AS total FROM vw_cash_flow WHERE DAY(tgl_cash)=DAY(curdate())", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For Pendapatan Kemarin
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Laba Pendapatan Lunas Kemarin = Rp." & toMoney(getSUMTotal("SELECT jual AS total FROM vw_cash_flow WHERE DAY(tgl_cash)=DAY(curdate()-1)", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For Pelunasna Pendapatan Kemarin
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pelunasan Piutang Kemarin = Rp." & toMoney(getSUMTotal("SELECT jual_sebelumnya AS total FROM vw_cash_flow WHERE DAY(tgl_cash)=DAY(curdate()-1)", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For Total Pendapatan Kemarin
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Laba Pendapatan Kemarin = Rp." & toMoney(getSUMTotal("SELECT jual+jual_sebelumnya AS total FROM vw_cash_flow WHERE DAY(tgl_cash)=DAY(curdate()-1)", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For sales this month
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Pendapatan Bulan Ini = Rp." & toMoney(getSUMTotal("SELECT (SUM(jual)+SUM(jual_sebelumnya)) AS total FROM vw_cash_flow WHERE MONTH(tgl_cash)=MONTH(curdate())", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
    
    '- For sales this month ago
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Pendapatan Bulan Lalu = Rp." & toMoney(getSUMTotal("SELECT (SUM(jual)+ SUM(jual_sebelumnya)) AS total FROM vw_cash_flow WHERE MONTH(tgl_cash)=MONTH(curdate())-1", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    '- For sales this year
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Pendapatan Tahun Ini = Rp." & toMoney(getSUMTotal("SELECT (SUM(jual)+SUM(jual_sebelumnya)) AS total FROM vw_cash_flow WHERE YEAR(tgl_cash)=YEAR(curdate())", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                         
    '- For sales this year
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Pendapatan Tahun Lalu = Rp." & toMoney(getSUMTotal("SELECT (SUM(jual)+ SUM(jual_sebelumnya)) AS total FROM vw_cash_flow WHERE YEAR(tgl_cash)=YEAR(curdate())-1", "total", CN)) & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
    
    '- Jumlah Pengguna
    strHTML = strHTML & "<font face='tahoma' size=2 color=#" & Hex$(255) & Hex$(147) & Hex$(31) & ">" & _
                            "<img src='ar.bmp'></img>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Jumlah Pengguna Program = " & getRecordCount("id", "tbl_pengguna") & "&nbsp;&nbsp;&nbsp;" & _
                        "</font>"
                        
    strHTML = strHTML & "</marquee>"
    
    ' Footer html
    strHTML = strHTML & "</b></body></html>"
    Open Environ$("TMP") & "\Klinik.tmp" For Output As #1
        Print #1, strHTML
    Close #1
    strHTML = vbNullString
    Call SavePicture(ig24x24.ListImages(1).Picture, Environ$("TMP") & "\ar.bmp")
    'WebAdvisory.Navigate Environ$("TMP") & "\Klinik.tmp"
    Screen.MousePointer = vbDefault
End Sub

Public Sub ClearInfoMsg()
    Dim strHTML As String
    Screen.MousePointer = vbHourglass
    ' Header html
    strHTML = "<html><body topmargin=9 leftmargin=0 bgcolor=#" & Hex$(80) & Hex$(80) & Hex$(80) & "><b>"
    ' Footer html
    strHTML = strHTML & "</b></body></html>"
    Open Environ$("TMP") & "\Klinik.tmp" For Output As #1
        Print #1, strHTML
    Close #1
    strHTML = vbNullString
    Call SavePicture(ig24x24.ListImages(1).Picture, Environ$("TMP") & "\ar.bmp")
    'WebAdvisory.Navigate Environ$("TMP") & "\Klinik.tmp"
    Screen.MousePointer = vbDefault
End Sub

Private Sub picAdvisory_Resize()
    On Error Resume Next
    'f WindowState <> vbMinimized Then WebAdvisory.Width = (picAdvisory.Width - WebAdvisory.Left) + 270
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    Frame1.Width = picLeft.ScaleWidth
    lvWin.Width = picLeft.ScaleWidth
    lvWin.Height = picLeft.ScaleHeight - lvWin.Top - 20
End Sub

Public Sub UnloadChilds()
Dim Form As Form
   For Each Form In Forms
      If Form.Name <> Me.Name And Form.Name <> "frmShortcuts" Then Unload Form
   Next Form
Set Form = Nothing
End Sub
