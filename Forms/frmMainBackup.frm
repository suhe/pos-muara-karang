VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote & Eksport Data"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   Icon            =   "frmMainBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   7695
      Begin VB.CommandButton cmdSave 
         Caption         =   "Connect"
         Height          =   315
         Left            =   5160
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMainBackup.frx":038A
         Left            =   120
         List            =   "frmMainBackup.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Day"
         Top             =   480
         Width           =   2370
      End
      Begin VB.ComboBox cmbHari 
         Height          =   315
         ItemData        =   "frmMainBackup.frx":038E
         Left            =   2640
         List            =   "frmMainBackup.frx":0390
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Day"
         Top             =   480
         Width           =   690
      End
      Begin VB.ComboBox cmbTahun 
         Height          =   315
         ItemData        =   "frmMainBackup.frx":0392
         Left            =   4080
         List            =   "frmMainBackup.frx":039F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Year"
         Top             =   480
         Width           =   1050
      End
      Begin VB.ComboBox cmbBulan 
         Height          =   315
         ItemData        =   "frmMainBackup.frx":03B3
         Left            =   3360
         List            =   "frmMainBackup.frx":03B5
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Month"
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "DSN"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         Height          =   240
         Index           =   8
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   630
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4035
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   7117
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
         Text            =   "ID Pasien"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "KD Pasien"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nm Pasien"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jk Pasien"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tgl Lahir"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Alamat"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Kota"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "No HP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "No Tlp"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Tgl Input"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Nm Pengguna"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmMainBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
