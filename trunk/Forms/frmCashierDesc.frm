VERSION 5.00
Begin VB.Form frmCashierDesc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Money Receive"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7140
   Icon            =   "frmCashierDesc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAMount 
      Caption         =   "Total Pembayaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.Label lblkembali 
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
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   4560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "KEMBALI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "......................................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   6480
      End
      Begin VB.Label lbltotal 
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
         Left            =   2160
         TabIndex        =   4
         Top             =   960
         Width           =   4560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DIBAYAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label lblbayar 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   4560
      End
   End
   Begin VB.Label lblKomisi 
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
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "KOMISI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lblPay 
      AutoSize        =   -1  'True
      Caption         =   "CASH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   1110
   End
End
Attribute VB_Name = "frmCashierDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    If (tbl.TABLE_PAY_TYPE <> "Credit") Then
        'lblPay.Caption = "Cash"
        lblPay.Visible = False
        lblkembali.Visible = True
        Call cetak_Faktur2
        Call cetak_Faktur4
        lblbayar.Caption = Format(tbl.TABLE_MONEY, "##,###0.00")
        lblkembali.Caption = Format(tbl.TABLE_CBACK, "##,###0.00")
        Label5.Visible = True
    Else
        'lblPay.Caption = "Credit"
        lblPay.Visible = False
        Call cetak_Faktur2
        Call cetak_Faktur3
        Call cetak_Faktur4
        Label5.Visible = False
        lblbayar.Caption = "Credit"
        lblkembali.Visible = False
    End If
    lbltotal.Caption = Format(tbl.TABLE_TOTAL, "##,###0.00")
End Sub

