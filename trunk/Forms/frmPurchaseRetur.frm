VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPurchaseRetur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retur Product"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   13995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel (F5)"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
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
      Left            =   2160
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print (F3)"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Save (F2)"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame fraCashBack 
      Caption         =   "Money Over"
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
      TabIndex        =   59
      Top             =   4680
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
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Text            =   "0"
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraPayment 
      Caption         =   "Grand Total Product"
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
      TabIndex        =   57
      Top             =   4680
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
         TabIndex        =   58
         Text            =   "0"
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraPickUp 
      Caption         =   "Order Pick Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      TabIndex        =   51
      Top             =   2280
      Width           =   2895
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdPickup 
         Caption         =   "..."
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
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
         TabIndex        =   56
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Pick ID"
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
         TabIndex        =   55
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         TabIndex        =   54
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblExpID 
         AutoSize        =   -1  'True
         Caption         =   "..."
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
         Left            =   1080
         TabIndex        =   53
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblEXPname 
         AutoSize        =   -1  'True
         Caption         =   "..."
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
         Left            =   1080
         TabIndex        =   52
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.Frame fraCustomer 
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4080
      TabIndex        =   36
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtRetur 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "0"
         ToolTipText     =   "Please Enter After Fill this Field"
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Disc"
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
         TabIndex        =   50
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblNama 
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
         Left            =   960
         TabIndex        =   49
         Top             =   480
         Width           =   1680
      End
      Begin VB.Label lblCode 
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
         Left            =   960
         TabIndex        =   48
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         TabIndex        =   47
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         Top             =   240
         Width           =   420
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
         TabIndex        =   45
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
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
         Top             =   960
         Width           =   300
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
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
         Left            =   960
         TabIndex        =   43
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label lblTotalP 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   960
         TabIndex        =   41
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Retur"
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
         TabIndex        =   40
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label lblDisc 
         AutoSize        =   -1  'True
         Caption         =   "Disc (%)"
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
         Left            =   960
         TabIndex        =   39
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Price"
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
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label lblPriceP 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   960
         TabIndex        =   37
         Top             =   1200
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   380
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   13995
      TabIndex        =   28
      Top             =   8310
      Width           =   13995
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   29
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   60
            Width           =   2535
         End
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   60
         Width           =   1365
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
      TabIndex        =   25
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtFak 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
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
         TabIndex        =   27
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
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   3855
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
         TabIndex        =   24
         Top             =   840
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
         TabIndex        =   23
         Top             =   3960
         Width           =   765
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
         TabIndex        =   22
         Top             =   1320
         Width           =   2385
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
         TabIndex        =   21
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Brand:"
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
         TabIndex        =   20
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Stock"
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
         TabIndex        =   19
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
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
         TabIndex        =   18
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblDiscount 
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
         TabIndex        =   17
         Top             =   1800
         Width           =   2385
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Discount:"
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
         TabIndex        =   16
         Top             =   1800
         Width           =   780
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
      TabIndex        =   14
      Top             =   120
      Width           =   6855
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
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmPurchaseRetur.frx":0000
         Left            =   240
         List            =   "frmPurchaseRetur.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image imgSearch 
         Height          =   480
         Left            =   6240
         Picture         =   "frmPurchaseRetur.frx":0004
         Top             =   120
         Width           =   480
      End
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
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   3480
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
         TabIndex        =   13
         Top             =   360
         Width           =   4560
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
      ScaleWidth      =   13995
      TabIndex        =   10
      Top             =   8280
      Width           =   13995
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13995
      TabIndex        =   0
      Top             =   8295
      Width           =   13995
   End
   Begin ComctlLib.ListView lstOrders 
      Height          =   2175
      Left            =   120
      TabIndex        =   61
      Top             =   6000
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   3836
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
         Size            =   8.25
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
         Text            =   "BarCode"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ProductCode"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ProductName"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Product Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Qty"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Disc ( % )"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sub Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4755
      Left            =   7080
      TabIndex        =   62
      Top             =   1080
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   8387
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
         Text            =   "Bar"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Disc(%)"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Stock"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmPurchaseRetur.frx":08CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseRetur.frx":15A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseRetur.frx":18F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseRetur.frx":4020
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseRetur.frx":59B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseRetur.frx":8BEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseRetur.frx":94C5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblstatus 
      Caption         =   "......"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   63
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   7080
      Top             =   840
      Width           =   6795
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmPurchaseRetur.frx":AE57
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgQty 
      Height          =   480
      Left            =   6240
      Picture         =   "frmPurchaseRetur.frx":BA9B
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
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
Attribute VB_Name = "frmPurchaseRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CURR_COL   As Integer
Dim rscashierRetur  As New Recordset
Dim RecordPage As New clsPaging
Dim SQLParser  As New clsSQLSelectParser
Dim rs         As New Recordset
Dim rsdetails  As New Recordset
Dim rsreturPurchaseCashier As New Recordset
Public PK      As Long


Private Sub CONTROL(Active As Boolean)
    fraFaktur.Enabled = Active
    fraCustomer.Enabled = Not Active
    fraAMount.Enabled = Active
    fraPayment.Enabled = Active
    fraCashBack.Enabled = Active
    fraPickUp.Enabled = Active
    fraProduct.Enabled = Active
    fraSearch.Enabled = Active
    lstOrders.Enabled = Active
    cmdNew.Enabled = Not Active
    'm_new_trans.Enabled = Not Active
    cmdPrint.Enabled = Not Active
    cmdCancel.Enabled = Active
    'm_cancel_transact.Enabled = Active
    cmdProcess.Enabled = Active
    'm_save_transaction.Enabled = Active
End Sub

Private Sub GeneratePK()
    PK = getIndex("HEAD_SALES")
    txtFak.Text = GenerateID(PK, "FAK-", "0000000000")
End Sub

Private Sub CancelGeneratePK()
    PK = getCancelIndex("HEAD_SALES")
    txtFak.Text = GenerateID(PK, "FAK-", "0000000000")
End Sub

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
                If isRecordExist("PRODUCTS", "PK", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
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
                frmSearch.cmbFields.AddItem "MDIMainMenu Product"
                .show vbModal
            End With
        Case "Delete"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("PRODUCTS", "PK", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    Dim ANS As Integer
                    ANS = MsgBox("Are you sure you want to delete the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation.", vbCritical + vbYesNo, "Confirm Record Delete")
                    Me.MousePointer = vbHourglass
                    If ANS = vbYes Then
                        DelRecwSQL "tbl_IC_Products", "PK", "", True, CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                        RefreshRecords
                        'MDIMainMenu.UpdateInfoMsg
                        MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
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
            'GenerateDSN
            'With MDIMainMenu.CR
             '   .Reset: MDIMainMenu.InitCrys
             '    .ReportFileName = App.Path & "\Reports\rptStocksInformation.rpt"
             '   .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"
            
             '   .WindowTitle = "Stocks Information List"
        
             '  .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
             '  .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
             '   .ParameterFields(2) = "prmTitle;STOCKS INFORMATION LIST;True"
                    
             '   .PageZoom 100
             '   .Action = 1
            'End With
            'RemoveDSN
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
    '-In this case I used SQL because it is faster than Filter function of VB
    '-when hundling millions of records.
    On Error GoTo err
    With rscashierRetur
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

'Private Sub cbpayment_Click()
'    If (cbpayment.Text = "Cash") Then
'        dcRekening.Enabled = False
'        fraPickUp.Enabled = False
'        dcRekening.Text = ""
'     Else
'        dcRekening.Enabled = True
'        fraPickUp.Enabled = True
'        'bind_dc "SELECT BankName FROM BANKS", "BankName", dcRekening, "PK", True
'        bind_dc "SELECT * FROM BANKS", "BankName", dcRekening, "PK", True
'    End If
'End Sub

'Private Sub cmdBrowse_Click()
 '   frmCashierCustomer.show vbModal
'End Sub

Private Sub cmdCancel_Click()
    Call txtSrchStr_Change
    CONTROL False
    Call Form_Activate
    lblstatus.Caption = "Cancel Transaction"
    Call clearText
    'Call CancelGeneratePK
    'txtFak.Text = ""
End Sub

Private Sub cmdNew_Click()
    If txtPayment.Text = "0" Then
        MsgBox "No Item Retur", vbOKOnly + vbCritical
    Else
        CONTROL True
        Call Form_Deactivate
        lblstatus.Caption = "New Transaction"
        Call clearText
        'Call GeneratePK
        txtSrchStr.SetFocus
   End If
End Sub

Private Sub clearText()
    lblTotal.Caption = 0
    'txtFak.Text = ""
    txtSrchStr.Text = ""
    'txtPayment.Text = "0"
    txtMoneyBack.Text = "0"
    lstOrders.ListItems.Clear
    lvList.ListItems.Clear
    'lblCodeCust.Caption = "CUS-00000"
    'lblNamaCust.Caption = "UMUM"
    'lblDiscPelanggan.Caption = "N"
End Sub

'Private Sub cmdPickup_Click()
    'frmCashierExpedition.show vbModal
'End Sub

Private Sub CmdPrint_Click()
    Dim intResponse As Integer
    If txtFak.Text = "" Then
        MsgBox "No Data Is Printed", vbOKOnly + vbCritical, "Warning"
    Else
        intResponse = MsgBox("Are you sure you want to Print!", vbYesNo + vbCritical, "Warning")
        If intResponse = vbYes Then
            Call printStruk
        End If
    End If
End Sub

Private Sub cmdProcess_Click()
    Dim i As Integer
    'Dim subtotal As Double
    Dim payment, ptotal As Double
    Dim intResponse As Integer
    
    'payment = FormatNumber(txtPayment.Text, True, True, True)
    'ptotal = FormatNumber(lblTotal.Caption, True, True, True)
    
 'If (payment >= ptotal) Then
        With rsreturPurchaseCashier
            'On Error Resume Next
            'subtotal = 0
            'For i = 1 To lstOrders.ListItems.Count
                '.AddNew
                '.Fields("PK") = PK
                'rsdetails.Fields("FakNo") = txtFak.Text
                'rsdetails.Fields("ProductCode") = lblCode.Caption
                '.Fields("Price_sell") = FormatCurrency(lstOrders.ListItems(i).SubItems(3), True, True, True)
                'rsdetails.Fields("qty_retur") = txtRetur.Text
                '.Fields("discount") = lstOrders.ListItems(i).SubItems(5)
                'rsdetails.Fields("total_retur") = Val(txtRetur.Text * lblPrice.Caption, 0)
                'rsdetails.Update
                'subtotal = subtotal + lstOrders.ListItems(i).SubItems(6)
            'Next i
            For i = 1 To lstOrders.ListItems.Count
                .AddNew
                '.Fields("PK") = PK
                .Fields("FakNo") = txtFak.Text
                .Fields("ProductCode") = lstOrders.ListItems(i).SubItems(1)
                .Fields("Status") = "Retur"
                .Fields("Price_buy") = FormatCurrency(lstOrders.ListItems(i).SubItems(3), True, True, True)
                .Fields("qty") = lstOrders.ListItems(i).SubItems(4)
                .Fields("qty_retur") = 0
                .Fields("discount") = lstOrders.ListItems(i).SubItems(5)
                .Fields("total") = FormatCurrency(lstOrders.ListItems(i).SubItems(6), True, True, True)
                .Fields("total_retur") = 0
                .Update
                'subtotal = subtotal + lstOrders.ListItems(i).SubItems(6)
                Dim qty As Byte
                qty = 0
                CN.Execute "INSERT INTO DETAILS_SALES(productCode,Qty) VALUES ('" & lstOrders.ListItems(i).SubItems(1) & "'," & qty & ") "
            Next i
            
            sql = "UPDATE DETAILS_PURCHASING SET qty_retur=" & txtRetur.Text & ",total_retur=" & Val(txtRetur.Text * lblPriceP.Caption) & " WHERE FakNo='" & txtFak.Text & "' AND productCode='" & lblCode.Caption & "'"
            'MsgBox sql
            CN.Execute sql
            
            'For i = 1 To lstOrders.ListItems.Count
            '     .AddNew
            '    ' .Fields("PK") = PK
            '     .Fields("FakNo") = txtFak.Text
            '     .Fields("ProductCode") = lstOrders.ListItems(i).SubItems(1)
            '     .Fields("Price_buy") = FormatCurrency(lstOrders.ListItems(i).SubItems(3), True, True, True)
            '     .Fields("qty") = txtRetur.Text
            '     .Fields("qty_retur") = 0
            '     .Fields("discount") = lstOrders.ListItems(i).SubItems(5)
            '     .Fields("status") = "Retur"
            '     .Fields("total") = FormatCurrency(lstOrders.ListItems(i).SubItems(6), True, True, True)
            '     .Fields("total_retur") = 0
            '     .Update
                'subtotal = subtotal + lstOrders.ListItems(i).SubItems(6)
            'Next i
            
        End With
        'rsdetails.Close
        'rsretur.Close
        
         'With rs
         '       .AddNew
         '       .Fields("PK") = PK
         '       .Fields("FakNo") = txtFak.Text
         '       .Fields("CustomerID") = lblCodeCust.Caption
         '       .Fields("DateAdded") = Now
         '       .Fields("AddedByFK") = CurrUser.USER_PK
         '       .Fields("DateModified") = Now
         '       .Fields("LastUserFK") = CurrUser.USER_PK
         '       .Fields("CashierID") = CurrUser.USER_PK
         '       .Fields("PaymentType") = cbpayment.Text
         '       .Fields("BankCode") = dcRekening.BoundText
         '       .Fields("ExpID") = lblExpID.Caption
         '       .Fields("ExpAmount") = FormatCurrency(txtAmount.Text, True, True, True)
         '       .Fields("Total") = FormatCurrency(lblTotal.Caption, True, True, True)
         '       .Fields("Payment") = FormatCurrency(txtPayment.Text, True, True, True)
         '       .Fields("CashBack") = FormatCurrency(txtPayment.Text) - FormatCurrency(lblTotal.Caption)
         '       .Update
         '       Dim total As Byte
         '       total = 0
         '       'Dim rssales As New Recordset
         '       CN.Execute "INSERT INTO HEAD_PURCHASING(DateAdded,Total) VALUES(#" & Now & "#," & total & " ) "
         '       CN.Execute "INSERT INTO HEAD_SALES(DateAdded,Total,PaymentType) VALUES(#" & Now & "#," & total & ",'Transfer') "
         '       CN.Execute "INSERT INTO HEAD_SALES(DateAdded,Total,PaymentType) VALUES(#" & Now & "#," & total & ",'Cash') "
        'End With
        'rs.Close
        
        'If (FormatNumber(lblTotal.Caption, True, True, True) >= 500000) Then
        '   If (lblCodeCust.Caption <> "CUS-00000") Then
        '      If (lblDiscPelanggan.Caption = "N") Then
        '        CN.Execute "UPDATE CUSTOMERS SET Discount='Y' WHERE CustomerID='" & lblCodeCust.Caption & "'"
        '        MsgBox "Congrulations Your Purchase Item is greater than Rp.500.000,- You Can Have Discount ", vbOKOnly + vbInformation
        '      End If
        '        'MsgBox "", vbOKOnly + vbInformation
        '   End If
        '
        '   If (lblStat.Caption = "None") Then
        '       CN.Execute "UPDATE CUSTOMERS SET Status='New' WHERE CustomerID='" & lblCodeCust.Caption & "'"
        '       MsgBox "Congrulations Your Level Is New Member !  ", vbOKOnly + vbInformation
        '   End If
           
       ' End If
        
        'Call Form_Load
        Call txtSrchStr_Change
        CONTROL False
        Call Form_Activate
        lblstatus.Caption = "Transaction Returned Is Saved !"
        MsgBox "Thanks You For Your Returned , Come Back Againt", vbOKOnly + vbInformation
        Unload Me
        frmSalesProductDetails.RefreshRecords
     'Else
     '   MsgBox "Sorry Money Not Enough", vbOKOnly + vbCritical, "Warning"
     'End If
End Sub

Private Sub ResetFields()
    'clearText Me
    'cmdNew.SetFocus
End Sub
'Private Sub btnRecOp_Click()
 '   frmCustomerRecOp.show vbModal
'End Sub

Private Sub Active()
    With MDIMainMenu
        .tbMenu.Buttons(9).Caption = "Close"
        .tbMenu.Buttons(9).Image = 7
    End With
End Sub

Private Sub Form_Activate()
    'HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fffffft"
    'Call Active
End Sub

Public Sub counttotal()
    Dim i As Integer
    Dim subtotal As Double
    On Error Resume Next
    subtotal = 0
    If (lstOrders.ListItems.Count > 0) Then
        For i = 0 To lstOrders.ListItems.Count
            subtotal = subtotal + lstOrders.ListItems(i).SubItems(6)
        Next i
    End If
    lblTotal.Caption = Format(subtotal + txtAmount.Text, "##,###0.00")
End Sub

Private Sub Form_Deactivate()
    'MDIMainMenu.HideTBButton "", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdNew_Click
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdProcess_Click
    ElseIf KeyCode = vbKeyF5 Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdNew_Click
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdProcess_Click
    ElseIf KeyCode = vbKeyF5 Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'Call listviewHeader
    MDIMainMenu.AddToWin Me.Caption, Name
    lstOrders.ListItems.Clear
    'Set the graphics for the controls
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
    
    With cboFilter
        .AddItem "Code"
        .AddItem "Name"
    End With
    'With cbpayment
    '    .AddItem "Cash"
    '    .AddItem "Transfer"
    'End With
    'If (cbpayment.Text = "Cash") Then
    '    dcRekening.Enabled = False
    '    fraPickUp.Enabled = False
    '    dcRekening.Text = ""
    ' Else
    '    dcRekening.Enabled = True
    '    fraPickUp.Enabled = True
    '    bind_dc "SELECT BankName FROM BANKS", "BankName", dcRekening, "PK"
        'bind_dc "SELECT * FROM EXPEDITIONS", "EXPID" & "-" & "NAME", dcpickup, "PK"
    'End If
    
    rs.Open "SELECT * FROM HEAD_PURCHASING WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
    Dim FAK As String
    FAK = txtFak.Text
    rsdetails.Open "SELECT * FROM DETAILS_PURCHASING WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
    rsreturPurchaseCashier.Open "SELECT * FROM DETAILS_PURCHASING WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
    cboFilter.Text = "Code"
    'lblCodeCust.Caption = "CUS-00000"
    'lblNamaCust.Caption = "UMUM"
    'lblDiscPelanggan.Caption = "N"
    'GeneratePK
    CONTROL False
End Sub

Private Sub FillList(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, rscashierRetur, RecordPage.PageStart, RecordPage.PageEnd, 16, 2, False, True, , , , "PK")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation
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
        'lvList.Width = Me.ScaleWidth
        'lvList.Height = (Me.ScaleHeight - Picture1.Height) - lvList.Top
        'list view resize
        lvList.Width = ScaleWidth - (lvList.Left + 100)
        lstOrders.Width = Me.ScaleWidth
        'text search produk and frame
        fraSearch.Width = ScaleWidth - (fraSearch.Left + 100)
        txtSrchStr.Width = fraSearch.Width - (txtSrchStr.Left + imgSearch.Width)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainMenu.RemToWin Me.Caption
    'MDIMainMenu.HideTBButton "", True
    Set frmPurchaseRetur = Nothing
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

Private Sub lblTotal_Change()
    If FormatCurrency(lblTotal.Caption, 2) >= FormatCurrency(txtPayment.Text, 2) Then
         fraCashBack.Caption = "Money Charge"
         txtMoneyBack.Text = Format(FormatNumber(lblTotal.Caption, 0) - FormatNumber(txtPayment.Text, 0), "##,###0.00")
    ElseIf FormatCurrency(lblTotal.Caption, 2) <= FormatCurrency(txtPayment.Text, 2) Then
    'Else
        fraCashBack.Caption = "Money Back"
        txtMoneyBack.Text = Format(FormatNumber(txtPayment.Text, 0) - FormatNumber(lblTotal.Caption, 0), "##,###0.00")
    End If
End Sub

Private Sub lstOrders_AfterLabelEdit(Cancel As Integer, NewString As String)
    'Call counttotal
End Sub

Private Sub lvList_Click()
     'Sort the listview
    If (lvList.ListItems.Count - 1 > 1) Then
        With lvList.SelectedItem
            lblBrand.Caption = .SubItems(1) & "(" & .SubItems(2) & ")"
            lblPrice.Caption = .SubItems(3)
            lblDiscount.Caption = .SubItems(4)
            lblStock.Caption = .SubItems(5)
        End With
    End If
End Sub

Private Sub lvList_DblClick()
    'CommandPass "Edit"
    On Error Resume Next
    If (lvList.SelectedItem.SubItems(5) = "" Or lvList.SelectedItem.SubItems(5) = 0) Then
        MsgBox "Sorry Not Enough Stock In Your Product Stock", vbOKOnly + vbCritical, "Out Of Stock"
    Else
        Call callBrand
        frmPurchaseReturAE.txtQty.SetFocus
        frmPurchaseReturAE.show vbModal
    End If
End Sub

Private Sub callBrand()
    With frmPurchaseReturAE
        .lblBarCode.Caption = lvList.SelectedItem.Text
        .lblCode.Caption = lvList.SelectedItem.SubItems(1)
        .lblName.Caption = lvList.SelectedItem.SubItems(2)
        .lblPrice.Caption = lvList.SelectedItem.SubItems(3)
        If (lvList.SelectedItem.SubItems(5) = "") Then
            .lblStock.Caption = 0
        Else
            .lblStock.Caption = lvList.SelectedItem.SubItems(5)
        End If
        'If (lblDiscPelanggan.Caption = "" Or lblDiscPelanggan.Caption = "N") Then
        '    .txtDisc.Text = 0
        'Else
        '    .txtDisc.Text = lvList.SelectedItem.SubItems(4)
        'End If
    End With
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Call callBrand
    frmPurchaseReturAE.txtQty.SetFocus
    frmPurchaseReturAE.show vbModal
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub m_cancel_transact_Click()
    Call cmdCancel_Click
End Sub

Private Sub m_new_trans_Click()
    Call cmdNew_Click
End Sub

Private Sub m_print_transaction_Click()
    Call CmdPrint_Click
End Sub

Private Sub m_save_transaction_Click()
    Call cmdProcess_Click
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then PopupMenu MDIMainMenu.mnuRecA
    'If Button = 2 Then PopupMenu MDIMainMenu.mnuRecA
End Sub

Private Sub InsertList()
   'Dim i As Long
   Dim itmX As ListItem
    With lstOrders.ListItems.Add
        .Text = "test"
        .SubItems(1) = "test"
    End With
   Set itmX = Nothing
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        Call counttotal
    End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        Call InsertList
    End If
End Sub

Private Sub txtPayment_Change()
    On Error Resume Next
    txtMoneyBack.Text = Format(txtPayment.Text - lblTotal.Caption, "##,###0.00")
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

Private Sub txtRetur_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        'Call counttotal
        If txtRetur.Text = "" Then MsgBox "Empty Retur", vbOKOnly + vbCritical: Exit Sub
        If txtRetur.Text = "0" Then MsgBox "Zero Retur", vbOKOnly + vbCritical: Exit Sub
        If Val(txtRetur.Text) <= Val(lblQty.Caption) Then
            MsgBox "Qty Retur Is Valid", vbOKOnly + vbQuestion
            txtPayment.Text = Format(FormatNumber(lblPriceP.Caption, 0) * FormatNumber(txtRetur.Text, 0), "##,###0.00")
        Else
            MsgBox "Qty Retur Is Not Valid", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub txtSrchStr_Change()
    Dim str As String
    On Error Resume Next
    
    If cboFilter.Text = "Code" Then
        str = "ProductCode"
    Else
        str = "ProductName"
    End If
    
    If txtSrchStr.Text <> "" Then
        With SQLParser
            .Fields = "PK,ProductCode,ProductName,Price_buy,Discount,Stock"
            .Tables = "QR_STOCKS_TOTAL"
            .wCondition = str & " Like '%" & txtSrchStr.Text & "%'"
            .SortOrder = " ProductName ASC,PK ASC"
            .SaveStatement
        End With
        
        'SQLParser =
        rscashierRetur.CursorLocation = adUseClient
        rscashierRetur.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
        
        With RecordPage
            .Start rscashierRetur, 10000
            FillList 1
        End With
        rscashierRetur.Close
    End If
End Sub


Private Sub txtSrchStr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        lvList.SetFocus
     End If
End Sub
