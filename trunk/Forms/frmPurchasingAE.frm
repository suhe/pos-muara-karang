VERSION 5.00
Begin VB.Form frmPurchasingAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add  To Buy"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmPurchasingAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Tag             =   "Zero"
      ToolTipText     =   "Enter For Calculate Purchasing"
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Frame Frame4 
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
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5895
      Begin VB.Label lblKemasan 
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
         Left            =   1680
         TabIndex        =   15
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kemasan"
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
         TabIndex        =   14
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblCode 
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
         Left            =   1680
         TabIndex        =   12
         Top             =   840
         Width           =   225
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
         TabIndex        =   11
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblName 
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
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblBarCode 
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
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   225
      End
      Begin VB.Label BarCode 
         AutoSize        =   -1  'True
         Caption         =   "ID"
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
         TabIndex        =   8
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   480
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
         Left            =   1680
         TabIndex        =   5
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Harga Obat"
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
         TabIndex        =   4
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label lblStock 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   2760
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stok Sisa"
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
         TabIndex        =   2
         Top             =   2760
         Width           =   780
      End
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Tag             =   "Zero"
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmPurchasingAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InsertList()
    Dim i As Long
    Dim subtotal As Double
    With frmPurchasing.lstOrders.ListItems.Add
            .Text = lblBarCode.Caption
            .SubItems(1) = lblCode.Caption
            .SubItems(2) = lblname.Caption
            .SubItems(3) = lblKemasan.Caption
            .SubItems(4) = txtPrice.Text
            .SubItems(5) = txtQty.Text
            .SubItems(6) = Format(Format(txtPrice.Text, "") * Val(txtQty.Text), "##,###0.00")
    End With
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

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmPurchasing.txtSrchStr.SetFocus
    Set frmPurchasingAE = Nothing
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        If txtPrice.Text = "" Then MsgBox "Empty Price", vbOKOnly + vbCritical: Exit Sub
        If txtPrice.Text < 0 Then MsgBox "Empty Price", vbOKOnly + vbCritical: Exit Sub
        'If is_zero(txtQty, True) = True Then Exit Sub
        Call InsertList
        Call frmPurchasing.counttotal
        txtPrice.Enabled = False
        Unload Me
        'txtPrice.Enabled = True
        'txtPrice.SetFocus
    End If
End Sub

Private Sub txtPrice_LostFocus()
    On Error Resume Next
    frmPurchasing.txtSrchStr.SetFocus
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        If txtQty.Text = "" Then MsgBox "Empty Qty", vbOKOnly + vbCritical: Exit Sub
        If is_zero(txtQty, True) = True Then Exit Sub
        'Call InsertList
        'Call frmPurchasing.counttotal
        'Unload Me
        txtPrice.Enabled = True
        txtPrice.SetFocus
    End If
End Sub

Private Sub txtQty_LostFocus()
    'On Error Resume Next
    'frmPurchasing.txtSrchStr.SetFocus
End Sub
