VERSION 5.00
Begin VB.Form frmPurchaseReturAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retur ADD"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1320
      TabIndex        =   13
      Text            =   "1"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtDisc 
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
      Left            =   4440
      TabIndex        =   12
      Text            =   "0"
      Top             =   3000
      Width           =   855
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Product Stock"
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
         TabIndex        =   11
         Top             =   2280
         Width           =   1185
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
         TabIndex        =   10
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Product Price"
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
         TabIndex        =   9
         Top             =   1800
         Width           =   1125
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
         TabIndex        =   8
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Product Name"
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
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Product Code"
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
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label BarCode 
         AutoSize        =   -1  'True
         Caption         =   "Bar Code :"
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
         TabIndex        =   5
         Top             =   360
         Width           =   840
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
         TabIndex        =   4
         Top             =   360
         Width           =   225
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
         TabIndex        =   3
         Top             =   1320
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
         TabIndex        =   2
         Top             =   3960
         Width           =   765
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
         TabIndex        =   1
         Top             =   840
         Width           =   225
      End
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
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "DISC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
End
Attribute VB_Name = "frmPurchaseReturAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InsertList()
    Dim i As Long
    Dim subtotal As Double
    'subtotal = frmPurchaseRetur.lblTotal.Caption
    With frmPurchaseRetur.lstOrders.ListItems.Add
        'If (frmPurchaseRetur.lstOrders.ListItems.Count - 1 < 1) Then
            .Text = lblBarCode.Caption
            .SubItems(1) = lblCode.Caption
            .SubItems(2) = lblName.Caption
            .SubItems(3) = lblPrice.Caption
            .SubItems(4) = txtQty.Text
            .SubItems(5) = txtDisc.Text
            '.SubItems(6) = Format(lblPrice.Caption - (lblPrice.Caption * Val(txtDisc.Text) / 100) * Val(txtQty.Text), "##,###0.00")
            .SubItems(6) = Format((lblPrice.Caption - ((lblPrice.Caption * txtDisc.Text) / 100)) * txtQty.Text, "##,###0.00")
    End With
    'kurangi di list
    frmPurchaseRetur.lvList.SelectedItem.SubItems(5) = Format(frmPurchaseRetur.lvList.SelectedItem.SubItems(5) - txtQty.Text, "##,###0.00")
    'subtotal = subtotal + (Val(lblPrice.Caption) * Val(txtQty.Text))
    'frmPurchaseRetur.lblTotal.Caption = Format(subtotal, "##,###0.00")
    'Call frmPurchaseRetur.counttotal
    Set frmPurchaseRetur.lstOrders = Nothing
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
    Set frmPurchaseReturAE = Nothing
End Sub

Private Sub txtDisc_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim a As Integer
    Dim B As Integer
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        a = FormatNumber(txtQty.Text, 0)
        B = FormatNumber(lblStock.Caption, 0)
     If (a <= B) Then
        Call InsertList
        Call frmPurchaseRetur.counttotal
        Unload Me
      Else
         MsgBox "Sorry Not Enough Stock In Your Product Stock", vbOKOnly + vbCritical, "Out Of Stock"
      End If
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim a As Integer
    Dim B As Integer
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        a = FormatNumber(txtQty.Text, 0)
        B = FormatNumber(lblStock.Caption, 0)
     If (a <= B) Then
        Call InsertList
        Call frmPurchaseRetur.counttotal
        Unload Me
      Else
         MsgBox "Sorry Not Enough Stock In Your Product Stock", vbOKOnly + vbCritical, "Out Of Stock"
      End If
    End If
End Sub


