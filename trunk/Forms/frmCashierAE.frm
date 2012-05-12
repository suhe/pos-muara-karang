VERSION 5.00
Begin VB.Form frmCashierAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Product To Sale"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "frmCashierAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHarga 
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
      Left            =   1320
      TabIndex        =   14
      Top             =   3600
      Width           =   4695
   End
   Begin VB.TextBox txtQty 
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
      Left            =   1320
      TabIndex        =   1
      Text            =   "10"
      Top             =   4200
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
      TabIndex        =   0
      Top             =   120
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
         TabIndex        =   17
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label3 
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
         TabIndex        =   16
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   12
         Top             =   2760
         Width           =   390
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
         TabIndex        =   11
         Top             =   2760
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Harga Jual"
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
         TabIndex        =   10
         Top             =   2280
         Width           =   900
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
         TabIndex        =   9
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label Label6 
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
         TabIndex        =   8
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Kode Obat"
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
         Width           =   870
      End
      Begin VB.Label BarCode 
         AutoSize        =   -1  'True
         Caption         =   "ID :"
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
         Top             =   360
         Width           =   285
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   840
         Width           =   225
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Harga"
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
      TabIndex        =   15
      Top             =   3600
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
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "frmCashierAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InsertList()
    Dim i As Long
    Dim subtotal As Double
    With frmCashier.lstOrders.ListItems.Add
            .Text = lblBarCode.Caption
            .SubItems(1) = lblCode.Caption
            .SubItems(2) = lblname.Caption
            .SubItems(3) = lblKemasan.Caption
            .SubItems(4) = Format(txtHarga.Text, "")
            .SubItems(5) = txtQty.Text
            .SubItems(6) = ""
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

Private Sub Form_Load()
    On Error Resume Next
    txtQty.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCashierAE = Nothing
End Sub

Private Sub txtHarga_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        If txtQty.Enabled = False Then
            On Error Resume Next
            If is_empty(txtQty, True) = True Then Exit Sub
            If is_empty(txtHarga, True) = True Then Exit Sub
            Call InsertList
            Call frmCashier.counttotal
            'On Error Resume Next
            frmCashier.lvListObat.SelectedItem.SubItems(5) = Val(frmCashier.lvListObat.SelectedItem.SubItems(5)) - Val(txtQty.Text)
            Unload Me
            frmCashier.cboFilter.SetFocus
        Else
            On Error Resume Next
            If (txtHarga.Text <> "") Then
                txtQty.Enabled = True
                txtQty.SetFocus
            Else
                txtQty.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If is_empty(txtQty, True) = True Then Exit Sub
    If is_empty(txtHarga, True) = True Then Exit Sub
    Dim a As Integer
    Dim b As Integer
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        If txtQty.Text = "" Then MsgBox "Empty Qty", vbOKOnly + vbCritical: Exit Sub
        a = FormatNumber(txtQty.Text, 0)
        b = FormatNumber(lblStock.Caption, 0)

        With frmCashier.lstOrders
            Call InsertList
            Dim i As Long
            For i = .ListItems.Count To 2 Step -1
                If lblCode.Caption = .ListItems(i - 1).SubItems(1) Then
                     .ListItems.Remove (i - 1)
                     MsgBox "Duplicate Data, System replace Data with last entry !", vbCritical + vbInformation
                End If
            Next i
        End With
        Call frmCashier.counttotal
        frmCashier.lvListObat.SelectedItem.SubItems(5) = Val(frmCashier.lvListObat.SelectedItem.SubItems(5)) - Val(txtQty.Text)
        Unload Me
        On Error Resume Next
        With frmCashier
            If .lvListObat.Visible = True Then
                .lvListObat.ListItems.Clear
                .txtSrchStr.Text = ""
                .txtSrchStr.SetFocus
            End If
        End With
    End If
End Sub
