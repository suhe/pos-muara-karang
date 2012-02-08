VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDebtsPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debt Payment"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
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
      Height          =   855
      Left            =   3360
      TabIndex        =   15
      Top             =   1200
      Width           =   3495
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "Cicilan 1"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
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
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Pay"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   14
      Top             =   3960
      Width           =   3015
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
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Over "
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
      Left            =   4560
      TabIndex        =   13
      Top             =   3960
      Width           =   2415
      Begin VB.TextBox txtOver 
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
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   11
      Top             =   2760
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
         TabIndex        =   12
         Top             =   240
         Width           =   4560
      End
   End
   Begin VB.Frame Frame7 
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
      TabIndex        =   8
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtFak 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   480
         Width           =   6615
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
         TabIndex        =   10
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
      Begin VB.ComboBox cbpayment 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Text            =   "Cash"
         Top             =   360
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dcRekening 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
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
         TabIndex        =   7
         Top             =   3960
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmDebtsPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs  As New Recordset
Dim i As Long


Private Sub cbpayment_Click()
    If (cbpayment.Text = "Cash") Then
        dcRekening.Enabled = True
        'fraPickUp.Enabled = True
        dcRekening.Text = ""
     Else
        dcRekening.Enabled = True
        'fraPickUp.Enabled = True
        bind_dc "SELECT * FROM BANKS", "BankName", dcRekening, "PK"
        'bind_dc "SELECT * FROM EXPEDITIONS", "EXPID" & "-" & "NAME", dcpickup, "PK"
    End If
End Sub

Private Sub cmdProcess_Click()
    If txtPayment.Text = "" Then MsgBox "Empty Payment", vbOKOnly + vbCritical: Exit Sub
    If txtPayment.Text = "0" Then MsgBox "Zero Payment", vbOKOnly + vbCritical: Exit Sub
    If txtOver.Text = "" Then MsgBox "Empty Cashback", vbOKOnly + vbCritical: Exit Sub
    If cbpayment.Text = "" Then MsgBox "Empty Payment", vbOKOnly + vbCritical: Exit Sub
    If (cbpayment.Text = "Transfer") Then
        If dcRekening.Text = "" Then MsgBox "Empty Type Of Payment", vbOKOnly + vbCritical: Exit Sub
    End If
    If (Combo1.Text = "") Then MsgBox "Type Debt Payment Is Empty", vbOKOnly + vbCritical: Exit Sub
    
    With rs
         .AddNew
        '.Fields("PK") = PK
        .Fields("FakNo") = txtFak.Text
        .Fields("Type") = Combo1.Text
        .Fields("DateAdded") = Now
        .Fields("AddedByFK") = CurrUser.USER_PK
        .Fields("DateModified") = Now
        .Fields("LastUserFK") = CurrUser.USER_PK
        .Fields("CashierID") = CurrUser.USER_PK
        .Fields("PaymentType") = cbpayment.Text
        If (dcRekening.Text <> "") Then
            .Fields("BankCode") = dcRekening.BoundText
        End If
        .Fields("Total") = FormatCurrency(lblTotal.Caption, True, True, True)
        .Fields("Payment") = FormatCurrency(txtPayment.Text, True, True, True)
        .Update
    End With
    rs.Close
    Unload Me
    Unload frmDebtDetails
    Unload frmDebt2
    frmDebt.RefreshRecords
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
    rs.Open "SELECT * FROM KREDITS ", CN, adOpenStatic, adLockOptimistic
    With cbpayment
        .AddItem "Cash"
        .AddItem "Transfer"
    End With
    
    If (cbpayment.Text = "Cash") Then
        dcRekening.Enabled = True
        dcRekening.Text = ""
     Else
        dcRekening.Enabled = True
        'fraPickUp.Enabled = True
        bind_dc "SELECT * FROM BANKS", "BankName", dcRekening, "PK"
    End If
    i = frmDebtDetails.lvList.ListItems.Count
    lblTotal.Caption = frmDebtDetails.lblTotal.Caption
    
    Dim K As Byte
    For K = 1 To 14
        Combo1.AddItem "Cicilan " & K
    Next K
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDebtsPayment = Nothing
    Set rs = Nothing
End Sub

Private Sub txtPayment_Change()
    On Error Resume Next
    txtOver.Text = Format(lblTotal.Caption - txtPayment.Text, "##,###0.00")
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
    NumberOnly KeyAscii
    If KeyAscii = 13 Then
        Call cmdProcess_Click
    End If
End Sub
