VERSION 5.00
Begin VB.Form frmStockChecking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Checking"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   150
   End
   Begin VB.Label lblCurrentRecord 
      AutoSize        =   -1  'True
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Business Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -1410
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Business Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -1440
      TabIndex        =   2
      Top             =   660
      Width           =   1455
   End
End
Attribute VB_Name = "frmStockChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSave_Click()
    If Text1.Text = "" Then MsgBox "Empty Text For + Correction", vbOKOnly + vbCritical: Exit Sub
    If Text2.Text = "" Then MsgBox "Empty Text For - Correction", vbOKOnly + vbCritical: Exit Sub
    sql = "UPDATE PRODUCTS SET QTY_PLUS=" & Text1.Text & " , QTY_MIN=" & Text2.Text & " WHERE PK=" & Me.Caption
    'MsgBox sql
    CN.Execute sql
    frmStock.RefreshRecords
    Unload Me
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
    Text1.Text = frmStock.lvList.SelectedItem.SubItems(11)
    Text2.Text = frmStock.lvList.SelectedItem.SubItems(12)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
    'If Index > 3 And Index < 13 Then
     '   KeyAscii = isNumber(KeyAscii)
    'End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
    'If Index > 3 And Index < 13 Then
        'KeyAscii = isNumber(KeyAscii)
    'End If
End Sub
