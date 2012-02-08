VERSION 5.00
Begin VB.Form frmBusinnesInfoTabel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabel Generator"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text3 
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
      Left            =   1560
      TabIndex        =   5
      Top             =   990
      Width           =   4335
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
      Left            =   1560
      TabIndex        =   4
      Top             =   540
      Width           =   4335
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -480
      ScaleHeight     =   30
      ScaleWidth      =   6690
      TabIndex        =   3
      Top             =   1410
      Width           =   6690
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
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox Text4 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1410
      Width           =   4335
   End
   Begin VB.TextBox Text5 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Tabel Expedition"
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
      Left            =   120
      TabIndex        =   12
      Top             =   990
      Width           =   1380
   End
   Begin VB.Label Label4 
      Caption         =   "Tabel Supplier"
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
      Left            =   120
      TabIndex        =   11
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tabel Customer"
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
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Faktur Jual"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1470
      Width           =   1620
   End
   Begin VB.Label Label5 
      Caption         =   "Faktur Beli"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1890
      Width           =   1395
   End
End
Attribute VB_Name = "frmBusinnesInfoTabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsSet As New Recordset
Dim rsgen As New Recordset

Private Sub Command1_Click()
    If is_empty(Text1) = True Then Exit Sub
    If is_empty(Text2) = True Then Exit Sub
    
    With rsgen
        .Fields("CUSTOMERS") = Text1.Text
        .Fields("SUPPLIERS") = Text2.Text
        .Fields("EXPEDITIONS") = Text3.Text
        .Fields("HEAD_SALES") = Text4.Text
        .Fields("HEAD_PURCHASING") = Text5.Text
        .Update
    End With
    
    'With CurrBiz
    '    .BUSINESS_ADDRESS = Text1.Text
    '    .BUSINESS_CONTACT_INFO = Text2.Text
    '    .BUSINNES_NAME = Text3.Text
    '    .BUSINNES_CITY = Text4.Text
    '    .BUSINNES_NOTE = Text5.Text
    'End With
    MsgBox "Changes has been successfully saved.", vbInformation
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rsgen.Open "SELECT * FROM TBL_GENERATOR WHERE TABLENAME='CUSTOMERS'", CN, adOpenStatic, adLockOptimistic
    Text1.Text = rsgen.Fields("NextNo")
    Text2.Text = rsgen.Fields("SUPPLIERS")
    Text3.Text = rsgen.Fields("EXPEDITIONS")
    Text4.Text = rsgen.Fields("HEAD_SALES")
    Text5.Text = rsgen.Fields("HEAD_PURCHASING")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBusinnesInfoTabel = Nothing
End Sub

Private Sub Text1_GotFocus()
    HLText Text1
End Sub

Private Sub Text2_GotFocus()
    HLText Text2
End Sub

Private Sub Text3_GotFocus()
    HLText Text3
End Sub

Private Sub Text4_GotFocus()
    HLText Text4
End Sub

Private Sub Text5_GotFocus()
    HLText Text5
End Sub


