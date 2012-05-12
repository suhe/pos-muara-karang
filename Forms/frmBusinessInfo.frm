VERSION 5.00
Begin VB.Form frmBusinessInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Klinik Information"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmBusinessInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
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
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   3
      Top             =   1380
      Width           =   4335
   End
   Begin VB.TextBox Text1 
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
      Top             =   120
      Width           =   4335
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -450
      ScaleHeight     =   30
      ScaleWidth      =   6690
      TabIndex        =   9
      Top             =   1380
      Width           =   6690
   End
   Begin VB.TextBox Text2 
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
      Top             =   510
      Width           =   4335
   End
   Begin VB.TextBox text3 
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
      Top             =   960
      Width           =   4335
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
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
      Left            =   4680
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Note :"
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
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "City :"
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
      Top             =   1440
      Width           =   1620
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
      Left            =   120
      TabIndex        =   10
      Top             =   90
      Width           =   1455
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
      Left            =   150
      TabIndex        =   8
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Contact Info:"
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
      Left            =   150
      TabIndex        =   7
      Top             =   960
      Width           =   1140
   End
End
Attribute VB_Name = "frmBusinessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_set As New Recordset

Private Sub Command1_Click()
    If is_empty(Text1) = True Then Exit Sub
    If is_empty(text2) = True Then Exit Sub
    If is_empty(text3) = True Then Exit Sub
    If is_empty(Text4) = True Then Exit Sub
    If is_empty(Text5) = True Then Exit Sub
    
        sql = "UPDATE tbl_business_info "
        sql = sql + "SET "
        sql = sql + " bussines_name='" & Text1.Text & "', "
        sql = sql + " bussines_address='" & text2.Text & "', "
        sql = sql + " bussines_cp='" & text3.Text & "', "
        sql = sql + " bussines_city='" & Text4.Text & "', "
        sql = sql + " bussines_note='" & Text5.Text & "' "
        CN.Execute sql
        
    With CurrBiz
        .BUSINNES_NAME = Text1.Text
        .BUSINESS_ADDRESS = text2.Text
        .BUSINESS_CONTACT_INFO = text3.Text
        .BUSINNES_CITY = Text4.Text
        .BUSINNES_NOTE = Text5.Text
    End With
    
    MsgBox "Changes has been successfully saved.", vbInformation
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rs_set.Open "SELECT * FROM tbl_business_info", CN, adOpenStatic, adLockOptimistic
    Text1.Text = rs_set.Fields("bussines_name")
    text2.Text = rs_set.Fields("bussines_address")
    text3.Text = rs_set.Fields("bussines_cp")
    Text4.Text = rs_set.Fields("bussines_city")
    Text5.Text = rs_set.Fields("bussines_note")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBusinessInfo = Nothing
End Sub
