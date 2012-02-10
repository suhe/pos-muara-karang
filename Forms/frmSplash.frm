VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3840
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":038A
   ScaleHeight     =   3840
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   5325
      Top             =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aplha V.2.5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   3000
      Width           =   3135
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DisableLoader As Boolean
Dim c As Integer

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim path As String
    
    If DisableLoader = False Then
        tmrUnload.Enabled = True
    End If
    'path = App.path & "\Images\logo\logo.jpg"
    'Me.Picture = path
End Sub

Private Sub Form_Unload(Cancel As Integer)
    c = 0
End Sub



Private Sub tmrUnload_Timer()
    c = c + 1
    If c = 5 Then MDIMainMenu.Enabled = True: Unload Me
    If c = 2 Then MDIMainMenu.Visible = True
End Sub
