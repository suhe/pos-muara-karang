VERSION 5.00
Begin VB.Form frmKeygen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keygen"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmKeygen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3840
      TabIndex        =   6
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1200
   End
   Begin VB.TextBox txtComputerID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   3795
   End
   Begin VB.TextBox txtLicense 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   1
      Top             =   960
      Width           =   3765
   End
   Begin VB.Label Label1 
      Caption         =   "Computer ID"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Licence Code:"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label6 
      Caption         =   "KeyGen for Standard Trial system with 30 secound delay."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmKeygen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objs As Object
Dim obj As Object
Dim WMI As Object
Dim Code1 As Single
Dim i As Integer
Dim zip, final As String

Function GetCpuID()
  Dim WMI, cpu, cpuid
  Set WMI = GetObject("winmgmts:")
  For Each cpu In WMI.InstancesOf("Win32_Processor")
   cpuid = cpuid + cpu.ProcessorID
  Next
  MsgBox cpuid
End Function

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdCreate_Click()
    Dim Code1 As Single
    If Len(txtComputerID.Text) < 4 Then
        MsgBox "The Name must be more than 4 characters.", vbInformation + vbOKOnly, "Ooops"
        Exit Sub
    End If
    For i = 1 To Len(txtComputerID.Text) - 1
        Code1 = Format(Asc(Right(txtComputerID.Text, Len(txtComputerID.Text) - i)) * 2 + (10 / i) + (i + 3 / 7), "#.#")
        zip = zip & Code1
    Next i
    zip = Right(zip, 8)
    For i = 1 To Len(zip) - 1
        Code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 4), "#00")
        final = final & Code1
    Next i
    final = Right(final, Len(final) - 4)
    final = final & Asc(txtComputerID.Text)
    txtLicense.Text = final
End Sub

Private Sub Form_Load()
    'GetCpuID
End Sub
