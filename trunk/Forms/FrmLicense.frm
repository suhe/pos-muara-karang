VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLicense 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License Software"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "FrmLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLicense 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "Update License"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txticense 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   4095
   End
   Begin RichTextLib.RichTextBox code 
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmLicense.frx":038A
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Bantuan Pengisian Serial Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   5925
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Untuk Mendapatkan License Number Hubungi 085.222.054.064  untuk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   5925
   End
   Begin VB.Label lblLicense 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   4
      Top             =   360
      Width           =   3045
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "License Number  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1605
   End
   Begin VB.Label lblComputerID 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   3045
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer ID        :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   240
      Picture         =   "FrmLicense.frx":040C
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "FrmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLicense_Click()
    'Dim final, zip As String
    Dim i As Integer
    Dim code1
    Dim final
    Dim zip
    
    For i = 1 To Len(lblComputerID.Caption) - 1
    code1 = Format(Asc(Right(lblComputerID.Caption, Len(lblComputerID.Caption) - i)) * 2 + (10 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
    Next i
    zip = Right(zip, 8)
    
    For i = 1 To Len(zip) - 1
        code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 4), "#00")
        final = final & code1
    Next i
    final = Right(final, Len(final) - 4)
    final = final & Asc(lblComputerID.Caption)
    
    If (final = txticense.Text) Then
        code.Text = final
        code.SaveFile "c:\windows\system\pos_license.rtf"
        MsgBox "License Anda Benar Selamat Menikmati Fitur Full Optimal dari Software Kami Terima Kasih!", vbInformation + vbOKOnly
    Else
        MsgBox "License Salah Coba Masukan License yang benar!", vbCritical + vbInformation
    End If
End Sub

Private Sub Form_Load()
  Dim WMI, cpu, cpuid
  Set WMI = GetObject("winmgmts:")
  For Each cpu In WMI.InstancesOf("Win32_Processor")
   cpuid = cpuid + cpu.ProcessorID
  Next
  lblComputerID.Caption = cpuid
  If Dir("c:\windows\system\pos_license.rtf") <> "" Then
    code.LoadFile "c:\windows\system\pos_license.rtf"
  Else
    code.SaveFile "c:\windows\system\pos_license.rtf"
    code.LoadFile "c:\windows\system\pos_license.rtf"
  End If
  lblLicense.Caption = code.Text
End Sub
