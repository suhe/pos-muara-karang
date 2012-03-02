VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSplashLicense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License Code"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmSplashLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTrial 
      Caption         =   "&Use Trial 30 Days"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdActive 
      Caption         =   "&Active Now"
      Default         =   -1  'True
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Licence Infomation"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   6015
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   22
         TabIndex        =   9
         Top             =   480
         Width           =   4575
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Licence Code not required when deactivating?"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin RichTextLib.RichTextBox text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   503
         _Version        =   393217
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmSplashLicense.frx":038A
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer ID"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Licence Code"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Question:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox regpic 
      BackColor       =   &H00E8AB2F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   495
      Left            =   360
      Picture         =   "frmSplashLicense.frx":040C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin RichTextLib.RichTextBox code 
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSplashLicense.frx":084E
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplashLicense.frx":08D0
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Activation Required"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E8AB2F&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmSplashLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LocationReg = "System\Windows\User"
Const PasswordReg = "kode"

Function GetInfoReg() As String
    On Error GoTo Ero
    Dim Reg As Object
    Set Reg = CreateObject("WScript.Shell")
    GetInfoReg = Reg.RegRead("HKEY_CLASSES_ROOT\" & LocationReg & "\")
    Exit Function
Ero:
    Reg.RegWrite "HKEY_CLASSES_ROOT\" & LocationReg & "\", Format(Now, "short date") 'memasukkan tgl sekarang
    GetInfoReg = Format(Now, "short date")
End Function

Function SuccessReg() As Boolean 'fungsi utk prosedur pemasukan kode registrasi
    Dim s As String
Lagi:
    s = InputBox("Masukkan kode registrasi", "Registrasi")
    If s = PasswordReg Then
    Dim Reg As Object
    Set Reg = CreateObject("WScript.Shell")
    Reg.RegWrite "HKEY_CLASSES_ROOT\" & LocationReg & "\", "Registered" 'mendaftarkan k registry
    MsgBox "Terima kasih telah melakukan registrasi", vbInformation, "Registrasi Sukses"
    SuccessReg = True
          
    ElseIf s = "" Then
    SuccessReg = False
      
    Else
    If MsgBox("Maaf kode registrasi salah, coba lagi ?", vbYesNo + vbExclamation, "Registrasi") = vbYes Then GoTo Lagi
    SuccessReg = False
    End If
End Function

Private Sub cmdActive_Click()
    Dim code1
    Dim final
    Dim zip
    Dim WMI, cpu, cpuid
    Dim i As Integer
    
    If Len(Text2.Text) < 17 Then
        MsgBox "The Licence Code you have entered is an invalid length."
        Exit Sub
    End If
    
    Set WMI = GetObject("winmgmts:")
    For Each cpu In WMI.InstancesOf("Win32_Processor")
     cpuid = cpuid + cpu.ProcessorID
    Next
    
    For i = 1 To Len(cpuid) - 1
    code1 = Format(Asc(Right(cpuid, Len(cpuid) - i)) * 2 + (10 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
    Next i
    zip = Right(zip, 8)
    
    For i = 1 To Len(zip) - 1
        code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 4), "#00")
        final = final & code1
    Next i
    final = Right(final, Len(final) - 4)
    final = final & Asc(cpuid)
    If (final = Text2.Text) Then
        code.Text = final
        code.SaveFile "c:\windows\system\pos_license.rtf"
        MsgBox "License Anda Benar Selamat Menikmati Fitur Full Optimal dari Software Kami Terima Kasih!", vbInformation + vbOKOnly
        CurrUser.USER_TRIAL = 0
        Unload Me
    Else
        MsgBox "License Salah Coba Masukan License yang benar!", vbCritical + vbInformation
    End If
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdTrial_Click()
    CurrUser.USER_TRIAL = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Dim WMI, cpu, cpuid
    Set WMI = GetObject("winmgmts:")
    For Each cpu In WMI.InstancesOf("Win32_Processor")
        cpuid = cpuid + cpu.ProcessorID
    Next
    text1.Text = cpuid
    RegTrial
End Sub

Private Sub RegTrial()
    Dim s As String, l As Byte
    Dim dayone As Byte
    dayone = 5
    s = GetInfoReg
    If s <> "Registered" Then 'jika belum terdaftar"
        l = dayone - (CDate(Format(Now, "short date")) - CDate(s)) 'max penggunaan trial 30 hari
        If l > 0 And l <= 30 Then 'jika masih ada sisa hari
            cmdTrial.Enabled = True
            cmdTrial.Caption = "&Use Trial " & l & " Days"
        Else 'jika kadaluarsa
            cmdTrial.Enabled = False
            cmdTrial.Caption = "&Expired"
        End If
    End If
End Sub
