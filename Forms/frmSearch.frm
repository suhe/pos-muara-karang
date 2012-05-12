VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Records"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Or"
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "And"
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.ComboBox cmbFields 
      Height          =   315
      ItemData        =   "frmSearch.frx":0A02
      Left            =   1800
      List            =   "frmSearch.frx":0A04
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5520
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Condition "
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      Begin VB.ComboBox cmbOperation 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   2
         ItemData        =   "frmSearch.frx":0A06
         Left            =   3960
         List            =   "frmSearch.frx":0A08
         TabIndex        =   17
         Text            =   "Semua Record"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSearch.frx":0A0A
         Left            =   240
         List            =   "frmSearch.frx":0A0C
         TabIndex        =   6
         Text            =   "Mengandung Kata"
         Top             =   1080
         Width           =   2470
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSearch.frx":0A0E
         Left            =   240
         List            =   "frmSearch.frx":0A10
         TabIndex        =   2
         Text            =   "Mengandung Kata"
         Top             =   360
         Width           =   2470
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   211288067
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   3
         Left            =   4920
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   211288067
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   211288065
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   1
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   211288065
         CurrentDate     =   38207
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   390
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Records Where?"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public srcColumnHeaders As ColumnHeaders 'Source column headers
Public srcNoOfCol As Long
Public srcform As Form 'Source form


Private Sub cmbOperation_Click(Index As Integer)
    If Index = 0 Then
        If cmbOperation(Index).Text = "Tanggal" Then
            dtpDate(0).Visible = True
            dtpDate(1).Visible = True
            txtFilter(0).Visible = False
        Else
            txtFilter(0).Visible = True
            dtpDate(0).Visible = False
            dtpDate(1).Visible = False
        End If
    Else
        If cmbOperation(Index).Text = "Tanggal" Then
            dtpDate(2).Visible = True
            dtpDate(3).Visible = True
            txtFilter(1).Visible = False
        Else
            txtFilter(1).Visible = True
            dtpDate(2).Visible = False
            dtpDate(3).Visible = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim varx  As String
    If cmbOperation(0).Text <> "Tanggal" Then If txtFilter(0).Text = "" Then txtFilter(0).SetFocus: Exit Sub
    On Error GoTo err
    Dim strFilter, strx1, strx2 As String
    strFilter = Replace(cmbFields.Text, "/", "")
    strFilter = Replace(cmbFields.Text, " ", "_")
    If (tbl.TABLE_TANGGAL = "Date") Then
        strx1 = ""
        strx2 = ""
    Else
        strx1 = "00:00:00"
        strx2 = "24:00:00"
    End If
    
    'MsgBox strx1
    Select Case cmbOperation(0).Text
        Case "Mengandung Kata": strFilter = strFilter & " LIKE '%" & txtFilter(0).Text & "%'"
        Case "=": strFilter = strFilter & " = '" & txtFilter(0).Text & "'"
        Case "<>": strFilter = strFilter & " <> '" & txtFilter(0).Text & "'"
        Case ">": strFilter = strFilter & " > '" & txtFilter(0).Text & "'"
        Case ">=": strFilter = strFilter & " >= '" & txtFilter(0).Text & "'"
        Case "<": strFilter = strFilter & " < '" & txtFilter(0).Text & "'"
        Case "<=": strFilter = strFilter & " <= '" & txtFilter(0).Text & "'"
        Case "Tanggal": strFilter = strFilter & " >= '" & Format(dtpDate(0).Value, "YYYY-MM-DD") & " " & strx1 & "' AND " & strFilter & "<='" & Format(dtpDate(1).Value, "YYYY-MM-DD") & " " & strx2 & "'"
        Dim day1, month1, year1, day2, month2, year2 As String
        
        day1 = Day(dtpDate(0).Value)
        month1 = Month(dtpDate(0).Value)
        year1 = Year(dtpDate(0).Value)
        
        day2 = Day(dtpDate(1).Value)
        month2 = Month(dtpDate(1).Value)
        year2 = Year(dtpDate(1).Value)
        
        tbl.TABLE_TANGGAL_AWAL = year1 & "-" & month1 & "-" & day1
        tbl.TABLE_TANGGAL_AKHIR = year2 & "-" & month2 & "-" & day2
        'MsgBox tbl.TABLE_TANGGAL_AWAL
    End Select
    
    If cmbOperation(1).Text <> "" Then
        '-Second operation
        If Option1.Value = True Then
            strFilter = strFilter & " AND "
        Else
            strFilter = strFilter & " OR "
        End If
        
        varx = Replace(cmbFields.Text, "/", "")
        varx = Replace(cmbFields.Text, " ", "_")
        
        Select Case cmbOperation(1).Text
            Case "Mengandung Kata": strFilter = strFilter & varx & " LIKE '%" & txtFilter(1).Text & "%'"
            Case "=": strFilter = strFilter & varx & " = '" & txtFilter(1).Text & "'"
            Case "<>": strFilter = strFilter & varx & " <> '" & txtFilter(1).Text & "'"
            Case ">": strFilter = strFilter & varx & " > '" & txtFilter(1).Text & "'"
            Case ">=": strFilter = strFilter & varx & " >= '" & txtFilter(1).Text & "'"
            Case "<": strFilter = strFilter & varx & " < '" & txtFilter(1).Text & "'"
            Case "<=": strFilter = strFilter & varx & " <= '" & txtFilter(1).Text & "'"
            Case "Tanggal": strFilter = strFilter & varx & " >= '" & dtpDate(2).Value & "' AND " & strFilter & "<= '" & dtpDate(3).Value & "'"
        End Select
    End If
    
    If CurrBiz.BUSINNES_SALE = 1 Then
        strFilter = strFilter & " AND "
        varx = Replace(cmbFields.Text, "/", "")
        varx = Replace(cmbFields.Text, " ", "_")
        
        Select Case cmbOperation(2).Text
            Case "Semua Record": strFilter = strFilter & varx & " LIKE '%" & txtFilter(1).Text & "%'"
            Case "Lunas Debitor dan Kreditor": strFilter = strFilter & " flag_debitor=0 AND flag_kreditor=0 "
            Case "Lunas Kreditor": strFilter = strFilter & " flag_debitor=1 AND flag_kreditor=0 "
            Case "Lunas Debitor": strFilter = strFilter & " flag_debitor=1 AND flag_kreditor=1"
            Case "Hutang Debitor dan Kreditor": strFilter = strFilter & " flag_debitor=1 AND flag_kreditor=1"
            Case "Hutang Kreditor": strFilter = strFilter & " flag_debitor=0 AND flag_kreditor=1 "
            Case "Hutang Debitor": strFilter = strFilter & " flag_debitor=1 AND flag_kreditor=0"
        End Select
    End If
    
    'MsgBox strFilter
    srcform.FilterRecord strFilter
    strFilter = vbNullString
    Unload Me
    Exit Sub
err:
        If err.Number = -2147352571 Then
            MsgBox "Invalid search operation.", vbExclamation
            Unload Me
        ElseIf err.Number = 3001 Then
            Resume Next
        Else
            prompt_err err, "frmFilter", "cmdOk_Click"
        End If
End Sub

Private Sub Form_Load()
    'Initialize values
    dtpDate(0).Value = Date
    dtpDate(1).Value = Date
    dtpDate(2).Value = Date
    dtpDate(3).Value = Date
    'Set the images for the controls
    If CurrBiz.BUSINNES_SALE = 1 Then
        cmbOperation(2).Visible = True
    Else
        cmbOperation(2).Visible = False
    End If
    With MDIMainMenu
        Image1.Picture = .i16x16.ListImages(7).Picture
        Image2.Picture = .i16x16.ListImages(7).Picture
    End With
    
    Dim i, K As Integer
    If srcNoOfCol = 0 Then srcNoOfCol = srcColumnHeaders.Count
     
    If srcform.Name = "frmDebt" Then
        K = 2
    Else
        K = 1
    End If
    
    For i = K To srcNoOfCol
       
        If srcColumnHeaders(i).Text <> "" Then cmbFields.AddItem srcColumnHeaders(i).Text
    Next i
    i = 0
    cmbFields.ListIndex = 0
    With cmbOperation(0)
        .AddItem "Mengandung Kata"
        .AddItem "="
        .AddItem "<>"
        .AddItem ">"
        .AddItem ">="
        .AddItem "<"
        .AddItem "<="
        .AddItem "Tanggal"
    End With
    With cmbOperation(1)
        .AddItem "Mengandung Kata"
        .AddItem "="
        .AddItem "<>"
        .AddItem ">"
        .AddItem ">="
        .AddItem "<"
        .AddItem "<="
        .AddItem "Tanggal"
    End With
    
    With cmbOperation(2)
        .AddItem "Semua Record"
        .AddItem "Lunas Debitor dan Kreditor"
        .AddItem "Lunas Kreditor"
        .AddItem "Lunas Debitor"
        .AddItem "Hutang Debitor dan Kreditor"
        .AddItem "Hutang Kreditor"
        .AddItem "Hutang Debitor"
    End With
     
     cmbOperation(0).ListIndex = 0
     cmbOperation(1).ListIndex = 0
     cmbOperation(2).Text = "Semua Record"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSearch = Nothing
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
    HLText txtFilter(Index)
End Sub
