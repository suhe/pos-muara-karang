VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGrafik 
   Caption         =   "Chart"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   9465
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9405
      TabIndex        =   1
      Top             =   7950
      Width           =   9465
      Begin VB.Frame Frame3 
         Caption         =   "| Per Year |"
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
         Left            =   3000
         TabIndex        =   11
         Top             =   0
         Width           =   2535
         Begin VB.ComboBox ComboYear 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Text            =   "2011"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdProses2 
            Caption         =   "&Proses"
            Height          =   375
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "| Per Month |"
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
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2895
         Begin VB.CommandButton cmdProses 
            Caption         =   "&Proses"
            Height          =   375
            Left            =   1800
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cmbTahun 
            Height          =   315
            Left            =   840
            TabIndex        =   9
            Text            =   "2011"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cmbBulan 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "| Ekspor |"
         Height          =   735
         Left            =   5640
         TabIndex        =   2
         Top             =   0
         Width           =   5415
         Begin VB.CommandButton CmdSaveas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Save As..."
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox Picture2 
            Height          =   375
            Left            =   4320
            ScaleHeight     =   315
            ScaleWidth      =   555
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exit"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton CmdPrint 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Print Chart"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4080
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   7935
      Left            =   0
      OleObjectBlob   =   "frmGrafikCashFlow.frx":0000
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   7935
      Left            =   0
      OleObjectBlob   =   "frmGrafikCashFlow.frx":2521
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmGrafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rss As New ADODB.Recordset ' Recordset
Dim RssY As New ADODB.Recordset ' Recordset
Dim ArrayChart() 'Array
Dim X As Long
Dim ssql As String

Private Sub CmdPrint_Click()
    On Error Resume Next
    If MSChart1.Enabled = True Then
        MSChart1.EditCopy 'This Makes MSChart Control to be Copied
    Else
        MSChart2.EditCopy
    End If
    DoEvents   ' may be needed for large datasets
    Printer.Print " "
    Picture2.PaintPicture Clipboard.GetData(), 0, 0
    Printer.EndDoc
End Sub

Private Sub cmdProses_Click()
    On Error Resume Next
    MSChart1.Visible = True
    MSChart2.Visible = False
    Set Rss = Nothing
    ssql = " SELECT * " & _
           " From QR_TOP_TEN WHERE Month = " & cmbBulan.Text & " AND Year=" & cmbTahun.Text & " ORDER BY Qty DESC"
    'MsgBox ssql
    Rss.Open ssql, CN, adOpenStatic, adLockReadOnly ' Making Recordset
    If Rss.RecordCount = 0 Then ' If no Record in Database, then Show an Error Msg and Exit the Sub
        MsgBox "No Data to Show on Chart!!!", vbCritical, "Chart": Exit Sub
    Else
        ReDim ArrayChart(1 To Rss.RecordCount, 1 To 2) ' Array
        'Puuting Records from Database to Array
        'MSChart1.Legend = "test"
        For X = 1 To Rss.RecordCount
        ArrayChart(X, 1) = Rss!ProductCode
        ArrayChart(X, 2) = Val(Rss!qty)
        'ArrayChart(X, 3) = Val(Rss!qty) / 1000
        'ArrayChart(X, 4) = Rss!payment
        Rss.MoveNext
        Next X
        
   With MSChart1.Legend
    .Location.Visible = True
    .VtFont.Name = "Arial"
    .VtFont.Size = 8
    .Location.LocationType = VtChLocationTypeTop
    .VtFont.Effect = VtFontStyleBold
  End With
  MSChart1.Plot.SeriesCollection(1).LegendText = "Penjualan"
  'MSChart1.Plot.SeriesCollection(2).LegendText = "Pembelian"
  'MSChart1.Plot.SeriesCollection(3).LegendText = "Pemby.Hutang"
'# Assigns our array to the MSChart control #
    MSChart1.ChartData = ArrayChart
    MSChart1.Refresh
    End If
    Frame1.Enabled = True
    Rss.Close
    Set Rss = Nothing
End Sub

Private Sub cmdProses2_Click()
    On Error Resume Next
    MSChart1.Visible = False
    MSChart2.Visible = True
    Set RssY = Nothing
    ssql = " SELECT * " & _
           " From QR_TOP_TEN WHERE Year = " & ComboYear.Text & " ORDER BY Month ASC"
    'MsgBox ssql
    RssY.Open ssql, CN, adOpenStatic, adLockReadOnly ' Making Recordset
    If RssY.RecordCount = 0 Then ' If no Record in Database, then Show an Error Msg and Exit the Sub
        MsgBox "No Data to Show on Chart!!!", vbCritical, "Chart": Exit Sub
    Else
        ReDim ArrayChart(1 To RssY.RecordCount, 1 To 2) ' Array
        'Puuting Records from Database to Array
        'MSChart1.Legend = "test"
        For X = 1 To RssY.RecordCount
        ArrayChart(X, 1) = RssY!ProductCode
        ArrayChart(X, 2) = Val(RssY!qty)
        'ArrayChart(X, 3) = Val(RssY!totalPurchase) / 1000
        RssY.MoveNext
        Next X
        
   With MSChart1.Legend
    .Location.Visible = True
    .VtFont.Name = "Arial"
    .VtFont.Size = 8
    .Location.LocationType = VtChLocationTypeTop
    .VtFont.Effect = VtFontStyleBold
  End With
  MSChart2.Plot.SeriesCollection(1).LegendText = "Penjualan"
  MSChart2.Plot.SeriesCollection(2).LegendText = "Pembelian"
  'MSChart1.Plot.SeriesCollection(3).LegendText = "Pemby.Hutang"
'# Assigns our array to the MSChart control #
    MSChart2.ChartData = ArrayChart
    MSChart2.Refresh
    End If
    Frame1.Enabled = True
    'Rss.Close
    'Set Rss = Nothing
End Sub

Private Sub CmdSaveas_Click()
    On Error GoTo Hell
    Dim strsavefile As String
    With CommonDialog1
        .Filter = "Pictures (*.bmp)|*.bmp" ' You can Also Save the Pic in JPG/GIF/TIFF
        .DefaultExt = "bmp"
        .CancelError = False
        .ShowSave
        strsavefile = .FileName
        If strsavefile = "" Then Exit Sub
    End With
    If MSChart1.Visible = True Then
        MSChart1.EditCopy
    Else
        MSChart2.EditCopy
    End If
    SavePicture Clipboard.GetData, strsavefile ' File Saved
    Exit Sub
Hell:
    MsgBox err.Description
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i, j As Integer
    For i = 1 To 12
        cmbBulan.AddItem i
    Next i
    
    For j = 2010 To 2022
        cmbTahun.AddItem j
    Next j
    
    For j = 2010 To 2022
        ComboYear.AddItem j
    Next j
    'Frame1.Enabled = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        MSChart1.Width = Me.ScaleWidth
        MSChart1.Height = (Me.ScaleHeight - Picture1.Height) - MSChart1.Top
        
        MSChart2.Width = Me.ScaleWidth
        MSChart2.Height = (Me.ScaleHeight - Picture1.Height) - MSChart2.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGrafik = Nothing
End Sub

