VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShortcuts 
   BackColor       =   &H80000005&
   Caption         =   "Shortcuts"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShortcuts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8700
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1_old 
      Left            =   2040
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2394
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":3070
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":4A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":6394
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":7D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":96B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":A392
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":B06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":BD46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":CA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":D6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":DFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":ECB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":F992
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1066E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":10F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":11C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1250A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":131E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":14B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1650E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1_oldest 
      Left            =   4440
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":16DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1AE3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1EE92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":22EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":26F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":29F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2DFE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":32036
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":3608A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":390DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":3D132
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":41186
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":451DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":4922E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":4D282
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":512D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":5532A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":5937E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":5D3D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":61426
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":6547A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":684CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":6C522
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   84
      ImageHeight     =   84
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":70576
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":765F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":7C593
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":81D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":878AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":8D4D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":932C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":99299
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":9FEBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":A5771
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":AB094
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":B114F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":B72C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":BE2BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":C3438
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":C91C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":CF5D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":D5905
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":DB0CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":DFF8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":E55B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":EB419
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":F1F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":F7D37
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":FCF16
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":103290
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1072E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":10C66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":112189
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":118251
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":11DD2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":123EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":129355
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":12FCC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":135953
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":13BB1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":14114D
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":147C4B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMenu 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   8916
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MousePointer    =   99
      MouseIcon       =   "frmShortcuts.frx":14CE2A
      OLEDragMode     =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CommandPass(ByVal srcPerformWhat As String)
    Select Case srcPerformWhat
        Case "New"
            '
        Case "Edit"
            'frmAbout.Show vbModal
    End Select
End Sub


Private Sub Active()
    'HighlightInWin Me.Name: Main.ShowTBButton "ttfffff"
    'With Main
     '   .tbMenu.Buttons(3).Caption = "User's Guide"
      '  .tbMenu.Buttons(3).Image = 10
        
       ' .tbMenu.Buttons(4).Caption = "About"
        '.tbMenu.Buttons(4).Image = 11
        
        '.mnuRACN.Caption = "User's Guide"
        '.mnuRAES.Caption = "About"
    'End With
End Sub

Private Sub Deactive()
    'Main.HideTBButton "", True
    'With Main
    '    .tbMenu.Buttons(3).Caption = "New"
     '   .tbMenu.Buttons(3).Image = 1
     '
       ' .tbMenu.Buttons(4).Caption = "Edit"
      '  .tbMenu.Buttons(4).Image = 2
    
        '.mnuRACN.Caption = "Create New"
        '.mnuRAES.Caption = "Edit Selected"
    'End With
End Sub

Private Sub Form_Activate()
    Active
    'HighlightInWin Name
End Sub

Private Sub Form_Deactivate()
    Deactive
End Sub

Private Sub Form_Load()
    
    With lvMenu
        Set .SmallIcons = ImageList1
        Set .Icons = ImageList1
        'For Sales
        .ListItems.Add , "frmCashier", "Kasir", 28, 28
        
        If (CurrUser.USER_ISADMIN = True) Or (CurrUser.USER_ISMANAGER = True) Then
             .ListItems.Add , "frmPurchasing", "Penambah Obat", 2, 2
             .ListItems.Add , "frmReturPurchase", "Retur Obat", 18, 18
        End If
         
        'If (CurrUser.USER_ISADMIN = True) Then
         '   .ListItems.Add , "frmProduct", "Stok Obat", 27, 27
        'End If
        
        
        If (CurrUser.USER_ISADMIN = True) Or (CurrUser.USER_ISMANAGER = True) Then
            .ListItems.Add , "frmProduct", "Stok Obat", 27, 27
            .ListItems.Add , "frmDebt", "Penagihan Kreditor", 15, 15
            .ListItems.Add , "frmPasien", "Data Pasien", 30, 30
            .ListItems.Add , "frmCategories", "Data Kategori Obat", 29, 29
            .ListItems.Add , "frmDepartement", "Data Departement", 31, 31
            .ListItems.Add , "frmSupplier", "Data Supplier", 9, 9
            .ListItems.Add , "frmKreditor", "Data Kreditor", 32, 32
            .ListItems.Add , "frmPurchase", "Laporan Transaksi Supplier", 35, 35
            .ListItems.Add , "frmProductList", "Laporan Banding Harga Obat", 16, 16
            .ListItems.Add , "frmSales", "Laporan Harian Pendapatan", 34, 34
            .ListItems.Add , "frmKomisi", "Laporan Harian Komisi Departement", 8, 8
            .ListItems.Add , "frmCashFlow", "Laporan Bulanan Cash Flow", 19, 19
        End If
        
        If (CurrUser.USER_ISADMIN = True) Then
            .ListItems.Add , "frmStockOpname", "Stok Opname", 37, 37
            .ListItems.Add , "frmGroup", "Cabang Klinik", 33, 33
            .ListItems.Add , "frmBusinessInfo", "info Klinik", 36, 36
        End If
        
        If (CurrUser.USER_ISADMIN = True) Then
            .ListItems.Add , "frmUserRec", "SDM Informatika Teknologi", 21, 21
        End If
        .ListItems.Add , "FrmLicense", "License POS", 38, 38
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Beep: Cancel = 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvMenu.Width = ScaleWidth
    lvMenu.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmShortcuts = Nothing
End Sub

Private Sub lvMenu_DblClick()
    Select Case lvMenu.SelectedItem.Key
        Case "frmCashier": LoadForm frmCashier
        Case "frmPurchasing"
              If CurrUser.USER_ISADMIN = False Then
              MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
              Else
                LoadForm frmPurchasing
              End If
        Case "frmProduct"
              If CurrUser.USER_ISADMIN = True Or CurrUser.USER_ISMANAGER = True Then
                     LoadForm frmProduct
                 Else
                    MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             End If
             
        Case "frmProductList"
              If CurrUser.USER_ISADMIN = True Or CurrUser.USER_ISMANAGER = True Then
                     LoadForm frmProductList
                 Else
                    MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             End If
             
        Case "frmStockOpname"
              If CurrUser.USER_ISADMIN = True Or CurrUser.USER_ISMANAGER = True Then
                     LoadForm frmStockOpname
                 Else
                    MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
              End If
             
        Case "frmCategories"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                   LoadForm frmCategories
             End If
        Case "frmPasien"
             'If CurrUser.USER_ISADMIN = False Then
                    ' MsgBox "Be Carefull Access this record.", vbCritical, "Access Denied"
             'Else
                    LoadForm frmPasien
             'End If
        Case "frmKreditor"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                   LoadForm frmKreditor
             End If
        Case "frmSupplier"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    LoadForm frmSupplier
             End If
        Case "frmSales"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                     LoadForm frmSales
             End If
         Case "frmKomisi"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                     LoadForm frmKomisi
             End If
         Case "frmDebt"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                     LoadForm frmDebt
             End If
        Case "frmCashFlow"
            If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    LoadForm frmCashFlow
            End If
        Case "frmDepartement"
            If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    LoadForm frmDepartement
            End If
        Case "frmPurchasing"
            If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    LoadForm frmPurchasing
            End If
        Case "frmPurchase"
            If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    LoadForm frmPurchase
            End If
        Case "frmSales"
            If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    LoadForm frmSales
            End If
        Case "frmReturPurchase"
            If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    LoadForm frmReturPurchase
            End If
        Case "frmBusinessInfo"
            If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
                 Else
                    frmBusinessInfo.show vbModal
            End If
        Case "frmGroup"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                frmGroup.OPEN_COMMAND = 1: frmGroup.show vbModal
             End If
        Case "frmCashFlow"
             If CurrUser.USER_ISADMIN = False Then
                     MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                LoadForm frmCashFlow
             End If
                 
        Case "frmUserRec"
             If CurrUser.USER_ISADMIN = False Then
                 MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
             Else
                frmUserRec.show vbModal
            End If
            
       Case "FrmLicense"
                FrmLicense.show vbModal
       End Select
End Sub

Private Sub lvMenu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lvMenu_DblClick
    End If
End Sub
