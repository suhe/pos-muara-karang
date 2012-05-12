VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPurchaseReturAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Retur Product"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   Icon            =   "frmSalesReturAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRetur 
      Caption         =   "&Retur"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   16
         Tag             =   "No Faktur"
         ToolTipText     =   "Masukan No Fak & Enter Untuk Melanjutkan"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Tag             =   "No Faktur"
         ToolTipText     =   "Masukan No Fak & Enter Untuk Melanjutkan"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Tag             =   "Keterangan"
         Text            =   "Obat Sudah Kadaluarsa "
         Top             =   4080
         Width           =   3015
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Tag             =   "Jumlah Retur"
         Text            =   "0"
         ToolTipText     =   "Masukan Jumlah Retur & Tekan Retur Untuk Meretur"
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Tag             =   "No Faktur"
         Text            =   "K"
         ToolTipText     =   "Masukan No Fak & Enter Untuk Melanjutkan"
         Top             =   360
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo dcObat 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Tag             =   "Kode Obat"
         Top             =   1800
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1560
         TabIndex        =   20
         Tag             =   "Qty"
         Top             =   2760
         Width           =   105
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Harga Beli"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "---------------------------------------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1560
         TabIndex        =   18
         Top             =   2280
         Width           =   2925
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Obat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nm Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Fak :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1560
         TabIndex        =   6
         Tag             =   "Qty"
         Top             =   3240
         Width           =   105
      End
      Begin VB.Label BarCode 
         AutoSize        =   -1  'True
         Caption         =   "No Fak :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Kode Obat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Netto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Retur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   3720
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   4200
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmPurchaseReturAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsKode As New Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRetur_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    If is_empty(txtEntry(3), True) = True Then Exit Sub
    If is_empty(dcObat, True) = True Then Exit Sub
    'If is_zero(lblQty, True) = True Then Exit Sub
    If is_zero(txtEntry(2), True) = True Then Exit Sub
    If Val(txtEntry(2).Text) > Val(lblQty.Caption) Then MsgBox "Data Melebihi Stok Beli", vbAbortRetryIgnore + vbInformation: Exit Sub
    'On Error Resume Next
    Dim rsobat As New Recordset
    If rsobat.State = 1 Then rsobat.Close
    rsobat.Open "SELECT * FROM tbl_beli_details d JOIN vw_stok b ON b.id_obat=d.id_obat WHERE d.id_obat=" & dcObat.BoundText, CN, adOpenStatic, adLockReadOnly
    If rsobat.RecordCount > 0 Then
        tbl.TABLE_KD_OBAT = rsobat.Fields("kd_obat")
        tbl.TABLE_NM_OBAT = rsobat.Fields("nm_obat")
        tbl.TABLE_TOTAL = rsobat.Fields("jumlah")
        tbl.TABLE_RETUR_OBAT = txtEntry(2).Text
        tbl.TABLE_SISA_OBAT = rsobat.Fields("sisa")
        tbl.TABLE_SISA_RETUR = Val(rsobat.Fields("sisa")) - Val(txtEntry(2).Text)
    End If
    
    sql = "UPDATE tbl_beli_details "
    sql = sql + "SET "
    sql = sql + " tgl_retur='" & Format(Date, "YYYY-MM-DD") & "', "
    sql = sql + " retur=" & txtEntry(2).Text & " "
    sql = sql + " WHERE no_beli='" & txtEntry(0).Text & "'"
    sql = sql + " AND id_obat=" & dcObat.BoundText & ""
    CN.Execute sql
    
    sql = "UPDATE tbl_obat "
    sql = sql + "SET "
    sql = sql + " stok_temp=stok_temp - " & txtEntry(2).Text
    sql = sql + " WHERE id_obat=" & dcObat.BoundText & ""
    CN.Execute sql
    
    Call ReturObat
    Call GeneratePK
    
    sql = "INSERT INTO tbl_beli(no_beli,tgl_beli,tgl_bayar,id_supplier,type,payment,bayar,hutang,flag_supplier,tgl_akhir,tgl_input,id_pengguna) "
            sql = sql + "VALUES( "
            sql = sql + " '" & tbl.TABLE_NO_FAK & "',"
            sql = sql + " '" & Format(Now, "YYYY-mm-dd h:m:s") & "', "
            sql = sql + " '-',"
            sql = sql + " " & tbl.TABLE_ID_SUPPLIER & ", "
            sql = sql + " 'Cash',"
            sql = sql + " 'Hutang',"
            sql = sql + " 0, "
            sql = sql + " -" & Val(txtEntry(2).Text) * Val(Label8.Caption) & ", "
            sql = sql + " 0, "
            sql = sql + " '" & Format(Now, "YYYY-mm-dd h:m:s") & "', "
            sql = sql + " '" & Format(Now, "YYYY-mm-dd h:m:s") & "', "
            sql = sql + " " & CurrUser.USER_PK & " "
            sql = sql + ") "
    'MsgBox sql
    CN.Execute sql
    sql = "INSERT INTO tbl_beli_details(no_beli,id_obat,harga_beli,jumlah,tgl_retur) "
            sql = sql + "VALUES( "
            sql = sql + " '" & tbl.TABLE_NO_FAK & "',"
            sql = sql + " " & tbl.TABLE_ID_OBAT & ", "
            sql = sql + " " & Label8.Caption & ", "
            sql = sql + " -" & Val(txtEntry(2).Text) & ", "
            sql = sql + " '" & Format(Now, "YYYY-mm-dd h:m:s") & "' "
            sql = sql + ") "
    'MsgBox sql
    CN.Execute sql
    Unload Me
    frmReturPurchase.RefreshRecords
End Sub

Private Sub GeneratePK()
    Dim PK As Integer
    PK = getIndex("id_beli", "tbl_beli")
    tbl.TABLE_NO_FAK = "K" & tbl.TABLE_GROUP & PK
End Sub


Private Sub dcObat_Click(Area As Integer)
    If dcObat.BoundText <> "" Then
    'On Error Resume Next
    'Dim app As String
    txtEntry(2).Enabled = True
    txtEntry(3).Enabled = True
    txtEntry(2).SetFocus
    Dim rsKode As New Recordset
    If rsKode.State = 1 Then rsKode.Close
    rsKode.Open "SELECT tbl_beli_details.id_obat,nm_obat,tbl_beli_details.harga_beli,(jumlah-retur) as total FROM tbl_beli_details JOIN tbl_obat ON tbl_obat.id_obat=tbl_beli_details.id_obat WHERE no_beli='" & txtEntry(0).Text & "' AND tbl_beli_details.id_obat=" & dcObat.BoundText, CN, adOpenStatic, adLockReadOnly
    If (rsKode.RecordCount > 0) Then
        'lblQty.Caption = rsKode.Fields("total")
        'rsKode.Close
        'Set rsKode = Nothing
        'app = MsgBox("Anda Setuju , Akan Meretur Barang dengan Kode " & rsKode.Fields("total"), vbYesNo + vbQuestion)
        'If app = vbYes Then
            tbl.TABLE_ID_OBAT = rsKode.Fields("id_obat")
            Label5.Caption = rsKode.Fields("nm_obat")
            Label8.Caption = rsKode.Fields("harga_beli")
            lblQty.Caption = rsKode.Fields("total")
            cmdRetur.Enabled = True
        'End If
        'dcObat.Enabled = False
    Else
        MsgBox "Data Tidak Ketemu !", vbCritical + vbInformation
    End If
    End If
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

Private Sub Form_Unload(Cancel As Integer)
    Set frmPurchaseReturAE = Nothing
End Sub

Private Sub Label8_Change()
    dcObat.Enabled = False
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Index = 2) Then
        NumberOnly KeyAscii
        If (KeyAscii = 13) Then
            Dim beli, retur As Integer
            beli = Val(lblQty.Caption)
            retur = Val(txtEntry(2).Text)
            If (retur > beli) Then
                MsgBox "Angka Retur Melebih Jumlah Beli!", vbCritical + vbInformation
                Exit Sub
            ElseIf (txtEntry(2).Text = "") Then
                MsgBox "Isi Angka jangan dikosongkan", vbCritical + vbInformation
                Exit Sub
            ElseIf (retur = 0) Then
                MsgBox "Angka Retur Tidak boleh 0 !", vbCritical + vbInformation
                Exit Sub
            Else
                cmdRetur.Enabled = True
                cmdRetur.SetFocus
            End If
        End If
    End If
   
    If (KeyAscii = 13) Then
        If (Index = 0) Then
            Dim total As Byte
            total = getRecordCount("id_beli", "tbl_beli", "WHERE no_beli ='" & txtEntry(0).Text & "' ")
            Dim rsreturning As New Recordset
            If rsreturning.State = 1 Then rsreturning.Close
            rsreturning.Open " SELECT *,s.id_supplier FROM tbl_beli b JOIN tbl_supplier s ON s.id_supplier=b.id_supplier WHERE b.no_beli='" & txtEntry(0).Text & "'", CN, adOpenStatic, adLockReadOnly
            If (rsreturning.RecordCount > 0) Then
                tbl.TABLE_NO_FAK = rsreturning.Fields("no_beli")
                tbl.TABLE_TANGGAL = rsreturning.Fields("tgl_beli")
                tbl.TABLE_ID_SUPPLIER = rsreturning.Fields("id_supplier")
                tbl.TABLE_NM_SUPPLIER = rsreturning.Fields("nm_supplier")
                txtEntry(1).Text = Format(rsreturning.Fields("tgl_beli"), "DD/MM/YYYY")
                txtEntry(4).Text = rsreturning.Fields("nm_supplier")
            End If
            
            If (total > 0) Then
                MsgBox "Data Ditemukan Silahkan Pilih Kode Obat !", vbOKCancel + vbInformation
                txtEntry(0).Enabled = False
                bind_dc "SELECT o.id_obat,o.nm_obat FROM tbl_beli_details d JOIN tbl_obat o ON o.id_obat=d.id_obat WHERE d.no_beli='" & Trim(txtEntry(0).Text) & "'", "nm_obat", dcObat, "id_obat"
                dcObat.Enabled = True
            Else
                dcObat.Enabled = False
                txtEntry(2).Text = ""
                txtEntry(3).Text = ""
                dcObat.Text = ""
                txtEntry(2).Enabled = False
                txtEntry(3).Enabled = False
                MsgBox "No No.Faktur This Data , Please Re-Enter No Faktur !", vbCritical + vbInformation
            End If
        End If
    End If
End Sub

