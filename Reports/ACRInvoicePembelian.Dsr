VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ACRInvoicePembelian 
   Caption         =   "Laporan Pembelian"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19315
   SectionData     =   "ACRInvoicePembelian.dsx":0000
End
Attribute VB_Name = "ACRInvoicePembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_Activate()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fffffft"
End Sub

Private Sub ActiveReport_Deactivate()
    MDIMainMenu.HideTBButton "", True
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        'Case vbKeyF1: MDIMainMenu.tbMenu.Button.Key = "Shortcut"
        Case vbKeyF2: CommandPass "New"
        Case vbKeyF3: CommandPass "Edit"
        Case vbKeyF4: CommandPass "Search"
        Case vbKeyF5: CommandPass "Delete"
        Case vbKeyF6: CommandPass "Refresh"
        Case vbKeyF8: CommandPass "Close"
    End Select
End Sub

Private Sub ActiveReport_ReportEnd()
    MDIMainMenu.RemToWin Me.Caption
        Dim total As Long
        Dim bayar As Double
        Dim rsPay As New Recordset
       
        sql = " SELECT b.id_beli,b.no_beli,b.tgl_beli,o.kd_obat,o.nm_obat,b.hutang,d.harga_beli,d.jumlah,(d.harga_beli * d.jumlah) as total,b.bayar"
        sql = sql + " FROM tbl_beli b"
        sql = sql + " INNER JOIN tbl_beli_details d ON d.no_beli=b.no_beli"
        sql = sql + " INNER JOIN tbl_obat o ON o.id_obat=d.id_obat"
        sql = sql + " WHERE b.flag_supplier=1 AND b.id_supplier = " & Trim(tbl.TABLE_ID_SUPPLIER) & " "
        
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If
        
        sql = sql + " ORDER BY b.id_beli ASC "
        
        Set rsPay = New Recordset
        If rsPay.State = 1 Then rsPay.Close
        rsPay.Open sql, CN, adOpenStatic, adLockReadOnly
        
        total = 0
        bayar = 0
        Do While Not rsPay.EOF
        'total = .lvList.ListItems.Count
        'If (total > 0) Then
            'For i = 1 To total
                sql = "UPDATE tbl_beli "
                sql = sql + " SET "
                sql = sql + " tgl_bayar='" & Format(Date, "YYYY-MM-DD") & "',"
                sql = sql + " payment='Lunas', "
                sql = sql + " flag_supplier= 0, "
                sql = sql + " hutang= 0, "
                sql = sql + " bayar=" & rsPay.Fields("hutang") & " "
                sql = sql + " WHERE id_beli=" & rsPay.Fields("id_beli") & ""
                bayar = bayar + Val(Format(rsPay.Fields("total"), ""))
                CN.Execute sql
            total = total + rsPay.Fields("jumlah")
            rsPay.MoveNext
           Loop
            
            rsPay.MoveFirst
            'Next i
            tbl.TABLE_TANGGAL_AWAL = rsPay.Fields("tgl_beli")
            rsPay.MoveLast
            tbl.TABLE_TANGGAL_AKHIR = rsPay.Fields("tgl_beli")
            tbl.TABLE_TOTAL = bayar
            tbl.TABLE_TOTAL_OBAT = total
            frmPurchase.RefreshRecords
End Sub

Private Sub ActiveReport_ReportStart()
    HighlightInWin Me.Name: MDIMainMenu.ShowTBButton "fffffft"
    MDIMainMenu.AddToWin Me.Caption, Name
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "Close"
            Unload Me
            cetak_Invoice_Pembelian
    End Select
    Exit Sub
    'Trap the error
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
End Sub
