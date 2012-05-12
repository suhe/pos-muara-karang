Attribute VB_Name = "Mod_Print"
Option Explicit
Dim i As Long

Public Sub cetak_Transfer()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim harga, total, bayar, kembali As Double
    On Error GoTo opps
     With Printer
        .Font.Name = "Arial Narrow"
        .Font.Size = 12
        .CurrentY = .CurrentY + 200 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Transfer Bank"; Spc(5); " "; Spc(5); ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal "; Spc(4); ":"; Spc(5); "" & tbl.TABLE_TGL_CASH & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Jam "; Spc(7); ":"; Spc(4); ""; Hour(Now); ":"; Minute(Now); ":"; Second(Now); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Laba Bersih "; Spc(0); ":"; Spc(5); "" & Format(tbl.TABLE_LABA_BERSIH, "##,###0") & ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Transfer  "; Spc(3); ":"; Spc(5); "" & Format(tbl.TABLE_TRANSFER, "##,###0") & ""; Tab(70); " "; Spc(5); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Sisa Kas"; Spc(4); ":"; Spc(5); "" & Format(tbl.TABLE_KAS_SISA, "##,###0") & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        Printer.Print ""
        .EndDoc
     End With
opps:
End Sub

Public Sub ReturObat()
    Dim Lines As Integer, Y As Long, OutStr As String
    On Error GoTo opps
     With Printer
        .Font.Name = "Arial Narrow"
        .Font.Size = 10
        .CurrentY = .CurrentY + 200 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Retur Obat "; Spc(4); " "; Spc(5); ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        Printer.Print ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " No.Faktur "; Spc(4); ":"; Spc(5); "" & tbl.TABLE_NO_FAK & ""; Tab(70); ""; Spc(7); ""; Spc(1); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal "; Spc(6); ":"; Spc(5); "" & tbl.TABLE_TANGGAL & ""; Tab(70); ""; Spc(7); ""; Spc(1); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Supplier "; Spc(5); ":"; Spc(5); "" & tbl.TABLE_NM_SUPPLIER & ""; Tab(70); ""; Spc(7); ""; Spc(1); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Obat "; Spc(3); ":"; Spc(5); "" & tbl.TABLE_KD_OBAT & ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Obat  "; Spc(1); ":"; Spc(5); "" & tbl.TABLE_NM_OBAT & ""; Tab(70); " "; Spc(5); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Beli "; Spc(10); ":"; Spc(5); "" & tbl.TABLE_TOTAL & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
       ' .CurrentX = .CurrentX + 500 ' Skip some space
       ' Printer.Print " Stok "; Spc(9); ":"; Spc(5); "" & tbl.TABLE_SISA_OBAT & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Retur "; Spc(8); ":"; Spc(5); "" & tbl.TABLE_RETUR_OBAT & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Stok "; Spc(9); ":"; Spc(5); "" & tbl.TABLE_SISA_RETUR & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        Printer.Print ""
        .EndDoc
     End With
opps:
     
End Sub


Public Sub cetak_Faktur()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim rscetak As New Recordset
    Dim harga, total, bayar, kembali As Double
    Dim rentang, jmlRentang, X As Byte
    rentang = 9
    On Error GoTo opps
     With Printer
        .Font.Name = "Arial Narrow"
        .Font.Size = 7
        .CurrentY = .CurrentY + 0 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " No Faktur"; Spc(5); ":"; Spc(5); "" & tbl.TABLE_NO_FAK & ""; Tab(70); "Tanggal "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TANGGAL & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Pasien"; Spc(2); ":"; Spc(5); "" & tbl.TABLE_KD_PASIEN & " / Umur : " & tbl.TABLE_UMUR_PASIEN & ""; Tab(70); "Jam "; Spc(7); ":"; Spc(1); "  "; Hour(Now); ":"; Minute(Now); ":"; Second(Now); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Pasien "; Spc(1); ":"; Spc(5); "" & tbl.TABLE_NM_PASIEN & ""; Tab(70); "Telepon "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TLP_PASIEN & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Dept  "; Spc(2); ":"; Spc(5); "" & tbl.TABLE_NM_DEPT & ""; Tab(70); "Rp :"
        .CurrentX = .CurrentX + 500 ' Skip some space
        'Printer.Print "   "; Spc(2); ""; Spc(5); "" & tbl.TABLE_NM_DEPT & ""
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentX = .CurrentX + 500
        Printer.Print " Nama Obat "; Tab(50); "Kemasan"; Tab(70); "Jumlah"; Tab(85); "Dosis";
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentY = .CurrentY + 0
        Dim xx As Byte
         For xx = 1 To 6 Step xx + 1
             .CurrentX = .CurrentX + 500
            Printer.Print "  "; Tab(50); "  "; Tab(65); " "; Tab(85); "  ";
            Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " ----------------------------------- ----------------------- -------------- ------------------- ---------------- "
            .CurrentY = .CurrentY + 1
        Next xx
        .EndDoc
     End With
opps:
End Sub

Public Sub cetak_Faktur2()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim rscetak As New Recordset
    Dim rsPlafon As New Recordset
    Dim harga, total, bayar, kembali As Double
    
    Dim rentang, jmlRentang, X As Byte
    rentang = 9
    On Error GoTo opps
     With Printer
        .Font.Name = "Arial Narrow"
        .Font.Size = 7
        .CurrentY = .CurrentY + 0 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " No Faktur"; Spc(5); ":"; Spc(5); "" & tbl.TABLE_NO_FAK & ""; Tab(70); "Tanggal "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TANGGAL & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Pasien"; Spc(2); ":"; Spc(5); "" & tbl.TABLE_KD_PASIEN & " "; Tab(70); "Jam "; Spc(7); ":"; Spc(1); "  "; Hour(Now); ":"; Minute(Now); ":"; Second(Now); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Pasien "; Spc(1); ":"; Spc(5); "" & tbl.TABLE_NM_PASIEN & ""; Tab(70); "Type"; Spc(7); ":"; Spc(2); " " & tbl.TABLE_TYPE & " "
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Dept  "; Spc(2); ":"; Spc(5); "" & tbl.TABLE_NM_DEPT & ""; Tab(70); ""; Spc(5); ""; Spc(2); ""
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentX = .CurrentX + 500
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        If tbl.TABLE_TYPE = "Credit" Then
            .CurrentX = .CurrentX + 500
            Printer.Print Tab(15); " Kreditor "; Tab(25); "" & tbl.TABLE_NM_KREDITUR & "";
            Printer.Print ""
        End If
        
        
        If tbl.TABLE_TYPE <> "Credit" Then
        .Font.Name = "Arial Narrow"
        .Font.Size = 10
         .CurrentX = .CurrentX + 500
        Printer.Print Tab(35); " Bayar "; Tab(45); "" & Format(tbl.TABLE_MONEY, "##,###0.00") & "";
        Printer.Print ""
        End If
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(35); " Total "; Tab(45); "" & Format(tbl.TABLE_TOTAL, "##,###0.00") & "";
        Printer.Print ""
        If tbl.TABLE_TYPE <> "Credit" Then
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(35); " Kembali "; Tab(45); "" & Format(tbl.TABLE_CBACK, "##,###0.00") & "";
        Printer.Print ""
        End If
        Printer.Print ""
        Printer.Print ""
        .Font.Name = "Arial Narrow"
        .Font.Size = 7
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINNES_NAME & ""
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINESS_ADDRESS & ""
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINNES_CITY & ""
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINNES_NOTE & ""
        .EndDoc
     End With
opps:
    
End Sub

Public Sub cetak_Faktur3()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim rscetak As New Recordset
    Dim rsPlafon As New Recordset
    Dim harga, total, bayar, kembali As Double
    
    Dim rentang, jmlRentang, X As Byte
    rentang = 9
    On Error GoTo opps
     With Printer
        .Font.Name = "Arial Narrow"
        .Font.Size = 7
        .CurrentY = .CurrentY + 0 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " No Faktur"; Spc(5); ":"; Spc(5); "" & tbl.TABLE_NO_FAK & ""; Tab(70); "Tanggal "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TANGGAL & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Pasien"; Spc(2); ":"; Spc(5); "" & tbl.TABLE_KD_PASIEN & " "; Tab(70); "Jam "; Spc(7); ":"; Spc(1); "  "; Hour(Now); ":"; Minute(Now); ":"; Second(Now); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Pasien "; Spc(1); ":"; Spc(5); "" & tbl.TABLE_NM_PASIEN & ""; Tab(70); "Type"; Spc(7); ":"; Spc(2); " " & tbl.TABLE_TYPE & " "
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Dept  "; Spc(2); ":"; Spc(5); "" & tbl.TABLE_NM_DEPT & ""; Tab(70); ""; Spc(5); ""; Spc(2); ""
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentX = .CurrentX + 500
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        If tbl.TABLE_TYPE = "Credit" Then
            .CurrentX = .CurrentX + 500
            Printer.Print Tab(15); " Kreditor "; Tab(25); "" & tbl.TABLE_NM_KREDITUR & "";
            Printer.Print ""
        End If
        
        
        If tbl.TABLE_TYPE <> "Credit" Then
        .Font.Name = "Arial Narrow"
        .Font.Size = 10
         .CurrentX = .CurrentX + 500
        Printer.Print Tab(35); " Bayar "; Tab(45); "" & Format(tbl.TABLE_MONEY, "##,###0.00") & "";
        Printer.Print ""
        End If
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(35); " Total "; Tab(45); "" & Format(tbl.TABLE_TOTAL, "##,###0.00") & "";
        Printer.Print ""
        If tbl.TABLE_TYPE <> "Credit" Then
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(35); " Kembali "; Tab(45); "" & Format(tbl.TABLE_CBACK, "##,###0.00") & "";
        Printer.Print ""
        End If
        
        If tbl.TABLE_TYPE = "Credit" Then
            sql = "SELECT j.no_jual,DATE_FORMAT(tgl_jual,'%Y-%m-%d') as tgl_jual,j.jw,DATE_ADD(DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),INTERVAL + j.jw DAY) as jatuh_tempo,j.piutang,(IF(j.flag_kreditor=0,(IF (j.id_kreditor>0,(IF(DATE_ADD(DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),INTERVAL+j.jw DAY)>CURDATE(),'Piutang','Tagih')),'Lunas')) ,'Lunas'))AS statusjual FROM tbl_jual j LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor INNER JOIN tbl_cabang c ON c.id_cabang=j.id_cabang WHERE j.id_kreditor=" & tbl.TABLE_ID_KREDITUR & " AND j.flag_kreditor=0  AND j.piutang>0 "
            If rscetak.State = 1 Then rscetak.Close
            Set rscetak = New ADODB.Recordset
            rscetak.Open sql, CN, adOpenStatic, adLockReadOnly
            If rscetak.RecordCount > 0 Then
                Printer.Print ""
                .CurrentX = .CurrentX + 500
                Printer.Print "------------ Info Tagihan ----------------"
                Do While (Not rscetak.EOF)
                    Printer.Print ""
                    .CurrentX = .CurrentX + 500
                    Printer.Print " " & rscetak.Fields("no_jual") & " "; Tab(20); " " & rscetak.Fields("tgl_jual") & "  "; Tab(35); " " & rscetak.Fields("jw") & " Hari"; Tab(55); " " & Format(rscetak.Fields("piutang"), "##,###0.00") & "  "; Tab(75); " " & rscetak.Fields("statusjual") & " ";
                    rscetak.MoveNext
                Loop
                Printer.Print ""
                Printer.Print ""
            End If
            
            If rsPlafon.State = 1 Then rsPlafon.Close
            sql = "SELECT k.plafon,(SELECT SUM(j.piutang) FROM tbl_jual j WHERE j.id_kreditor=k.id_kreditor AND j.flag_kreditor=1 AND j.piutang > 0) as limitHutang,(k.plafon-(SELECT SUM(j.piutang) FROM tbl_jual j WHERE j.id_kreditor=k.id_kreditor AND j.flag_kreditor=1 AND j.piutang > 0)) as sisa FROM tbl_kreditor k WHERE k.id_kreditor=" & tbl.TABLE_ID_KREDITUR & " AND (SELECT SUM(j.piutang) FROM tbl_jual j WHERE j.id_kreditor=k.id_kreditor AND j.flag_kreditor=1 AND j.piutang > 0)>k.plafon  "
            rsPlafon.Open sql, CN, adOpenStatic, adLockReadOnly
            If (rsPlafon.RecordCount > 0) Then
                Printer.Print ""
                .CurrentX = .CurrentX + 500
                Printer.Print " Plafon "; Tab(20); " " & Format(rsPlafon.Fields("plafon"), "##,###0.00") & "  ";
                Printer.Print ""
                .CurrentX = .CurrentX + 500
                Printer.Print " Piutang "; Tab(20); " " & Format(rsPlafon.Fields("limitHutang"), "##,###0.00") & "  ";
                Printer.Print ""
                .CurrentX = .CurrentX + 500
                Printer.Print " Sisa "; Tab(20); " " & Format(rsPlafon.Fields("sisa"), "##,###0.00") & "  ";
            End If
            rsPlafon.Close
        End If
        Printer.Print ""
        Printer.Print ""
        .Font.Name = "Arial Narrow"
        .Font.Size = 7
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINNES_NAME & ""
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINESS_ADDRESS & ""
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINNES_CITY & ""
        .CurrentX = .CurrentX + 500
        Printer.Print Tab(40); "" & CurrBiz.BUSINNES_NOTE & ""
        .EndDoc
     End With
opps:
    'MsgBox "Printer Error.", "Printer", "Error Message"
End Sub

Public Sub cetak_Faktur4()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim harga, total, bayar, kembali As Double
    Dim rentang, jmlRentang, X As Byte
    Dim rsobat As New Recordset
    rentang = 9
    On Error GoTo opps
     With Printer
        .Font.Name = "Arial Narrow"
        .Font.Size = 7
        .CurrentY = .CurrentY + 0 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " No Faktur"; Spc(5); ":"; Spc(5); "" & tbl.TABLE_NO_FAK & ""; Tab(70); "Tanggal "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TANGGAL & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Pasien"; Spc(2); ":"; Spc(5); "" & tbl.TABLE_KD_PASIEN & " / Umur : " & tbl.TABLE_UMUR_PASIEN & ""; Tab(70); "Jam "; Spc(7); ":"; Spc(1); "  "; Hour(Now); ":"; Minute(Now); ":"; Second(Now); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Pasien "; Spc(1); ":"; Spc(5); "" & tbl.TABLE_NM_PASIEN & ""; Tab(70); "Telepon "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TLP_PASIEN & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Dept  "; Spc(2); ":"; Spc(5); "" & tbl.TABLE_NM_DEPT & ""; Tab(70); "Rp "; Spc(9); ":"; Spc(2); "" & Format(tbl.TABLE_TOTAL, "##,###0.00") & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentX = .CurrentX + 500
        Printer.Print " Nama Obat "; Tab(50); "Kemasan"; Tab(70); "Jumlah"; Tab(85);
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentY = .CurrentY + 0
        Set rsobat = New Recordset
        If rsobat.State = 1 Then rsobat.Close
        sql = "SELECT o.nm_obat,o.kemasan,FORMAT(d.jumlah,0) as jumlah FROM tbl_jual_details d INNER JOIN tbl_obat o ON o.id_obat=d.id_obat WHERE d.no_jual='" & tbl.TABLE_NO_FAK & "' ORDER BY o.kd_obat LIMIT 100 "
        rsobat.Open sql, CN, adOpenStatic, adLockReadOnly
        If rsobat.RecordCount > 0 Then
            Do While (Not rsobat.EOF)
                 .CurrentX = .CurrentX + 500
                Printer.Print " " & rsobat.Fields("nm_obat") & "  "; Tab(50); " " & rsobat.Fields("kemasan") & "  "; Tab(75); " " & rsobat.Fields("jumlah") & "  ";
                Printer.Print " "
                .CurrentX = .CurrentX + 500
                Printer.Print " ----------------------------------- ----------------------- -------------- ------------------- ---------------- "
                .CurrentY = .CurrentY + 1
                rsobat.MoveNext
            Loop
        End If
        rsobat.Close
        .EndDoc
     End With
opps:
End Sub

Public Sub cetak_FakturBeli()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim rscetak As New Recordset
    Dim harga, total, bayar, kembali As Double
    Dim rentang, jmlRentang, X As Byte
    rentang = 9
    On Error GoTo opps
     With Printer
        .Font.Name = "Times New Roman"
        .Font.Size = 8
        .CurrentY = .CurrentY + 200 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " No Faktur"; Spc(7); ":"; Spc(5); "" & tbl.TABLE_NO_FAK & ""; Tab(70); "Tanggal "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TANGGAL & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Supplier"; Spc(2); ":"; Spc(5); "" & tbl.TABLE_ID_SUPPLIER & ""; Tab(70); "Jam "; Spc(7); ":"; Spc(1); "  "; Hour(Now); ":"; Minute(Now); ":"; Second(Now); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Supplier "; Spc(1); ":"; Spc(5); "" & tbl.TABLE_NM_SUPPLIER & ""; Tab(70); "Telepon "; Spc(3); ":"; Spc(2); "" & tbl.TABLE_TLP_SUPPLIER & ""
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentX = .CurrentX + 500
        Printer.Print " Nama Obat "; Tab(40); "Kemasan"; Tab(60); "Harga"; Tab(70); "Jumlah"; Tab(85); "Total";
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .CurrentY = .CurrentY + 0
        'sql = "SELECT *,FORMAT((j.harga_jual * j.jumlah),0) as total FROM tbl_beli_details j INNER JOIN tbl_obat O ON O.id_obat=j.id_obat"
        
        'If rscetak.State = 1 Then rscetak.Close
        'Set rscetak = New Recordset
        'rscetak.CursorLocation = adUseClient
        'rscetak.Open sql, CN, adOpenStatic, adLockReadOnly
        Dim xx As Byte
        Dim ax As Integer
        ax = frmPurchasing.lstOrders.ListItems.Count
            
        
        If (ax > 0) Then
            
            For xx = 1 To ax
                .CurrentX = .CurrentX + 500
                Printer.Print " " & frmPurchasing.lstOrders.ListItems(xx).SubItems(2) & " "; Tab(40); " " & frmPurchasing.lstOrders.ListItems(xx).SubItems(3) & " "; Tab(55); " " & frmPurchasing.lstOrders.ListItems(xx).SubItems(4) & " "; Tab(75); " " & frmPurchasing.lstOrders.ListItems(xx).SubItems(5) & " "; Tab(85); " " & frmPurchasing.lstOrders.ListItems(xx).SubItems(6) & " ";
                Printer.Print ""
                .CurrentY = .CurrentY + 50
            Next xx
            Printer.Print ""
            .CurrentX = .CurrentX + 500
            Printer.Print " --------------------------------------------------------------------------------------------------------------- "
            Printer.Print ""
            .CurrentX = .CurrentX + 500
            Printer.Print Tab(75); " Bayar "; Tab(85); "" & Format(tbl.TABLE_TOTAL, "##,###0.00") & "";
            Printer.Print ""
            .CurrentX = .CurrentX + 500
            .EndDoc
        Else
            MsgBox "No Data"
        End If
     End With
opps:
    'MsgBox "Printer Error.", "Printer", "Error Message"
End Sub

Public Sub printStock()
    Dim rpt As ACRStock
    Set rpt = New ACRStock
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .lblTgl.Caption = "Tgl Print : " & Format(Date, "DD/MM/YYYY")
        
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
    
        sql = "SELECT *,@max:=(stok_min * 4) as max,(@max-sisa) as permintaan FROM vw_stok "
        .DataControl1.Source = sql
        .GroupHeader1.DataField = "id_kategori"
        .txtCategoryName.DataField = "nm_kategori"
        .txtKodeProduct.DataField = "kd_obat"
        .txtName.DataField = "nm_obat"
        .txtMerk.DataField = "kemasan"
        .txtGudang.DataField = "stok"
        .txtBeli.DataField = "beli"
        .txtJual.DataField = "jual"
        .txtRetur.DataField = "rugi"
        .txtStock.DataField = "sisa"
        .txtMin.DataField = "stok_min"
        .txtMax.DataField = "max"
        .txtPermintaan.DataField = "permintaan"
     End With
End Sub

Public Sub printStockMin()
    Dim rpt As ACRStock
    Set rpt = New ACRStock
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .lblTgl.Caption = "Tgl Print : " & Format(Date, "DD/MM/YYYY")
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
        sql = "select *,@max:=(stok_min*4) as max,(@max-sisa) as permintaan from vw_stok_min order by id_kategori ASC "
        .DataControl1.Source = sql
        .GroupHeader1.DataField = "id_kategori"
        .txtCategoryName.DataField = "nm_kategori"
        .txtKodeProduct.DataField = "kd_obat"
        .txtName.DataField = "nm_obat"
        .txtMerk.DataField = "kemasan"
        .txtGudang.DataField = "stok"
        .txtBeli.DataField = "beli"
        .txtJual.DataField = "jual"
        .txtRetur.DataField = "rugi"
        .txtStock.DataField = "sisa"
        .txtMin.DataField = "stok_min"
        .txtMax.DataField = "max"
        .txtPermintaan.DataField = "permintaan"
     End With
End Sub

Public Sub printPasien()
    Dim rpt As ACRPasien
    Set rpt = New ACRPasien
    Dim baw, bak As Long
    Dim gb, gk As String
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    baw = Val(frmPasien.lvList.ListItems(1).Text)
    bak = Val(frmPasien.lvList.ListItems(frmPasien.lvList.ListItems.Count).Text)
    If (baw <= bak) Then
        gb = ">="
        gk = "<="
    Else
        gb = "<="
        gk = ">="
    End If
    
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
        sql = "SELECT * FROM tbl_pasien WHERE id_pasien " & gb & " " & baw & " AND  id_pasien  " & gk & " " & bak & ""
        .DataControl1.Source = sql
        .txtKode.DataField = "kd_pasien"
        .txtName.DataField = "nm_pasien"
        .txtTglLahir.DataField = "tgl_lahir"
        .txtTmptLahir.DataField = "tmpt_lahir"
        .txtAlamat.DataField = "alamat"
        .txtHP.DataField = "no_hp"
        .txtKota.DataField = "kota"
     End With
End Sub

Public Sub printCashFlow()
    Dim rpt As ACRCashFlow
    Set rpt = New ACRCashFlow
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
         sql = "SELECT tgl_cash,(jual+jual_sebelumnya) as jual,(beli+beli_sebelumnya) as beli FROM vw_cash_flow  ORDER BY tgl_cash ASC "
        
        .DataControl1.Source = sql
        .txtDate.DataField = "tgl_cash"
        .txtPenjualan.DataField = "jual"
        .txtPembelian.DataField = "beli"
        .txtRetur.DataField = "retur"
        .txtKomisi.DataField = "komisi"
        .txtLaba.DataField = "laba"
        .txtTransfer.DataField = "cash"
        .txtKasSisa.DataField = "kas_total"
        .txtTotalPasien.DataField = "pasien"
        'grandtotal
        .txtSumPenjualan.DataField = "jual"
        .txtSumPembelian.DataField = "beli"
        .txtSumRetur.DataField = "retur"
        .txtSumKomisi.DataField = "komisi"
        .txtSumLaba.DataField = "laba"
        .txtSumTransfer.DataField = "cash"
        .txtSumKasSisa.DataField = "kas_total"
     End With
End Sub

Public Sub printCashFlowdetails()
    Dim rpt As ACRCashFlowDetails
    Set rpt = New ACRCashFlowDetails
    DBPath = "DSN=" + CurrUser.User_DSN + ""

    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .lblTanggal.Caption = "Dari Tanggal " & tbl.TABLE_TANGGAL_AWAL & " Sampai " & tbl.TABLE_TANGGAL_AKHIR
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
         sql = "SELECT *,vf.tgl_cash,(vf.jual+vf.jual_sebelumnya) as total_jual,(vf.beli+vf.beli_sebelumnya) as total_beli,vk.komisi as komisidep,vk.pasien as vkpasien FROM vw_cash_flow vf "
         sql = sql & " LEFT JOIN vw_komisi vk ON vk.tgl_jual=vf.tgl_cash WHERE MONTH(vf.tgl_cash)=MONTH(CURDATE()) "
         If tbl.TABLE_TANGGAL_AWAL <> "" Then
            sql = sql & " AND vf.tgl_cash >='" & tbl.TABLE_TANGGAL_AWAL & "' AND vf.tgl_cash <='" & tbl.TABLE_TANGGAL_AKHIR & "'  "
            tbl.TABLE_TANGGAL_AWAL = ""
            tbl.TABLE_TANGGAL_AKHIR = ""
         End If
         sql = sql & " ORDER BY vf.id ASC "
        .DataControl1.Source = sql
        .GroupHeader1.DataField = "tgl_cash"
        .txtDate.DataField = "tgl_cash"
        
        .txtLabaBerjalan.DataField = "jual"
        .txtPelunasanPiutang.DataField = "jual_sebelumnya"
        .txtPenjualan.DataField = "total_jual"
        
        .txtPengeluaran.DataField = "beli"
        .txtByarHutang.DataField = "beli_sebelumnya"
        .txtPembelian.DataField = "total_beli"
        
        
        .txtRetur.DataField = "retur"
        .txtKomisi.DataField = "komisi"
        .txtLaba.DataField = "laba"
        .txtTransfer.DataField = "cash"
        .txtKasSisa.DataField = "kas_total"
        .txtTotalPasien.DataField = "pasien"
        .txtByr.DataField = "lunas"
        
        'details
        .txtKdDep.DataField = "kd_departement"
        .txtNmDep.DataField = "nm_departement"
        .txtKomisiDep.DataField = "komisidep"
        .txtPasienDep.DataField = "vkpasien"
        'grandtotal
        .txtSumLabaBerjalan.DataField = "jual"
        .txtSumLabaBerjalan.SummaryDistinctField = "tgl_cash"
        .txtSUMPelunasanPiutang.DataField = "jual_sebelumnya"
        .txtSUMPelunasanPiutang.SummaryDistinctField = "tgl_cash"
        .txtSumPenjualan.DataField = "total_jual"
        .txtSumPenjualan.SummaryDistinctField = "tgl_cash"
        
        .txtSUMPengeluaran.txtSumPembelian.DataField = "beli"
        .txtSUMPengeluaran.SummaryDistinctField = "tgl_cash"
        .txtSUMBayarHutang.DataField = "beli_sebelumnya"
        .txtSUMBayarHutang.SummaryDistinctField = "tgl_cash"
        .txtSumPembelian.DataField = "total_beli"
        .txtSumPembelian.SummaryDistinctField = "tgl_cash"
        
        .txtSumRetur.DataField = "retur"
        .txtSumRetur.SummaryDistinctField = "tgl_cash"
        .txtSumKomisi.DataField = "komisi"
        .txtSumKomisi.SummaryDistinctField = "tgl_cash"
        .txtSumLaba.DataField = "laba"
        .txtSumLaba.SummaryDistinctField = "tgl_cash"
        .txtSumTransfer.DataField = "cash"
        .txtSumTransfer.SummaryDistinctField = "tgl_cash"
        .txtSumKasSisa.DataField = "kas_total"
        .txtSumKasSisa.SummaryDistinctField = "tgl_cash"
        .txtSumPasien.DataField = "pasien"
        .txtSumPasien.SummaryDistinctField = "tgl_cash"
        .txtSumBYr.DataField = "lunas"
        .txtSumBYr.SummaryDistinctField = "tgl_cash"
        .show
     End With
End Sub

Public Sub printSalesSummary()
    Dim rpt As ACRSalesSummary
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    Set rpt = New ACRSalesSummary
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
        .lblTanggal.Caption = "Dari Tanggal " & tbl.TABLE_TANGGAL_AWAL & " Sampai " & tbl.TABLE_TANGGAL_AKHIR
        
        sql = " SELECT j.no_jual,j.tgl_jual,(IF(j.flag_kreditor=1,(IF (j.id_kreditor>0,(IF(DATE_ADD(DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),INTERVAL + j.jw DAY)>CURDATE(),'Piutang','Jatuh Tempo')),'Lunas')) ,'Lunas'))AS status,j.tgl_bayar,j.kd_pasien,p.nm_pasien,k.nm_kreditor,d.kd_departement,d.nm_departement,c.nm_cabang,j.flag_kreditor,j.flag_debitor,j.bayar,j.piutang,j.komisi,(j.bayar-j.komisi) as total "
        sql = sql + " From "
        sql = sql + " tbl_jual j"
        sql = sql + " INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien"
        sql = sql + " LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor"
        sql = sql + " INNER JOIN tbl_departement d ON d.id_departement=j.id_departement"
        sql = sql + " INNER JOIN tbl_cabang c ON c.id_cabang=j.id_cabang"
        sql = sql + " WHERE j.no_jual<>'' "
        
        If (tbl.TABLE_ID_KREDITUR <> "") Then
            sql = sql + " AND j.id_kreditor= " & tbl.TABLE_ID_KREDITUR & " AND j.flag_kreditor=1 AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')< CURDATE() "
        End If
        
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If
        'MsgBox "Print Data Sales !", vbOKOnly + vbInformation
         
        .DataControl1.Source = sql
        .GroupHeader1.DataField = "tgl_bayar"
        .txtDate.DataField = "tgl_bayar"
        .txtFak.DataField = "no_jual"
        .txtTglBayar.DataField = "tgl_jual"
        .txtCustID.DataField = "kd_pasien"
        .txtCustomerName.DataField = "nm_pasien"
        .txtKreditor.DataField = "nm_kreditor"
        .txtNmDep.DataField = "nm_departement"
        .txtKas.DataField = "bayar"
        .txtPiutang.DataField = "piutang"
        .txtKasTotal.DataField = "total"
        'Group 1
        .txtSubKas.DataField = "bayar"
        .txtSubPiutang.DataField = "piutang"
        .txtSubTotal.DataField = "total"
        'All
        .txtGrandKas.DataField = "bayar"
        .txtGrandPiutang.DataField = "piutang"
        .txtGrandTotal.DataField = "total"
        
        .lblTgl1.Caption = " Tgl." & Format(Date, "DD/MM/YYYY")
        .lblTgl2.Caption = " Tgl." & Format(Date, "DD/MM/YYYY")
    End With
End Sub


Public Sub printSalesCommision()
    Dim rpt As ACRListKomisi
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    Set rpt = New ACRListKomisi
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
         
        .lblTanggal.Caption = "Dari Tanggal " & tbl.TABLE_TANGGAL_AWAL & " Sampai " & tbl.TABLE_TANGGAL_AKHIR
        
        sql = "SELECT j.no_jual,j.tgl_jual,(IF(j.flag_kreditor=1,(IF (j.id_kreditor>0,(IF(DATE_ADD(DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),INTERVAL + j.jw DAY)>CURDATE(),'Piutang','Jatuh Tempo')),'Lunas')) ,'Lunas'))AS status,j.tgl_komisi,j.kd_pasien,p.nm_pasien,k.nm_kreditor,d.kd_departement,d.nm_departement,c.nm_cabang,j.flag_kreditor,j.flag_debitor,j.bayar,j.piutang,j.komisi,(j.bayar-j.komisi) as total "
        sql = sql & " FROM tbl_jual j "
        sql = sql & " INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien"
        sql = sql & " LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor "
        sql = sql & " INNER JOIN tbl_departement d ON d.id_departement=j.id_departement "
        sql = sql & " INNER JOIN tbl_cabang c ON c.id_cabang=j.id_cabang "
        sql = sql & " INNER JOIN tbl_pengguna pp ON pp.id=j.id_pengguna "
        sql = sql & " WHERE j.no_jual<>'' AND j.flag_kreditor=0 "
        
        If (tbl.TABLE_ID_DEPT <> "") Then
            sql = sql + " AND j.id_departement= " & tbl.TABLE_ID_DEPT & " AND j.flag_debitor=1  AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= CURDATE() "
        End If
        
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If
        
        sql = sql & " ORDER BY j.tgl_komisi ASC,j.no_jual ASC,j.kd_pasien ASC " 'DATE_FORMAT(j.tgl_jual,'%Y-%m-%d') ASC
        
        .DataControl1.Source = sql
        .GroupHeader1.DataField = "tgl_komisi"
        .txtDate.DataField = "tgl_komisi"
        .txtFak.DataField = "no_jual"
        .txtTglBayar.DataField = "tgl_jual"
        .txtCustID.DataField = "kd_pasien"
        .txtCustomerName.DataField = "nm_pasien"
        .txtKreditor.DataField = "nm_kreditor"
        .txtNmDep.DataField = "nm_departement"
        .txtCabang.DataField = "nm_cabang"
        .txtKomisi.DataField = "komisi"
        .txtStatus.DataField = "status"
        'Group 1
        .txtSubKomisi.DataField = "komisi"
        'All
        .txtGrandKomisi.DataField = "komisi"
    End With
End Sub

Public Sub printPurchaseSummary()
    Dim rpt As ACRPurchaseSummary
    Set rpt = New ACRPurchaseSummary
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
         
        sql = " SELECT b.no_beli,b.tgl_beli,b.tgl_bayar,b.id_supplier,s.nm_supplier,b.payment,b.bayar,b.hutang,(b.bayar-b.hutang) as sisa "
        sql = sql + " FROM tbl_beli b"
        sql = sql + " INNER JOIN tbl_supplier s ON s.id_supplier=b.id_supplier "
        sql = sql + " WHERE b.no_beli <> '' "
        
        If (tbl.TABLE_ID_SUPPLIER <> "") Then
            sql = sql + " AND b.flag_supplier=1 AND b.id_supplier = " & Trim(tbl.TABLE_ID_SUPPLIER) & " "
        End If
        
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If

        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If
        
        .DataControl1.Source = sql
        .GroupHeader1.DataField = "tgl_bayar"
        .txtDate.DataField = "tgl_bayar"
        .txtFak.DataField = "no_beli"
        .txtTglBayar.DataField = "tgl_beli"
        .txtSuppID.DataField = "id_supplier"
        .txtSupplierName.DataField = "nm_supplier"
        .txtPaymentType.DataField = "payment"
        .txtBayar.DataField = "bayar"
        .txtHutang.DataField = "hutang"
        .txtSisaHutang.DataField = "sisa"
        'Group 1
        .txtSubBayar.DataField = "bayar"
        .txtSubHutang.DataField = "hutang"
        .txtSubSisa.DataField = "sisa"
        'All
        .txtGrandBayar.DataField = "bayar"
        .txtGrandHutang.DataField = "hutang"
        .txtGrandSisa.DataField = "sisa"
    End With
End Sub

Public Sub printPurchaseDetails()
    Dim rpt As ACRPurchaseDetails
    Set rpt = New ACRPurchaseDetails
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
         
        sql = " SELECT *,(d.harga_beli * d.jumlah) as total "
        sql = sql + " FROM tbl_beli b"
        sql = sql + " INNER JOIN tbl_supplier s ON s.id_supplier=b.id_supplier "
        sql = sql + " INNER JOIN tbl_beli_details d ON d.no_beli=b.no_beli "
        sql = sql + " INNER JOIN tbl_obat o ON o.id_obat=d.id_obat "
        sql = sql + " WHERE b.no_beli <> '' "
        
        If (tbl.TABLE_ID_SUPPLIER <> "") Then
            sql = sql + " AND b.flag_supplier=1 AND b.id_supplier = " & Trim(tbl.TABLE_ID_SUPPLIER) & " "
        End If
        
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If

        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If
        
        
        .DataControl1.Source = sql
        .GroupHeader1.DataField = "kd_obat"
        .txtKodeProduct.DataField = "kd_obat"
        .txtName.DataField = "nm_obat"
        .txtIlmiah.DataField = "nm_ilmiah"
        .txtKemasan.DataField = "kemasan"
        
        .txtKdFaktur.DataField = "no_beli"
        .txtTanggal.DataField = "tgl_beli"
        .txtFlag.DataField = "flag_supplier"
        .txtKode.DataField = "id_supplier"
        .txtNmSupplier.DataField = "nm_supplier"
        .txtHargaBeli.DataField = "harga_beli"
        .txtQty.DataField = "jumlah"
        .txtTotal.DataField = "total"
    
        'All
        .txtGrandQty.DataField = "jumlah"
        .txtGrandTotal.DataField = "total"
    End With
End Sub

Public Sub printRetur()
    Dim rpt As ACRRetur
    Set rpt = New ACRRetur
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        
        .ADO1.CursorLocation = ddADOUseClient
        .ADO1.ConnectionString = DBPath
        sql = "SELECT *,(d.harga_beli*d.retur)as total "
        sql = sql & " FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli "
        sql = sql & " INNER JOIN tbl_obat o ON o.id_obat=d.id_obat INNER JOIN tbl_kategori k ON k.id_kategori=o.id_kategori "
        sql = sql & " WHERE d.retur > 0 "
        sql = sql & " ORDER BY o.id_kategori,o.kd_obat ASC"
        .ADO1.Source = sql
        .txtKdFaktur.DataField = "no_beli"
        .txtTanggal.DataField = "tgl_retur"
        .txtKodeProduct.DataField = "kd_obat"
        .txtName.DataField = "nm_obat"
        .txtMerk.DataField = "kemasan"
        .txtHargaBeli.DataField = "harga_beli"
        .txtRetur.DataField = "retur"
        .txtTotal.DataField = "total"
        .txtSumHarga.DataField = "harga_beli"
        .txtSumRetur.DataField = "retur"
        .txtSumTotal.DataField = "total"
     End With
End Sub

Public Sub printInvoice()
    Dim rpt As ACRInvoice
    Set rpt = New ACRInvoice
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    MDIMainMenu.HideTBButton "", True
    With rpt
        .lblname.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblkota.Caption = CurrBiz.BUSINNES_CITY
        .lblTgl.Caption = Format(Date, "DD/MM/YYYY")
        
        .lblNamaKreditor.Caption = tbl.TABLE_NM_KREDITUR
        .lblAlamatKreditor.Caption = tbl.TABLE_ALMT_KREDITUR
        .lblKotaKreditor.Caption = tbl.TABLE_KOTA_KREDITUR
        .lblTlpKreditor.Caption = tbl.TABLE_TLP_KREDITOR
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
        sql = " SELECT j.no_jual,j.tgl_jual,j.kd_pasien,p.nm_pasien,d.nm_departement,(j.piutang+j.bayar) as piutang"
        sql = sql + " From "
        sql = sql + " tbl_jual j"
        sql = sql + " INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien"
        sql = sql + " LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor"
        sql = sql + " INNER JOIN tbl_departement d ON d.id_departement=j.id_departement"
        sql = sql + " WHERE j.id_kreditor= " & tbl.TABLE_ID_KREDITUR & " AND j.flag_kreditor=1 AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')< CURDATE() "
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If
        .DataControl1.Source = sql
        .txtKdFaktur.DataField = "no_jual"
        .txtTanggal.DataField = "tgl_jual"
        .txtKodePasien.DataField = "kd_pasien"
        .txtNamaPasien.DataField = "nm_pasien"
        .txtNmDepartement.DataField = "nm_departement"
        .txtTotal.DataField = "piutang"
        .txtSumTotal.DataField = "piutang"
        .txtTotalPasien.DataField = "kd_pasien"
     End With
     'tbl.TABLE_TANGGAL_AWAL = ""
     'tbl.TABLE_TANGGAL_AKHIR = ""
     MsgBox "Print Invoice Pendapatan, Akan Melunasi Seluruh Faktur yang telah ada di Invoice !  ", vbCritical + vbInformation
End Sub

Public Sub Invoice_lunas()
    Dim total As Integer
    Dim bayar As Double
    Dim rsJual As New Recordset
    
    sql = " SELECT j.id_jual,j.no_jual,j.tgl_jual,j.kd_pasien,p.nm_pasien,d.nm_departement,(j.piutang+j.bayar) as piutang"
    sql = sql + " From "
    sql = sql + " tbl_jual j"
    sql = sql + " INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien"
    sql = sql + " LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor"
    sql = sql + " INNER JOIN tbl_departement d ON d.id_departement=j.id_departement"
    sql = sql + " WHERE j.id_kreditor= " & tbl.TABLE_ID_KREDITUR & " AND j.flag_kreditor=1 AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')< CURDATE() "
    If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
    End If
    If rsJual.State = 1 Then rsJual.Close
    Set rsJual = New ADODB.Recordset
    rsJual.CursorLocation = adUseClient
    rsJual.Open sql, CN, adOpenStatic, adLockReadOnly
    
    bayar = 0
    total = 0
    Do While Not rsJual.EOF
        sql = "UPDATE tbl_jual "
        sql = sql + "SET "
        sql = sql + " tgl_bayar='" & Format(Date, "YYYY-MM-DD") & "',"
        sql = sql + " payment='Lunas', "
        sql = sql + " flag_kreditor= 0 , "
        sql = sql + " bayar=" & Format(rsJual.Fields("piutang"), "") & ", "
        sql = sql + " dibayar=" & Format(rsJual.Fields("piutang"), "") & ", "
        sql = sql + " piutang= 0  "
        sql = sql + " WHERE id_jual=" & Format(rsJual.Fields("id_jual"), "") & ""
        bayar = bayar + Val(Format(Format(rsJual.Fields("piutang"), "")))
        CN.Execute sql
        rsJual.MoveNext
        total = total + 1
    Loop
            
    rsJual.MoveFirst
    tbl.TABLE_TANGGAL_AWAL = rsJual.Fields("tgl_jual")
    rsJual.MoveLast
    tbl.TABLE_TANGGAL_AKHIR = rsJual.Fields("tgl_jual")
    tbl.TABLE_TOTAL = bayar
    tbl.TABLE_TOTAL_PASIEN = total
End Sub

Public Sub cetak_Invoice()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim harga, total, bayar, kembali As Double
    On Error GoTo opps
     With Printer
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .CurrentY = .CurrentY + 200 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Pelunasan Penjualan"; Spc(5); " "; Spc(5); ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal Awal "; Spc(6); ":"; Spc(5); "" & tbl.TABLE_TANGGAL_AWAL & ""
         .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal Akhir "; Spc(5); ":"; Spc(5); "" & tbl.TABLE_TANGGAL_AKHIR & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Kreditor "; Spc(5); ":"; Spc(5); "" & tbl.TABLE_ID_KREDITUR & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Kreditor  "; Spc(4); ":"; Spc(5); "" & tbl.TABLE_NM_KREDITUR & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Total Pasien  "; Spc(7); ":"; Spc(5); "" & tbl.TABLE_TOTAL_PASIEN & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Total Uang"; Spc(9); ":"; Spc(5); "" & Format(tbl.TABLE_TOTAL, "##,###0") & ""
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .EndDoc
     End With
opps:
     
End Sub

Public Sub printInvoicePembelian()
    Dim rpt As ACRInvoicePembelian
    Set rpt = New ACRInvoicePembelian
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    MDIMainMenu.HideTBButton "", True
    With rpt
        .lblname.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblkota.Caption = CurrBiz.BUSINNES_CITY
        .lblTgl.Caption = Format(Date, "DD/MM/YYYY")
        
        .lblNamaSupplier.Caption = tbl.TABLE_NM_SUPPLIER
        .lblAlamatSupplier.Caption = tbl.TABLE_ALMT_SUPPLIER
        .lblKotaSupplier.Caption = tbl.TABLE_KOTA_SUPPLIER
        .lblTlpSupplier.Caption = tbl.TABLE_TLP_SUPPLIER
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
        
        sql = " SELECT b.no_beli,b.tgl_beli,o.kd_obat,o.nm_obat,d.harga_beli,d.jumlah,(d.harga_beli * d.jumlah) as total"
        sql = sql + " FROM tbl_beli b"
        sql = sql + " INNER JOIN tbl_beli_details d ON d.no_beli=b.no_beli"
        sql = sql + " INNER JOIN tbl_obat o ON o.id_obat=d.id_obat"
        sql = sql + " WHERE b.flag_supplier=1 AND b.id_supplier = " & Trim(tbl.TABLE_ID_SUPPLIER) & " "
        
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(b.tgl_beli,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If

        .DataControl1.Source = sql
        .txtKdFaktur.DataField = "no_beli"
        .txtTanggal.DataField = "tgl_beli"
        .txtKodeObat.DataField = "kd_obat"
        .txtNamaMerk.DataField = "nm_obat"
        .txtHargaBeli.DataField = "harga_beli"
        .txtJumlah.DataField = "jumlah"
        .txtTotal.DataField = "total"
        .txtTotalObat.DataField = "jumlah"
        .txtSumTotal.DataField = "total"
     End With
     MsgBox "Print Invoice Pembelian, Akan Melunasi Seluruh Hutang yang telah ada di Invoice !  ", vbCritical + vbInformation
End Sub

Public Sub cetak_Invoice_Pembelian()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim harga, total, bayar, kembali As Double
    On Error GoTo opps
     With Printer
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .CurrentY = .CurrentY + 200 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Pelunasan Pembelian"; Spc(5); " "; Spc(5); ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal Awal "; Spc(6); ":"; Spc(5); "" & tbl.TABLE_TANGGAL_AWAL & ""
         .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal Akhir "; Spc(5); ":"; Spc(5); "" & tbl.TABLE_TANGGAL_AKHIR & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Supplier "; Spc(5); ":"; Spc(5); "" & tbl.TABLE_ID_SUPPLIER & ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Supplier  "; Spc(4); ":"; Spc(5); "" & tbl.TABLE_NM_SUPPLIER & ""; Tab(70); " "; Spc(5); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Total Obat"; Spc(10); ":"; Spc(5); "" & Format(tbl.TABLE_TOTAL_OBAT, "#,##0") & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Total Uang"; Spc(9); ":"; Spc(5); "" & Format(tbl.TABLE_TOTAL, "#,##0") & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        Printer.Print " --------------------------------------------------------------------------------------------------------------- "
        .EndDoc
     End With
     tbl.TABLE_TANGGAL_AWAL = ""
     tbl.TABLE_TANGGAL_AKHIR = ""
opps:

End Sub

Public Sub printKomisi()
    Dim rpt As ACRCommision
    Set rpt = New ACRCommision
    DBPath = "DSN=" + CurrUser.User_DSN + ""

    MDIMainMenu.HideTBButton "", True
    With rpt
        .lblname.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblkota.Caption = CurrBiz.BUSINNES_CITY
        .lblKdDebitor.Caption = tbl.TABLE_KD_DEPT
        .lblNamaDebitor.Caption = tbl.TABLE_NM_DEPT
        .lblTgl.Caption = Format(Date, "DD/MM/YYYY")
        
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
        sql = " SELECT j.no_jual,j.tgl_jual,j.kd_pasien,p.nm_pasien,k.nm_kreditor,j.komisi"
        sql = sql + " From "
        sql = sql + " tbl_jual j"
        sql = sql + " INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien"
        sql = sql + " LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor"
        sql = sql + " INNER JOIN tbl_departement d ON d.id_departement=j.id_departement"
        sql = sql + " WHERE j.id_departement= " & tbl.TABLE_ID_DEPT & " AND j.flag_kreditor=0 AND j.flag_debitor=1  AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= CURDATE() "
        
        If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
            sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
        End If
        
        .DataControl1.Source = sql
        .txtKdFaktur.DataField = "no_jual"
        .txtTanggal.DataField = "tgl_jual"
        .txtKodePasien.DataField = "kd_pasien"
        .txtNamaPasien.DataField = "nm_pasien"
        .txtNmKreditor.DataField = "nm_kreditor"
        .txtTotal.DataField = "komisi"
        .txtTotalPasien.DataField = "kd_pasien"
        .txtSumTotal.DataField = "komisi"
     End With
     
     'tbl.TABLE_TANGGAL_AWAL = ""
     'tbl.TABLE_TANGGAL_AKHIR = ""
     MsgBox "Print Invoice Komisi, Akan Melunasi Seluruh Faktur yang telah ada di Invoice !  ", vbCritical + vbInformation
End Sub

Public Sub LunasKomisi()
    Dim total As Integer
    Dim bayar As Double
    Dim rsKom As New Recordset
     
    sql = " SELECT j.no_jual,j.tgl_jual,j.kd_pasien,p.nm_pasien,k.nm_kreditor,j.komisi"
    sql = sql + " From "
    sql = sql + " tbl_jual j"
    sql = sql + " INNER JOIN tbl_pasien p ON p.kd_pasien=j.kd_pasien"
    sql = sql + " LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor"
    sql = sql + " INNER JOIN tbl_departement d ON d.id_departement=j.id_departement"
    sql = sql + " WHERE j.id_departement= " & tbl.TABLE_ID_DEPT & " AND j.flag_kreditor=0 AND j.flag_debitor=1  AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= CURDATE() "
    
    If ((tbl.TABLE_TANGGAL_AWAL <> "") And (tbl.TABLE_TANGGAL_AKHIR <> "")) Then
        sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')>= '" & tbl.TABLE_TANGGAL_AWAL & "' "
        sql = sql + " AND DATE_FORMAT(j.tgl_jual,'%Y-%m-%d')<= '" & tbl.TABLE_TANGGAL_AKHIR & "' "
    End If
    
    Set rsKom = New Recordset
    If rsKom.State = 1 Then rsKom.Close
    rsKom.Open sql, CN, adOpenStatic, adLockReadOnly
    total = 0
    bayar = 0
    total = 0
    Do While Not rsKom.EOF
        sql = "UPDATE tbl_jual "
        sql = sql + "SET "
        sql = sql + " flag_debitor= 0,  "
        sql = sql + " tgl_komisi= '" & Format(Date, "YYYY-MM-DD") & "'"
        sql = sql + " WHERE no_jual='" & rsKom.Fields("no_jual") & "'"
        CN.Execute sql
        bayar = bayar + Val(Format(rsKom.Fields("komisi"), ""))
        total = total + Val(Format(rsKom.Fields("kd_pasien"), ""))
        rsKom.MoveNext
        total = total + 1
    Loop
    
    rsKom.MoveFirst
    tbl.TABLE_TANGGAL_AWAL = rsKom.Fields("tgl_jual")
    rsKom.MoveLast
    tbl.TABLE_TANGGAL_AKHIR = rsKom.Fields("tgl_jual")
    tbl.TABLE_TOTAL = bayar
    tbl.TABLE_TOTAL_PASIEN = total
End Sub
    
Public Sub cetak_Invoice_Komisi()
    Dim Lines As Integer, Y As Long, OutStr As String
    Dim harga, total, bayar, kembali As Double
    On Error GoTo opps
     With Printer
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .CurrentY = .CurrentY + 200 ' Skip some space
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Pelunasan Komisi"; Spc(5); " "; Spc(5); ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal Awal "; Spc(6); ":"; Spc(5); "" & tbl.TABLE_TANGGAL_AWAL & ""
         .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Tanggal Akhir "; Spc(5); ":"; Spc(5); "" & tbl.TABLE_TANGGAL_AKHIR & ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Kode Departement "; Spc(1); ":"; Spc(5); "" & tbl.TABLE_KD_DEPT & ""; Tab(70); ""; Spc(3); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Nama Departement  "; Spc(0); ":"; Spc(5); "" & tbl.TABLE_NM_DEPT & ""; Tab(70); " "; Spc(5); ""; Spc(2); ""
        .CurrentX = .CurrentX + 500 ' Skip some space
        'Format(tbl.TABLE_KAS_SISA, "##,###0")
        Printer.Print " Total Pasien"; Spc(8); ":"; Spc(5); "" & Format(tbl.TABLE_TOTAL_PASIEN, "##,###0") & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        .CurrentX = .CurrentX + 500 ' Skip some space
        Printer.Print " Total Komisi"; Spc(7); ":"; Spc(5); "" & Format(tbl.TABLE_TOTAL, "##,###0") & ""; Tab(70); ""; Spc(8); ""; Spc(9); " "
        Printer.Print ""
        .CurrentX = .CurrentX + 500
        .EndDoc
     End With
     With tbl
        .TABLE_TANGGAL_AWAL = Empty
        .TABLE_TANGGAL_AKHIR = Empty
        .TABLE_KD_DEPT = Empty
        .TABLE_NM_DEPT = Empty
        .TABLE_TOTAL_PASIEN = Empty
        .TABLE_TOTAL = Empty
     End With
opps:
End Sub

Public Sub printStockOpname()
    Dim rpt As ACRStockOpname
    Set rpt = New ACRStockOpname
    DBPath = "DSN=" + CurrUser.User_DSN + ""
    With rpt
        .lblNama.Caption = CurrBiz.BUSINNES_NAME
        .lblALamat.Caption = CurrBiz.BUSINESS_ADDRESS
        .lblCity.Caption = CurrBiz.BUSINNES_CITY
        .lblTelepon.Caption = CurrBiz.BUSINESS_CONTACT_INFO
        .lblTgl.Caption = "Tgl Print : " & Format(Date, "DD/MM/YYYY")
        
        .DataControl1.CursorLocation = ddADOUseClient
        .DataControl1.ConnectionString = DBPath
        
        sql = "SELECT *,DATE_FORMAT(op.tgl_input,'%d-%m-%Y') as dateop,DATE_FORMAT(op.tgl_input,'%H:%i:%s') as timeop,@sf:=((o.box_sedang * op.kem_sedang)+(o.box_kecil*op.kem_kecil)+op.satuan) as sf,@s:=(@sf-op.stok_sblm) as sisa,(@s*o.harga_beli) as total,p.nm_pengguna "
        sql = sql & " FROM tbl_opname op "
        sql = sql & " INNER JOIN tbl_obat o ON o.id_obat=op.id_obat "
        sql = sql & " INNER JOIN tbl_pengguna p ON p.id=op.id_pengguna "
        sql = sql & " WHERE op.flag_opname=1 "
        sql = sql & " ORDER BY DATE_FORMAT(op.tgl_input,'%d-%m-%Y') ASC,DATE_FORMAT(op.tgl_input,'%H:%i:%s') ASC,op.id_obat ASC "
        
        .DataControl1.Source = sql
        .GH.DataField = "dateop"
        .txtTanggal.DataField = "dateop"
        .txtJam.DataField = "timeop"
        .txtKode.DataField = "kd_obat"
        .txtnama.DataField = "nm_obat"
        .txtHargaBeli.DataField = "harga_beli"
        .txtStokDatabase.DataField = "stok_sblm"
        .txtStokFisik.DataField = "sf"
        .txtStokSisa.DataField = "sisa"
        .txtTotalKerugian.DataField = "total"
        .txtGrandTotal.DataField = "total"
        .txtNmPemeriksa.DataField = "nm_pengguna"
        .show
     End With
End Sub
