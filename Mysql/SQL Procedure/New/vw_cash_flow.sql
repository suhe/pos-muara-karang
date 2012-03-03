CREATE
    /*[ALGORITHM = {UNDEFINED | MERGE | TEMPTABLE}]
    [DEFINER = { user | CURRENT_USER }]
    [SQL SECURITY { DEFINER | INVOKER }]*/
    VIEW `pos_db`.`vw_cash_flow` 
    AS
(
SELECT 
c.id,c.tgl_cash,c.money_cash,
/* Jual Hari Sebelumnya Lunas Hari Ini */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS jual_sebelumnya,
/* Jual Hari Sebelumnya Lunas Hari Ini */

/* Jual Hari Ini Lunas Hari Ini */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS jual,
/* Jual Hari Ini Lunas Hari Ini */

/* Beli Hari Sebelumnya Lunas Hari Ini */ 
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS beli_sebelumnya,
 /* Beli Hari Sebelumnya Lunas Hari Ini */

/* Beli Hari ini Lunas Hari Ini */
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=b.tgl_bayar AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS beli,
/* Beli Hari ini Lunas Hari Ini */

(IF((SELECT COUNT(*) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier=0 AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(d.retur*d.harga_beli) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier=0 AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS retur,  
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.komisi) FROM tbl_jual j  WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS komisi,

(
(
/* Laba Hari Ini */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Laba Hari Ini */

+
/* Laba Hari lalu */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Laba Hari lalu */

)
 -

(


(
/* Beli Hari ini */
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=b.tgl_bayar AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Beli Hari Ini */
+
/* Beli Hari Lalu */
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Beli Hari Lalu */
)


+ 

(IF((SELECT COUNT(*) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier=0 AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(d.retur*d.harga_beli) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))+
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.komisi) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
)
)
 AS laba ,
c.cash,
((
(


(
/* Laba Hari Ini */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Laba Hari Ini */
+
/* Laba Hari lalu */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Laba Hari lalu */
)

 -
(

(
/* Beli Hari ini */
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=b.tgl_bayar AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Beli Hari Ini */
+
/* Beli Hari Lalu */
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Beli Hari Lalu */
)
+ 
(IF((SELECT COUNT(*) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier=0 AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(d.retur*d.harga_beli) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier=0 AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))+
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.komisi) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
)
) - c.cash )) AS kas,

((
(

(
/* Laba Hari Ini */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=j.tgl_bayar AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Laba Hari Ini */
+
/* Laba Hari lalu */
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.bayar+j.piutang) FROM tbl_jual j WHERE j.flag_kreditor=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(j.tgl_input,'%Y-%m-%d') AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Laba Hari lalu */
)
 -
(


(
/* Beli Hari ini */
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=b.tgl_bayar AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(b.tgl_input,'%Y-%m-%d')=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Beli Hari Ini */
+
/* Beli Hari Lalu */
(IF((SELECT COUNT(*) FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(b.bayar+b.hutang)  FROM tbl_beli b WHERE b.flag_supplier=0 AND DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')<>DATE_FORMAT(b.tgl_input,'%Y-%m-%d') AND b.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
/* Beli Hari Lalu */
)

+ 
(IF((SELECT COUNT(*) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier=0 AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(d.retur*d.harga_beli) FROM tbl_beli b JOIN tbl_beli_details d ON d.no_beli=b.no_beli WHERE b.flag_supplier=0 AND d.tgl_retur=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))+
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT SUM(j.komisi) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor AND j.tgl_komisi=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0))
)
) - c.cash )+ c.money_cash) AS kas_total,
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT COUNT(j.kd_pasien) FROM tbl_jual j WHERE  DATE_FORMAT(j.tgl_input,'%Y-%m-%d')=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS pasien
,
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT COUNT(j.kd_pasien) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS lunas
,
(IF((SELECT COUNT(*) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d'))>0,(SELECT COUNT(j.kd_pasien) FROM tbl_jual j WHERE j.flag_kreditor=0 AND j.flag_debitor=0 AND j.tgl_bayar=DATE_FORMAT(c.tgl_cash,'%Y-%m-%d')),0)) AS pay_komisi
,c.flag
FROM tbl_cash c

);

SELECT * FROM vw_cash_flow


