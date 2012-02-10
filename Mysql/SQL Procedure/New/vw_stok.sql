CREATE
    /*[ALGORITHM = {UNDEFINED | MERGE | TEMPTABLE}]
    [DEFINER = { user | CURRENT_USER }]
    [SQL SECURITY { DEFINER | INVOKER }]*/
    VIEW `pos_db`.`vw_stok` 
    AS
    
(
SELECT o.id_obat,o.kd_obat,o.nm_obat,kemasan,o.box_sedang,o.box_kecil,o.harga_jual,o.harga_beli,(o.harga_jual-o.harga_beli)AS profit,o.stok,o.flag_opname,
(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0)) AS beli,
(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0)) AS jual,
(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0)) AS rugi,
((IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))-
(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0))+ (o.stok) - 
(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))
) AS sisa,o.stok_min,
o.tgl_input
FROM tbl_obat o
) 