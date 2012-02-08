
CREATE
    /*[ALGORITHM = {UNDEFINED | MERGE | TEMPTABLE}]
    [DEFINER = { user | CURRENT_USER }]
    [SQL SECURITY { DEFINER | INVOKER }]*/
    VIEW `pos_db`.`vw_stok_min` 
    AS
(
SELECT o.id_obat,o.kd_obat ,o.nm_obat,o.id_kategori,k.nm_kategori,o.kemasan,o.harga_jual,o.harga_beli,(o.harga_jual-o.harga_beli)AS profit,o.stok,
(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0)) AS beli,
(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0)) AS jual,
(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0)) AS rugi,
((IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))-
(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0))+ (o.stok) - 
(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))
) AS sisa,o.stok_min,
o.tgl_input,p.nm_pengguna
FROM tbl_obat o 
LEFT JOIN tbl_kategori k ON k.id_kategori=o.id_kategori 
LEFT JOIN tbl_pengguna p ON p.id=o.id_pengguna
WHERE o.stok_min >= 
((IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))-
(IF((SELECT COUNT(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat)>0,(SELECT SUM(j.jumlah) FROM tbl_jual_details j WHERE j.id_obat=o.id_obat),0))+ (o.stok) - 
(IF((SELECT COUNT(b.jumlah) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat)>0,(SELECT SUM(b.retur) FROM tbl_beli_details b WHERE b.id_obat=o.id_obat),0))
)
);
