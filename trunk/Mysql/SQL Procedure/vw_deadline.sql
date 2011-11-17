SELECT *,
(IF(j.flag_kreditor=0,(IF (j.id_kreditor>0,(IF(DATE_ADD(j.tgl_jual, INTERVAL - k.jangka_waktu DAY)>CURDATE(),"Piutang","Tagih")),'Lunas')) ,'Lunas'))AS statusjual
FROM tbl_jual j
LEFT JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor