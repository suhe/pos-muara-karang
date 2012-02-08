SELECT DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),CURDATE(),DATE_ADD(DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),INTERVAL+30 DAY) AS tgl_jw
FROM tbl_jual j


SELECT *,DATE_ADD(DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),INTERVAL+k.jangka_waktu DAY) AS tgl_jw,(IF(j.flag_kreditor=0,(IF (j.id_kreditor>0,(IF(DATE_ADD(DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),INTERVAL+k.jangka_waktu DAY)>CURDATE(),'Piutang','Tagih')),'Lunas')) ,'Lunas'))AS statusjual,DATE_FORMAT(tgl_jual,'%Y-%m-%d') AS tgl_jual FROM tbl_jual j JOIN tbl_kreditor k ON k.id_kreditor=j.id_kreditor WHERE j.flag_kreditor=0  AND j.piutang>0