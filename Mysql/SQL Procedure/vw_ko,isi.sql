SELECT DATE_FORMAT(j.tgl_jual,'%Y-%m-%d') AS tgl_jual,d.kd_departement,d.nm_departement,d.an,d.bn,d.rn,SUM(j.komisi) AS komisi,COUNT(j.kd_pasien) AS pasien
FROM tbl_jual j
JOIN tbl_departement d ON d.id_departement=j.id_departement
GROUP BY DATE_FORMAT(j.tgl_jual,'%Y-%m-%d'),j.id_departement

