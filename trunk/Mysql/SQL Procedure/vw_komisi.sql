
CREATE
    /*[ALGORITHM = {UNDEFINED | MERGE | TEMPTABLE}]
    [DEFINER = { user | CURRENT_USER }]
    [SQL SECURITY { DEFINER | INVOKER }]*/
    VIEW `pos_db`.`vw_komisi` 
    AS
(
SELECT j.tgl_komisi AS tgl_jual,d.kd_departement,d.nm_departement,d.an,d.bn,d.rn,SUM(j.komisi) AS komisi,COUNT(j.kd_pasien) AS pasien
FROM tbl_jual j
JOIN tbl_departement d ON d.id_departement=j.id_departement
WHERE j.flag_debitor=1
GROUP BY j.tgl_komisi,j.id_departement
);
