/*
SQLyog Enterprise - MySQL GUI v8.18 
MySQL - 5.5.8 : Database - pos_db
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;
CREATE DATABASE /*!32312 IF NOT EXISTS*/`pos_db` /*!40100 DEFAULT CHARACTER SET latin1 */;

USE `pos_db`;

/*Table structure for table `tbl_beli` */

DROP TABLE IF EXISTS `tbl_beli`;

CREATE TABLE `tbl_beli` (
  `id_beli` int(11) NOT NULL AUTO_INCREMENT,
  `no_beli` varchar(255) NOT NULL,
  `tgl_beli` datetime DEFAULT NULL,
  `tgl_bayar` varchar(20) DEFAULT NULL,
  `id_supplier` int(11) NOT NULL,
  `type` enum('Cash','Transfer') NOT NULL,
  `payment` enum('Lunas','Hutang') NOT NULL,
  `bayar` double NOT NULL,
  `hutang` double NOT NULL,
  `flag_supplier` tinyint(1) NOT NULL,
  `tgl_akhir` date NOT NULL,
  `tgl_input` datetime NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  PRIMARY KEY (`id_beli`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_beli` */

insert  into `tbl_beli`(`id_beli`,`no_beli`,`tgl_beli`,`tgl_bayar`,`id_supplier`,`type`,`payment`,`bayar`,`hutang`,`flag_supplier`,`tgl_akhir`,`tgl_input`,`id_pengguna`) values (1,'KAA1','2011-08-20 11:19:35','2011-08-22',4,'Cash','Lunas',125000,0,1,'2011-08-20','2011-08-20 11:19:35',1),(2,'KAA2','2011-08-20 11:20:54','2011-08-20',14,'Cash','Lunas',70000,0,1,'2011-08-20','2011-08-20 11:20:54',1),(3,'KAA3','2011-08-20 11:22:38','2011-08-20',4,'Transfer','Lunas',17000,0,1,'2011-08-20','2011-08-20 11:22:38',1),(4,'KAA4','2011-08-22 13:45:25','2011-08-23',4,'Cash','Lunas',90000,0,1,'2011-08-22','2011-08-22 13:45:25',1),(5,'KAA5','2011-08-22 13:45:54','2011-08-22',11,'Cash','Lunas',12500,0,1,'2011-08-22','2011-08-22 13:45:54',1),(6,'KAA6','2011-08-23 14:49:51','2011-08-23',4,'Cash','Lunas',820000,0,1,'2011-08-23','2011-08-23 14:49:51',1);

/*Table structure for table `tbl_beli_details` */

DROP TABLE IF EXISTS `tbl_beli_details`;

CREATE TABLE `tbl_beli_details` (
  `no_beli` varchar(255) NOT NULL,
  `id_obat` int(11) NOT NULL,
  `harga_beli` double NOT NULL,
  `jumlah` int(11) NOT NULL,
  `tgl_retur` varchar(20) DEFAULT NULL,
  `retur` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `tbl_beli_details` */

insert  into `tbl_beli_details`(`no_beli`,`id_obat`,`harga_beli`,`jumlah`,`tgl_retur`,`retur`) values ('KAA1',16,500,10,'2011-08-23',2),('KAA1',18,12000,10,'2011-08-23',2),('KAA2',13,700,100,NULL,0),('KAA3',18,12000,1,NULL,0),('KAA3',16,500,10,NULL,0),('KAA4',16,500,1,NULL,0),('KAA4',11,8000,1,NULL,0),('KAA4',15,0,10,NULL,0),('KAA4',17,7000,10,NULL,0),('KAA4',27,1500,1,NULL,0),('KAA4',25,1000,10,NULL,0),('KAA5',16,500,1,NULL,0),('KAA5',18,12000,1,NULL,0),('KAA6',18,12000,1,NULL,0),('KAA6',10,8000,1,NULL,0),('KAA6',1,8000,100,NULL,0);

/*Table structure for table `tbl_business_info` */

DROP TABLE IF EXISTS `tbl_business_info`;

CREATE TABLE `tbl_business_info` (
  `id` int(11) NOT NULL,
  `id_cabang` int(11) NOT NULL,
  `bussines_name` varchar(255) NOT NULL,
  `bussines_cp` varchar(255) NOT NULL,
  `bussines_address` varchar(255) NOT NULL,
  `bussines_city` varchar(255) NOT NULL,
  `bussines_note` varchar(255) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `tbl_business_info` */

insert  into `tbl_business_info`(`id`,`id_cabang`,`bussines_name`,`bussines_cp`,`bussines_address`,`bussines_city`,`bussines_note`) values (1,0,'Klinik Dokter Dr Keluarga Muara','DR Dr. Darwis Hartono','Jl.Muara Karang A5 No.29','DKI Jakarta','Kunjungi www.cy99.com');

/*Table structure for table `tbl_cabang` */

DROP TABLE IF EXISTS `tbl_cabang`;

CREATE TABLE `tbl_cabang` (
  `id_cabang` int(11) NOT NULL AUTO_INCREMENT,
  `kd_cabang` char(2) DEFAULT NULL,
  `nm_cabang` varchar(255) NOT NULL,
  `almt_cabang` varchar(255) NOT NULL,
  `kota_cabang` varchar(255) NOT NULL,
  `jw_waktu` int(2) NOT NULL,
  `plafon_default` double NOT NULL,
  `tgl_input` datetime NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  PRIMARY KEY (`id_cabang`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_cabang` */

insert  into `tbl_cabang`(`id_cabang`,`kd_cabang`,`nm_cabang`,`almt_cabang`,`kota_cabang`,`jw_waktu`,`plafon_default`,`tgl_input`,`id_pengguna`) values (0,'AA','Cabang Muara Karang','Jl.Muara Karang No.90','DKI Jakarata',60,1000000,'2011-07-10 14:44:19',1),(1,'AB','Cabang Muara Angke','Jakarta Raya','jakarta Raya',60,1000000,'2011-07-10 14:44:37',1);

/*Table structure for table `tbl_cash` */

DROP TABLE IF EXISTS `tbl_cash`;

CREATE TABLE `tbl_cash` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `tgl_cash` date DEFAULT NULL,
  `money_cash` double NOT NULL,
  `flag` smallint(6) NOT NULL,
  `cash` double DEFAULT NULL,
  `tgl_input` date DEFAULT NULL,
  `id_pengguna` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_cash` */

insert  into `tbl_cash`(`id`,`tgl_cash`,`money_cash`,`flag`,`cash`,`tgl_input`,`id_pengguna`) values (1,'2011-08-20',0,0,0,'2011-08-20',3),(2,'2011-08-21',313800,0,100000,'2011-08-21',1),(3,'2011-08-22',23680,0,0,'2011-08-22',1),(4,'2011-08-23',92500,0,100000,'2011-08-23',1),(5,'2011-08-24',937400,0,0,'2011-08-24',1),(7,'2011-08-25',937400,0,0,'2011-08-25',1),(8,'2011-08-26',12212400,0,0,'2011-08-26',1);

/*Table structure for table `tbl_departement` */

DROP TABLE IF EXISTS `tbl_departement`;

CREATE TABLE `tbl_departement` (
  `id_departement` int(11) NOT NULL AUTO_INCREMENT,
  `kd_departement` varchar(255) NOT NULL,
  `nm_departement` varchar(255) NOT NULL,
  `parent_id` int(11) NOT NULL,
  `group_departement` enum('1','2','3','4','5') NOT NULL,
  `bn` double NOT NULL,
  `an` double NOT NULL,
  `vn` double NOT NULL,
  `rn` int(11) NOT NULL,
  `pn` double NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  `tgl_input` datetime NOT NULL,
  PRIMARY KEY (`id_departement`)
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_departement` */

insert  into `tbl_departement`(`id_departement`,`kd_departement`,`nm_departement`,`parent_id`,`group_departement`,`bn`,`an`,`vn`,`rn`,`pn`,`id_pengguna`,`tgl_input`) values (1,'120','Departement Gigi',0,'1',0,0,0,75,0,1,'2011-07-10 13:57:14'),(2,'100','Departement Umum',0,'2',0,0,0,0,7500,3,'2011-07-20 00:00:00'),(3,'145','Departement Lab',0,'1',0,0,0,75,0,1,'2011-08-11 02:48:21'),(5,'2020','Departement Tindakan',0,'2',200000,300000,30000,40,0,1,'2011-07-24 18:38:48'),(6,'2021','Departement Tindakan 1',5,'2',0,199000,0,40,0,1,'2011-07-24 18:40:09'),(7,'2022','Departement Tindakan 2',5,'2',300001,400000,40000,40,0,1,'2011-07-24 18:41:27'),(8,'101','Departement Obat',0,'1',0,0,0,0,0,1,'2011-07-30 12:22:30'),(9,'2023','Departement Tindakan III',5,'2',400001,10000000,50000,40,0,1,'2011-07-30 12:28:12');

/*Table structure for table `tbl_jual` */

DROP TABLE IF EXISTS `tbl_jual`;

CREATE TABLE `tbl_jual` (
  `id_jual` int(11) NOT NULL AUTO_INCREMENT,
  `no_jual` varchar(255) NOT NULL,
  `tgl_jual` datetime DEFAULT NULL,
  `tgl_bayar` varchar(20) DEFAULT NULL,
  `tgl_komisi` varchar(20) DEFAULT NULL,
  `tgl_akhir` date DEFAULT NULL,
  `kd_pasien` varchar(11) NOT NULL,
  `id_kreditor` int(11) NOT NULL,
  `id_cabang` int(11) NOT NULL,
  `id_departement` int(11) NOT NULL,
  `type` enum('Cash','Transfer') NOT NULL,
  `payment` enum('Lunas','Piutang') NOT NULL,
  `bayar` double NOT NULL,
  `dibayar` double NOT NULL,
  `piutang` double NOT NULL,
  `komisi` double NOT NULL,
  `flag_debitor` tinyint(1) NOT NULL,
  `flag_kreditor` tinyint(1) NOT NULL,
  `jw` tinyint(2) NOT NULL,
  `tgl_input` datetime NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  PRIMARY KEY (`id_jual`)
) ENGINE=InnoDB AUTO_INCREMENT=22 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_jual` */

insert  into `tbl_jual`(`id_jual`,`no_jual`,`tgl_jual`,`tgl_bayar`,`tgl_komisi`,`tgl_akhir`,`kd_pasien`,`id_kreditor`,`id_cabang`,`id_departement`,`type`,`payment`,`bayar`,`dibayar`,`piutang`,`komisi`,`flag_debitor`,`flag_kreditor`,`jw`,`tgl_input`,`id_pengguna`) values (1,'MAB1','2011-08-20 10:52:54','2011-08-21','2011-08-21','2011-08-20','D1',18,1,5,'Cash','Lunas',125000,125000,0,50000,1,1,56,'2011-08-20 10:52:54',3),(2,'MAB2','2011-08-20 10:53:24','2011-08-20','2011-08-21','2011-08-20','B4',0,1,5,'Cash','Lunas',205000,300000,0,70000,1,1,0,'2011-08-20 10:53:24',3),(3,'MAA3','2011-08-20 11:05:37','2011-08-20','2011-08-21','2011-08-20','D2',0,0,5,'Cash','Lunas',195800,200000,0,78320,1,1,0,'2011-08-20 11:05:37',1),(4,'MAA4','2011-08-20 11:05:50','2011-08-21',NULL,'2011-08-20','A1',15,0,8,'Cash','Lunas',125000,125000,0,0,0,1,5,'2011-08-20 11:05:50',1),(5,'MAA5','2011-08-20 11:06:04','2011-08-21','2011-08-21','2011-08-20','B3',15,0,5,'Cash','Lunas',120000,120000,0,48000,1,1,5,'2011-08-20 11:06:04',1),(6,'MAA6','2011-08-22 14:03:32','2011-08-22',NULL,'2011-08-22','D1',0,0,8,'Cash','Lunas',320000,320000,0,0,0,1,0,'2011-08-22 14:03:32',1),(7,'MAA7','2011-08-22 14:09:30','2011-08-23','2011-08-23','2011-08-22','K1',2,0,5,'Cash','Lunas',134000,134000,0,53600,1,1,30,'2011-08-22 14:09:30',1),(8,'MAA8','2011-08-22 14:15:07','2011-08-23','2011-08-23','2011-08-22','M2',2,0,5,'Cash','Lunas',120000,120000,0,48000,1,1,8,'2011-08-22 14:15:07',1),(9,'MAA9','2011-08-23 14:48:04','2011-08-23',NULL,'2011-08-23','R1',0,0,5,'Cash','Lunas',120000,120000,0,48000,0,1,0,'2011-08-23 14:48:04',1),(10,'MAA10','2011-08-23 14:50:23','2011-08-23',NULL,'2011-08-23','K1',0,0,5,'Cash','Lunas',1700000,1800000,0,660000,0,1,0,'2011-08-23 14:50:23',1),(11,'MAA11','2011-08-24 10:38:25','2011-08-25',NULL,'2011-08-24','B3',20,0,5,'Cash','Lunas',6640000,6640000,0,2636000,0,1,60,'2011-08-24 10:38:25',1),(12,'MAA12','2011-08-24 10:45:25','2011-08-25',NULL,'2011-08-24','B3',15,0,5,'Cash','Lunas',1050000,1050000,0,400000,0,1,60,'2011-08-24 10:45:25',1),(13,'MAA13','2011-08-24 10:46:46','2011-08-25',NULL,'2011-08-24','K1',20,0,5,'Cash','Lunas',2005000,2005000,0,782000,0,1,39,'2011-08-24 10:46:46',1),(14,'MAA14','2011-08-24 10:48:00','2011-08-25',NULL,'2011-08-24','B4',20,0,5,'Cash','Lunas',1580000,1580000,0,612000,0,1,32,'2011-08-24 10:48:00',1),(15,'MAA15','2011-08-25 11:08:40','2011-08-26',NULL,'2011-08-25','K1',15,0,5,'Cash','Lunas',870000,870000,0,328000,0,1,47,'2011-08-25 11:08:40',1),(16,'MAA16','2011-08-25 11:09:25','-',NULL,'2011-08-25','B3',2,0,5,'Cash','',0,0,179800,71920,0,0,47,'2011-08-25 11:09:25',1),(17,'MAA17','2011-08-25 11:10:27','-',NULL,'2011-08-25','B2',20,0,5,'Cash','',0,0,1922000,748800,0,0,47,'2011-08-25 11:10:27',1),(18,'MAA18','2011-08-25 11:11:17','-',NULL,'2011-08-25','K1',20,0,5,'Cash','',0,0,170000,68000,0,0,47,'2011-08-25 11:11:17',1),(19,'MAA19','2011-08-26 11:13:42','-',NULL,'2011-08-26','B4',20,0,5,'Cash','',0,0,6290000,2496000,0,0,45,'2011-08-26 11:13:42',1),(20,'MAA20','2011-08-26 11:15:09','-',NULL,'2011-08-26','K1',20,0,5,'Cash','',0,0,8250000,3280000,0,0,45,'2011-08-26 11:15:09',1),(21,'MAA21','2011-08-26 11:22:07','-',NULL,'2011-08-26','K1',20,0,5,'Cash','',0,0,1200000,460000,0,0,24,'2011-08-26 11:22:07',1);

/*Table structure for table `tbl_jual_details` */

DROP TABLE IF EXISTS `tbl_jual_details`;

CREATE TABLE `tbl_jual_details` (
  `no_jual` varchar(255) NOT NULL,
  `id_obat` int(11) NOT NULL,
  `harga_jual` double NOT NULL,
  `jumlah` int(11) NOT NULL,
  `dosis` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `tbl_jual_details` */

insert  into `tbl_jual_details`(`no_jual`,`id_obat`,`harga_jual`,`jumlah`,`dosis`) values ('MAB1',18,12000,10,''),('MAB1',16,500,10,''),('MAB2',18,12000,10,''),('MAB2',16,500,10,''),('MAB2',11,8000,10,''),('MAA3',18,12000,10,''),('MAA3',16,500,10,''),('MAA3',12,7000,10,''),('MAA3',10,800,1,''),('MAA4',18,12000,10,''),('MAA4',16,500,10,''),('MAA5',18,12000,10,''),('MAA6',18,12000,10,''),('MAA6',16,500,100,''),('MAA6',27,1500,100,''),('MAA7',18,12000,10,''),('MAA7',16,500,10,''),('MAA7',10,9000,1,''),('MAA8',18,12000,10,''),('MAA9',18,12000,10,''),('MAA10',18,12000,100,''),('MAA10',16,500,1000,''),('MAA11',18,12000,10,''),('MAA11',16,500,10,''),('MAA11',27,1500,10,''),('MAA11',2,5000,100,''),('MAA11',9,60000,100,''),('MAA12',18,12000,10,''),('MAA12',16,500,100,''),('MAA12',1,8000,100,''),('MAA12',24,800,100,''),('MAA13',16,500,10,''),('MAA13',18,12000,100,''),('MAA13',24,800,1000,''),('MAA14',16,500,100,''),('MAA14',18,12000,100,''),('MAA14',24,800,100,''),('MAA14',26,1000,100,''),('MAA14',22,1500,100,''),('MAA15',18,12000,10,''),('MAA15',16,500,100,''),('MAA15',12,7000,100,''),('MAA16',18,12000,10,''),('MAA16',16,500,100,''),('MAA16',10,9800,1,''),('MAA17',20,2200,10,''),('MAA17',18,12000,100,''),('MAA17',17,7000,100,''),('MAA18',18,12000,10,''),('MAA18',16,500,100,''),('MAA19',18,12000,100,''),('MAA19',16,500,10000,''),('MAA19',10,90000,1,''),('MAA20',18,12000,100,''),('MAA20',16,500,100,''),('MAA20',12,7000,1000,''),('MAA21',18,12000,100,'');

/*Table structure for table `tbl_kategori` */

DROP TABLE IF EXISTS `tbl_kategori`;

CREATE TABLE `tbl_kategori` (
  `id_kategori` int(11) NOT NULL AUTO_INCREMENT,
  `nm_kategori` varchar(255) NOT NULL,
  `desk_kategori` varchar(255) NOT NULL,
  `tgl_input` datetime NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  PRIMARY KEY (`id_kategori`)
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_kategori` */

insert  into `tbl_kategori`(`id_kategori`,`nm_kategori`,`desk_kategori`,`tgl_input`,`id_pengguna`) values (1,'Century 21','Centrury','2011-07-10 14:25:00',1),(2,'Apotek','Mantap','2011-07-10 14:25:22',2),(4,'Mantap','Mantap','2011-07-10 16:26:56',1),(5,'Tortoise','Tortoise','2011-07-10 16:27:06',1),(6,'rtee','erereg','2011-07-10 16:49:06',1),(7,'kaget','kaget','2011-07-10 18:36:38',1),(8,'kA','DA','2011-07-10 18:36:46',1),(9,'test','test','2011-08-11 05:26:26',1);

/*Table structure for table `tbl_kreditor` */

DROP TABLE IF EXISTS `tbl_kreditor`;

CREATE TABLE `tbl_kreditor` (
  `id_kreditor` int(11) NOT NULL AUTO_INCREMENT,
  `nm_kreditor` varchar(255) NOT NULL,
  `almt_kreditor` varchar(255) NOT NULL,
  `kota_kreditor` varchar(255) NOT NULL,
  `tlp_kreditor` varchar(15) NOT NULL,
  `cp_kreditor` varchar(255) NOT NULL,
  `email_kreditor` varchar(255) NOT NULL,
  `jangka_waktu` int(2) NOT NULL,
  `plafon` double NOT NULL,
  `tgl_input` datetime NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  PRIMARY KEY (`id_kreditor`)
) ENGINE=InnoDB AUTO_INCREMENT=21 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_kreditor` */

insert  into `tbl_kreditor`(`id_kreditor`,`nm_kreditor`,`almt_kreditor`,`kota_kreditor`,`tlp_kreditor`,`cp_kreditor`,`email_kreditor`,`jangka_waktu`,`plafon`,`tgl_input`,`id_pengguna`) values (1,'PT Lacone','Jl.Yuang No.90','Bandung','08432222','Haris','iptek_suhe@yahoo.co.id',30,10000000,'2011-07-10 14:31:40',1),(2,'PT Kreston International Indonesia','Jl.HR Rasuna Said Kav5 Blox X2','Bandung','084223232','Bpk Haris Maulana','',30,10000000,'2011-07-10 22:05:24',1),(3,'PT Arya Makmur Sentosa 2','0853343432','Bandung','0853343432','Suhendar','',30,10000000,'2011-07-18 04:15:02',0),(4,'PT Ekasapata','adad','adasd','adad','adasd','',30,10000000,'2011-07-18 04:26:50',0),(5,'PT,Sari Mulya','Kuningan Timur','Jakarta','08544765','Bpk.Ahmad','',30,10000000,'2011-08-13 22:32:03',1),(6,'PT.Jaya Sakti No.50','Jln.Jendral Sudirman No.80','Jakarta','0987654588','Ibu.Zahra','',30,10000000,'2011-08-13 22:33:17',1),(7,'PT.Maju Mundur','Jln,Widuri No.90','Jakarta','9987646677','Bpk.Andi','',30,10000000,'2011-08-13 22:38:16',1),(8,'PT.Angin Beliung','Jln.Kedoya Duri No.50','Jakarta','0435775678','Ibu.Pupu','',30,10000000,'2011-08-13 22:39:24',1),(9,'PT.Sarinah Murti Jaya','Jln.Perintis No.20','Jakarta','0556777888','Bpk.Parhan','',30,10000000,'2011-08-13 22:40:23',1),(10,'PT.Sentosa Raya','Jln.Mampang No.80','Jakarta','0213457878','Bpk.Anton','',30,10000000,'2011-08-13 22:41:21',1),(11,'PT.Angkasa Raya','Jln.Kuningan Barat No.50','Jakarta','0214556788','Ibu.Dewi','',30,10000000,'2011-08-13 22:42:30',1),(12,'PT.Murti Angkasa','Jln.Perintis No.40','Jakarta','0456733455','Ibu.Sandra','',30,10000000,'2011-08-13 22:43:24',1),(13,'PT.Laut Barat','Jln.Karang Mulya No.80','Tangerang','0856754456','Bpk.Sodikin','',30,10000000,'2011-08-13 22:44:28',1),(14,'PT.Bumi Pertiwi','Jln.Mampang No.20','Jakarta','0546787687','Bpk.Burhan','',30,10000000,'2011-08-13 22:45:22',1),(15,'PT.Berlian Raya','Jln.Raden Saleh No70','Jakarta','0213456789','Bpk.Zulfikar','',30,10000000,'2011-08-13 22:46:52',1),(16,'PT.Sanjawijaya','Jln.Perintis.No.70','Jakarta','0216678899','Ibu.Ratna','',30,10000000,'2011-08-13 22:47:57',1),(17,'PT.Sanjaya  Sakti','Jln.Kuningan Timur No.10','DKI Jakarta','0212234567','Ibu Linda','',30,10000000,'2011-08-14 20:13:34',1),(18,'PT.Permata Raya','Jln.Kuningan Barat','Jakarta','0213456567','Bpk.Syaripudin','',30,10000000,'2011-08-14 20:14:30',1),(19,'PT.Sarijaya','Jln.Kedoya Duri','Tangerang','0213456789','Bpk.Hery','',30,10000000,'2011-08-14 20:15:42',1),(20,'PT.Sarinah Jaya','Jln.Kuningan Timur No.30','Jakarta','0213456765','Bpk.Zulkarnaen','',30,10000000,'2011-08-14 20:16:37',1);

/*Table structure for table `tbl_log` */

DROP TABLE IF EXISTS `tbl_log`;

CREATE TABLE `tbl_log` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `tgl_akses` datetime DEFAULT NULL,
  `id_pengguna` int(11) DEFAULT NULL,
  `id_cabang` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=40 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_log` */

insert  into `tbl_log`(`id`,`tgl_akses`,`id_pengguna`,`id_cabang`) values (1,'2011-08-20 10:52:29',3,1),(2,'2011-08-20 10:55:52',3,1),(3,'2011-08-20 10:57:26',3,1),(4,'2011-08-20 11:00:48',3,1),(5,'2011-08-20 11:02:40',1,0),(6,'2011-08-20 11:19:00',1,0),(7,'2011-08-20 11:22:12',1,0),(8,'2011-08-20 11:24:00',1,0),(9,'2011-08-20 11:25:57',1,0),(10,'2011-08-20 11:30:14',1,0),(11,'2011-08-20 11:32:04',1,0),(12,'2011-08-20 11:33:28',1,0),(13,'2011-08-20 11:34:25',1,0),(14,'2011-08-20 11:37:49',1,0),(15,'2011-08-20 11:38:41',1,0),(16,'2011-08-20 11:39:24',1,0),(17,'2011-08-20 13:10:03',1,0),(18,'2011-08-21 13:17:51',1,0),(19,'2011-08-21 13:25:38',1,0),(20,'2011-08-21 13:28:19',1,0),(21,'2011-08-21 13:32:47',1,0),(22,'2011-08-22 13:41:04',1,0),(23,'2011-08-22 14:01:56',1,0),(24,'2011-08-22 14:14:57',1,0),(25,'2011-08-22 14:21:55',1,0),(26,'2011-08-23 14:43:24',1,0),(27,'2011-08-23 14:47:44',1,0),(28,'2011-08-23 15:00:27',1,0),(29,'2011-08-24 10:26:38',1,0),(30,'2011-08-24 10:27:17',1,0),(31,'2011-08-24 10:38:14',1,0),(32,'2011-08-24 10:45:19',1,0),(33,'2011-08-24 11:00:49',1,0),(34,'2011-08-25 11:01:27',1,0),(35,'2011-08-25 11:03:44',1,0),(36,'2011-08-25 11:06:31',1,0),(37,'2011-08-25 11:08:33',1,0),(38,'2011-08-26 11:12:46',1,0),(39,'2011-08-26 11:22:00',1,0);

/*Table structure for table `tbl_obat` */

DROP TABLE IF EXISTS `tbl_obat`;

CREATE TABLE `tbl_obat` (
  `id_obat` int(11) NOT NULL AUTO_INCREMENT,
  `pl_obat` char(1) DEFAULT NULL,
  `pk_obat` int(11) DEFAULT NULL,
  `kd_obat` varchar(255) NOT NULL,
  `id_kategori` int(11) NOT NULL,
  `nm_obat` varchar(255) NOT NULL,
  `nm_ilmiah` varchar(255) DEFAULT NULL,
  `harga_beli` double NOT NULL,
  `harga_jual` double NOT NULL,
  `kemasan` varchar(255) NOT NULL,
  `stok` double NOT NULL,
  `stok_min` double NOT NULL,
  `tgl_input` datetime NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  PRIMARY KEY (`id_obat`)
) ENGINE=InnoDB AUTO_INCREMENT=28 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_obat` */

insert  into `tbl_obat`(`id_obat`,`pl_obat`,`pk_obat`,`kd_obat`,`id_kategori`,`nm_obat`,`nm_ilmiah`,`harga_beli`,`harga_jual`,`kemasan`,`stok`,`stok_min`,`tgl_input`,`id_pengguna`) values (1,'B',1,'B1',1,'Batuk','test',1000,8000,'Tablet',2,1,'2011-07-10 14:26:39',2),(2,'F',1,'F1',1,'Furakui','test',4500,5000,'Tablet',2,2,'2011-08-11 04:06:49',1),(9,'F',2,'F2',4,'Farugajimon','testing',6000,60000,'Tablet',2,2,'2011-08-11 05:31:36',1),(10,'A',1,'A999',1,'Obat Batuk','testing',7000,8000,'Tablet',2,2,'2011-07-16 14:41:01',1),(11,'B',2,'B2',1,'Burayak','Buarayak',7000,8000,'Tablet',5,2,'2011-07-20 01:55:47',1),(12,'B',3,'B3',1,'Bantalan Kapuk',NULL,5000,7000,'Tablet',4,2,'2011-07-20 01:56:07',1),(13,'A',2,'A2',1,'Amoxilin',NULL,500,700,'Tablet',5,2,'2011-07-24 11:17:04',1),(14,'P',1,'P1',7,'Paramex',NULL,5000,8000,'Tablet',10,2,'2011-07-24 14:04:19',1),(15,'P',2,'P2',1,'Panadol',NULL,0,0,'Tablet',0,0,'2011-08-09 22:54:32',1),(16,'B',4,'B4',4,'Bodrex','xianida',400,500,'Tablet',50,15,'2011-08-13 21:59:27',1),(17,'L',1,'L1',2,'Laserin','lsn',3500,7000,'Sirup',45,10,'2011-08-13 22:00:23',1),(18,'O',1,'O1',1,'OBH Combi Plus','ochi',7500,12000,'Sirup',100,25,'2011-08-13 22:01:26',1),(19,'K',1,'K1',5,'Komix','kmx',1000,1700,'Sirup',125,25,'2011-08-13 22:02:18',1),(20,'A',3,'A3',6,'Anastan','anstn',1300,2200,'Tablet',170,30,'2011-08-13 22:03:23',1),(21,'P',3,'P3',6,'Paramex','pmx',400,1000,'Tablet',200,50,'2011-08-13 22:05:46',1),(22,'K',2,'K2',1,'Konidin','xkjhn',400,1500,'Tablet',250,56,'2011-08-13 22:06:30',1),(23,'N',1,'N1',2,'Nelco','nlc',5700,9000,'Sirup',56,20,'2011-08-13 22:07:26',1),(24,'D',1,'D1',1,'Decolgen','Dcolgn',350,800,'Tablet',160,20,'2011-08-13 22:09:02',1),(25,'M',1,'M1',2,'Mixagrip','mixp',450,1000,'Tablet',300,50,'2011-08-13 22:09:51',1),(26,'D',2,'D2',2,'Dialet','Dltz',200,1000,'Tablet',300,60,'2011-08-13 22:10:29',1),(27,'U',1,'U1',2,'Ultraflu','ultf',450,1500,'Tablet',360,80,'2011-08-13 22:11:27',1);

/*Table structure for table `tbl_pasien` */

DROP TABLE IF EXISTS `tbl_pasien`;

CREATE TABLE `tbl_pasien` (
  `id_pasien` int(11) NOT NULL AUTO_INCREMENT,
  `pl_pasien` char(1) DEFAULT NULL,
  `pk_pasien` int(11) DEFAULT NULL,
  `kd_pasien` varchar(255) NOT NULL,
  `nm_pasien` varchar(255) NOT NULL,
  `relasi` varchar(255) DEFAULT NULL,
  `jk_pasien` enum('Pria','Wanita') NOT NULL,
  `tgl_lahir` date NOT NULL,
  `tmpt_lahir` varchar(255) NOT NULL,
  `alamat` varchar(255) NOT NULL,
  `kota` varchar(255) NOT NULL,
  `no_hp` varchar(15) NOT NULL,
  `no_tlp` varchar(15) NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  `tgl_input` datetime NOT NULL,
  PRIMARY KEY (`id_pasien`)
) ENGINE=InnoDB AUTO_INCREMENT=27 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_pasien` */

insert  into `tbl_pasien`(`id_pasien`,`pl_pasien`,`pk_pasien`,`kd_pasien`,`nm_pasien`,`relasi`,`jk_pasien`,`tgl_lahir`,`tmpt_lahir`,`alamat`,`kota`,`no_hp`,`no_tlp`,`id_pengguna`,`tgl_input`) values (1,'J',1,'J1','Joko Widodo','yada','Pria','2011-07-10','DKI Jakarta','Jl.Perintis No.39','DKI Jakarta','085222054064','130828',1,'2011-07-10 13:26:26'),(2,'K',1,'K1','Karunia Bakti','yada','Pria','2011-07-10','DKI Jakarta','Jl.Perintis No.90','DKI Jakarta','0853332222','021-9092211',1,'2011-07-10 13:29:01'),(3,'A',1,'A1','Andreas Iniesta','Xavi Hernandes','Pria','2011-07-24','Andalusia','Jl.Barcelona No.90 ','Melanisia','0832323232','021-90001212',1,'2011-08-11 00:58:24'),(4,'B',1,'B1','Bangun Purba','Saudara Kandung','Pria','2011-08-12','Bandung','Jl.Pelita Harapan Jaya','Bandung','08334343','0211212',1,'2011-08-12 04:15:29'),(5,'M',1,'M1','Marmo','keluarga','Pria','2011-07-16','Bandung','test','tet','08567788','021-0988811',0,'2011-07-16 14:36:12'),(6,'M',2,'M2','Marni','test','Pria','2011-07-16','test','Jl.Ba','Bandung','test','test',0,'2011-07-16 14:38:51'),(7,'B',2,'B2','Barugamuri Daichi','Suhendar','Pria','2011-07-18','121212','Jl.Pejompongan No.90','Bandung','212121','212121',0,'2011-07-18 00:49:22'),(8,'A',2,'A2','Anita yuan','Sastrowardoyo','Wanita','2004-05-29','Jakarta','Jln.Gatot Subroto','Jakarta','085714466195','021445765',1,'2011-07-24 11:13:40'),(9,'B',3,'B3','Bondan Prakorso','Agung Permana','Pria','2004-05-29','Bandung','Jl.Hamid Awaludin No.80','Bandung','083444222','021-90012121',1,'2011-07-24 11:15:09'),(10,'J',2,'J2','Juminten','Suhendar','Pria','1903-04-04','','Jl.Pegangsaan Timur No.9','Tanjung Periok','08544433','03222',1,'2011-08-12 20:18:14'),(11,'D',1,'D1','Dede Watanabe','keluarga','Pria','1902-04-02','','Jl.Raya Pekan Mandura','Tanjung Periok','0853330999','021-0911111',1,'2011-08-13 08:11:22'),(12,'A',3,'A3','anita yuan','Yuan','Wanita','1982-03-02','','jln.Gatot subroto','Jakarta','0857324232','02156785',1,'2011-08-13 21:41:10'),(13,'D',2,'D2','Doni Kusuma ','Atmawijaya','Pria','1983-05-04','','jln.Oma Angga wisastra','bandung','085222453267','021768549',1,'2011-08-13 21:43:03'),(14,'R',1,'R1','Rani Widuriningsih','Sastrowardoyo','Wanita','1988-08-06','','jln.Kuningan Timur','Jakarta','0857123224542','021654327',1,'2011-08-13 21:45:16'),(15,'S',1,'S1','Santi Susilawati','Rudianto','Wanita','1989-06-09','','Jln.Kuningan Barat','Jakarta','081223456672','0213222567',1,'2011-08-13 21:47:18'),(16,'B',4,'B4','Budianto','Mahmudin','Pria','1990-02-01','','Jln.Kuningan Timur','Jakarta','085222342421','02177665543',1,'2011-08-13 21:50:30'),(17,'L',1,'L1','Lania Purwanti','Suryadi','Wanita','1993-05-08','','Jln.Jendral sudirman','Jakarta','085722345621','0213455673',1,'2011-08-13 21:52:09'),(18,'L',2,'L2','Larasati','Adipura','Wanita','1990-03-07','','Jln.Mega Kuningan','Jakarta','081234455632','0213988754',1,'2011-08-13 21:53:45'),(19,'K',2,'K2','Kirana','Anggawidura','Wanita','1994-04-04','','Jln.Jendral Sudirman','Jakarta','085712324155','02134556432',1,'2011-08-13 21:55:18'),(20,'P',1,'P1','Pupu Puspitasari','Syaripudin','Wanita','1987-06-05','','Jln.Guru Mughni','Jakarta','081222346677','02134256787',1,'2011-08-13 21:56:53'),(21,'G',1,'G1','Gyugaada','saca','Pria','1903-07-05','','sada','Tanjung Periok','','',1,'2011-08-20 10:04:11'),(22,'D',3,'D3','dad','asa','Pria','1905-04-08','','ss','Tanjung Periok','','',1,'2011-08-20 10:07:32'),(23,'S',2,'S2','ssss','sss','Pria','1901-06-03','','sass','Pluit','','',1,'2011-08-20 10:15:06'),(24,'B',5,'B5','bbbb','ada','Pria','1946-03-02','','afafa','Tanjung Periok','343','43434',1,'2011-08-20 10:20:24'),(25,'E',1,'E1','erwin','hendar','Pria','1978-05-02','','mk blok 4','Pluit','8768787','6767',1,'2011-08-25 20:24:46'),(26,'T',1,'T1','Tanti Maria Hutapea','Tuti','Pria','1904-05-03','','Jl.Lengsir Wengi No.80 Jakarta','Pluit','0992232','2323232323',1,'2011-08-20 11:24:43');

/*Table structure for table `tbl_pengguna` */

DROP TABLE IF EXISTS `tbl_pengguna`;

CREATE TABLE `tbl_pengguna` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `nm_pengguna` varchar(255) NOT NULL,
  `password` varchar(255) NOT NULL,
  `jk_pengguna` enum('Pria','Wanita') NOT NULL,
  `level` enum('Administrator','Manager','User') NOT NULL,
  `user_cabang` int(11) DEFAULT NULL,
  `id_admin` int(11) NOT NULL,
  `tgl_input` datetime NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_pengguna` */

insert  into `tbl_pengguna`(`id`,`nm_pengguna`,`password`,`jk_pengguna`,`level`,`user_cabang`,`id_admin`,`tgl_input`) values (1,'Suhendar','12345','Pria','Administrator',0,1,'2011-07-30 13:49:49'),(2,'Prameswara','12345','Pria','Manager',1,1,'2011-07-10 13:50:36'),(3,'Rika Van Houten','12345','Wanita','User',1,1,'2011-07-10 13:51:22'),(4,'Administrator','12345','Pria','Administrator',1,1,'2011-07-10 13:14:12'),(5,'test','test','Pria','User',1,0,'2011-07-16 14:33:39');

/*Table structure for table `tbl_supplier` */

DROP TABLE IF EXISTS `tbl_supplier`;

CREATE TABLE `tbl_supplier` (
  `id_supplier` int(11) NOT NULL AUTO_INCREMENT,
  `nm_supplier` varchar(255) NOT NULL,
  `almt_supplier` varchar(255) NOT NULL,
  `tlp_supplier` varchar(15) NOT NULL,
  `cp_supplier` varchar(255) NOT NULL,
  `kota_supplier` varchar(255) NOT NULL,
  `negara_supplier` varchar(255) NOT NULL,
  `id_pengguna` int(11) NOT NULL,
  `tgl_input` datetime NOT NULL,
  PRIMARY KEY (`id_supplier`)
) ENGINE=InnoDB AUTO_INCREMENT=17 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_supplier` */

insert  into `tbl_supplier`(`id_supplier`,`nm_supplier`,`almt_supplier`,`tlp_supplier`,`cp_supplier`,`kota_supplier`,`negara_supplier`,`id_pengguna`,`tgl_input`) values (1,'PT Mutiara Laut Perkasa No.90','Jl.Pelaut No.90','021098999','Paul Gascogini','DKI Jakarta','Indonesia',1,'2011-07-10 13:34:53'),(2,'PT Arkansa Obat Pratama','Jl.Perintis No.90','021-0901111','Paul Van Der Sar','Jakarta','Indonesia',1,'2011-07-10 13:35:48'),(3,'PT Ekayasa Perwira Utama','Jl.Kebon Kopi Raya','021-0299912','Suhendar','Bandung','Indonesia',1,'2011-07-10 17:10:21'),(4,'PT Petukangan Barat','Jl.kebon Jeruk Raya No.80','021-9829901','Bpk.Anwar','DKI Jakarta','Indonesia',0,'2011-07-18 04:21:32'),(5,'PT Vilio','Jln.Oma Angga Wisastra','022334456227','Bpk.Rudi','Bandung','Indonesia',1,'2011-08-13 22:13:47'),(6,'PT Jasatex','Jln.Ciparay-Majalaya','02234456673','Bpk.Rudianto','Bandung','Indonesia',1,'2011-08-13 22:15:18'),(7,'PT. Karina Jaya','Jln.Kuningan Timur','021345567437','Ibu.Desy','Jakarta','Indonesia',1,'2011-08-13 22:16:28'),(8,'PT.Sarimatex','Jln Puri Jaya','021345667843','Bpk.Setiawan','Jakarta','Indonesia',1,'2011-08-13 22:18:35'),(9,'PT.Sarijaya Sakti','JLn.Jendral Sudirman','0219887665545','Ibu.Sumiati','Jakarta','Indonesia',1,'2011-08-13 22:20:19'),(10,'PT.Gunung Galunggung','Jln.Perintis','02134556327','Bpk.Santoso','Jakarta','Indonesia',1,'2011-08-13 22:22:40'),(11,'PT.Garuda Jaya','Jln.Kuningan Barat','021344567876','Bpk.Romli','Jakarta','Indonesia',1,'2011-08-13 22:24:16'),(12,'PT.Perkasa Sakti','Jln.Jendral Sudirman','0216678543987','Ibu.mega','Jakarta','Indonesia',1,'2011-08-13 22:25:08'),(13,'PT.Mutiara Jaya','Jln.Perintis','02177789654324','Ibu.Zahwa','Jakarta','Indonesia',1,'2011-08-13 22:26:02'),(14,'PT.Intan Sakti','Jln.Oma Angga Wisastra No.30','0223445654377','Bpk.Rahmat','Bandung','Indonesia',1,'2011-08-13 22:27:55'),(15,'PT.Karang Taruna','Jln.Karang Mulya No.45','021334567647','Bpk.Hidayat','Jakarta','Indonesia',1,'2011-08-13 22:29:16'),(16,'PT.Sansan','Jln.Oma Angga Wisastra No.60','02287654398','Ibu.Sinta','Bandung','Indonesia',1,'2011-08-13 22:30:29');

/*Table structure for table `vw_cash_flow` */

DROP TABLE IF EXISTS `vw_cash_flow`;

/*!50001 DROP VIEW IF EXISTS `vw_cash_flow` */;
/*!50001 DROP TABLE IF EXISTS `vw_cash_flow` */;

/*!50001 CREATE TABLE  `vw_cash_flow`(
 `id` int(11) ,
 `tgl_cash` date ,
 `money_cash` double ,
 `jual_sebelumnya` double ,
 `jual` double ,
 `beli_sebelumnya` double ,
 `beli` double ,
 `retur` double ,
 `komisi` double ,
 `laba` double ,
 `cash` double ,
 `kas` double ,
 `kas_total` double ,
 `pasien` bigint(21) ,
 `lunas` bigint(21) ,
 `pay_komisi` bigint(21) ,
 `flag` smallint(6) 
)*/;

/*Table structure for table `vw_cash_flow_before` */

DROP TABLE IF EXISTS `vw_cash_flow_before`;

/*!50001 DROP VIEW IF EXISTS `vw_cash_flow_before` */;
/*!50001 DROP TABLE IF EXISTS `vw_cash_flow_before` */;

/*!50001 CREATE TABLE  `vw_cash_flow_before`(
 `id` int(11) ,
 `tgl_cash` varchar(10) ,
 `jual_sebelumnya` double ,
 `jual` double ,
 `beli_sebelumnya` double ,
 `beli` double ,
 `retur` double ,
 `komisi` double ,
 `laba` double ,
 `cash` double ,
 `kas` double ,
 `pasien` bigint(21) ,
 `lunas` bigint(21) ,
 `pay_komisi` bigint(21) ,
 `flag` smallint(6) 
)*/;

/*Table structure for table `vw_komisi` */

DROP TABLE IF EXISTS `vw_komisi`;

/*!50001 DROP VIEW IF EXISTS `vw_komisi` */;
/*!50001 DROP TABLE IF EXISTS `vw_komisi` */;

/*!50001 CREATE TABLE  `vw_komisi`(
 `tgl_jual` varchar(20) ,
 `kd_departement` varchar(255) ,
 `nm_departement` varchar(255) ,
 `an` double ,
 `bn` double ,
 `rn` int(11) ,
 `komisi` double ,
 `pasien` bigint(21) 
)*/;

/*Table structure for table `vw_stok_min` */

DROP TABLE IF EXISTS `vw_stok_min`;

/*!50001 DROP VIEW IF EXISTS `vw_stok_min` */;
/*!50001 DROP TABLE IF EXISTS `vw_stok_min` */;

/*!50001 CREATE TABLE  `vw_stok_min`(
 `id_obat` int(11) ,
 `kd_obat` varchar(255) ,
 `nm_obat` varchar(255) ,
 `id_kategori` int(11) ,
 `nm_kategori` varchar(255) ,
 `kemasan` varchar(255) ,
 `harga_jual` double ,
 `harga_beli` double ,
 `profit` double ,
 `stok` double ,
 `beli` decimal(32,0) ,
 `jual` decimal(32,0) ,
 `rugi` decimal(32,0) ,
 `sisa` double ,
 `stok_min` double ,
 `tgl_input` datetime ,
 `nm_pengguna` varchar(255) 
)*/;

/*View structure for view vw_cash_flow */

/*!50001 DROP TABLE IF EXISTS `vw_cash_flow` */;
/*!50001 DROP VIEW IF EXISTS `vw_cash_flow` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vw_cash_flow` AS (select `c`.`id` AS `id`,`c`.`tgl_cash` AS `tgl_cash`,`c`.`money_cash` AS `money_cash`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `jual_sebelumnya`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (convert(date_format(`j`.`tgl_input`,'%Y-%m-%d') using latin1) = `j`.`tgl_bayar`) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (convert(date_format(`j`.`tgl_input`,'%Y-%m-%d') using latin1) = `j`.`tgl_bayar`) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `jual`,if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `beli_sebelumnya`,if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) AS `beli`,if(((select count(0) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`d`.`retur` * `d`.`harga_beli`)) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `retur`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum(`j`.`komisi`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `komisi`,((if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0)) - (((if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) + if(((select count(0) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`d`.`retur` * `d`.`harga_beli`)) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where (`b`.`flag_supplier` and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0)) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum(`j`.`komisi`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0))) AS `laba`,`c`.`cash` AS `cash`,(((if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) - (((if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) + if(((select count(0) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`d`.`retur` * `d`.`harga_beli`)) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0)) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum(`j`.`komisi`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and `j`.`flag_debitor` and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0))) - `c`.`cash`) AS `kas`,((((if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) - (((if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) + if(((select count(0) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`d`.`retur` * `d`.`harga_beli`)) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0)) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum(`j`.`komisi`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and `j`.`flag_debitor` and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0))) - `c`.`cash`) + `c`.`money_cash`) AS `kas_total`,if(((select count(0) from `tbl_jual` `j` where (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d'))) > 0),(select count(`j`.`kd_pasien`) from `tbl_jual` `j` where (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d'))),0) AS `pasien`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select count(`j`.`kd_pasien`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) AS `lunas`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select count(`j`.`kd_pasien`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) AS `pay_komisi`,`c`.`flag` AS `flag` from `tbl_cash` `c`) */;

/*View structure for view vw_cash_flow_before */

/*!50001 DROP TABLE IF EXISTS `vw_cash_flow_before` */;
/*!50001 DROP VIEW IF EXISTS `vw_cash_flow_before` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vw_cash_flow_before` AS (select `c`.`id` AS `id`,date_format(`c`.`tgl_cash`,'%Y-%m-%d') AS `tgl_cash`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `jual_sebelumnya`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (convert(date_format(`j`.`tgl_input`,'%Y-%m-%d') using latin1) = `j`.`tgl_bayar`) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (convert(date_format(`j`.`tgl_input`,'%Y-%m-%d') using latin1) = `j`.`tgl_bayar`) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `jual`,if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `beli_sebelumnya`,if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) AS `beli`,if(((select count(0) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`d`.`retur` * `d`.`harga_beli`)) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `retur`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum(`j`.`komisi`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) AS `komisi`,((if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0)) - (((if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) + if(((select count(0) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`d`.`retur` * `d`.`harga_beli`)) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where (`b`.`flag_supplier` and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0)) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum(`j`.`komisi`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0))) AS `laba`,`c`.`cash` AS `cash`,(((if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`j`.`tgl_input`,'%Y-%m-%d')) and (`j`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`j`.`bayar` + `j`.`piutang`)) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) - (((if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`c`.`tgl_cash`,'%Y-%m-%d') <> date_format(`b`.`tgl_input`,'%Y-%m-%d')) and (`b`.`tgl_bayar` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0) + if(((select count(0) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select sum((`b`.`bayar` + `b`.`hutang`)) from `tbl_beli` `b` where ((`b`.`flag_supplier` = 1) and (date_format(`b`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0)) + if(((select count(0) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum((`d`.`retur` * `d`.`harga_beli`)) from (`tbl_beli` `b` join `tbl_beli_details` `d` on((`d`.`no_beli` = `b`.`no_beli`))) where ((`b`.`flag_supplier` = 1) and (`d`.`tgl_retur` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0)) + if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))) > 0),(select sum(`j`.`komisi`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and `j`.`flag_debitor` and (`j`.`tgl_komisi` = convert(date_format(`c`.`tgl_cash`,'%Y-%m-%d') using latin1)))),0))) - `c`.`cash`) AS `kas`,if(((select count(0) from `tbl_jual` `j` where (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d'))) > 0),(select count(`j`.`kd_pasien`) from `tbl_jual` `j` where (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d'))),0) AS `pasien`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select count(`j`.`kd_pasien`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) AS `lunas`,if(((select count(0) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))) > 0),(select count(`j`.`kd_pasien`) from `tbl_jual` `j` where ((`j`.`flag_kreditor` = 1) and (`j`.`flag_debitor` = 1) and (date_format(`j`.`tgl_input`,'%Y-%m-%d') = date_format(`c`.`tgl_cash`,'%Y-%m-%d')))),0) AS `pay_komisi`,`c`.`flag` AS `flag` from `tbl_cash` `c`) */;

/*View structure for view vw_komisi */

/*!50001 DROP TABLE IF EXISTS `vw_komisi` */;
/*!50001 DROP VIEW IF EXISTS `vw_komisi` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vw_komisi` AS (select `j`.`tgl_komisi` AS `tgl_jual`,`d`.`kd_departement` AS `kd_departement`,`d`.`nm_departement` AS `nm_departement`,`d`.`an` AS `an`,`d`.`bn` AS `bn`,`d`.`rn` AS `rn`,sum(`j`.`komisi`) AS `komisi`,count(`j`.`kd_pasien`) AS `pasien` from (`tbl_jual` `j` join `tbl_departement` `d` on((`d`.`id_departement` = `j`.`id_departement`))) where (`j`.`flag_debitor` = 1) group by `j`.`tgl_komisi`,`j`.`id_departement`) */;

/*View structure for view vw_stok_min */

/*!50001 DROP TABLE IF EXISTS `vw_stok_min` */;
/*!50001 DROP VIEW IF EXISTS `vw_stok_min` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vw_stok_min` AS (select `o`.`id_obat` AS `id_obat`,`o`.`kd_obat` AS `kd_obat`,`o`.`nm_obat` AS `nm_obat`,`o`.`id_kategori` AS `id_kategori`,`k`.`nm_kategori` AS `nm_kategori`,`o`.`kemasan` AS `kemasan`,`o`.`harga_jual` AS `harga_jual`,`o`.`harga_beli` AS `harga_beli`,(`o`.`harga_jual` - `o`.`harga_beli`) AS `profit`,`o`.`stok` AS `stok`,if(((select count(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)),0) AS `beli`,if(((select count(`j`.`jumlah`) from `tbl_jual_details` `j` where (`j`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`j`.`jumlah`) from `tbl_jual_details` `j` where (`j`.`id_obat` = `o`.`id_obat`)),0) AS `jual`,if(((select count(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`b`.`retur`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)),0) AS `rugi`,(((if(((select count(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)),0) - if(((select count(`j`.`jumlah`) from `tbl_jual_details` `j` where (`j`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`j`.`jumlah`) from `tbl_jual_details` `j` where (`j`.`id_obat` = `o`.`id_obat`)),0)) + `o`.`stok`) - if(((select count(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`b`.`retur`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)),0)) AS `sisa`,`o`.`stok_min` AS `stok_min`,`o`.`tgl_input` AS `tgl_input`,`p`.`nm_pengguna` AS `nm_pengguna` from ((`tbl_obat` `o` left join `tbl_kategori` `k` on((`k`.`id_kategori` = `o`.`id_kategori`))) left join `tbl_pengguna` `p` on((`p`.`id` = `o`.`id_pengguna`))) where (`o`.`stok_min` >= (((if(((select count(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)),0) - if(((select count(`j`.`jumlah`) from `tbl_jual_details` `j` where (`j`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`j`.`jumlah`) from `tbl_jual_details` `j` where (`j`.`id_obat` = `o`.`id_obat`)),0)) + `o`.`stok`) - if(((select count(`b`.`jumlah`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)) > 0),(select sum(`b`.`retur`) from `tbl_beli_details` `b` where (`b`.`id_obat` = `o`.`id_obat`)),0)))) */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
