CREATE DATABASE  IF NOT EXISTS `library_system` /*!40100 DEFAULT CHARACTER SET latin1 */;
USE `library_system`;
-- MySQL dump 10.13  Distrib 5.6.13, for Win32 (x86)
--
-- Host: 127.0.0.1    Database: library_system
-- ------------------------------------------------------
-- Server version	5.6.12-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `items`
--

DROP TABLE IF EXISTS `items`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `items` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `NAME` varchar(255) DEFAULT NULL,
  `DESCRIPTION` varchar(2000) DEFAULT NULL,
  `CREATED_BY` varchar(50) DEFAULT NULL,
  `CREATED_DATE` datetime DEFAULT NULL,
  `LAST_MOD_BY` varchar(50) DEFAULT NULL,
  `LAST_MOD_DATE` datetime DEFAULT NULL,
  `ITEM_CODE` varchar(255) DEFAULT NULL,
  `LOCATION_ID` int(11) DEFAULT NULL,
  `ITEM_TYPE_ID` int(11) DEFAULT NULL,
  `CATEGORY_ID` int(11) DEFAULT NULL,
  `DONATED_BY` varchar(255) DEFAULT NULL,
  `AUTHOR` varchar(255) DEFAULT NULL,
  `STATUS` varchar(30) DEFAULT NULL,
  `PURCHASE_COST` int(11) DEFAULT NULL,
  `AQUISITION_TYPE` varchar(50) DEFAULT NULL,
  `PUBLISHER` varchar(255) DEFAULT NULL,
  `COPYRIGHT_YEAR` varchar(255) DEFAULT NULL,
  `VOLUME` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `ID_UNIQUE` (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=28 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `items`
--

LOCK TABLES `items` WRITE;
/*!40000 ALTER TABLE `items` DISABLE KEYS */;
INSERT INTO `items` VALUES (11,'Gospel 01','dcdsfd','admin','2014-01-18 13:19:20','admin','2014-02-01 19:44:43','SK001',9,1,2,'sdfsdf','sdsafc','Borrowed',232323232,NULL,NULL,NULL,NULL),(12,'Gospel 02','dcdsfd','admin','2014-01-18 13:20:11','admin','2014-01-24 14:27:43','SK002',9,1,2,'sdfsdf','sdsafc','Loss',2321123,NULL,NULL,NULL,NULL),(13,'Gospel 03','dcdsfd','admin','2014-01-18 13:20:24','admin','2014-02-04 14:49:00','SK003',9,8,2,'sdfsdf','sdsafc','Available',2312321,NULL,NULL,NULL,NULL),(14,'Math for Dummy Vol 1','dsdfdsf','admin','2014-01-18 13:21:09','admin','2014-02-07 17:32:43','MATH01',10,1,1,'Icha','dsdfdsf','Available',NULL,'Donated','','',NULL),(15,'Math for Dummy Vol 2','dsdfdsf','admin','2014-01-18 13:21:29','admin','2014-02-07 17:09:27','MATH02',10,1,1,NULL,'dsdfdsf','Obsolete',123232,'Purchased','','',NULL),(16,'Math for Dummy Vol 3','ssssssssss','admin','2014-01-18 13:21:39','admin','2014-02-01 19:43:34','MATH03',10,1,1,'sfdsfd','dsdfdsfsss','Loss',999999999,NULL,NULL,NULL,NULL),(17,'Gospel 01','dcdsfd','admin','2014-01-22 23:25:09','admin','2014-02-04 21:44:37','SK001',9,1,2,'sdfsdf','sdsafc','Available',55555,NULL,NULL,NULL,NULL),(18,'adasd','212','admin','2014-01-24 11:39:20','admin','2014-01-24 14:34:14','adasdsa',7,1,1,'1232312323.1232','adasdsa','Available',12355,NULL,NULL,NULL,NULL),(19,'sdsadsa','1232132','admin','2014-02-04 15:15:56','admin','2014-02-04 15:15:56','12345',18,8,1,NULL,'adssadas','Available',12323231,'Purchased','asdsad','123232','1'),(20,'Gospel 01','dcdsfd','admin','2014-02-07 17:09:38','admin','2014-02-07 17:09:38','SK001',9,1,2,NULL,'sdsafc','Available',55555,'Purchased','','',NULL),(21,'Gospel 01','dcdsfd','admin','2014-02-07 17:09:49','admin','2014-02-07 17:23:47','SK001',9,1,2,NULL,'sdsafc','Damaged',55555,'Purchased','','',NULL),(22,'adasd','212','admin','2014-02-07 17:10:00','admin','2014-02-07 17:10:00','adasdsa',7,1,1,NULL,'adasdsa','Available',12355,'Purchased','','',NULL),(23,'Gospel 01','dcdsfd','admin','2014-02-07 17:10:16','admin','2014-02-07 17:23:55','SK001',9,1,2,NULL,'sdsafc','Damaged',55555,'Purchased','','',NULL),(24,'sdsadsa','1232132','admin','2014-02-07 17:10:39','admin','2014-02-07 17:10:39','12345',18,8,1,NULL,'adssadas','Available',12323231,'Purchased','asdsad','123232','1'),(25,'Book of Math','xcsxc','admin','2014-02-07 17:11:23','admin','2014-02-07 17:11:23','UU24',7,1,1,NULL,'','Available',12323,'Purchased','asdsad','123232',NULL),(26,'Book of Math','xcsxc','admin','2014-02-07 17:11:49','admin','2014-02-07 17:11:49','UU24',7,1,1,NULL,'','Available',12323,'Purchased','asdsad','123232',NULL),(27,'Book of Math','xcsxc','admin','2014-02-07 17:12:02','admin','2014-02-07 17:23:03','UU24',7,1,1,NULL,'','Loss',12323,'Purchased','asdsad','123232',NULL);
/*!40000 ALTER TABLE `items` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-02-08  7:40:08
