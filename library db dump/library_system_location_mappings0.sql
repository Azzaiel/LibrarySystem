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
-- Table structure for table `location_mappings`
--

DROP TABLE IF EXISTS `location_mappings`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `location_mappings` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `NAME` varchar(255) DEFAULT NULL,
  `FILE_NAME` varchar(2000) DEFAULT NULL,
  `CREATED_BY` varchar(50) DEFAULT NULL,
  `CREATED_DATE` datetime DEFAULT NULL,
  `LAST_MOD_BY` varchar(50) DEFAULT NULL,
  `LAST_MOD_DATE` datetime DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `ID_UNIQUE` (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=17 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `location_mappings`
--

LOCK TABLES `location_mappings` WRITE;
/*!40000 ALTER TABLE `location_mappings` DISABLE KEYS */;
INSERT INTO `location_mappings` VALUES (1,'loc1','loc1.jpg','System','2014-01-04 19:27:39','System','2014-01-04 21:57:21'),(3,'loc2','loc2.jpg','System','2014-01-04 19:38:05','System','2014-01-05 13:22:46'),(5,'loc12','loc12.jpg','System','2014-01-04 19:43:57','System','2014-01-04 19:46:56'),(6,'loc10','loc10.jpg','System','2014-01-04 19:45:45','System','2014-01-05 13:22:58'),(7,'loc14','loc14.jpg','System','2014-01-04 19:45:31','System','2014-01-04 19:47:37'),(9,'loc3','loc3.jpg','System','2014-01-04 19:50:08','System','2014-01-04 19:50:49'),(10,'loc5','loc5.jpg','System','2014-01-04 19:50:54','System','2014-01-04 19:54:22'),(14,'loc4','loc4.jpg','System','2014-01-04 19:54:06','System','2014-01-05 13:24:18'),(15,'Missing','','System','2014-01-04 20:04:00','System','2014-01-05 13:24:13'),(16,'loc6','loc6.jpg','System','2014-01-05 14:54:26',NULL,NULL);
/*!40000 ALTER TABLE `location_mappings` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-08 22:04:09
