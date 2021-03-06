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
-- Table structure for table `students`
--

DROP TABLE IF EXISTS `students`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `students` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `LRN` varchar(200) NOT NULL DEFAULT '',
  `FIRST_NAME` varchar(200) NOT NULL,
  `MIDDLE_NAME` varchar(200) NOT NULL,
  `LAST_NAME` varchar(200) NOT NULL,
  `SECTION_ID` int(11) NOT NULL,
  `CREATED_BY` varchar(255) NOT NULL,
  `CREATED_DATE` datetime DEFAULT NULL,
  `LAST_MOD_BY` varchar(255) DEFAULT NULL,
  `LAST_MOD_DATE` datetime DEFAULT NULL,
  `status` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`ID`,`LRN`),
  UNIQUE KEY `ID_UNIQUE` (`ID`),
  KEY `IDssdsd_idx` (`SECTION_ID`),
  CONSTRAINT `SECTION_ID` FOREIGN KEY (`SECTION_ID`) REFERENCES `sections` (`ID`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `students`
--

LOCK TABLES `students` WRITE;
/*!40000 ALTER TABLE `students` DISABLE KEYS */;
INSERT INTO `students` VALUES (1,'182532','James','Platon','Guerra',10,'System','2013-12-28 12:24:54','admin','2014-02-06 21:38:32','Not Enrolled'),(4,'123123','asdasd','dasdasdas','asdas',16,'System','2013-12-29 10:43:57','System','2013-12-29 11:50:27','Enrolled'),(6,'22589245','Jomar','D','Abaygar',16,'System','2013-12-29 10:50:39','System','2013-12-29 11:07:19','Enrolled'),(7,'98865','Jhonel','S','Abaygar',16,'System','2013-12-29 10:51:24','System','2013-12-29 11:09:31','Enrolled'),(9,'123232','asdasd','dasdasdas','adasd',17,'System','2013-12-29 11:08:11','System','2013-12-29 11:47:40','Enrolled'),(12,'32435','fhgfh','B','Boinky',16,'System','2013-12-29 11:48:31','System','2014-01-12 15:59:43','Enrolled'),(15,'5135151','francis','dg','adasd',17,'System','2014-01-12 16:00:27','System','2014-01-12 16:01:32','Enrolled'),(16,'123232','baog','dg','adasd',17,'System','2014-01-12 16:00:52','System','2014-01-12 16:00:52','Enrolled'),(17,'123','sss','dsd','wsedwe',10,'System','2014-01-13 15:46:17','System','2014-01-13 15:46:17','Enrolled'),(18,'69','Adrian','BRU','No',27,'admin','2014-01-17 14:57:36','admin','2014-02-06 21:38:11','Not Enrolled');
/*!40000 ALTER TABLE `students` ENABLE KEYS */;
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
