CREATE 
    ALGORITHM = UNDEFINED 
    DEFINER = `root`@`localhost` 
    SQL SECURITY DEFINER
VIEW `iodd`.`projects_view` AS
    SELECT 
        `iodd`.`projects`.`Id` AS `Id`,
        `iodd`.`projects`.`Name` AS `Name`,
        `iodd`.`projects`.`Client` AS `Client`,
        `iodd`.`projects`.`ClientWeb` AS `ClientWeb`,
        `iodd`.`projects`.`ProjectWeb` AS `ProjectWeb`,
        `iodd`.`projects`.`Location` AS `Location`,
        `iodd`.`projects`.`ProjectType` AS `ProjectType`,
        `iodd`.`projects`.`Industry` AS `Industry`,
        `iodd`.`projects`.`Description` AS `Description`
    FROM
        `iodd`.`projects`