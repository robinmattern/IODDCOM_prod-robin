CREATE 
    ALGORITHM = UNDEFINED 
    DEFINER = `root`@`localhost` 
    SQL SECURITY DEFINER
VIEW `iodd`.`members_projects_view` AS
    SELECT 
        `iodd`.`members`.`Id` AS `Id`,
        `iodd`.`members`.`MemberNo` AS `MemberNo`,
        `iodd`.`members`.`TitleName` AS `TitleName`,
        `iodd`.`members`.`FirstName` AS `FirstName`,
        `iodd`.`members`.`MiddleName` AS `MiddleName`,
        `iodd`.`members`.`LastName` AS `LastName`,
        `iodd`.`members`.`PostName` AS `PostName`,
        `iodd`.`members`.`Company` AS `Company`,
        `iodd`.`members`.`Address1` AS `Address1`,
        `iodd`.`members`.`Address2` AS `Address2`,
        `iodd`.`members`.`City` AS `City`,
        `iodd`.`members`.`State` AS `State`,
        `iodd`.`members`.`Zip` AS `Zip`,
        `iodd`.`members`.`Country` AS `Country`,
        `iodd`.`members`.`Phone1` AS `Phone1`,
        `iodd`.`members`.`Phone2` AS `Phone2`,
        `iodd`.`members`.`Fax` AS `Fax`,
        `iodd`.`members`.`WebSite` AS `WebSite`,
        `iodd`.`members`.`Email` AS `Email`,
        `iodd`.`members`.`Skills` AS `Skills`,
        `iodd`.`members`.`Active` AS `Active`,
        `iodd`.`members`.`Bio` AS `Bio`,
        `iodd`.`members_projects`.`Sort` AS `Sort`,
        `iodd`.`members_projects`.`Role` AS `Role`,
        `iodd`.`members_projects`.`Duration` AS `Duration`,
        `iodd`.`members_projects`.`Dates` AS `Dates`,
        `iodd`.`projects`.`Name` AS `ProjectName`,
        `iodd`.`projects`.`Client` AS `Client`,
        `iodd`.`projects`.`ClientWeb` AS `ClientWeb`,
        `iodd`.`projects`.`ProjectWeb` AS `ProjectWeb`,
        `iodd`.`projects`.`Location` AS `Location`,
        `iodd`.`projects`.`ProjectType` AS `ProjectType`,
        `iodd`.`projects`.`Industry` AS `Industry`,
        `iodd`.`projects`.`Description` AS `Description`
    FROM
        ((`iodd`.`members`
        JOIN `iodd`.`members_projects` ON ((`iodd`.`members`.`Id` = `iodd`.`members_projects`.`MemberId`)))
        JOIN `iodd`.`projects` ON ((`iodd`.`members_projects`.`ProjectId` = `iodd`.`projects`.`Id`)))
    ORDER BY `iodd`.`members`.`LastName` , `iodd`.`members`.`FirstName` , `iodd`.`members_projects`.`Dates` DESC