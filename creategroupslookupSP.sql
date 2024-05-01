USE [BCDWH]
GO

/****** Object:  StoredProcedure [dbo].[Get_AD_AllUsersWithGroups]    Script Date: 5/1/2024 5:57:03 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROC [dbo].[Get_AD_AllUsersWithGroups]
AS
DECLARE @Query NVARCHAR(1024), @Path NVARCHAR(1024)
DECLARE @distinguishedName nvarchar(256)
DECLARE @SAMAccountName nvarchar(256)

CREATE TABLE #users (distinguishedName nvarchar(1000), SAMAccountName nvarchar(100))
CREATE TABLE #results(SAMAccountName nvarchar(100), DistinguishedName nvarchar(1000), GroupName nvarchar(1000), ActiveDirectoryPath nvarchar(1000))

-- Get all the users from AD
SET @Query = '
   SELECT distinguishedName, SAMAccountName
   FROM OPENQUERY(ADSI, ''
       SELECT distinguishedName , SAMAccountName
       FROM ''''LDAP://DC=CONTOSO,DC=LOCAL''''
       WHERE 
           objectClass = ''''user'''' 
   '')
'



INSERT INTO #users
EXEC Master.sys.SP_EXECUTESQL @Query

-- For each user in #users, get a list of groups they belong to
DECLARE cUsers CURSOR FOR
    SELECT distinguishedName, SAMAccountName from dbo.#users u 
        order by u.distinguishedName

OPEN cUsers

FETCH NEXT FROM cUsers
INTO @distinguishedName, @SAMAccountName

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @distinguishedName = REPLACE(@distinguishedName, '''', '''''')
    SET @SAMAccountName = REPLACE(@SAMAccountName, '''', '''''')
    
    SET @Query = '
        INSERT INTO #results
        SELECT ''' + @SAMAccountName + ''', ''' + @distinguishedName + ''', cn as GroupName, AdsPath AS ActiveDirectoryPath
        FROM OPENQUERY (ADSI, ''<LDAP://DC=CONTOSO,DC=LOCAL>;(&(objectClass=group)(member:1.2.840.113556.1.4.1941:=' 
       + @distinguishedName +'));cn, adspath;subtree'')'

    EXEC Master.sys.SP_EXECUTESQL @Query  


    FETCH NEXT FROM cUsers
    INTO @distinguishedName, @SAMAccountName
END

CLOSE cUsers
DEALLOCATE cUsers

SELECT * FROM dbo.#results r

GO


