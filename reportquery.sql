declare @users TABLE(
[User OU Type] nvarchar(30),distinguishedName nvarchar(1000),SAMAccountName nvarchar(100),[Status] nvarchar(30),
givenName nvarchar(max),sn nvarchar(max),DisplayName nvarchar(max),title nvarchar(max),department nvarchar(max),telephoneNumber nvarchar(max),userPrincipalName nvarchar(max),lastlogon datetime,manager nvarchar(max)
)
declare @groups TABLE(
distinguishedName nvarchar(1000)
,Groups nvarchar(1000))
DECLARE @distinguishedName nvarchar(256), @Query NVARCHAR(1024)
  
 insert into @users select case
 when upper([distinguishedName]) like '%OU=%CONTRACT%' then 'Contractor OU'
 when upper([distinguishedName]) like '%OU=%DISABLED%' then 'Disabled OU'
 when upper([distinguishedName]) like '%OU=%SERVICE%' then 'Service Acct OU'
 when upper([distinguishedName]) like '%OU=%ADMIN%' then 'Admin OU'
 when upper([distinguishedName]) like '%OU=%USERS%' then 'Users OU'
 else '' end as [User OU Type]
 
 ,distinguishedName , SAMAccountName, [Status],givenName,sn,DisplayName,title,department,telephoneNumber,userPrincipalName,
 case when convert(bigint, lastlogon) = 0 then null when convert(bigint, lastlogon) > 2650467743999999716 then null else CAST((convert(bigint, lastlogon) / 864000000000.0 - 109207) AS DATETIME) end as lastlogon,
 manager 
 from (
 SELECT distinguishedName , SAMAccountName,givenName,sn,DisplayName,title,department,telephoneNumber,userPrincipalName,lastlogon,manager,'Enabled' as [Status]
   FROM OPENQUERY(ADSI, '
       SELECT distinguishedName , SAMAccountName,givenName,sn,DisplayName,title,department,telephoneNumber,userPrincipalName,lastlogon,manager
       FROM ''LDAP://DC=CONTOSO,DC=LOCAL''
       WHERE 
           objectCategory = ''Person'' 
           AND objectClass = ''user''
		   AND ''userAccountControl:1.2.840.113556.1.4.803:''<>2
   ') 
   union all
 SELECT distinguishedName , SAMAccountName,givenName,sn,DisplayName,title,department,telephoneNumber,userPrincipalName,lastlogon,manager,'Disabled' as [Status]
   FROM OPENQUERY(ADSI, '
       SELECT distinguishedName , SAMAccountName,givenName,sn,DisplayName,title,department,telephoneNumber,userPrincipalName,lastlogon,manager
       FROM ''LDAP://DC=CONTOSO,DC=LOCAL''
       WHERE 
           objectCategory = ''Person'' 
           AND objectClass = ''user''
		   AND ''userAccountControl:1.2.840.113556.1.4.803:''=2
   ')) as l


DECLARE cUsers CURSOR FOR
    SELECT distinguishedName from @users u

OPEN cUsers

FETCH NEXT FROM cUsers
INTO @distinguishedName

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @distinguishedName = REPLACE(@distinguishedName, '''', '''''')
    
    SET @Query = '
        SELECT ''' + @distinguishedName + ''', cn as Groups
        FROM OPENQUERY (ADSI, ''<LDAP://DC=CONTOSO,DC=LOCAL>;(&(objectClass=group)(member:1.2.840.113556.1.4.1941:=' 
       + @distinguishedName +'));cn, adspath;subtree'')'

    
        INSERT INTO @groups EXEC Master.sys.SP_EXECUTESQL @Query


    FETCH NEXT FROM cUsers
    INTO @distinguishedName
END

CLOSE cUsers
DEALLOCATE cUsers

select
[User OU Type]
,case when userPrincipalName is not null then userPrincipalName else SAMAccountName end as [E-Mail/Login]
,isnull(DisplayName,'') as DisplayName
,isnull(givenName,'') as [First Name]
,isnull(sn,'') as [Last Name]
,isnull(title,'') as [Job Title]
,isnull(department,'') as [Department]
--,isnull(manager,'') as Manager
,case when manager is not null then (select top 1 userPrincipalName from @users where distinguishedName = u.manager) else '' end as [Manager E-Mail]
,isnull(telephoneNumber,'') as telephoneNumber
,[Status] as [Account Status]
,lastlogon
,Groups as [Groups List]
from @users as u inner join
(select distinguishedName,string_agg([Groups],', ') as Groups from @groups group by distinguishedName) as g
on g.[distinguishedName] = u.[distinguishedName]
order by [User OU Type] desc,[Account Status] desc, [Last Name] asc