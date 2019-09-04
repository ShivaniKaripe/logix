-- $Id: users.sql 33871 2011-07-12 20:40:08Z rob $
-- Build Version: 7.3.1.138972

-- ************************************************************************************
-- This script assumes you've created the 4 databases and named them 
--  - LogixEX
--  - LogixRT
--  - LogixWH
--  - LogixXS
-- 
--  It creates the Copient_Logix login, and adds it to the 4 databases
-- ************************************************************************************
--IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name = N'Copient_Logix')
--    CREATE LOGIN [Copient_Logix] WITH PASSWORD = 'App0mattox', DEFAULT_DATABASE = [master], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF
--GO



use [master] 
IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name = N'LOGIX-TEST-VM\CopientSVC')
	CREATE LOGIN [LOGIX-TEST-VM\CopientSVC] FROM WINDOWS WITH DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english]
GO
EXEC sys.sp_addsrvrolemember @loginame = [LOGIX-TEST-VM\CopientSVC], @rolename = N'bulkadmin'
GO

use TestLogixEX
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientSVC')
    CREATE USER [LOGIX-TEST-VM\CopientSVC] FOR LOGIN [LOGIX-TEST-VM\CopientSVC] WITH DEFAULT_SCHEMA=[dbo]
GO
exec sp_addrolemember db_owner,  [LOGIX-TEST-VM\CopientSVC]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientSVC] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientSVC]
GO


use TestLogixRT
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientSVC')
    CREATE USER [LOGIX-TEST-VM\CopientSVC] FOR LOGIN [LOGIX-TEST-VM\CopientSVC] WITH DEFAULT_SCHEMA=[dbo]
GO

exec sp_addrolemember db_datareader,  [LOGIX-TEST-VM\CopientSVC]
exec sp_addrolemember db_datawriter,  [LOGIX-TEST-VM\CopientSVC]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientSVC] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientSVC]
GO



use TestLogixWH
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientSVC' )
    CREATE USER [LOGIX-TEST-VM\CopientSVC] FOR LOGIN [LOGIX-TEST-VM\CopientSVC] WITH DEFAULT_SCHEMA=[dbo]
GO
exec sp_addrolemember db_datareader,  [LOGIX-TEST-VM\CopientSVC]
exec sp_addrolemember db_datawriter,  [LOGIX-TEST-VM\CopientSVC]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientSVC] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientSVC]
GO

use TestLogixXS
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientSVC')
    CREATE USER [LOGIX-TEST-VM\CopientSVC] FOR LOGIN [LOGIX-TEST-VM\CopientSVC] WITH DEFAULT_SCHEMA=[dbo]
GO
exec sp_addrolemember db_datareader,  [LOGIX-TEST-VM\CopientSVC]
exec sp_addrolemember db_datawriter,  [LOGIX-TEST-VM\CopientSVC]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientSVC] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientSVC]
GO


use [master] 
IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name = N'LOGIX-TEST-VM\CopientWEB')
	CREATE LOGIN [LOGIX-TEST-VM\CopientWEB] FROM WINDOWS WITH DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english]
GO
EXEC sys.sp_addsrvrolemember @loginame = [LOGIX-TEST-VM\CopientWEB], @rolename = N'bulkadmin'
GO

use TestLogixEX
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientWEB')
    CREATE USER [LOGIX-TEST-VM\CopientWEB] FOR LOGIN [LOGIX-TEST-VM\CopientWEB] WITH DEFAULT_SCHEMA=[dbo]
GO
exec sp_addrolemember db_owner,  [LOGIX-TEST-VM\CopientWEB]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientWEB] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientWEB]
GO


use TestLogixRT
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientWEB')
    CREATE USER [LOGIX-TEST-VM\CopientWEB] FOR LOGIN [LOGIX-TEST-VM\CopientWEB] WITH DEFAULT_SCHEMA=[dbo]
GO

exec sp_addrolemember db_datareader,  [LOGIX-TEST-VM\CopientWEB]
exec sp_addrolemember db_datawriter,  [LOGIX-TEST-VM\CopientWEB]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientWEB] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientWEB]
GO



use TestLogixWH
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientWEB' )
    CREATE USER [LOGIX-TEST-VM\CopientWEB] FOR LOGIN [LOGIX-TEST-VM\CopientWEB] WITH DEFAULT_SCHEMA=[dbo]
GO
exec sp_addrolemember db_datareader,  [LOGIX-TEST-VM\CopientWEB]
exec sp_addrolemember db_datawriter,  [LOGIX-TEST-VM\CopientWEB]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientWEB] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientWEB]
GO

use TestLogixXS
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'LOGIX-TEST-VM\CopientWEB')
    CREATE USER [LOGIX-TEST-VM\CopientWEB] FOR LOGIN [LOGIX-TEST-VM\CopientWEB] WITH DEFAULT_SCHEMA=[dbo]
GO
exec sp_addrolemember db_datareader,  [LOGIX-TEST-VM\CopientWEB]
exec sp_addrolemember db_datawriter,  [LOGIX-TEST-VM\CopientWEB]
GRANT CONNECT TO [LOGIX-TEST-VM\CopientWEB] AS [dbo]
exec sp_addrolemember 'Copient_Logix_uspRole', [LOGIX-TEST-VM\CopientWEB]
GO
