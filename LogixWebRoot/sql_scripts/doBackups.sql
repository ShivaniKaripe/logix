-- $Id: doBackups.sql 25519 2010-12-29 19:30:12Z rob $
-- Build Version: 7.3.1.138972

BACKUP DATABASE [LogixEX] TO  DISK = N'C:\dev\Logix\trunk\SQLscripts\complete\LogixEX.bak' 
    WITH NOFORMAT, NOINIT,  NAME = N'LogixEX-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10
GO

BACKUP DATABASE [LogixRT] TO  DISK = N'C:\dev\Logix\trunk\SQLscripts\complete\LogixRT.bak' 
    WITH NOFORMAT, NOINIT,  NAME = N'LogixRT-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10
GO

BACKUP DATABASE [LogixWH] TO  DISK = N'C:\dev\Logix\trunk\SQLscripts\complete\LogixWH.bak' 
    WITH NOFORMAT, NOINIT,  NAME = N'LogixWH-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10
GO

BACKUP DATABASE [LogixXS] TO  DISK = N'C:\dev\Logix\trunk\SQLscripts\complete\LogixXS.bak' 
    WITH NOFORMAT, NOINIT,  NAME = N'LogixXS-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10
GO
