-- $Id: LogixRT_EnableCM.sql 29215 2011-04-06 19:47:56Z rob $

Declare @defaultEngine nvarchar(255)
set @defaultEngine = 'CM'



print CURRENT_TIMESTAMP;
print 'Enabling ' + @defaultEngine + ' Engine'
print 'Using database ' + DB_NAME() + ' on ' + @@SERVERNAME;
print 'Running as user ' + SYSTEM_USER;

update PromoEngines set DefaultEngine = 0 WHERE DefaultEngine = 1 AND [Description] <> @defaultEngine; -- ensure no other engine is set as default
update PromoEngines set DefaultEngine = 1, Installed = 1 WHERE [Description] = @defaultEngine;         -- enable the default engine and make it default

IF     NOT EXISTS (SELECT TOP 1 * FROM PromoEngines WHERE  [EngineID] =  2  AND   [Installed] =  1  ) -- check if either the CPE or UE engines are installed
   AND NOT EXISTS (SELECT TOP 1 * FROM PromoEngines WHERE  [EngineID] =  9  AND   [Installed] =  1  )

update PromoEngines set Installed = 0 WHERE [EngineID] = 3; -- set the website engine to disabled if not

GO

