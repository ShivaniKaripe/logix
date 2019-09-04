-- $Id: LogixRT_EnableCPE.sql 29215 2011-04-06 19:47:56Z rob $

Declare @defaultEngine nvarchar(255)
set @defaultEngine = 'CPE'



print CURRENT_TIMESTAMP;
print 'Enabling ' + @defaultEngine + ' Engine'
print 'Using database ' + DB_NAME() + ' on ' + @@SERVERNAME;
print 'Running as user ' + SYSTEM_USER;

update PromoEngines set DefaultEngine = 0 WHERE DefaultEngine = 1 AND [Description] <> @defaultEngine; -- ensure no other engine is set as default
update PromoEngines set DefaultEngine = 1, Installed = 1 WHERE [Description] = @defaultEngine;         -- enable the default engine and make it default
update PromoEngines set Installed = 1                    WHERE [Description] = 'Website';              -- enable CustWeb connector

GO
