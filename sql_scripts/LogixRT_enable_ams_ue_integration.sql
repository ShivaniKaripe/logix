-- $Id: LogixRT_EnableUE.sql 37184 2011-09-14 14:26:02Z mark $

Declare @defaultEngine nvarchar(255)
set @defaultEngine = 'UE'



print CURRENT_TIMESTAMP;
print 'Enabling ' + @defaultEngine + ' Engine'
print 'Using database ' + DB_NAME() + ' on ' + @@SERVERNAME;
print 'Running as user ' + SYSTEM_USER;

update PromoEngines set DefaultEngine = 0 WHERE DefaultEngine = 1 AND [Description] <> @defaultEngine; -- ensure no other engine is set as default
update PromoEngines set DefaultEngine = 1, Installed = 1 WHERE [Description] = @defaultEngine;         -- enable the default engine and make it default
update PromoEngines set Installed = 1                    WHERE [Description] = 'Website';              -- enable CustWeb connector

IF EXISTS(select * from [dbo].[PrinterTypes] where PrinterTypeID=10)
	print 'Enabling printer(s) for the ' + @defaultEngine + ' Engine';
	update PrinterTypes set Installed=1 where PrinterTypeID=10;
GO
	
IF EXISTS(select 1 from [dbo].[LastSync] where AppID=1001)
	Update dbo.LastSync Set SendAlert=1,LogAvailable=1 Where AppID=1001; 
GO
