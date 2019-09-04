/*
--	This index is not needed in base code. So Uday Sangepu is commeting it out. At some point this file needs to be removed from base code and all branches.

USE [LogixRT]
GO

--Dropping the NonClusetered Indexes from the ActivityLog table 

IF EXISTS(SELECT * FROM sys.indexes WHERE name = 'IX_ActivityLog_ActivityTypeID' AND object_id = OBJECT_ID('dbo.ActivityLog'))
BEGIN
	DROP INDEX IX_ActivityLog_ActivityTypeID ON dbo.ActivityLog
END
GO

IF EXISTS(SELECT * FROM sys.indexes WHERE name = 'IX_ActivityLog_AdminID' AND object_id = OBJECT_ID('dbo.ActivityLog'))
BEGIN
	DROP INDEX IX_ActivityLog_AdminID ON dbo.ActivityLog
END
GO

IF EXISTS(SELECT * FROM sys.indexes WHERE name = 'IX_ActivityLog_LinkID' AND object_id = OBJECT_ID('dbo.ActivityLog'))
BEGIN
	DROP INDEX IX_ActivityLog_LinkID ON dbo.ActivityLog
END
GO

IF EXISTS(SELECT * FROM sys.indexes WHERE name = 'IX_ActivityLog_SessionID' AND object_id = OBJECT_ID('dbo.ActivityLog'))
BEGIN
	DROP INDEX IX_ActivityLog_SessionID ON dbo.ActivityLog
END	
GO

--Adding the Covering Index NonClustered on the ActivityLog table
IF NOT EXISTS(SELECT * FROM sys.indexes WHERE name = 'IX_ActivityLog_LinkID_ActivityTypeID_AdminID_SessionID' AND object_id = OBJECT_ID('dbo.ActivityLog'))
BEGIN
	CREATE NONCLUSTERED INDEX [IX_ActivityLog_LinkID_ActivityTypeID_AdminID_SessionID] ON [dbo].[ActivityLog] 
(
	[LinkID] ASC,
	[ActivityTypeID] ASC,
	[AdminID] ASC,
	[SessionID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 80) ON [PRIMARY]

END	
GO
*/