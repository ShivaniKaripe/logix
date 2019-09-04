--
-- $Id: LogixRT_enable_ams_epm_integration.sql 120924 2018-02-23 05:18:09Z rs185275 $
--
-- version:7.3.1.138972.Official Build (SUSDAY10202)


-- added in 5.4b03
IF NOT EXISTS ( SELECT 1 FROM RemoteDataTypes WHERE RemoteDataTypeID = 1 AND EngineID = 2 )
Insert into RemoteDataTypes (RemoteDataTypeID, EngineID, Description) values (1, 2, 'Issued Rewards')
GO



IF NOT EXISTS ( SELECT 1 FROM RemoteDataStyles WHERE StyleID = 15 )
INSERT INTO RemoteDataStyles (StyleID, RemoteDataTypeID, Description, FilePathRequired) VALUES (15,1,'Preference as reward',0);
GO


IF NOT EXISTS ( SELECT 1 FROM remotedataoptions WHERE StyleID = 15 )
INSERT INTO RemoteDataOptions (RemoteDataTypeID, StyleID, Enabled, OutputPath, LastUpdate, CPEUpdateFlag, Deleted) VALUES (1, 15, 1, '', GETDATE(), 1, 0);
GO


-- Make IssuanceUpload Max Records Per Batch CPE_SystemOption visible
Update CPE_SystemOptions set Visible=1 where OptionID=71;
GO


-- Enable the integration from AMS to EPM
Update Integrations set Installed=1 where IntegrationID=1;
GO
