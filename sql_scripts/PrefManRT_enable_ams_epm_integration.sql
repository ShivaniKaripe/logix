--
-- $Id: PrefManRT_enable_ams_epm_integration.sql 55174 2012-09-07 15:55:10Z mark $
--
-- version:7.3.1.138972.Official Build (SUSDAY10202)
-- Enable the integration from EPM to AMS
Update Channels set Installed=1 where ChannelID=1;
GO


