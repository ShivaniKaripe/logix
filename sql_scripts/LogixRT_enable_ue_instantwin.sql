-- $Id: LogixRT_enable_ue_instantwin.sql 29215 2011-04-06 19:47:56Z rob $

-- Enabling instant win condition for UE
UPDATE PromoEngineComponentTypes SET Enabled=1 WHERE EngineID=9 and ComponentTypeID=1 and LinkID=8;
GO
