
-- $Id: addCopient_Logix_uspRole.sql 29656 2011-04-13 17:52:04Z rob $
-- Build Version: 7.3.1.138972

print 'Adding Copient_Logix_uspRole'
GO
                                                    
sp_addrole 'Copient_Logix_uspRole'
GO

sp_addrolemember 'Copient_Logix_uspRole', 'Copient_Logix'
GO
