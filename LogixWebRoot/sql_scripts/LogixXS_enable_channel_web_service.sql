

print CURRENT_TIMESTAMP;
print 'Enable channel web service properties'
print 'Using database ' + DB_NAME() + ' on ' + @@SERVERNAME;
print 'Running as user ' + SYSTEM_USER;

--Enable login for altid and customer card
update ChannelCustIDTypes set LogonEnabled = 1 WHERE CardTypeID = 3;
update ChannelCustIDTypes set LogonEnabled = 1 WHERE CardTypeID = 0;

--Disable register flag for altid and customer card
update ChannelCustIDTypes set RegisterEnabled = 0 WHERE CardTypeID = 3;
update ChannelCustIDTypes set RegisterEnabled = 0 WHERE CardTypeID = 0;

--Enable PINSettingID for all card types, set value to 2 SHARED
--UNDEFINED = 0
--NOT_USED = 1 <- No pin/password required
--SHARED = 2 <- pin/password shared across all cardtypes
--EXPLICIT = 3 <- pin/password distinct across all cardtypes
--EXPLICT_READ_ONLY = 4

update CardTypes set PinSettingID = 2;

GO

