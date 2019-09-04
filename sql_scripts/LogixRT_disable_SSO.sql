print CURRENT_TIMESTAMP;
print 'Disabling SSO'
print 'Using database ' + DB_NAME() + ' on ' + @@SERVERNAME;
print 'Running as user ' + SYSTEM_USER;

IF NOT EXISTS ( SELECT TOP 1 * FROM SystemOptions WHERE  [OptionID] = 326)
    print 'Upgrade to Logix 7.0 or above to Enable/Disable SSO'
ELSE
	BEGIN
		UPDATE SystemOptions SET OptionValue = '' WHERE [OptionID] = 326
		UPDATE SystemOptions SET OptionValue = '' WHERE [OptionID] = 327
		UPDATE SystemOptions SET OptionValue = '' WHERE [OptionID] = 328

		print 'Completed Disabling SSO Succesfully ';
	END
GO

