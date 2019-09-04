
--Single Sign-On -(**Required**)
DECLARE @Enable_SSO VARCHAR(40);
SET @Enable_SSO = ''; -- 0 for Disable/ 1 for Enable

--NEP Login Page (SSO) -(**Required**)
DECLARE @NEP_Login_Url VARCHAR(100);
SET @NEP_Login_Url = '';

--NEP Portal URL (SSO) -(**Required**)
DECLARE @NEP_Portal_Url VARCHAR(100);
SET @NEP_Portal_Url = '';


print CURRENT_TIMESTAMP;
print 'Enabling SSO'
print 'Using database ' + DB_NAME() + ' on ' + @@SERVERNAME;
print 'Running as user ' + SYSTEM_USER;

IF (@Enable_SSO = '' OR @NEP_Login_Url = '' OR @NEP_Login_Url = '')
    print 'Please Set all (**Required**) fields to proceed'
ELSE IF NOT EXISTS ( SELECT TOP 1 * FROM SystemOptions WHERE  [OptionID] = 326)
    print 'Upgrade to Logix 7.0 or above to Enable/Disable SSO'
ELSE
	BEGIN
		UPDATE SystemOptions SET OptionValue = @Enable_SSO WHERE [OptionID] = 326
		UPDATE SystemOptions SET OptionValue = @NEP_Login_Url WHERE [OptionID] = 327
		UPDATE SystemOptions SET OptionValue = @NEP_Portal_Url WHERE [OptionID] = 328

		print 'Completed Enabling SSO Succesfully ';
	END
GO

