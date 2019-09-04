--Create the new Catalina Coupon passthru
IF NOT EXISTS ( SELECT TOP 1 * FROM PassThruRewards WHERE [PassThruRewardID] = 6 )
    INSERT INTO PassThruRewards (PassThruRewardID, Name,             PhraseID, LSInterfaceID, ActionTypeID, DataTemplate, Presentation, PresentationPhraseID, LastUpdate ) 
                         VALUES (6,                'Catalina Coupon', NULL,    2,             7, 
'<Coupon>
  <Type>^61;</Type>
</Coupon>', '<table border=0 cellpadding=2 cellspacing=0>
  <tr><td align=right><label>MCLU:</label> </td><td> ^61; </td></tr>
</table>', NULL, GETDATE() );
GO
--  Update the ActionTypeID to be 7 for CPE's sake
UPDATE PassThruRewards SET ActionTypeID = 7 WHERE PassThruRewardID = 6

IF NOT EXISTS ( SELECT TOP 1 * FROM PassThruPresTags WHERE [PassThruPresTagID] = 61 )
    INSERT INTO PassThruPresTags (PassThruPresTagID, TokenValueSelector, ReplacementText, ParamName, ParamDataTypeID, MaxLength) VALUES (61, 0, '^', 'MCLU', 1, 7);
GO


--Create RemoteData entries for passthru reward type
IF NOT EXISTS ( SELECT TOP 1 * FROM RemoteDataStyles WHERE [StyleID] = 12 )
    insert into RemoteDataStyles (StyleID, RemoteDataTypeID, Description, FilePathRequired) values (12, 1, 'Passthru', 0);
GO
IF NOT EXISTS ( SELECT TOP 1 * FROM RemoteDataOptions WHERE [StyleID] = 12 )
    insert into RemoteDataOptions (RemoteDataTypeID, StyleID, Enabled, OutputPath, LastUpdate, CPEUpdateFlag, Deleted) values (1, 12, 1, '', GETDATE(), 0, 0);
GO
