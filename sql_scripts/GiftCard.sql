-- $Id: LogixRT_EnableCPE.sql 29215 2011-04-06 19:47:56Z rob $

-- UE-specific passthru per BZ2987:

DELETE FROM PassThruRewards WHERE PassThruRewardID = 14;
GO
INSERT INTO [dbo].[PassThruRewards] ([PassThruRewardID], [Name], [LSInterfaceID], [ActionTypeID], [DataTemplate], [Presentation], [LastUpdate])
VALUES (14, 'Gift card', 2, 100,
'<GiftCard>
  <MonetaryValue>^36;</MonetaryValue>
  <NameOfCard>^37;</NameOfCard>
  <BuyDescription>^38;</BuyDescription>
  <ChargebackDepartment>^39;</ChargebackDepartment>
</GiftCard>',
'<table border=0 cellpadding=2 cellspacing=0>  <tr><td align=right>Monetary value:</td><td> ^36; </td></tr>  <tr><td align=right>Name of card:</td><td> ^37; </td></tr>  <tr><td align=right>Buy description:</td><td> ^38; </td></tr>  <tr><td align=right>Chargeback department:</td><td> ^39; </td></tr>  </table>', GETDATE());
GO

DELETE FROM PassThruPresTags WHERE PassThruPresTagID in (36,37,38,39);
GO
INSERT INTO PassThruPresTags (PassThruPresTagID, TokenValueSelector, ReplacementText, ParamName, ParamDataTypeID, MaxLength)
VALUES (36, 0, '^', 'MonetaryValue', 2, 8);
INSERT INTO PassThruPresTags (PassThruPresTagID, TokenValueSelector, ReplacementText, ParamName, ParamDataTypeID, MaxLength)
VALUES (37, 0, '^', 'NameOfCard', 4, 256);
INSERT INTO PassThruPresTags (PassThruPresTagID, TokenValueSelector, ReplacementText, ParamName, ParamDataTypeID, MaxLength)
VALUES (38, 0, '^', 'BuyDescription', 4, 255);
INSERT INTO PassThruPresTags (PassThruPresTagID, TokenValueSelector, ReplacementText, ParamName, ParamDataTypeID, MaxLength)
VALUES (39, 1, '^', 'ChargebackDepartment', 1, NULL);
GO

DELETE FROM PassThruTokenValues WHERE PKID in (41);
GO
INSERT INTO PassThruTokenValues (PKID, PassThruPresTagID, DisplayOrder, OptionDescription, OptionValue, SourceDatabase, SourceTable, SourceWhereClause, SourceOrderByClause) VALUES (41, 39, 1, 'Name', 'ChargeBackDeptID', 'RT', 'ChargeBackDepts', 'Deleted=0 and ChargeBackDeptID not in (10,14)', 'Name');
GO

DELETE FROM PromoEnginePassThrus WHERE PassThruRewardID = 14 AND EngineID = 2 AND ComponentTypeID = 2;
GO
INSERT INTO PromoEnginePassThrus (PassThruRewardID, EngineID, ComponentTypeID, Enabled, Singular)
VALUES (14, 2, 2, 1, 0)
GO