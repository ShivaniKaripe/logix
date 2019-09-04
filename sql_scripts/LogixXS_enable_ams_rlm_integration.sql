-- LogixXS ams-rlm integration script

-- 20130417.AS: 2 new card types for AMS-RLM integration

if not exists (select * from CardTypes where CardTypeID=9) 
  INSERT INTO [dbo].[CardTypes] ([CardTypeID], [Description], [CustTypeID], [PhraseID], [ExtCardTypeID], [OnePerCustomer], [LastUpdate], [PhraseTerm], [PaddingLength], [MaxIDLength], [CreateCardsInUI], [DeleteCardsInUI], [UpdateCardsInUI], [CanNotRemoveLastID], [UpdateCardStatusInUI], [NumericOnly]) VALUES (9, N'Account Number', 0, 6815, '9', 0, GETDATE(), 'term.accountnumber', 0, 19, 1, 1, 1, 0, 1, 0);
GO
if not exists (select * from CardTypes where CardTypeID=10) 
  INSERT INTO [dbo].[CardTypes] ([CardTypeID], [Description], [CustTypeID], [PhraseID], [ExtCardTypeID], [OnePerCustomer], [LastUpdate], [PhraseTerm], [PaddingLength], [MaxIDLength], [CreateCardsInUI], [DeleteCardsInUI], [UpdateCardsInUI], [CanNotRemoveLastID], [UpdateCardStatusInUI], [NumericOnly]) VALUES (10, N'Barcode', 0, 6816, '10', 0, GETDATE(), 'term.cardtypebarcode', 0, 19, 1, 1, 1, 0, 1, 0);
GO