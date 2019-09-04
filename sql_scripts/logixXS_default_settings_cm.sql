

UPDATE CardTypes with (RowLock) SET PaddingLength = 0, OnePerCustomer = 1, LastUpdate = GetDate();
GO

print 'LogixXS Default settings script for CM engine is complete'
