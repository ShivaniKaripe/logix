<script runat="server">
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
  Dim CopientFileName As String = "sv-adjust-code.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""

  Function AdjustStoredValue(ByVal AdminUserID As Long, ByVal ProgramID As Long, ByVal CustomerPK As Long, _
                        ByVal Adjust As String, Optional ByVal RevokeLocalID As Long = 0) As String
    Dim dt As DataTable
    Dim dtUnits As DataTable = Nothing
    Dim Units() As Decimal = Nothing
    Dim AdjustQty As Long
    Dim Value As Decimal
    Dim ExpirePeriod As Long = 0
    Dim ExpireDate As String = ""
    Dim StoreInUnits As Boolean = False
    Dim i As Long = 0
    Dim j As Long = 0
    Dim LogText As String = ""
    Dim ExternalIDs As String()
    Dim ExternalID As String = ""
    Dim LocalIDs As Long()
    Dim LocalID As Long
    Dim RetMsg As String = ""
    Dim SVTypeID As Long = 1
    Dim UOMLimit As Long = 1
    Dim ProgramName As String = ""
    
    Try
      If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
      If (MyCommon.LXSadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixXS()
            
      MyCommon.QueryStr = "select Name, Value, ExpirePeriod, OneUnitPerRec, SVExpireType, SVExpirePeriodType, " & _
                          "ExpireTOD, ExpireDate, SVTypeID, UnitOfMeasureLimit from StoredValuePrograms with (NoLock) " & _
                          "where SVProgramID=" & ProgramID & " and Deleted=0;"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        ProgramName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
        Value = MyCommon.NZ(dt.Rows(0).Item("Value"), 0D)
        ExpirePeriod = MyCommon.NZ(dt.Rows(0).Item("ExpirePeriod"), 0)
        StoreInUnits = MyCommon.NZ(dt.Rows(0).Item("OneUnitPerRec"), False)
        ExpireDate = GetExpireDate(ProgramID)
        SVTypeID = MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 1)
        If (SVTypeID = 3 Or SVTypeID = 5) Then
          UOMLimit = MyCommon.NZ(dt.Rows(0).Item("UnitOfMeasureLimit"), 1)
        Else
          UOMLimit = 1
        End If
      End If
      
      If (StoreInUnits) Then
        'create table needed for batch inserts
        dtUnits = New DataTable()
        dtUnits.Columns.Add("TableNum", Type.GetType("System.String"))
        dtUnits.Columns.Add("Operation", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col1", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col2", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col3", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col4", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col5", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col6", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col7", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col8", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col9", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col10", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col11", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col12", Type.GetType("System.String"))
        dtUnits.Columns.Add("Col13", Type.GetType("System.String"))
        dtUnits.Columns.Add("LocalServerID", Type.GetType("System.Int32"))
        dtUnits.Columns.Add("LocationID", Type.GetType("System.Int64"))
        dtUnits.Columns.Add("WaitingACK", Type.GetType("System.Int32"))
        dtUnits.Columns.Add("AdminUserID", Type.GetType("System.Int32"))
        dtUnits.Columns.Add("Replayed", Type.GetType("System.Bit"))
        dtUnits.Columns.Add("ReplayedDate", Type.GetType("System.DateTime"))
      End If
      
      If (Long.TryParse(Adjust, AdjustQty) OrElse RevokeLocalID > 0) Then
        If (StoreInUnits) Then
          If (AdjustQty > 0) Then
            dtUnits.Clear()
            ReDim ExternalIDs(AdjustQty)
            ExternalIDs = GetExternalIDs(AdjustQty)
            ReDim LocalIDs(AdjustQty)
            LocalIDs = GetLocalIDs(AdjustQty)
            For j = 0 To AdjustQty - 1
              dtUnits.Rows.Add(New Object() {"10", "1", LocalIDs(j).ToString, "-9", ProgramID.ToString, 0, CustomerPK.ToString, 1, 0, _
                                              Value.ToString, Now.ToString("yyyy-MM-ddThh:mm:ss"), _
                                              ExpireDate, ExternalIDs(j), (Value * UOMLimit).ToString, 0, _
                                              -9, -9, 0, AdminUserID})
            Next
            BatchInsert(dtUnits, dtUnits.Rows.Count)
          Else
            RevokeStoredValue(AdminUserID, CustomerPK, ProgramID, 0, AdjustQty, UOMLimit, RevokeLocalID)
          End If
        Else
          If (AdjustQty > 0) Then
            ExternalID = GetExternalID()
            LocalID = GetLocalID()
            ' Add the adjustment to the pending queue for the agent to process it on the next pass  
            MyCommon.QueryStr = "pa_CPE_TU_InsertData_SV"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@TableNum", SqlDbType.VarChar, 4).Value = "10"
            MyCommon.LXSsp.Parameters.Add("@Operation", SqlDbType.VarChar, 2).Value = "1"
            MyCommon.LXSsp.Parameters.Add("@Col1", SqlDbType.VarChar, 255).Value = LocalID.ToString ' Local ID
            MyCommon.LXSsp.Parameters.Add("@Col2", SqlDbType.VarChar, 255).Value = "-9" ' Server Serial
            MyCommon.LXSsp.Parameters.Add("@Col3", SqlDbType.VarChar, 255).Value = ProgramID.ToString 'SVProgramID
            MyCommon.LXSsp.Parameters.Add("@Col4", SqlDbType.VarChar, 255).Value = 0 'OfferID
            MyCommon.LXSsp.Parameters.Add("@Col5", SqlDbType.VarChar, 255).Value = CustomerPK.ToString ' CustomerPK
            MyCommon.LXSsp.Parameters.Add("@Col6", SqlDbType.VarChar, 255).Value = AdjustQty ' QtyEarned
            MyCommon.LXSsp.Parameters.Add("@Col7", SqlDbType.VarChar, 255).Value = "0" ' QtyUsed
            MyCommon.LXSsp.Parameters.Add("@Col8", SqlDbType.VarChar, 255).Value = Value.ToString 'Value
            MyCommon.LXSsp.Parameters.Add("@Col9", SqlDbType.VarChar, 255).Value = Now.ToString("yyyy-MM-ddTHH:mm:ss") ' EarnedDate
            MyCommon.LXSsp.Parameters.Add("@Col10", SqlDbType.VarChar, 255).Value = ExpireDate ' Expire Date
            MyCommon.LXSsp.Parameters.Add("@Col11", SqlDbType.VarChar, 255).Value = ExternalID ' ExternalID
            MyCommon.LXSsp.Parameters.Add("@Col12", SqlDbType.VarChar, 255).Value = (AdjustQty * Value * UOMLimit).ToString ' TotalValueEarned
            MyCommon.LXSsp.Parameters.Add("@Col13", SqlDbType.VarChar, 255).Value = "0" ' unused
            MyCommon.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = -9
            MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = -9
            MyCommon.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = 0
            MyCommon.LXSsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
            MyCommon.LXSsp.ExecuteNonQuery()
            MyCommon.Close_LXSsp()
          Else
            RevokeStoredValue(AdminUserID, CustomerPK, ProgramID, 0, AdjustQty, UOMLimit, RevokeLocalID)
          End If
        End If
        
        LogText = Copient.PhraseLib.Lookup("history.sv-adjust", LanguageID)
        LogText &= " " & MyCommon.Extract_Val(ProgramID)
        If ProgramName <> "" Then
          LogText &= " (""" & ProgramName & """)"
        End If
        LogText &= " " & StrConv(Copient.PhraseLib.Lookup("term.by", LanguageID), VbStrConv.Lowercase) & " "
        LogText &= MyCommon.Extract_Val(Adjust)
        MyCommon.Activity_Log2(25, 13, CustomerPK, AdminUserID, LogText, MyCommon.Extract_Val(ProgramID), MyCommon.Extract_Val(Adjust))
      End If
      
    Catch ex As Exception
      RetMsg = ex.Message
    End Try
    
    Return RetMsg
  End Function
  
  Public Sub RevokeStoredValue(ByVal AdminUserID As Long, ByVal CustomerPK As Long, ByVal ProgramID As Long, _
                               ByVal OfferID As Long, ByVal Adjustment As Long, ByVal UOMLimit As Long, Optional ByVal RevokeLocalID As Long = 0)
    Dim dt As DataTable
    Dim row As DataRow
    Dim qtyEarned As Long
    Dim qtyUsed As Long
    Dim availableQty As Long
    Dim currentAdj As Long
    Dim RedemptionPriority As Long = 1
    Dim adjustQty As Long = 0
    Dim value As Decimal = 0D
    Dim NewExternalID As String = ""
    Dim StatusFlag As String = "2"
    Dim ExternalID As String = ""
    
    ' check if any pending revocations are processing, if so disallow any further revocations until these are first processed
    If (PendingRevokesExist(CustomerPK, ProgramID)) Then
      Throw New Exception(Copient.PhraseLib.Lookup("sv-adjust.revokespending", LanguageID))
    End If
    
    currentAdj = Math.Abs(Adjustment)
    
    Long.TryParse(MyCommon.Fetch_SystemOption(59), RedemptionPriority)
    
    If (RevokeLocalID = 0) Then
      MyCommon.QueryStr = "select StoredValueID, LocalID, QtyEarned, QtyUsed, ExternalID, Value, ServerSerial " & _
                          "from StoredValue with (NoLock) where CustomerPK=" & CustomerPK & " " & _
                          " and SVProgramID=" & ProgramID & " and StatusFlag=1 and Deleted=0 " & _
                          "order by ExpireDate "
      MyCommon.QueryStr &= IIf(RedemptionPriority = 1, "asc", "desc")
    Else
      ' this is a full revoke of a specific Local ID, first find the remaining quantity and replace the adjustment with it
      MyCommon.QueryStr = "select SUM(QtyEarned) - SUM(QtyUsed) as QtyAvail " & _
                          "from StoredValue SV with (NoLock) where StatusFlag=1 and Deleted=0 and LocalID=" & RevokeLocalID
      dt = MyCommon.LXS_Select
      If (dt.Rows.Count > 0) Then
        currentAdj = MyCommon.NZ(dt.Rows(0).Item("QtyAvail"), currentAdj)
      End If
      
      ' now find all earned records with that external id
      MyCommon.QueryStr = "select StoredValueID, LocalID, QtyEarned, QtyUsed, ExternalID, Value, ServerSerial " & _
                          "from StoredValue with (NoLock) where CustomerPK=" & CustomerPK & " " & _
                          " and SVProgramID=" & ProgramID & " and StatusFlag=1 and Deleted=0 " & _
                          " and LocalID=" & RevokeLocalID & " order by ExpireDate "
      
    End If
    ' get all the earned records for this program sorted with most recent expiration date first
    
    dt = MyCommon.LXS_Select
    
    If (dt.Rows.Count > 0) Then
      For Each row In dt.Rows
        value = MyCommon.NZ(row.Item("Value"), 0D)
        qtyEarned = MyCommon.NZ(row.Item("QtyEarned"), 0)
        qtyUsed = MyCommon.NZ(row.Item("QtyUsed"), 0)
        availableQty = qtyEarned - qtyUsed
      
        If (availableQty > 0 AndAlso availableQty <= currentAdj) Then
          ' full record revoke
          adjustQty = availableQty
          currentAdj -= availableQty
          NewExternalID = MyCommon.NZ(row.Item("ExternalID"), "")
          StatusFlag = "2"
        Else
          ' partial record revoke
          adjustQty = currentAdj
          currentAdj = 0
          NewExternalID = GetExternalID()
          StatusFlag = "1"
        End If

        ' Add the adjustment to the pending queue for the agent to process it on the next pass  
        MyCommon.QueryStr = "pa_CPE_TU_InsertData_SV"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@TableNum", SqlDbType.VarChar, 4).Value = "10" 'TableNum
        MyCommon.LXSsp.Parameters.Add("@Operation", SqlDbType.VarChar, 2).Value = "3" 'Operation
        MyCommon.LXSsp.Parameters.Add("@Col1", SqlDbType.VarChar, 255).Value = MyCommon.NZ(row.Item("LocalID"), "") ' Local ID
        MyCommon.LXSsp.Parameters.Add("@Col2", SqlDbType.VarChar, 255).Value = MyCommon.NZ(row.Item("ServerSerial"), "-9") ' Server Serial
        MyCommon.LXSsp.Parameters.Add("@Col3", SqlDbType.VarChar, 255).Value = CustomerPK.ToString ' CustomerPK
        MyCommon.LXSsp.Parameters.Add("@Col4", SqlDbType.VarChar, 255).Value = NewExternalID ' NewExternalID
        MyCommon.LXSsp.Parameters.Add("@Col5", SqlDbType.VarChar, 255).Value = StatusFlag 'StatusFlag
        MyCommon.LXSsp.Parameters.Add("@Col6", SqlDbType.VarChar, 255).Value = adjustQty.ToString ' QtyUsed
        MyCommon.LXSsp.Parameters.Add("@Col7", SqlDbType.VarChar, 255).Value = MyCommon.NZ(row.Item("ExternalID"), "") ' PrevExternalID
        MyCommon.LXSsp.Parameters.Add("@Col8", SqlDbType.VarChar, 255).Value = "2" 'HistoryStatusFlag
        MyCommon.LXSsp.Parameters.Add("@Col9", SqlDbType.VarChar, 255).Value = Now.ToString("yyyy-MM-ddTHH:mm:ss") ' LastUpdate
        MyCommon.LXSsp.Parameters.Add("@Col10", SqlDbType.VarChar, 255).Value = value.ToString  ' Value
        MyCommon.LXSsp.Parameters.Add("@Col11", SqlDbType.VarChar, 255).Value = ProgramID.ToString ' SVProgramID
        MyCommon.LXSsp.Parameters.Add("@Col12", SqlDbType.VarChar, 255).Value = 0 ' RedeemedValue
        MyCommon.LXSsp.Parameters.Add("@Col13", SqlDbType.VarChar, 255).Value = (adjustQty * value * UOMLimit).ToString ' BreakageValue
        MyCommon.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = -9
        MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = -9
        MyCommon.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = 0
        MyCommon.LXSsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
              
        MyCommon.LXSsp.ExecuteNonQuery()
        MyCommon.Close_LXSsp()
      
        'check if we're done revoking
        If (currentAdj = 0) Then Exit For
      Next
    Else
      ' no records to revoke so alert the user
      Throw New Exception(Copient.PhraseLib.Lookup("sv.excessiverevoke", LanguageID))
    End If
    
  End Sub
  
  Function PendingRevokesExist(ByVal CustomerPK As Long, ByVal ProgramID As Long) As Boolean
    Dim PendingExist As Boolean = False
    Dim dt As DataTable
    
    MyCommon.QueryStr = "select Col3 as CustomerPK from CPE_UploadTemp_SV with (NoLock) " & _
                        "where TableNum = '10' and Operation = '3' and Col3='" & CustomerPK & "' " & _
                        "and Col11 = '" & ProgramID & "';"
    dt = MyCommon.LXS_Select
    
    PendingExist = (dt.Rows.Count > 0)
    
    Return PendingExist
  End Function
  
  Public Sub BatchInsert(ByVal dataTable As DataTable, ByVal batchSize As Int32)
    Dim adapter As New SqlDataAdapter()

    adapter.InsertCommand = New SqlCommand("pa_CPE_TU_InsertData_SV", MyCommon.LXSadoConn)
    adapter.InsertCommand.CommandType = CommandType.StoredProcedure

    adapter.InsertCommand.Parameters.Add("@TableNum", SqlDbType.VarChar, 4, "TableNum")
    adapter.InsertCommand.Parameters.Add("@Operation", SqlDbType.VarChar, 2, "Operation")
    adapter.InsertCommand.Parameters.Add("@Col1", SqlDbType.VarChar, 255, "Col1")
    adapter.InsertCommand.Parameters.Add("@Col2", SqlDbType.VarChar, 255, "Col2")
    adapter.InsertCommand.Parameters.Add("@Col3", SqlDbType.VarChar, 255, "Col3")
    adapter.InsertCommand.Parameters.Add("@Col4", SqlDbType.VarChar, 255, "Col4")
    adapter.InsertCommand.Parameters.Add("@Col5", SqlDbType.VarChar, 255, "Col5")
    adapter.InsertCommand.Parameters.Add("@Col6", SqlDbType.VarChar, 255, "Col6")
    adapter.InsertCommand.Parameters.Add("@Col7", SqlDbType.VarChar, 255, "Col7")
    adapter.InsertCommand.Parameters.Add("@Col8", SqlDbType.VarChar, 255, "Col8")
    adapter.InsertCommand.Parameters.Add("@Col9", SqlDbType.VarChar, 255, "Col9")
    adapter.InsertCommand.Parameters.Add("@Col10", SqlDbType.VarChar, 255, "Col10")
    adapter.InsertCommand.Parameters.Add("@Col11", SqlDbType.VarChar, 255, "Col11")
    adapter.InsertCommand.Parameters.Add("@Col12", SqlDbType.VarChar, 255, "Col12")
    adapter.InsertCommand.Parameters.Add("@Col13", SqlDbType.VarChar, 255, "Col13")
    adapter.InsertCommand.Parameters.Add("@LocalServerID", SqlDbType.Int, 4, "LocalServerID")
    adapter.InsertCommand.Parameters.Add("@LocationID", SqlDbType.BigInt, 8, "LocationID")
    adapter.InsertCommand.Parameters.Add("@WaitingACK", SqlDbType.Int, 4, "WaitingACK")
    adapter.InsertCommand.Parameters.Add("@AdminUserID", SqlDbType.Int, 4, "AdminUserID")
    
    adapter.InsertCommand.UpdatedRowSource = UpdateRowSource.None
    adapter.UpdateBatchSize = batchSize
 
    adapter.Update(dataTable)
  End Sub
  
  Function GetExternalID() As String
    Dim ExternalID(1) As String
    
    ExternalID = GetExternalIDs(1)
    
    Return ExternalID(0)
  End Function
  
  Function GetExternalIDs(ByVal BlockSize As Long) As String()
    Dim dt As DataTable
    Dim row As DataRow
    Dim XidStyleID As Long = 0
    Dim i As Long
    Dim MaxSize As Long = 0
    
    If (BlockSize < 1) Then BlockSize = 1
    Dim ExternalIDs(BlockSize - 1) As String
    Dim SequenceIDs(BlockSize - 1) As Integer
    
    'initialize arrays
    For i = 0 To BlockSize - 1
      ExternalIDs(i) = ""
      SequenceIDs(i) = 0
    Next
    
    XidStyleID = MyCommon.Fetch_SystemOption(58)
    MaxSize = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(76))
    
    If (XidStyleID = 0) Then
      For i = 0 To BlockSize - 1
        SequenceIDs(i) = 0
      Next
      ExternalIDs = MyCommon.Format_SV_ExternalIDs(SequenceIDs)
    Else
      ' get the next sequence number to use as part of the External ID
      MyCommon.QueryStr = "pa_GetNextStoredValueSequence"
      MyCommon.Open_LXSsp()
      MyCommon.LXSsp.Parameters.Add("@MaxSize", SqlDbType.Int).Value = MaxSize
      MyCommon.LXSsp.Parameters.Add("@BlockSize", SqlDbType.Int).Value = BlockSize
      dt = MyCommon.LXSsp_select
      If (dt.Rows.Count > 0) Then
        i = 0
        For Each row In dt.Rows
          If (i <= SequenceIDs.GetUpperBound(0)) Then
            SequenceIDs(i) = MyCommon.NZ(row.Item("spid"), 0)
          End If
          i += 1
        Next
        ExternalIDs = MyCommon.Format_SV_ExternalIDs(SequenceIDs)
      End If
      MyCommon.Close_LXSsp()
    End If
    
    Return ExternalIDs
  End Function
  
  Function GetLocalID() As Long
    Dim LocalIDs As Long()
    Dim LocalID As Long = 0
    
    LocalIDs = GetLocalIDs(1)
    
    If (LocalIDs.Length > 0) Then
      LocalID = LocalIDs(0)
    End If
    
    Return LocalID
  End Function
  
  Function GetLocalIDs(ByVal BlockSize As Long) As Long()
    Dim dt As DataTable
    Dim row As DataRow
    Dim i As Long

    If (BlockSize < 1) Then BlockSize = 1
    Dim SequenceIDs(BlockSize - 1) As Long
        
    ' get the next Local ID
    MyCommon.QueryStr = "pa_GetNextStoredValueLocalID"
    MyCommon.Open_LXSsp()
    MyCommon.LXSsp.Parameters.Add("@BlockSize", SqlDbType.Int).Value = BlockSize
    dt = MyCommon.LXSsp_select
    If (dt.Rows.Count > 0) Then
      i = 0
      For Each row In dt.Rows
        If (i <= SequenceIDs.GetUpperBound(0)) Then
          SequenceIDs(i) = MyCommon.NZ(row.Item("localid"), 0)
        End If
        i += 1
      Next
    End If
    MyCommon.Close_LXSsp()
    
    Return SequenceIDs
  End Function

  Function GetExpireDate(ByVal ProgramID As Long) As String
    Dim ExpirePeriod As Long = 0
    Dim SVExpireType As Long = 0
    Dim SVExpirePeriodType As Long = 0
    Dim ExpireTOD As String = ""
    Dim ExpireDate As Date = Now
    Dim dt As DataTable
    Dim RetDate As String = ""
    
    Try
      If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
      
      MyCommon.QueryStr = "select ExpirePeriod, SVExpireType, SVExpirePeriodType, " & _
                          "ExpireTOD, ExpireDate from StoredValuePrograms with (NoLock) " & _
                          "where SVProgramID = " & ProgramID & " and Deleted=0;"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        ExpirePeriod = MyCommon.NZ(dt.Rows(0).Item("ExpirePeriod"), 0)
        SVExpireType = MyCommon.NZ(dt.Rows(0).Item("SVExpireType"), 0)
        SVExpirePeriodType = MyCommon.NZ(dt.Rows(0).Item("SVExpirePeriodType"), 0)
        ExpireTOD = MyCommon.NZ(dt.Rows(0).Item("ExpireTOD"), "")
        ExpireDate = MyCommon.NZ(dt.Rows(0).Item("ExpireDate"), Now)
      End If
      
      Select Case SVExpireType
        Case 1
          RetDate = ExpireDate.ToString("yyyy-MM-ddTHH:mm:ss")
        Case 2
          If (ExpireTOD = "" OrElse ExpireTOD.Length <> 5) Then ExpireTOD = "00:00"
          RetDate = Now.AddDays(ExpirePeriod).ToString("yyyy-MM-dd") & "T" & ExpireTOD & ":00"
          If (Date.TryParse(RetDate, ExpireDate)) Then
            RetDate = ExpireDate.ToString("yyyy-MM-ddTHH:mm:ss")
          Else
            RetDate = Now.ToString("yyyy-MM-ddTHH:mm:ss")
          End If
        Case 3
          If (SVExpirePeriodType = 1) Then
            RetDate = Now.AddDays(ExpirePeriod).ToString("yyyy-MM-ddTHH:mm:ss")
          Else
            RetDate = Now.AddHours(ExpirePeriod).ToString("yyyy-MM-ddTHH:mm:ss")
          End If
        Case Else
          RetDate = Now.ToString("yyyy-MM-ddTHH:mm:ss")
      End Select

    Catch ex As Exception
      Throw New Exception(ex.ToString)
    End Try

    Return RetDate
  End Function
  
</script>
