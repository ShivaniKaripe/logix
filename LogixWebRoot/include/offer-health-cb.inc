<script runat="server">
' version:7.3.1.138972.Official Build (SUSDAY10202)
  Function GetOfferHealthTable(ByRef MyCommon As Copient.CommonInc, ByVal SearchClause As String, ByVal ExitOnWarning As Boolean) As System.Data.DataTable
    Dim dst As System.Data.DataTable = Nothing
    Dim dtHealth As System.Data.DataTable = Nothing
    Dim row, rowHealth As System.Data.DataRow
    Dim ValidLocations As Integer = 0
    Dim WatchLocations As Integer = 0
    Dim WarningLocations As Integer = 0
    Dim ValidOfferComponents As Boolean = False
    
    ' store all the data into one datatable and filter it as necessary            
    dtHealth = New System.Data.DataTable
    dtHealth.Columns.Add("IncentiveID", System.Type.GetType("System.Int64"))
    dtHealth.Columns.Add("EngineId", System.Type.GetType("System.Int32"))
    dtHealth.Columns.Add("IncentiveName", System.Type.GetType("System.String"))
    dtHealth.Columns.Add("WatchLocations", System.Type.GetType("System.Int64"))
    dtHealth.Columns.Add("ValidLocations", System.Type.GetType("System.Int64"))
    dtHealth.Columns.Add("WarningLocations", System.Type.GetType("System.Int64"))
    dtHealth.Columns.Add("ValidOfferComponents", System.Type.GetType("System.Boolean"))
    
    MyCommon.QueryStr = "select Distinct IncentiveID, IncentiveName from CPE_Incentives I WITH (NoLock) " & _
                        "inner join OfferLocations AS OL WITH (NoLock) ON I.IncentiveID = OL.OfferID AND I.Deleted = 0 AND OL.Deleted = 0 " & _
                        "inner join  LocGroupItems AS LGI WITH (NoLock) ON LGI.LocationGroupID = OL.LocationGroupID AND LGI.Deleted = 0 " & _
                        "where IsTemplate=0 and StartDate<= getDate() and EndDate >= getDate() " & SearchClause & _
                        "union " & _
                        "select  Distinct I.IncentiveID, I.IncentiveName from CPE_Incentives I WITH (NoLock) " & _
                        "inner join OfferLocations AS OL WITH (NoLock) ON I.IncentiveID = OL.OfferID AND I.Deleted = 0 AND OL.Deleted = 0 " & _
                        "inner join LocationGroups AS LG WITH (NoLock) ON LG.LocationGroupID = OL.LocationGroupID AND LG.Deleted = 0 AND LG.AllLocations = 1 " & _
                        "where IsTemplate=0 and StartDate<= getDate() and EndDate >= getDate() " & SearchClause
    'Send(MyCommon.QueryStr)
    dst = MyCommon.LRT_Select
    If (dst.Rows.Count > 0) Then
      ' build a new datatable to allow sorting on computed columns
      For Each row In dst.Rows
        ValidateOfferHeader(MyCommon, row.Item("IncentiveID"), ValidLocations, WatchLocations, WarningLocations)
        ValidOfferComponents = OfferComponentsValid(MyCommon, row.Item("IncentiveID"))
        rowHealth = dtHealth.NewRow
        rowHealth.Item("IncentiveID") = row.Item("IncentiveID")
        rowHealth.Item("EngineID") = 2
        rowHealth.Item("IncentiveName") = MyCommon.NZ(row.Item("IncentiveName"), "")
        rowHealth.Item("ValidLocations") = ValidLocations
        rowHealth.Item("WatchLocations") = WatchLocations
        rowHealth.Item("WarningLocations") = WarningLocations
        rowHealth.Item("ValidOfferComponents") = ValidOfferComponents
		If(not ValidOfferComponents) then
			dtHealth.Rows.Add(rowHealth)
		End If
        
        ' when we just want to know if there are any offer health problems (e.g. status page system warnings)
        ' and once we find one then exit forthwith
        If (ExitOnWarning) AndAlso (WarningLocations > 0 OrElse Not ValidOfferComponents) Then
          Exit For
        End If
      Next
    End If
    
    MyCommon.QueryStr = "select Distinct O.OfferID, O.Name from Offers O WITH (NoLock) " & _
                        "inner join CM_ST_OfferLocations AS OL WITH (NoLock) ON O.OfferID = OL.OfferID AND O.Deleted = 0 AND OL.Deleted = 0 " & _
                        "inner join  LocGroupItems AS LGI WITH (NoLock) ON LGI.LocationGroupID = OL.LocationGroupID AND LGI.Deleted = 0 " & _
                        "where IsTemplate=0 and ProdStartDate<= getDate() and ProdEndDate >= getDate() " & SearchClause & _
                        "union " & _
                        "select Distinct O.OfferID, O.Name from Offers O WITH (NoLock) " & _
                        "inner join CM_ST_OfferLocations AS OL WITH (NoLock) ON O.OfferID = OL.OfferID AND O.Deleted = 0 AND OL.Deleted = 0 " & _
                        "inner join LocationGroups AS LG WITH (NoLock) ON LG.LocationGroupID = OL.LocationGroupID AND LG.Deleted = 0 AND LG.AllLocations = 1 " & _
                        "where IsTemplate=0 and ProdStartDate<= getDate() and ProdEndDate >= getDate() " & SearchClause
    'Send(MyCommon.QueryStr)
    dst = MyCommon.LRT_Select
    If (dst.Rows.Count > 0) Then
      ' build a new datatable to allow sorting on computed columns
      For Each row In dst.Rows
        ValidateCmOfferHeader(MyCommon, row.Item("OfferID"), ValidLocations, WatchLocations, WarningLocations)
        ValidOfferComponents = CmOfferComponentsValid(MyCommon, row.Item("OfferID"))
        rowHealth = dtHealth.NewRow
        rowHealth.Item("IncentiveID") = row.Item("OfferID")
        rowHealth.Item("EngineID") = 0
        rowHealth.Item("IncentiveName") = MyCommon.NZ(row.Item("Name"), "")
        rowHealth.Item("ValidLocations") = ValidLocations
        rowHealth.Item("WatchLocations") = WatchLocations
        rowHealth.Item("WarningLocations") = WarningLocations
        rowHealth.Item("ValidOfferComponents") = ValidOfferComponents
		If(not ValidOfferComponents) then
			dtHealth.Rows.Add(rowHealth)
		End If
        
        ' when we just want to know if there are any offer health problems (e.g. status page system warnings)
        ' and once we find one then exit forthwith
        If (ExitOnWarning) AndAlso (WarningLocations > 0 OrElse Not ValidOfferComponents) Then
          Exit For
        End If
      Next
    End If
    Return dtHealth
  End Function
  
  Sub ValidateOfferHeader(ByRef MyCommon As Copient.CommonInc, ByVal IncentiveID As Integer, ByRef ValidLocations As Integer, _
                          ByRef WatchLocations As Integer, ByRef WarningLocations As Integer)
    Dim dtValid As System.Data.DataTable
    Dim rowOK(), rowWatches(), rowWarnings() As System.Data.DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceCount As Double
    
    objTemp = MyCommon.Fetch_CPE_SystemOption(41)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If
    
    objTemp = MyCommon.Fetch_CPE_SystemOption(42)
    If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
      GraceCount = 0.1D
    End If
    
    MyCommon.QueryStr = "dbo.pa_ValidationReport_Incentive"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@IncentiveID", System.Data.SqlDbType.Int).Value = IncentiveID
    MyCommon.LRTsp.Parameters.Add("@GraceHours", System.Data.SqlDbType.Int).Value = GraceHours
    MyCommon.LRTsp.Parameters.Add("@GraceCount", System.Data.SqlDbType.Decimal, 2).Value = GraceCount
    dtValid = MyCommon.LRTsp_select()
    
    RemoveInactiveLocations(MyCommon, dtValid, IncentiveID)
    
    rowOK = dtValid.Select("Status=0", "LocationName")
    rowWatches = dtValid.Select("Status=1", "LocationName")
    rowWarnings = dtValid.Select("Status=2", "LocationName")
    ValidLocations = rowOK.Length
    WatchLocations = rowWatches.Length
    WarningLocations = rowWarnings.Length
    
    MyCommon.Close_LRTsp()
  End Sub
  
  Function OfferComponentsValid(ByRef MyCommon As Copient.CommonInc, ByVal IncentiveID As Integer) As Boolean
    Dim IsValid As Boolean = True
    Dim InvalidCount As Integer = 0
    Dim dtComponents As System.Data.DataTable = Nothing
    Dim rowComp As System.Data.DataRow = Nothing
    
    MyCommon.QueryStr = "dbo.pa_ValidationReport_OfferComponents"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.Int).Value = IncentiveID
    dtComponents = MyCommon.LRTsp_select
    
    For Each rowComp In dtComponents.Rows
      InvalidCount = InvalidComponents(MyCommon, rowComp)
      If (InvalidCount > 0) Then
        IsValid = False
        Exit For
      End If
    Next
    
    MyCommon.Close_LRTsp()
    Return IsValid
  End Function
  
  Function InvalidComponents(ByRef MyCommon As Copient.CommonInc, ByVal rowComp As System.Data.DataRow) As Integer
    Dim RecordType As String = ""
    Dim ID As Integer
    Dim StoredProcName As String = ""
    Dim IDParmName As String = ""
    Dim TypeCode As String = ""
    Dim PageName As String = ""
    Dim dtValid As System.Data.DataTable
    Dim rowOK(), rowWatches(), rowWarnings() As System.Data.DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceCount As Double
    Dim InvalidCt As Integer = 0
    Dim IgnoreID As Boolean
    
    objTemp = MyCommon.Fetch_CPE_SystemOption(41)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If
    
    objTemp = MyCommon.Fetch_CPE_SystemOption(42)
    If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
      GraceCount = 0.1D
    End If
    
    RecordType = MyCommon.NZ(rowComp.Item("RecordType"), "")
    ID = MyCommon.NZ(rowComp.Item("ID"), -1)
    
    Select Case RecordType
      Case "term.customergroup"
        StoredProcName = "dbo.pa_ValidationReport_CustGroup"
        IDParmName = "@CustomerGroupID"
        TypeCode = "cg"
        PageName = "cgroup-edit.aspx?CustomerGroupID="
        IgnoreID = IIf(ID = 1 OrElse ID = 2, True, False)
      Case "term.productgroup"
        StoredProcName = "dbo.pa_ValidationReport_ProdGroup"
        IDParmName = "@ProductGroupID"
        TypeCode = "pg"
        PageName = "pgroup-edit.aspx?ProductGroupID="
        IgnoreID = IIf(ID = 1, True, False)
      Case "term.graphics"
        StoredProcName = "dbo.pa_ValidationReport_Graphic"
        IDParmName = "@OnScreenAdID"
        TypeCode = "gr"
        PageName = "graphic-edit.aspx?OnScreenAdID="
    End Select
    
    If (Not IgnoreID) Then
      MyCommon.QueryStr = StoredProcName
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add(IDParmName, System.Data.SqlDbType.Int).Value = ID
      MyCommon.LRTsp.Parameters.Add("@GraceHours", System.Data.SqlDbType.Int).Value = GraceHours
      MyCommon.LRTsp.Parameters.Add("@GraceCount", System.Data.SqlDbType.Decimal, 2).Value = GraceCount
      dtValid = MyCommon.LRTsp_select()
      
      rowOK = dtValid.Select("Status=0", "LocationName")
      rowWatches = dtValid.Select("Status=1", "LocationName")
      rowWarnings = dtValid.Select("Status=2", "LocationName")
      InvalidCt = rowWarnings.Length
      
      MyCommon.Close_LRTsp()
    End If
    Return InvalidCt
  End Function
  
  Sub RemoveInactiveLocations(ByRef MyCommon As Copient.CommonInc, ByRef dt As System.Data.DataTable, ByVal IncentiveID As Integer)
    Dim dtLoc As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim LocationID As Integer
    Dim LocGroupList As String = ""
    Dim IsAllLocs As Boolean
    Dim LocTable As New Hashtable
    
    If (Not dt Is Nothing And dt.Rows.Count > 0) Then
      ' first find all locations is selected
      MyCommon.QueryStr = "select LocationGroupID from OfferLocations with (NoLock) where OfferID=" & IncentiveID & " and Deleted=0;"
      dtLoc = MyCommon.LRT_Select
      For Each row In dtLoc.Rows
        If (LocGroupList <> "") Then LocGroupList += ","
        LocGroupList = LocGroupList + MyCommon.NZ(row.Item("LocationGroupID"), "-1").ToString
        IsAllLocs = MyCommon.NZ(row.Item("LocationGroupID"), -1) = 1
      Next
      If (Not IsAllLocs) Then
        ' find all the locations for the given location groups
        MyCommon.QueryStr = "select LocationID from LocGroupItems with (NoLock) where Deleted = 0 " & _
                            "and LocationGroupID in (" & LocGroupList & ");"
        dtLoc = MyCommon.LRT_Select
        For Each row In dtLoc.Rows
          LocationID = MyCommon.NZ(row.Item("LocationID"), "-1")
          If (Not LocTable.ContainsKey(LocationID.ToString)) Then
            LocTable.Add(MyCommon.NZ(row.Item("LocationID"), "-1").ToString, MyCommon.NZ(row.Item("LocationID"), "-1").ToString)
          End If
        Next
        ' remove the location if it doesn't currently exist for the incentive
        For Each row In dt.Rows
          LocationID = MyCommon.NZ(row.Item("LocationID"), "-1")
          If (Not LocTable.ContainsKey(LocationID.ToString)) Then
            row.Delete()
          End If
        Next
      End If
    End If
  End Sub
  
  Sub ValidateCmOfferHeader(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByRef ValidLocations As Integer, _
                          ByRef WatchLocations As Integer, ByRef WarningLocations As Integer)
    Dim dtValid As System.Data.DataTable
    Dim rowOK(), rowWatches(), rowWarnings() As System.Data.DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceHoursWarn As Integer
    
    objTemp = MyCommon.Fetch_CM_SystemOption(10)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If
    
    objTemp = MyCommon.Fetch_CM_SystemOption(11)
    If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
      GraceHoursWarn = 24
    End If
    
    MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_Offer"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.Int).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@GraceHours", System.Data.SqlDbType.Int).Value = GraceHours
    MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", System.Data.SqlDbType.Int).Value = GraceHoursWarn
    dtValid = MyCommon.LRTsp_select()
    
    rowOK = dtValid.Select("Status=0", "LocationName")
    rowWatches = dtValid.Select("Status=2", "LocationName")
    rowWarnings = dtValid.Select("Status=3", "LocationName")
    ValidLocations = rowOK.Length
    WatchLocations = rowWatches.Length
    WarningLocations = rowWarnings.Length
    
    MyCommon.Close_LRTsp()
  End Sub
  
  Function CmOfferComponentsValid(ByRef MyCommon As Copient.CommonInc, ByVal IncentiveID As Integer) As Boolean
    Dim IsValid As Boolean = True
    Dim InvalidCount As Integer = 0
    Dim dtComponents As System.Data.DataTable = Nothing
    Dim rowComp As System.Data.DataRow = Nothing
    
    MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_OfferComponents"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.Int).Value = IncentiveID
    dtComponents = MyCommon.LRTsp_select
    
    For Each rowComp In dtComponents.Rows
      InvalidCount = InvalidCmComponents(MyCommon, rowComp)
      If (InvalidCount > 0) Then
        IsValid = False
        Exit For
      End If
    Next
    
    MyCommon.Close_LRTsp()
    Return IsValid
  End Function
  
  Function InvalidCmComponents(ByRef MyCommon As Copient.CommonInc, ByVal rowComp As System.Data.DataRow) As Integer
    Dim RecordType As String = ""
    Dim ID As Integer
    Dim StoredProcName As String = ""
    Dim IDParmName As String = ""
    Dim TypeCode As String = ""
    Dim PageName As String = ""
    Dim dtValid As System.Data.DataTable
    Dim rowOK(), rowWatches(), rowWarnings() As System.Data.DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceHoursWarn As Integer
    Dim InvalidCt As Integer = 0
    Dim IgnoreID As Boolean
    
    objTemp = MyCommon.Fetch_CM_SystemOption(10)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If
    
    objTemp = MyCommon.Fetch_CM_SystemOption(11)
    If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
      GraceHoursWarn = 24
    End If
    
    RecordType = MyCommon.NZ(rowComp.Item("RecordType"), "")
    ID = MyCommon.NZ(rowComp.Item("ID"), -1)
    
    Select Case RecordType
      Case "term.customergroup"
        StoredProcName = "dbo.pa_CM_ValidationReport_CustGroup"
        IDParmName = "@CustomerGroupID"
        TypeCode = "cg"
        PageName = "cgroup-edit.aspx?CustomerGroupID="
        IgnoreID = IIf(ID = 1 OrElse ID = 2, True, False)
      Case "term.productgroup"
        StoredProcName = "dbo.pa_CM_ValidationReport_ProdGroup"
        IDParmName = "@ProductGroupID"
        TypeCode = "pg"
        PageName = "pgroup-edit.aspx?ProductGroupID="
        IgnoreID = IIf(ID = 1, True, False)
    End Select
    
    If (Not IgnoreID) Then
      MyCommon.QueryStr = StoredProcName
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add(IDParmName, System.Data.SqlDbType.Int).Value = ID
      MyCommon.LRTsp.Parameters.Add("@GraceHours", System.Data.SqlDbType.Int).Value = GraceHours
      MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", System.Data.SqlDbType.Int).Value = GraceHoursWarn
      dtValid = MyCommon.LRTsp_select()
      
      rowOK = dtValid.Select("Status=0", "LocationName")
      rowWatches = dtValid.Select("Status=2", "LocationName")
      rowWarnings = dtValid.Select("Status=3", "LocationName")
      InvalidCt = rowWarnings.Length
      
      MyCommon.Close_LRTsp()
    End If
    Return InvalidCt
  End Function
</script>
