Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml
Imports System.Data
Imports System.IO
Imports System.Collections.Generic
Imports System.Xml.Serialization
Imports System.Xml.Schema
Imports System.Data.SqlClient
Imports System.Collections.Specialized.NameValueCollection
Imports Copient.CommonInc
Imports Copient
Imports CMS
Imports CMS.AMS
Imports CMS.AMS.Models
Imports CMS.AMS.Contract
Imports CMS.Contract
Imports System.Diagnostics
Imports System.Globalization

<WebService(Namespace:="http://www.copienttech.com/AMSActivityLogData/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class AMSActivityLogData
    Inherits System.Web.Services.WebService
    ' Return error codes
    Private Enum ERROR_CODES As Integer
        INVALID_GUID = -1
        ERROR_NO_ACTIVITY_FOUND_FOR_SEARCH_CRITERIA = -2
        INVALID_XML_DOCUMENT = -3
        INCORRECT_DATERANGE = -4
        INCORRECT_STARTDATEFORMAT = -5
        INCORRECT_ENDDATEFORMAT = -6
        APPLICATION_EXCEPTION = -9999
        ERROR_NONE = 0
    End Enum
    Private Const CONNECTOR_ID As Integer = 71
    Private ActivityLogFile As String = "ActivityLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private ActivityErrorLogFile As String = "ActivityErrorLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private MyCommon As New Copient.CommonInc

    <WebMethod()> _
    Public Function GetActivityLogData(ByVal Guid As String, ByVal AdminID As String, ByVal StartDate As String, ByVal EndDate As String) As XmlDocument
        Dim MethodName As String = "GetActivityLogData"
        Dim ResponseXml As New XmlDocument()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Static RequestTime As DateTime = Now
        Dim DaysSpan As Int64 = -1
        Dim dtStart As DateTime
        Dim dtEnd As DateTime
        Dim dtfi As New DateTimeFormatInfo()
        dtfi.ShortDatePattern = "MM/dd/yyyy"
        dtfi.DateSeparator = "/"

        Try

            'setting the startdate and enddate if not provided
            'if enddate is not provided then calculate the enddate based on startdate or default it currentdate
            If EndDate = String.Empty Then
                If (IsDate(StartDate)) Then
                    dtStart = Convert.ToDateTime(StartDate, dtfi)
                    dtEnd = dtStart.AddDays(30)
                    EndDate = dtEnd.ToString(dtfi.ShortDatePattern)
                Else
                    dtEnd = DateTime.Now
                    EndDate = dtEnd.ToString(dtfi.ShortDatePattern)
                End If

            Else
                'check if the start date provided is valid
                If (IsDate(EndDate)) Then
                    dtEnd = Convert.ToDateTime(EndDate, dtfi)
                Else
                    ErrorCode = ERROR_CODES.INCORRECT_ENDDATEFORMAT
                    ErrorMsg = "Incorrect EndDate Format"
                    Throw New Exception(ErrorMsg)
                End If
            End If

            'if startdate is not provided then calculate the startdate based on enddate
            If StartDate = String.Empty Then
                dtStart = dtEnd.AddDays(-30)
                StartDate = dtStart.ToString(dtfi.ShortDatePattern)
            Else
                'check if the start date provided is valid
                If (IsDate(StartDate)) Then
                    dtStart = Convert.ToDateTime(StartDate, dtfi)
                Else
                    ErrorCode = ERROR_CODES.INCORRECT_STARTDATEFORMAT
                    ErrorMsg = "Incorrect StartDate Format"
                    Throw New Exception(ErrorMsg)

                End If
            End If

            'calculate the daterange 
            If (StartDate.Length >= 10 And EndDate.Length >= 10) Then
                DaysSpan = DateDiff("d", dtStart, dtEnd)
            End If

            If (DaysSpan <= 30 And DaysSpan <> -1) Then

                'Make sure StartDate starts at beginning of the date if no time entered
                If IsDate(StartDate) Then
                    StartDate = StartDate & " 00:00:00"
                End If

                'Make sure EndDate ends at the last second of the date if no time entered
                If IsDate(EndDate) Then
                    EndDate = EndDate & " 23:59:59"
                End If

                If IsValidGUID(Guid, MethodName, MyCommon) Then
                    ResponseXml = ActivityLogXml(AdminID, StartDate, EndDate)
                    If (ResponseXml.InnerXml = String.Empty) Then
                        ErrorCode = ERROR_CODES.ERROR_NO_ACTIVITY_FOUND_FOR_SEARCH_CRITERIA
                        ErrorMsg = "No records found for the search criteria entered."
                    Else
                        ErrorMsg = "Search Criteria - StartDate :" & StartDate & ", EndDate:" & EndDate & ", AdminID:" & AdminID
                    End If
                Else
                    ErrorCode = ERROR_CODES.INVALID_GUID
                    ErrorMsg = "Invalid GUID"
                End If
            Else
                ErrorCode = ERROR_CODES.INCORRECT_DATERANGE
                ErrorMsg = "Incorrect Date range - should not be more than 30 days"
            End If
        Catch ex As Exception
            If ErrorCode = ERROR_CODES.ERROR_NONE Then
                ErrorCode = ERROR_CODES.APPLICATION_EXCEPTION
                ErrorMsg = "Application error encountering: " & ex.ToString
            End If
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If Not (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = BuildErrorResponse(ErrorCode, ErrorMsg)
        End If
        AppendToLog("GetActivityLogData", RequestTime, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function

    Private Sub AppendToLog(ByVal FunctionName As String, ByVal LogText As String, ByVal ErrorCode As ERROR_CODES, ByVal LogMsg As String)
        Dim LogString As String = ""
        Dim AndLogText As String = ""

        If LogText <> "" Then AndLogText = LogText & ";"

        Select Case ErrorCode

            Case ERROR_CODES.ERROR_NONE  'No Errors logging to ActivityLogFile

                LogString = "Success=" & FunctionName & "; " & AndLogText & ErrorCode.ToString & " " & LogMsg & "; Server= " & Environment.MachineName
                Copient.Logger.Write_Log(ActivityLogFile, LogString, True)

            Case Else   'Errors logging to ActivityErrorLogFile

                LogString = "Error=" & FunctionName & "; " & AndLogText & "; Error_encountered=" & ErrorCode.ToString & " " & LogMsg & "; Server=" & Environment.MachineName
                Copient.Logger.Write_Log(ActivityErrorLogFile, LogString, True)

        End Select

    End Sub

    Private Function ActivityLogXml(ByVal AdminID As String, ByVal StartDate As String, ByVal EndDate As String) As XmlDocument

        Dim ResponseXml As New XmlDocument()
        Dim DateRange As String = ""
        Dim AdminIDNotZero As String = ""
        Dim TopLimit As Integer = 5000
        Dim STopLimit As String = ""
        Dim SActivityLogXml As New StringBuilder()
        Dim dt As DataTable = Nothing
        Dim row As DataRow
        Dim startTime As Date = Date.Now
        Dim SqlFilter As String = " Where "


        If StartDate <> String.Empty And EndDate <> String.Empty Then
            DateRange = SqlFilter & "AL.ActivityDate between @StartDate and @EndDate"
            MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
            MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        End If

        If AdminID <> String.Empty Then
            AdminIDNotZero = " and AdminID = @AdminID "
            MyCommon.DBParameters.Add("@AdminID", SqlDbType.Int).Value = Convert.ToInt64(AdminID)
        End If

        If TopLimit > 0 Then
            STopLimit = " TOP (" & TopLimit & ")"
        End If

        Try
            If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select" & STopLimit & " AL.AdminID, AL.ActivityDate Date, AT.Name Type, AL.Description Description from ActivityLog AL with (NoLock)" & _
                " inner Join ActivityTypes AT with (NoLock) on AT.ActivityTypeID=AL.ActivityTypeID" & DateRange & AdminIDNotZero & " order by Date Desc"
            dt = MyCommon.ExecuteQuery(DataBases.LogixRT)

            Dim ms As New MemoryStream
            Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
            If (dt.Rows.Count > 0) Then
                Init_XMLWriter(Writer, "Success")
                For Each row In dt.Rows
                    Writer.WriteStartElement("Activity")
                    generateXML(row, Writer)
                    Writer.WriteEndElement()
                Next
                End_XMLWriter(Writer)
                ms.Seek(0, SeekOrigin.Begin)
                ResponseXml.Load(ms)
                ms.Flush()
                Writer.Close()
            End If
            If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Close_LogixRT()
        Catch ex As Exception
            Throw ex
        End Try
        Return ResponseXml
    End Function

    Private Sub generateXML(ByVal row As DataRow, ByRef Writer As XmlTextWriter)

        For Each col As DataColumn In row.Table.Columns
            Writer.WriteElementString(col.ColumnName, row(col.ColumnName).ToString())
        Next

    End Sub

    Private Sub Init_XMLWriter(ByRef writer As XmlTextWriter, ByVal returnCode As String)

        writer.Formatting = Formatting.Indented
        writer.Indentation = 4
        writer.WriteStartDocument()
        writer.WriteStartElement("ActivityData")
        writer.WriteAttributeString("returnCode", returnCode)
        writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
        writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

    End Sub

    Private Sub End_XMLWriter(ByRef Writer As XmlWriter)

        Writer.WriteEndDocument()
        Writer.Flush()

    End Sub

    Private Function BuildErrorResponse(ByVal ErrorCode As ERROR_CODES, ByVal ErrorMsg As String) As XmlDocument

        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim ErrorXMl As New XmlDocument()

        Init_XMLWriter(Writer, ErrorCode.ToString())
        Writer.WriteElementString("ErrorMessage", ErrorMsg)
        Writer.WriteEndElement()
        End_XMLWriter(Writer)
        ms.Seek(0, SeekOrigin.Begin)
        ErrorXMl.Load(ms)
        ms.Flush()
        Writer.Close()
        Return ErrorXMl

    End Function

    Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String, ByRef Common As Copient.CommonInc) As Boolean
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc
        Dim MsgBuf As New StringBuilder()

        Try
            IsValid = ConnInc.IsValidConnectorGUID(Common, CONNECTOR_ID, GUID)
        Catch ex As Exception
            IsValid = False
        End Try

        ' Log the call
        Try
            MsgBuf.Append(IIf(IsValid, "Validated call to ", "Invalid call to "))
            MsgBuf.Append(MethodName)
            MsgBuf.Append(" from GUID: ")
            MsgBuf.Append(GUID)
            MsgBuf.Append(" and IP: " & HttpContext.Current.Request.UserHostAddress)
            Copient.Logger.Write_Log(ActivityLogFile, MsgBuf.ToString, True)
        Catch ex As Exception
            ' ignore
        End Try

        Return IsValid
    End Function
End Class
