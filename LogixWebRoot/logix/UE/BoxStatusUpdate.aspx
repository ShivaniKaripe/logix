<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: configuration.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2010.  All rights reserved by:
  ' *
  ' * NCR Corporation
  ' * 1435 Win Hentschel Boulevard
  ' * West Lafayette, IN  47906
  ' * voice: 888-346-7199  fax: 765-464-1369
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' *
  ' * PROJECT : NCR Advanced Marketing Solution
  ' *
  ' * MODULE  : Preference Manager
  ' *
  ' * PURPOSE : 
  ' *
  ' * NOTES   : 
  ' *
  ' * Version : 1.1b1.0 
  ' *
  ' *****************************************************************************
%>
<script runat="server">

  
  Dim Common As New Copient.CommonInc
  Dim UIInc As New Copient.LogixInc
  Dim Handheld As Boolean = False

  
  
    
  '-------------------------------------------------------------------------------------------------------------  
  
  
  Sub Update_Box_Status()
    
    Dim dst As DataTable
    Dim BoxID As String
    Dim BoxOpen As String
    Dim TargetUser As String = ""
    Dim DupKey As Boolean
    Dim ex As System.Data.SqlClient.SqlException

    On Error GoTo ErrorHandler
    
    BoxID = GetCgiValue("boxid")
    BoxOpen = GetCgiValue("boxopen")
    TargetUser = Common.Extract_Val(GetCgiValue("targetuser"))
    If TargetUser = 0 Then
      Send("TargetUser parameter not specified")
      Exit Sub
    End If

    If BoxID = "" Or BoxOpen = "" Then
      Send("Missing box or boxopen parameter")
    Else
      BoxID = Common.Extract_Val(BoxID)
      If BoxOpen < 0 Or BoxOpen > 1 Then
        Send("Invalid boxopen value!")
        Exit Sub
      End If
      BoxOpen = Common.Extract_Val(BoxOpen)
      Common.QueryStr = "Update AdminUserBoxStates set BoxOpen=" & BoxOpen & " where BoxID=" & BoxID & " and AdminUserID=" & TargetUser & ";"
      Common.LRT_Execute()
      'Send("Rows Affected=" & Common.RowsAffected)
      If Common.RowsAffected = 0 Then
        'the update didn't work, so we need to insert the row into the table
        Common.QueryStr = "Insert into AdminUserBoxStates with (RowLock) (BoxID, AdminUserID, BoxOpen) values (" & BoxID & ", " & TargetUser & ", " & BoxOpen & ");"
        Common.LRT_Execute()
      End If
      Send("OK")
    End If
    Exit Sub
 
ErrorHandler:
    DupKey = False

    If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
      ex = Err.GetException
      If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
        DupKey = True
      End If
    End If

    If DupKey Then 'we got a duplicate key warning - ignore it
      Err.Clear()
      Resume Next
    Else
      Response.Write("<pre>" & Common.Error_Processor() & "</pre>")
    End If

    
  End Sub
  
  
  
  
  
</script>
<%
  '-------------------------------------------------------------------------------------------------------------    
  ' Execution starts here ... 
  
  Dim Mode As String
  Dim CustomerPK As Long
  Dim CustomerTypeID As Integer
  
  Common.AppName = "BoxStatusUpdate.aspx"
  
  Response.Expires = 0
  On Error GoTo ErrorTrap
  If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
  
  AdminUserID = Verify_AdminUser(Common, UIInc)
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Mode = GetCgiValue("mode")
  Select Case Mode
    Case Else
      Update_Box_Status()
  End Select
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
  UIInc = Nothing

  Response.End()


ErrorTrap:
  Response.Write("<pre>" & Common.Error_Processor() & "</pre>")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
  UIInc = Nothing
  
%>
