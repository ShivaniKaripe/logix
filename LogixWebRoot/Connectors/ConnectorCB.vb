'
' *****************************************************************************
' * FILENAME: ConnectorCB.vb 
' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' * Copyright © 2002 - 2009.  All rights reserved by:
' *
' * NCR Corporation
' * 1435 Win Hentschel Boulevard
' * West Lafayette, IN  47906
' * voice: 888-346-7199  fax: 765-464-1369
' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' *
' * PROJECT : NCR Advanced Marketing Solution
' *
' * MODULE  : Logix
' *
' * PURPOSE : 
' *
' * NOTES   : 
' *
' * Version : 7.3.1.138972 
' *
' *****************************************************************************
'

Public Class ConnectorCB
  Inherits System.Web.UI.Page

  Public LSVersion As String
  Public LSBuild As String

  Public Sub Send_Response_Header(ByVal ResponseStr As String, ByVal MajorVersion As String, ByVal MinorVersion As String, ByVal Build As String, ByVal BuildRevision As String)
    Response.ContentType = "text/html"
    Send(ResponseStr & "," & Trim(CStr(MajorVersion)) & "," & Trim(CStr(MinorVersion)) & "," & Trim(CStr(Build)) & "," & Trim(CStr(BuildRevision)))
  End Sub

  Public Sub Send(ByVal WebText As String)
    Response.Write(WebText & vbCrLf)
  End Sub

  Public Sub Sendb(ByVal WebText As String)
    Response.Write(WebText)
  End Sub

  Public Function Get_Raw_Form(ByVal RawStream As System.IO.Stream) As String
    Dim Index As Long
    Dim RawLen, strRead As Long
    Dim RawRequest As String

    RawRequest = ""
    ' Find number of bytes in stream.
    RawLen = CInt(RawStream.Length)
    ' Create a byte array.
    Dim RawArray(RawLen) As Byte
    ' Read stream into byte array.
    strRead = RawStream.Read(RawArray, 0, RawLen)
    ' Convert byte array to a text string.
    For Index = 0 To RawLen - 1
      RawRequest = RawRequest & Chr(RawArray(Index))
    Next Index
    Get_Raw_Form = RawRequest
  End Function

  Public Function Get_Page_Value(ByVal ParamName As String) As String
    'This function returns the value associated with the Form Field or Get Parameter from the calling web page
    Dim TempVal As String

    tempval = Request.QueryString(ParamName)
    If tempval = "" Then
      tempval = Request.Form(ParamName)
    End If

    Get_Page_Value = TempVal

  End Function

  Public Function GetCgiValue(ByVal VarName As String) As String
    Dim TempVal As String
    TempVal = ""
    TempVal = Request.QueryString(VarName)
    If TempVal = "" Then TempVal = Request.Form(VarName)
    GetCgiValue = TempVal
  End Function

End Class
