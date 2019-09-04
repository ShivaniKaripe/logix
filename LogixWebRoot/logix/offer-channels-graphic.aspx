<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-channels.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2013.  All rights reserved by:
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
 %>

<script runat="server">
  Dim Common As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim Handheld As Boolean = False
  Dim MediaTypeID As Integer
  Dim OfferID As Long
  Dim ChannelID As Integer
  Dim AssetLanguageID As Integer
  Dim Mode As String = ""
  Dim InfoMessage As String = ""
  Dim ShowImg As Boolean = True
  
  Public Structure ImageResults
    Public Valid As Boolean
    Public Message As String
    Public Format As System.Drawing.Imaging.ImageFormat
    Public Width As Integer
    Public Height As Integer
  End Structure
  
  Private Sub Send_Page()
    Send("<html>")
    Send("  <head>")
    Send_Page_Script()
    Send("  </head>")
    Send("  <body>")
    Send("    <form id=""uploadform"" name=""uploadform"" method=""post"" enctype=""multipart/form-data"">")
    Send("      <input type=""hidden"" name=""Mode"" id=""Mode"" value=""Preview"" />")
    Send("      <input type=""hidden"" name=""ChannelID"" id=""ChannelID"" value=""" & ChannelID & """ />")
    Send("      <input type=""hidden"" name=""OfferID"" id=""OfferID"" value=""" & OfferID & """ />")
    Send("      <input type=""hidden"" name=""MediaTypeID"" id=""MediaTypeID"" value=""" & MediaTypeID & """ />")
    Send("      <input type=""hidden"" name=""LanguageID"" id=""LanguageID"" value=""" & AssetLanguageID & """ />")
    Send("      <div style=""width:98%;font-size: 12px; font-family: Arial, Helvetica, sans-serif;"">")
    Send("      <div style=""float: left; width:75%;"">")
    Send("        " & GetUploadLabelText(MediaTypeID) & "<br />")
    Send("        <input type=""file"" name=""GraphicFile"" style=""width: 80%;"" accept=""image/*"" onchange=""document.uploadform.submit();"" value="""" />")
        '    Send("<div id=""divfile"" style=""height:0px;overflow:hidden"">")
        'Send("<input type=""file"" id=""browse"" name=""GraphicFile"" style=""width: 80%;"" accept=""image/*"" onchange=""fileonclick();"" />")
        'Send("</div>")
        'Send("<button type=""button"" onclick=""chooseFile();"">"&Copient.PhraseLib.Lookup("term.browse", LanguageID) &"</button>")
        'Send("<label id=""lblfileupload"" name=""lblfileupload"">"&Copient.PhraseLib.Lookup("term.nofilesselected", LanguageID)&"</label>")
    Send("      </div>")
    Send("      <div style=""float: right;"">")
    Send("        <input type=""submit"" name=""Save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ onclick=""document.getElementById('Mode').value = 'SaveSelection';""" & IIf(Mode = "" OrElse InfoMessage <> "", " disabled=""disabled""", "") & " />")
    Send("        <input type=""submit"" name=""Cancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""document.getElementById('Mode').value = 'CancelSelection';"" />")
    Send("      </div>")
    Send("      </div>")
    Send("      <div style=""float: left;clear: left; width: 100%;margin-top: 5px;; margin-bottom: 5px;"">")
    Send("        <hr />")
    Send("      </div>")
    Send("      <div style=""float: left; clear: left; width: 98%;"">")
    Send("        <span id=""infoMsg""" & IIf(InfoMessage = "", " style=""display:none;""", "") & ">" & InfoMessage & "</span>")
    If ShowImg Then Send_Graphic()
    Send("      </div>")
    Send("    </form>")
    Send("  </body>")
    Send("</html>")
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------  

  Private Sub Save_Preview()
    Dim MediaData As String = ""
    Dim Results As ImageResults
    
    If Request.Files.Count > 0 Then
      
      Using ms As New System.IO.MemoryStream()
        Request.Files(0).InputStream.CopyTo(ms)
        ' validate that this is an image and in a allowable format.
        Results = ValidateImageStream(ms, MediaTypeID)
        If Results.Valid Then
          MediaData = Convert.ToBase64String(ms.GetBuffer)
        End If
      End Using

      If Not Results.Valid Then
        InfoMessage = Results.Message
        ShowImg = False

      ElseIf OfferID > 0 AndAlso ChannelID > 0 AndAlso MediaTypeID > 0 AndAlso AssetLanguageID > 0 AndAlso MediaData.Trim().Length > 0 Then
        ' save the graphic to the preview media data column 
        Common.QueryStr = "dbo.pt_ChannelOfferAssets_SavePreview"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@MediaTypeID", SqlDbType.Int).Value = MediaTypeID
        Common.LRTsp.Parameters.Add("@PreviewMediaData", SqlDbType.NVarChar, -1).Value = MediaData
        Common.LRTsp.Parameters.Add("@MediaFormatID", SqlDbType.Int).Value = GetMediaFormatID(Results.Format)
        Common.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int).Value = AssetLanguageID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
      End If
    End If

  End Sub
  
  '-------------------------------------------------------------------------------------------------------------  

  Private Sub Save_Selection()

    If OfferID > 0 AndAlso ChannelID > 0 AndAlso MediaTypeID > 0 AndAlso AssetLanguageID > 0 Then
      ' copy the contents of the preview media data column into the media data column
      Common.QueryStr = "dbo.pt_ChannelOfferAssets_SaveSelected"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
      Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
      Common.LRTsp.Parameters.Add("@MediaTypeID", SqlDbType.Int).Value = MediaTypeID
      Common.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int).Value = AssetLanguageID
      Common.LRTsp.ExecuteNonQuery()
      Common.Close_LRTsp()
      
      Common.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Detokenize("offer-channels.createdGraphic", LanguageID, GetMediaTypeName(MediaTypeID), GetChannelName(ChannelID)))
    End If

  End Sub
  
  '-------------------------------------------------------------------------------------------------------------  

  Private Sub Cancel_Selection()

    If OfferID > 0 AndAlso ChannelID > 0 AndAlso MediaTypeID > 0 AndAlso AssetLanguageID > 0 Then
      ' delete the contents of the preview media data column, if there is no data in media data then delete the entire row.
      Common.QueryStr = "dbo.pt_ChannelOfferAssets_DeletePreview"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
      Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
      Common.LRTsp.Parameters.Add("@MediaTypeID", SqlDbType.Int).Value = MediaTypeID
      Common.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int).Value = AssetLanguageID
      Common.LRTsp.ExecuteNonQuery()
      Common.Close_LRTsp()
    End If

  End Sub
  
  '-------------------------------------------------------------------------------------------------------------  

  Private Sub Send_Graphic()
    If OfferID > 0 AndAlso ChannelID > 0 AndAlso MediaTypeID > 0 AndAlso IsImageData() Then
      Send("<center>")
      Send("  <img src=""show-image.aspx?caller=channels&full=1" & IIf(Mode = "Preview", "&preview=1", "") & _
           "&channelid=" & ChannelID & "&offerid=" & OfferID & "&mediatypeid=" & MediaTypeID & "&languageid=" & AssetLanguageID & "&t=" & Date.Now.Ticks & """ id=""imgGraphic"" />")
      Send("</center>")
    End If
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------  

  Private Sub Send_Page_Script()
    Send("<script type=""text/javascript"">")
    Send("")
    Select Case Mode
      Case "SaveSelection", "CancelSelection"
        Send("  window.onload = hideSelf;")
    End Select
    Send("")
    Send("  function hideSelf() {")
    If Mode = "SaveSelection" Then
      Send("    parent.reloadGraphic('" & "assetCh" & ChannelID & "Mt" & MediaTypeID & "L" & AssetLanguageID & "', " & ChannelID & ");")
    End If
    Send("    parent.hideGraphicSelector();")
    Send("  }")
    Send("")
  '   Send("  function chooseFile() {")
  ' Send("   document.getElementById(""browse"").click();")
  ' Send(" }")
  ' Send(" function fileonclick()")
  ' Send(" {")
  ' Send(" var filename=document.getElementById(""browse"").value;")
  ' Send(" document.getElementById(""lblfileupload"").innerText = filename.replace(""C:\\fakepath\\"", """");")
  '   Send("try{")
  '      Send("document.uploadform.submit();")
  '        Send("  } catch(err) { for(var i=0;i<1000;i++){} document.uploadform.submit();}")
  'Send(" }")
    Send("<" & "/script>")
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------  

  Private Function IsImageData() As Boolean
    Dim dt As DataTable
    
    Common.QueryStr = "select ChannelID from ChannelOfferAssets with (NoLock) " & _
                      "where ChannelID=" & ChannelID & " and OfferID=" & OfferID & _
                      " and MediaTypeID=" & MediaTypeID & "and LanguageID=" & AssetLanguageID & ";"
    dt = Common.LRT_Select()
    
    Return (dt.Rows.Count > 0)
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  

  Private Sub GetFormValues()
    OfferID = Common.Extract_Val(GetCgiValue("OfferID"))
    ChannelID = Common.Extract_Val(GetCgiValue("ChannelID"))
    MediaTypeID = Common.Extract_Val(GetCgiValue("MediaTypeID"))
    AssetLanguageID = Common.Extract_Val(GetCgiValue("LanguageID"))
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------  

  Function GetChannelName(ByVal ChannelID As Integer) As String
    Dim Name As String = ""
    Dim dt As DataTable
    
    Common.QueryStr = "select Name, PhraseTerm from Channels with (NoLock) where ChannelID=" & ChannelID
    dt = Common.LRT_Select()
    If dt.Rows.Count > 0 Then
      Name = Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("PhraseTerm"), "").ToString(), LanguageID, Common.NZ(dt.Rows(0).Item("Name"), ""))
    End If
    
    Return Name
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  

  Function GetMediaTypeName(ByVal MediaTypeID As Integer) As String
    Dim Name As String = ""
    Dim dt As DataTable
    
    Common.QueryStr = "select Name, PhraseTerm from ChannelMediaTypes with (NoLock) where MediaTypeID=" & MediaTypeID
    dt = Common.LRT_Select()
    If dt.Rows.Count > 0 Then
      Name = Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("PhraseTerm"), "").ToString(), LanguageID, Common.NZ(dt.Rows(0).Item("Name"), ""))
    End If
    
    Return Name
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  

  Function ValidateImageStream(ByVal ms As System.IO.MemoryStream, ByVal MediaTypeID As Integer) As ImageResults
    Dim Results As New ImageResults
    Dim img As System.Drawing.Image
    Dim Formats As List(Of System.Drawing.Imaging.ImageFormat)
    
    If ms IsNot Nothing Then
      Try
        ' first check if the stream is even an image
        img = System.Drawing.Image.FromStream(ms)

        With Results
          .Format = img.RawFormat
          .Width = img.Width
          .Height = img.Height
        End With
        
        ' next check if the image meets all the restrictions (e.g. size, format)
        Formats = GetAllowableImageFormats(MediaTypeID)
        If Formats IsNot Nothing AndAlso Formats.Count > 0 AndAlso Not GetAllowableImageFormats(MediaTypeID).Contains(img.RawFormat) Then
          Results.Valid = False
          Results.Message = Copient.PhraseLib.Detokenize("offer-channels-graphic.invalidFileFormat", LanguageID, String.Join(",", Formats))
        ElseIf Not IsValidSize(MediaTypeID, img.Width, img.Height) Then
          Results.Valid = False
          Results.Message = Copient.PhraseLib.Detokenize("offer-channels-graphic.invalidFileSize", LanguageID, Results.Width, Results.Height)
        Else
          Results.Valid = True
          Results.Message = ""
        End If
        
      Catch aex As ArgumentException
        Results.Valid = False
        Results.Message = Copient.PhraseLib.Lookup("offer-channels-graphic.notImage", LanguageID)
        
      Catch ex As Exception
        Results.Valid = False
        Results.Message = ex.ToString
      End Try
    End If
    
    Return Results
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  

  Function GetAllowableImageFormats(ByVal MediaTypeID As Integer) As List(Of System.Drawing.Imaging.ImageFormat)
    Dim Formats As New List(Of System.Drawing.Imaging.ImageFormat)
    Dim dt As DataTable
    
    Common.QueryStr = "select IsNull(MF.Name, '') as FormatName from MediaTypeFormats as MTF with (NoLock) " & _
                      "inner join MediaFormats as MF with (NoLock) on MF.MediaFormatID = MTF.MediaFormatID " & _
                      "where MediaTypeID = " & MediaTypeID
    dt = Common.LRT_Select()
    For Each row As DataRow In dt.Rows
      Select Case row.Item("FormatName").ToString().ToUpper()
        Case "JPG", "JPEG"
          Formats.Add(System.Drawing.Imaging.ImageFormat.Jpeg)
        Case "GIF"
          Formats.Add(System.Drawing.Imaging.ImageFormat.Gif)
        Case "PNG"
          Formats.Add(System.Drawing.Imaging.ImageFormat.Png)
        Case Is = "BMP"
          Formats.Add(System.Drawing.Imaging.ImageFormat.Bmp)
      End Select
    Next
    
    Return Formats
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  

  Function GetMediaFormatID(ByVal Format As System.Drawing.Imaging.ImageFormat) As Integer
    Dim MediaFormatID As Integer = 0
    Dim dt As DataTable
    Dim MediaFormatName As String = ""
    
    If Format.Equals(System.Drawing.Imaging.ImageFormat.Jpeg) Then
      MediaFormatName = "JPEG"
    ElseIf Format.Equals(System.Drawing.Imaging.ImageFormat.Png) Then
      MediaFormatName = "PNG"
    ElseIf Format.Equals(System.Drawing.Imaging.ImageFormat.Gif) Then
      MediaFormatName = "GIF"
    End If
    
    Common.QueryStr = "select MediaFormatID from MediaFormats with (NoLock) where Name = '" & MediaFormatName & "';"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      MediaFormatID = Common.NZ(dt.Rows(0).Item("MediaFormatID"), 0)
    End If
    
    Return MediaFormatID
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  
  
  Function GetRestrictionsTable(ByVal MediaTypeID As Integer) As DataTable
    Dim dt As DataTable
    
    ' if no size restrictions are in place, then it is considered a valid size.
    Common.QueryStr = "select MPT.Name, MPT.PhraseTerm, MPT.MinValue, MPT.MaxValue from MediaParamTypes as MPT with (NoLock) " & _
                      "inner join ChannelMediaParams as CMP with (NoLock) on CMP.ParamTypeID = MPT.ParamTypeID " & _
                      "inner join ChannelMedia as CM with (NoLock) on CM.ChannelMediaID = CMP.ChannelMediaID " & _
                      "where CM.MediaTypeID =" & MediaTypeID
    dt = Common.LRT_Select
    
    Return dt
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  

  Function IsValidSize(ByVal MediaTypeID As Integer, ByVal Width As Integer, ByVal Height As Integer) As Boolean
    Dim Valid As Boolean = True
    Dim dt As DataTable
    
    ' if no size restrictions are in place, then it is considered a valid size.
    dt = GetRestrictionsTable(MediaTypeID)
    If dt.Rows.Count > 0 Then
      For Each row As DataRow In dt.Rows
        Select Case Common.NZ(row.Item("Name"), "").ToString.ToUpper
          Case "GRAPHIC WIDTH"
            Valid = (Width >= Common.NZ(row.Item("MinValue"), Integer.MinValue) AndAlso Width <= Common.NZ(row.Item("MaxValue"), Integer.MaxValue))
          Case "GRAPHIC HEIGHT"
            Valid = (Height >= Common.NZ(row.Item("MinValue"), Integer.MinValue) AndAlso Height <= Common.NZ(row.Item("MaxValue"), Integer.MaxValue))
        End Select
        
        If Not Valid Then Exit For
      Next
    End If
   
    Return Valid
  End Function
  
  '-------------------------------------------------------------------------------------------------------------  

  Function GetUploadLabelText(ByVal MediaTypeID As Integer) As String
    Dim LabelText As New StringBuilder()
    Dim Formats As List(Of System.Drawing.Imaging.ImageFormat) = GetAllowableImageFormats(MediaTypeID)
    Dim dt As DataTable
    
    LabelText.Append(Copient.PhraseLib.Detokenize("offer-channels.graphicNote", LanguageID, GetMediaTypeName(MediaTypeID)))
    
    If Formats IsNot Nothing AndAlso Formats.Count > 0 Then
      LabelText.Append("<br />" & Copient.Detokenize("offer-channels-graphic.acceptedFileTypes", LanguageID, String.Join(",", Formats)))
    End If
    
    dt = GetRestrictionsTable(MediaTypeID)
    If dt.Rows.Count > 0 Then
      LabelText.Append("&nbsp;&nbsp;&nbsp;")
      For Each row As DataRow In dt.Rows
        LabelText.Append("<BR />" & Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseTerm"), ""), LanguageID, Common.NZ(row.Item("Name"), "")) & ": ")
        LabelText.Append("[" & Copient.PhraseLib.Lookup("term.min", LanguageID) & ": " & CInt(Common.NZ(row.Item("MinValue"), 0)) & "&nbsp;&nbsp;")
        LabelText.Append(Copient.PhraseLib.Lookup("term.max", LanguageID) & ": " & CInt(Common.NZ(row.Item("MaxValue"), 0)) & "];")
      Next
    End If
    
    Return LabelText.ToString
  End Function
  
  </script>

  <%
    Common.AppName = "offer-channels-graphic.aspx"
    Response.Expires = 0
    On Error GoTo ErrorTrap

    If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
      Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    AdminUserID = Verify_AdminUser(Common, Logix)
    GetFormValues()
    
    Mode = GetCgiValue("Mode")
    Select Case Mode
      Case "Preview"
        Save_Preview()

      Case "SaveSelection"
        Save_Selection()
        
      Case "CancelSelection"
        Cancel_Selection()
        
    End Select

    Send_Page()
    
    If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
    Common = Nothing
    Logix = Nothing

    Response.End()


ErrorTrap:
    Response.Write("<pre>" & Common.Error_Processor() & "</pre>")
    If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
    Common = Nothing
    Logix = Nothing
  
    
%>