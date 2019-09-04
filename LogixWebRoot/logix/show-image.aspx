<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Drawing" %>

<%
  ' *****************************************************************************
  ' * FILENAME: show-image.aspx 
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
  Dim UIInc As New Copient.LogixInc
  
  Const STANDARD_WIDTH As Integer = 570
  Const STANDARD_HEIGHT As Integer = 150
  
  Const THUMBNAIL_WIDTH As Integer = 150
  Const THUMBNAIL_HEIGHT As Integer = 150
  
  
  '-------------------------------------------------------------------------------------------------------------
  
  Sub Send_Image(ByVal Caller As String)
    Dim ImageBytes(-1) As Byte
    Dim ImgExt As String = "png"
    
    Try
      Select Case Caller.ToLower
        Case "channels"
          ImageBytes = Get_Channel_Image()
        Case "udf"
          ImageBytes = Get_Udf_Image()
        Case Else
          ImageBytes = Get_Image_Not_Found(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT, "Bad Caller") 'Copient.PhraseLib.Lookup("term.imagenotfound", LanguageID))
      End Select
      
    Catch ex As Exception
      ImageBytes = Get_Image_Not_Found(THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT, ex.ToString)

    End Try
    
    Response.ContentType = "image/" & ImgExt
    Response.BinaryWrite(ImageBytes)
  End Sub

  '-------------------------------------------------------------------------------------------------------------
  
  Function Get_Channel_Image() As Byte()
    Dim ImageBytes(-1) As Byte
    Dim DTImage As DataTable
    Dim IsFullSized, IsPreviewImage As Boolean
    Dim ChannelID, MediaTypeID, AssetLanguageID As Integer
    Dim OfferID As Long
    
    ' determine if the request is for the full-sized image or a thumbnail
    IsFullSized = (GetCgiValue("full") = "1")
    IsPreviewImage = (GetCgiValue("preview") = "1")
    Integer.TryParse(GetCgiValue("channelid"), ChannelID)
    Integer.TryParse(GetCgiValue("mediatypeid"), MediaTypeID)
    Integer.TryParse(GetCgiValue("languageid"), AssetLanguageID)
    Long.TryParse(GetCgiValue("offerid"), OfferID)
    
    Common.QueryStr = "select " & IIf(IsPreviewImage, "Preview", "") & "MediaData as MediaData from ChannelOfferAssets with (NoLock) " & _
                      "where ChannelID = " & ChannelID & " and OfferID = " & OfferID & " and MediaTypeID = " & MediaTypeID & " and LanguageID=" & AssetLanguageID
    DTImage = Common.LRT_Select
    If DTImage.Rows.Count > 0 Then
      If Not IsDBNull(DTImage.Rows(0).Item("MediaData")) Then
        ImageBytes = Convert.FromBase64String(DTImage.Rows(0).Item("MediaData"))
      End If
    End If
      
    If ImageBytes Is Nothing OrElse ImageBytes.Length = 0 Then
      ImageBytes = Get_Image_Not_Found(115, 86, Copient.PhraseLib.Lookup("term.imagenotfound", LanguageID))
    ElseIf Not IsFullSized Then
      ImageBytes = GetThumbnailFromBytes(ImageBytes)
    End If
      
    Return ImageBytes
  End Function

  Function Get_Udf_Image() As Byte()
    Dim ImageBytes(-1) As Byte
    Dim sFullPath As String
    
    ' determine if the request is for the full-sized image or a thumbnail
    sFullPath = GetCgiValue("src")
    
    If sFullPath.Length > 0 Then
      Dim webClient = New System.Net.WebClient()
      Try
        ImageBytes = webClient.DownloadData(sFullPath)
      Catch ex As Exception
        ImageBytes = Nothing
      End Try
    End If
      
    If ImageBytes Is Nothing OrElse ImageBytes.Length = 0 Then
      ImageBytes = Get_Image_Not_Found(115, 86, Copient.PhraseLib.Lookup("term.imagenotfound", LanguageID))
    End If
      
    Return ImageBytes
  End Function
  
  Public Function ConvertImageFileToBytes(ByVal ImageFilePath As String) As Byte()
    Dim _tempByte() As Byte = Nothing
    If String.IsNullOrEmpty(ImageFilePath) = True Then
      Return Nothing
    End If
    Try
      Dim _fileInfo As New IO.FileInfo(ImageFilePath)
      Dim _NumBytes As Long = _fileInfo.Length
      Dim _FStream As New IO.FileStream(ImageFilePath, IO.FileMode.Open, IO.FileAccess.Read)
      Dim _BinaryReader As New IO.BinaryReader(_FStream)
      _tempByte = _BinaryReader.ReadBytes(Convert.ToInt32(_NumBytes))
      _fileInfo = Nothing
      _NumBytes = 0
      _FStream.Close()
      _FStream.Dispose()
      _BinaryReader.Close()
      Return _tempByte
    Catch ex As Exception
      Return Nothing
    End Try
  End Function


  '-------------------------------------------------------------------------------------------------------------
      
  Function Get_Image_Not_Found(ByVal Width As Integer, ByVal Height As Integer, ByVal ImageText As String) As Byte()
    Dim bm As Bitmap = Nothing
    Dim ms As New MemoryStream()
    Dim objGraphics As Graphics = Nothing
    Dim ImageBytes(-1) As Byte
    
    Try
      'bm = New Bitmap(570, 150)
      bm = New Bitmap(Width, Height)
      objGraphics = Graphics.FromImage(bm)
      objGraphics.Clear(Color.White)
      objGraphics.DrawRectangle(Pens.Black, 0, 0, 113, 84)
      objGraphics.DrawString("?", New Font("Courier New", 38, FontStyle.Bold), Brushes.Crimson, 35, 15)
      objGraphics.DrawString(ImageText, New Font("Courier New", 8, FontStyle.Bold), Brushes.DarkSlateGray, 5, 5)

      bm.Save(ms, ImageFormat.Png)
      ImageBytes = ms.ToArray
    Catch ex As Exception
      ReDim ImageBytes(-1)
    Finally
      If objGraphics IsNot Nothing Then objGraphics.Dispose()
      If bm IsNot Nothing Then bm.Dispose()
      ms.Close()
      ms.Dispose()
    End Try

    Return ImageBytes
  End Function
  
  '-------------------------------------------------------------------------------------------------------------
  
  Function Get_Image_No_Data() As Byte()
    Dim bm As New Bitmap(570, 150)
    Dim ms As New MemoryStream()
    Dim objGraphics As Graphics = Nothing
    Dim ImageBytes(-1) As Byte
    
    Try
      objGraphics = Graphics.FromImage(bm)
      objGraphics.Clear(Color.White)
      objGraphics.DrawRectangle(Pens.Black, 0, 0, 113, 84)
      objGraphics.DrawString("?", New Font("Courier New", 38, FontStyle.Bold), Brushes.Crimson, 35, 15)
      objGraphics.DrawString(Copient.PhraseLib.Lookup("term.nodata", LanguageID), _
                             New Font("Courier New", 8, FontStyle.Bold), Brushes.DarkSlateGray, 5, 5)

      bm.Save(ms, ImageFormat.Png)
      ImageBytes = ms.ToArray
    Catch ex As Exception
      ReDim ImageBytes(-1)
    Finally
      If objGraphics IsNot Nothing Then objGraphics.Dispose()
      bm.Dispose()
      ms.Close()
      ms.Dispose()
    End Try

    Return ImageBytes
  End Function
  
  Sub Modify_Image_For_No_FullSized(ByRef ImageBytes As Byte())
    Dim bm As Bitmap = Nothing
    Dim bmNew As Bitmap = Nothing
    Dim NewWidth As Integer
    Dim g As Graphics
    Dim StringSize As New SizeF
    Dim ImageText As String = Copient.PhraseLib.Lookup("show-image.NoFullSizedImage", LanguageID, "Full-Sized Image Not Found")
    Dim ImageFont As New Font("Arial", 9, FontStyle.Regular)
    Dim ImageLeft As Integer = 0
    Dim TextLeftOffset As Integer = 3
    
    Using ms As System.IO.MemoryStream = New System.IO.MemoryStream(ImageBytes)
      bm = New Bitmap(ms)
      g = Graphics.FromImage(bm)
      StringSize = g.MeasureString(ImageText, ImageFont)
      NewWidth = IIf(bm.Width < StringSize.Width, StringSize.Width + TextLeftOffset, bm.Width)
      g.Dispose()
      
      ImageLeft = CInt((NewWidth - bm.Width) / 2)
      bmNew = New Bitmap(NewWidth, bm.Height + 30)
      g = Graphics.FromImage(bmNew)
      g.Clear(Color.FromArgb(192, 192, 192))
      g.DrawImage(bm, ImageLeft, 5)

      g.DrawString(ImageText, ImageFont, Brushes.DarkRed, TextLeftOffset, bmNew.Height - 22)
      
      Using msNew As System.IO.MemoryStream = New System.IO.MemoryStream()
        bmNew.Save(msNew, Imaging.ImageFormat.Jpeg)
        ImageBytes = msNew.ToArray
      End Using
      
    End Using

    If bm IsNot Nothing Then bm.Dispose()
    If bmNew IsNot Nothing Then bmNew.Dispose()
    
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------
      
  Private Function GetThumbnailFromBytes(ByVal b As Byte()) As Byte()
    Dim Thumbnail As System.Drawing.Image
    Dim ImageBytes(-1) As Byte
    Dim AspectRect As Rectangle
    
    Using ms As New System.IO.MemoryStream(b)
      Using img As System.Drawing.Image = System.Drawing.Image.FromStream(ms)
        AspectRect = GetImageRect(img.Width, img.Height)
        Thumbnail = img.GetThumbnailImage(AspectRect.Width, AspectRect.Height, Function() False, IntPtr.Zero)
      End Using
    End Using

    Using ms As New System.IO.MemoryStream()
      Thumbnail.Save(ms, Imaging.ImageFormat.Png)
      ImageBytes = ms.GetBuffer()
    End Using

    Return ImageBytes
  End Function
  
  '-------------------------------------------------------------------------------------------------------------
      
  Private Function GetImageRect(ByVal Width As Integer, ByVal Height As Integer) As Rectangle
    Dim rect As New Rectangle
    
    rect.Width = Width
    rect.Height = Height
    
    While (rect.Width > THUMBNAIL_WIDTH OrElse rect.Height > THUMBNAIL_HEIGHT)
      rect.Height /= 1.1
      rect.Width /= 1.1
    End While

    Return rect
  End Function
  
</script>

<%
  
  '-------------------------------------------------------------------------------------------------------------
  ' Main code - execution starts here ...   
  Dim Caller As String = ""
  Dim PKID As Long
  
  Common.AppName = "show-image.aspx"
  Response.Expires = 0
  
  On Error GoTo ErrorTrap
  If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
  
  AdminUserID = Verify_AdminUser(Common, UIInc)
  'If LanguageID <= 0 Then LanguageID = 1
  
  Response.Expires = 0
  
  Caller = GetCgiValue("caller")
  
  Send_Image(Caller)
  
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
