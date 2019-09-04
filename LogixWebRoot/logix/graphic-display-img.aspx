<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%@ Import Namespace="System.IO" %>
<%
  Dim CopientFileName As String = "graphic-display-img.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim GraphicPath As String = ""
  Dim objBitmap As New Bitmap(350, 208)
  Dim objGraphics As Graphics = Graphics.FromImage(objBitmap)
  Dim font As New Font("Courier New", 10, FontStyle.Regular)
  Dim fileExt As String = ""
  Dim MemStream As New System.IO.MemoryStream()
  Dim LangID As Integer = 1
  Dim AdminUserID As Long
  
  MyCommon.AppName = "graphic-display-img"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  GraphicPath = Request.QueryString("path")
  LangID = IIf((Request.QueryString("lang") <> ""), CInt(Request.QueryString("lang")), 1)
  
  If (GraphicPath <> "" And GraphicPath.Length > 3) Then
    fileExt = Right(GraphicPath, 3)
    If (File.Exists(GraphicPath)) Then
      fileExt = Right(GraphicPath, 3)
      If (fileExt = "gif") Then
        Response.ContentType = "image/gif"
      ElseIf (fileExt = "jpg" Or fileExt = "jpeg") Then
        Response.ContentType = "image/jpeg"
      Else
        Response.ContentType = "image/" & fileExt
      End If
      Response.BinaryWrite(File.ReadAllBytes(GraphicPath))
    Else
      Response.ContentType = "image/jpeg"
      objGraphics.Clear(Color.White)
      If (GraphicPath.IndexOf("_tn.") > -1) Then
        objBitmap = New Bitmap(115, 86)
        objGraphics = Graphics.FromImage(objBitmap)
        objGraphics.Clear(Color.White)
        objGraphics.DrawRectangle(Pens.Black, 0, 0, 113, 84)
        objGraphics.DrawString("?", New Font("Courier New", 38, FontStyle.Bold), Brushes.Crimson, 35, 5)
        objGraphics.DrawString(Copient.PhraseLib.Lookup("graphics.imagenotfound", LangID), New Font("Courier New", 8, FontStyle.Bold), Brushes.Crimson, 5, 55)
      Else
        objGraphics.DrawRectangle(Pens.Black, 0, 0, 349, 207)
        objGraphics.DrawString(Copient.PhraseLib.Lookup("graphics.imagenotfound", LangID) & ":" & vbCrLf & vbCrLf & ParsePath(GraphicPath), font, Brushes.DarkSlateGray, 5, 5)
      End If
      objBitmap.Save(MemStream, ImageFormat.Jpeg)
      MemStream.WriteTo(Response.OutputStream)
    End If
  End If
%>
<script runat="server">
  Function ParsePath(ByVal path As String) As String
    Dim pathBuf As New StringBuilder()
    Dim tempStr As String
    Dim i As Integer
    Dim slashPos As Integer = 0
    Dim prevSlashPos As Integer = 0
    Dim startPos As Integer = 0
    
    If (path.Length > 40) Then
      For i = 1 To path.Length
        tempStr = Mid(path, i, 40)
        If (tempStr.Length < 40) Then
          pathBuf.Append(tempStr)
          i += tempStr.Length
        Else
          slashPos = tempStr.LastIndexOf("\")
          If (slashPos <= 0) Then
            tempStr = Mid(tempStr, i, 37) & "..."
            i += 39
          Else
            tempStr = Mid(tempStr, i, slashPos)
            i += (slashPos - 1)
          End If
          pathBuf.Append(tempStr)
          pathBuf.Append(vbCrLf)
        End If
      Next
    Else
      pathBuf.Append(path)
    End If
    
    Return pathBuf.ToString
  End Function
</script>
<%
done:
  MyCommon = Nothing
  Logix = Nothing
  objGraphics.Dispose()
  objBitmap.Dispose()
  Response.Flush()
  Response.End()
%>
