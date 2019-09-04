<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    '-----------------------------------------------------------------------------
    'Execution starts here ... 
  
    MyCommon.AppName = "UEfolderFeeds.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
    If (Request.QueryString("buyerID") <> "") Then
        Response.Expires = 0
        Response.Clear()
        Response.ContentType = "text/html"
        GenerateFolderList(Request.QueryString("buyerID"))
    End If
    Response.Flush()
    Response.End()
  
%>
<script runat="server">
    Public MyCommon As New Copient.CommonInc
    Dim CopientFileName As String = "UEfolderFeeds.aspx"
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
    Dim Logix As New Copient.LogixInc
  
    Sub GenerateFolderList(ByVal buyerID As Integer)
        Dim rst As DataTable
        Dim row As DataRow
        Dim i As Integer
        Dim folderNames As String = ""
        Dim folderList As String = ""
        
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Try
            MyCommon.QueryStr = "select FI.FolderID,F.FolderName from FolderItems FI " & _
                                 "inner join Folders F on FI.FolderID=F.FolderID where(LinkID = " & buyerID & "And LinkTypeID = 2)"
            
            rst = MyCommon.LRT_Select
            'If Buyer is not set with default folder, then get the default UE folder if set any.
    If rst.Rows.Count = 0 Then
        MyCommon.QueryStr = "select FolderID,FolderName from Folders where DefaultUEFolder=1"
        rst = MyCommon.LRT_Select
      End If
            If rst.Rows.Count > 0 Then
                folderNames += "<ul>"
                For Each row In rst.Rows
                    folderList += MyCommon.NZ(row.Item("FolderID"), "").ToString()
                    folderNames += "<li>"
                    folderNames += MyCommon.NZ(row.Item("FolderName"), "")
                    folderNames += "</li>"
                Next
                folderNames += "</ul>"
            Else
                folderNames = Copient.PhraseLib.Lookup("term.noneselected", LanguageID)
            End If
                
            folderNames += "<input type=""hidden"" id=""tempfolderList"" name=""tempfolderList"" value=""" & folderList & """/>"
            Send(folderNames)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub
</script>
