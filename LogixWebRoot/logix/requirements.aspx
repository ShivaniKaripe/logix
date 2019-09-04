<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%
  ' *****************************************************************************
  ' * FILENAME: requirements.aspx
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

  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""

  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If

  'Response.Expires = 0
  MyCommon.AppName = "requirements.aspx"
  MyCommon.Open_LogixRT()

  LanguageID = MyCommon.Fetch_SystemOption("1")

  Send_HeadBegin("term.requirements")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Logos()
  Send("<div class=""tabs"" id=""tabs"">")
  Send("<hr class=""hidden"" />")
  Send("</div>")
  Send("")
  Send_Subtabs(Logix, 0, 2)
%>
<script type="text/javascript">
function testJavaScript() {
	input_box=alert("<% Sendb(Copient.PhraseLib.Lookup("term.success", LanguageID)) %>");
}
function testPopup(url) {
    popW = 260;
    popH = 100;
    var popup = window.open("about:blank","Blank","width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no");
    if (popup == null) {
        alert("<% Sendb(Copient.PhraseLib.Lookup("term.failed", LanguageID)) %>");
    } else {
        popup.close();
        alert("<% Sendb(Copient.PhraseLib.Lookup("term.success", LanguageID)) %>");
    }
}
</script>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.systemrequirements", LanguageID))%>
  </h1>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <div id="column1">
    <div class="box" id="browser">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.browser", LanguageID))%>
        </span>
      </h2>
      <b>
        <% Sendb(Copient.PhraseLib.Lookup("term.detected", LanguageID))%>:</b>
      <%= Request.Browser.Browser %>
      <%= Request.Browser.Version %>
      <% Send(StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase))%>
      <%= Request.Browser.Platform %>
      <br />
      <br class="half" />
      <% Sendb(Copient.PhraseLib.Lookup("requirements.browser", LanguageID))%>
      <br />
      <hr class="hidden" />
    </div>
    <div class="box" id="javascript">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.javascript", LanguageID))%>
        </span>
      </h2>
      <b>
        <% Sendb(Copient.PhraseLib.Lookup("term.detected", LanguageID))%>:</b>
      <%=Request.Browser.JavaScript%>
      <br />
      <br class="half" />
      <% Sendb(Copient.PhraseLib.Lookup("requirements.javascript", LanguageID))%>
      <br />
      <hr class="hidden" />
    </div>
    <div class="box" id="display">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>
        </span>
      </h2>
      <b>
        <% Sendb(Copient.PhraseLib.Lookup("term.detected", LanguageID))%>:</b>
      <script type="text/javascript">
        document.write(screen.width + 'x' + screen.height + ', ' + screen.colorDepth);
      </script>
      <% Send(Copient.PhraseLib.Lookup("term.bit", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.color", LanguageID), VbStrConv.Lowercase))%>
      <br />
      <br class="half" />
      <% Sendb(Copient.PhraseLib.Lookup("requirements.display", LanguageID))%>
      <br />
      <hr class="hidden" />
    </div>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="unicode">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.unicode", LanguageID))%>
        </span>
      </h2>
      <% Sendb(Copient.PhraseLib.Lookup("requirements.unicode", LanguageID))%>
      <hr class="hidden" />
    </div>
  </div>
  <br clear="all" />
</div>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>