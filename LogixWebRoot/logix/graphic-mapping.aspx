<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable = Nothing
  Dim row As DataRow = Nothing
  Dim AdID As Long
  Dim ImgWidth As Integer
  Dim ImgHeight As Integer
  Dim ImgType As Integer
  Dim LabelsStr As String = ""
  Dim GraphicPath As String = ""
  Dim imgExt As String = ""
  Dim ControlsTop As Integer = 0
  Const DEFAULT_GRAPHIC_PATH As String = "C:\"
  Dim htmlBuf As New StringBuilder
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  AdID = MyCommon.Extract_Val(Request.QueryString("adId"))
  
  Send_HeadBegin("term.logix", "term.map")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  
  MyCommon.QueryStr = "select Name, Width, Height, ImageType from OnScreenAds with (NoLock) where OnScreenAdID=" & AdID & " and deleted=0;"
  dt = MyCommon.LRT_Select
  If (dt.Rows.Count > 0) Then
    ImgWidth = MyCommon.NZ(dt.Rows(0).Item("Width"), 1)
    ImgHeight = MyCommon.NZ(dt.Rows(0).Item("Height"), 1)
    ImgType = MyCommon.NZ(dt.Rows(0).Item("ImageType"), 1)
    ControlsTop = ImgHeight + 10
  End If
  
  ' Build the graphic file path string
  GraphicPath = MyCommon.Fetch_SystemOption(47)
  If (GraphicPath.Trim().Length = 0) Then
    GraphicPath = DEFAULT_GRAPHIC_PATH
  End If
  If Not (Right(GraphicPath, 1) = "\") Then
    GraphicPath = GraphicPath & "\"
  End If
  If (ImgType = "1") Then
    imgExt = "jpg"
  ElseIf (ImgType = "2") Then
    imgExt = "gif"
  End If
  GraphicPath = GraphicPath & CStr(AdID) & "img." & imgExt
%>

<script type="text/javascript">
// Process mouse
document.onmousedown = doMouseDown
document.onmousemove = doMouseMove
document.onmouseup = doMouseUp

var curEl = null  // Track current item.
var startX = 0, startY = 0;
var bGrabbingPos = false;
var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;
var imgWidth=0, imgHeight=0;

function doMouseMove(e) {
    var elemBox = document.getElementById("imgBox");
    var newleft = 0, newtop = 0;
    var scrTop = (isIE) ? parseInt(document.body.scrollTop) : parseInt(window.pageYOffset);
    var scrLeft = (isIE) ? parseInt(document.body.scrollLeft) : parseInt(window.pageXOffset);
    var xPos = 0, yPos = 0;
    if (e == null) e = event;
    // Check if mouse button is down
    //if ((leftButtonClicked(e)) && (curEl != null)) {
    if (bGrabbingPos && (curEl != null)) {
        if (null != curEl) {
            elemBox.style.display = "inline";
            elemBox.style.left = (startX + scrLeft) + 'px';
            elemBox.style.top = (startY + scrTop) + 'px';
            xPos = (isIE) ? e.clientX : e.pageX;
            yPos = (isIE) ? e.clientY : e.pageY;

            if (xPos < startX) {
                elemBox.style.width = (isIE) ? startX - xPos : (startX - xPos) + 'px';
                startX = xPos;
            } else {
                elemBox.style.width = (isIE) ? xPos - startX : (xPos - startX) + 'px';
                if (startX + parseInt(elemBox.style.width) > parseInt(document.getElementById("imgDiv").style.width)) {
                    elemBox.style.width = (isIE) ? parseInt(document.getElementById("imgDiv").style.width) - startX : parseInt(document.getElementById("imgDiv").style.width) - startX + 'px';
                }
            }            
            if (yPos < startY) {
                elemBox.style.height = (isIE) ? startY - yPos : (startY - yPos) + 'px';
                startY = yPos;
            } else {
                elemBox.style.height = (isIE) ? yPos - startY :  (yPos - startY) + 'px';
                if (startY + parseInt(elemBox.style.height) > parseInt(document.getElementById("imgDiv").style.height)) {
                    elemBox.style.height = (isIE) ? parseInt(document.getElementById("imgDiv").style.height) - startY : parseInt(document.getElementById("imgDiv").style.height) - startY + 'px';
                }
            }
            if (isIE) {
                e.returnValue = false
            } else {
                return false;
            }
        }
    }
}

function doMouseDown(e) {
    var box = document.getElementById("imgBox");
    var srcElem = null;
    var scrTop = (isIE) ? parseInt(document.body.scrollTop) : parseInt(window.pageYOffset);
    var scrLeft = (isIE) ? parseInt(document.body.scrollLeft) : parseInt(window.pageXOffset);
    if (e == null) {
        e = event;
    }
    srcElem = (!isIE) ?  e.target : event.srcElement;
    if (leftButtonClicked(e) && srcElem.name != null) {
        if (srcElem.name.length >=5 && srcElem.name.substring(0,5)=="tpImg")  {
            bGrabbingPos = true;
            if (!isIE) {
                box.style.opacity = .7;
            }
            resetBox();
            curEl = srcElem.offsetParent;
            startX = e.clientX + scrLeft;
            startY = e.clientY + scrTop;
            box.style.left = (isIE) ? startX : startX + 'px;';
            box.style.top = (isIE) ? startY : startY + 'px;';
            box.style.display = "inline";
            return false;
        }
    }
}

function doMouseUp(e) {
    var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;                
    var box = document.getElementById("imgBox");
    if ( (!isIE && e.which==1) || (isIE) ) {
        //alert("imgWidth: " + document.getElementById("imgDiv").style.width)
        //alert(startX + parseInt(document.getElementById("imgBox").style.width));
        curEl=null
        assignPoints();
        bGrabbingPos = false;
    }
}

function leftButtonClicked(e) {
    var retVal = false;
    var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;                
    if (isIE) {
      retVal = (e.button == 1);
    } else {
      retVal = (e.button == 0 && e.which == 1);  
    }
    return retVal;
}

function resetBox() {
    var box = document.getElementById("imgBox");
    startX = 0;
    startY = 0;
    box.style.left = (isIE) ? startX : startX + 'px;';
    box.style.top = (isIE) ? startY : startY + 'px;';
    box.style.width = (isIE) ? 0 : '0px;';
    box.style.height = (isIE) ? 0 : '0px';
    box.style.display = "none";
}      

function assignPoints() {
    var elemX = document.getElementById("tpX");
    var elemY = document.getElementById("tpY");
    var elemWidth = document.getElementById("tpWidth");
    var elemHeight = document.getElementById("tpHeight");
    var box = document.getElementById("imgBox");
    var width = 0, height = 0;
    var scrTop = (isIE) ? parseInt(document.body.scrollTop) : parseInt(window.pageYOffset);
    var scrLeft = (isIE) ? parseInt(document.body.scrollLeft) : parseInt(window.pageXOffset);
    var imgDiv = document.getElementById("imgDiv");
    width = parseInt(box.style.width);
    height = parseInt(box.style.height);
    if (bGrabbingPos && width > 0 && height > 0) {
        elemX.value = startX - parseInt(imgDiv.style.left) + parseInt(scrLeft);
        elemY.value = startY - parseInt(imgDiv.style.top) + parseInt(scrTop);
        elemWidth.value = parseInt(box.style.width);
        elemHeight.value = parseInt(box.style.height);
    }
}

function clearEntry() {
    var elemName = document.getElementById("tpName");
    var elemX = document.getElementById("tpX");
    var elemY = document.getElementById("tpY");
    var elemWidth = document.getElementById("tpWidth");
    var elemHeight = document.getElementById("tpHeight");
    elemName.value = "";
    elemX.value = "";
    elemY.value = "";
    elemWidth.value = "";
    elemHeight.value = "";
}

function addTouchpoint() {
    if (opener!=null) {
        var tpName = opener.document.getElementById("txtAreaName");
        if (tpName != null) {
            tpName.value = document.getElementById("tpName").value;
        }
        var tpX = opener.document.getElementById("txtXPos");
        if (tpX != null) {
            tpX.value = document.getElementById("tpX").value;
        }
        var tpY = opener.document.getElementById("txtYPos");
        if (tpY != null) {
            tpY.value = document.getElementById("tpY").value;
        }
        var tpWidth = opener.document.getElementById("txtWidth");
        if (tpWidth != null) {
            tpWidth.value = document.getElementById("tpWidth").value;
        }
        var tpHeight = opener.document.getElementById("txtHeight");
        if (tpHeight != null) {
            tpHeight.value = document.getElementById("tpHeight").value;
        }
    window.close();
    }
}

function redraw() {
    var box = document.getElementById("imgBox");
    var elemX = document.getElementById("tpX");
    var elemY = document.getElementById("tpY");
    var elemWidth = document.getElementById("tpWidth");
    var elemHeight = document.getElementById("tpHeight");
    var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;                
    var scrTop = (isIE) ? parseInt(document.body.scrollTop) : parseInt(window.pageYOffset);
    var scrLeft = (isIE) ? parseInt(document.body.scrollLeft) : parseInt(window.pageXOffset);
    box.style.display = "inline";
    
    if (!isIE) {
        box.style.opacity = .7;
    }
    box.style.left = (isIE) ? elemX.value : elemX.value + 'px;';
    box.style.top = (isIE) ? elemX.value : elemY.value + 'px;';
    box.style.width = (isIE) ? elemWidth.value : elemWidth.value + 'px;';
    box.style.height = (isIE) ? elemHeight.value : elemHeight.value + 'px;';
}

function focusOnName() {
    var elemName = document.getElementById("tpName");
    
    if (elemName != null) {
        elemName.focus();
    }
}

function initPage() {
    var elemImage = document.getElementById("imgDiv");
    var imgWidth = 0, imgHeight = 0;
    
    if (elemImage != null) {
        imgWidth = parseInt(elemImage.style.width);
        imgHeight = parseInt(elemImage.style.height);
    }
    focusOnName();
    drawWithCurrentValues()
}

function drawWithCurrentValues() {
    var bValueSet = false;
    if (opener!=null) {
        var tpName = opener.document.getElementById("txtAreaName");
        if (tpName != null) {
            document.getElementById("tpName").value = tpName.value;
            bValueSet = bValueSet || (tpName.value != "");
        }
        var tpX = opener.document.getElementById("txtXPos");
        if (tpX != null) {
            document.getElementById("tpX").value = tpX.value ;
            bValueSet = bValueSet || (tpX.value != "");
        }
        var tpY = opener.document.getElementById("txtYPos");
        if (tpY != null) {
            document.getElementById("tpY").value = tpY.value;
            bValueSet = bValueSet || (tpY.value != "");
        }
        var tpWidth = opener.document.getElementById("txtWidth");
        if (tpWidth != null) {
            document.getElementById("tpWidth").value = tpWidth.value;
            bValueSet = bValueSet || (tpWidth.value != "");
        }
        var tpHeight = opener.document.getElementById("txtHeight");
        if (tpHeight != null) {
            document.getElementById("tpHeight").value = tpHeight.value;
            bValueSet = bValueSet || (tpHeight.value != "");
        }
    }
    if (bValueSet) {
        redraw();
    }
}
</script>
<style type="text/css">
        html { overflow: auto; }
        body { overflow: auto; }
</style>
<%
  Send_Scripts()
  Send_HeadEnd()
%>
<body class="popup" onload="initPage();" style="background-color: #e0e0e0;">
  <div id="wrap">
    <%
      If (Logix.UserRoles.AccessGraphics = False) Then
        Send_Denied(2, "perm.graphics-access")
        GoTo done
      End If
    %>
    <div id="imgDiv" style="position:absolute; top:0; left:0; width:<%Sendb(ImgWidth) %>px; height:<%Sendb(ImgHeight) %>px; z-index:1; border:0;">
      <img src="graphic-display-img.aspx?path=<% Sendb(Server.UrlEncode(GraphicPath)) %>&amp;lang=<%Sendb(LanguageID) %>" alt="" title="" name="tpImgKeypad" />
    </div>
    <%
      MyCommon.QueryStr = "select AreaID, Name, X, Y, Width, Height from TouchAreas with (NoLock) where OnScreenAdID=" & AdID & " and Deleted=0;"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        For Each row In dt.Rows
          'send the horizontal lines
          htmlBuf.Append("<img src=""/images/blackdot.png"" style=""z-index:100;position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) & "px;"" width=""" & MyCommon.NZ(row.Item("Width"), 1) & """ height=""1"" alt="""" />")
          htmlBuf.Append("<img src=""/images/blackdot.png"" style=""z-index:100;position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) + MyCommon.NZ(row.Item("Height"), 1) & "px;LEFT:" & MyCommon.NZ(row.Item("X"), 0) & "px;"" width=""" & MyCommon.NZ(row.Item("Width"), 1) & """ height=""1"" alt="""" />")
          'send the vertical lines
          htmlBuf.Append("<img src=""/images/blackdot.png"" style=""z-index:100;position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) & "px;"" width=""1"" height=""" & MyCommon.NZ(row.Item("Height"), 1) & """ alt="""" />")
          htmlBuf.Append("<img src=""/images/blackdot.png"" style=""z-index:100;position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) + MyCommon.NZ(row.Item("Width"), 1) & "px;"" width=""1"" height=""" & MyCommon.NZ(row.Item("Height"), 1) & """ alt="""" />")
          'If (ShowDeliverables > 0) Then
          '    LabelsStr = LabelsStr & "<div onclick=""javascript:showNextDeliverable(" & MyCommon.NZ(row.Item("AreaID"), "") & ", '" & MyCommon.NZ(row.Item("Name"), "") & "');"" style=""position: absolute; top: " & top + MyCommon.NZ(row.Item("Y"), 0) & "px; left: " & MyCommon.NZ(row.Item("X"), 0) & "px; width: " & MyCommon.NZ(row.Item("Width"), 1) & "px; height: " & MyCommon.NZ(row.Item("Height"), 1) & "px;"">" & vbCrLf
          'Else
          LabelsStr = LabelsStr & "<div style=""z-index:100;position: absolute; top: " & 30 + MyCommon.NZ(row.Item("Y"), 0) & "px; left: " & MyCommon.NZ(row.Item("X"), 0) & "px; width: " & MyCommon.NZ(row.Item("Width"), 1) & "px; height: " & MyCommon.NZ(row.Item("Height"), 1) & "px;"">" & vbCrLf
          'End If
          LabelsStr = LabelsStr & "    <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""height:100%;"" summary="""">" & vbCrLf
          LabelsStr = LabelsStr & "     <tr>" & vbCrLf
          LabelsStr = LabelsStr & "      <td valign=""middle"">" & vbCrLf
          LabelsStr = LabelsStr & "       <center>" & vbCrLf
          LabelsStr = LabelsStr & "       <table border=""0"" cellpadding=""2"" cellspacing=""0"" summary="""">" & vbCrLf
          LabelsStr = LabelsStr & "        <tr>" & vbCrLf
          LabelsStr = LabelsStr & "         <td valign=""middle"" bgcolor=""white"">" & vbCrLf
          LabelsStr = LabelsStr & "          <center><span style=""color:#000000;font-face:arial;"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</span></center>" & vbCrLf
          LabelsStr = LabelsStr & "         </td>" & vbCrLf
          LabelsStr = LabelsStr & "        </tr>" & vbCrLf
          LabelsStr = LabelsStr & "       </table>" & vbCrLf
          LabelsStr = LabelsStr & "       </center>" & vbCrLf
          LabelsStr = LabelsStr & "      </td>" & vbCrLf
          LabelsStr = LabelsStr & "     </tr>" & vbCrLf
          LabelsStr = LabelsStr & "    </table>" & vbCrLf
          LabelsStr = LabelsStr & "  </div>" & vbCrLf
        Next
      End If
      htmlBuf.Append(LabelsStr)
      Sendb(htmlBuf.ToString)
    %>
    <div id="imgBox" style="border:dashed 1px #ff0000; position:absolute; display:none; z-index:100; line-height:1px;">
      <div style="height:100%; width:100%; filter:alpha(opacity=70); background-color:#e0e0e0;">
      </div>
    </div>
    <div id="touchpoints" style="position:absolute; top:<% Sendb(ControlsTop) %>px; left:0; width:100%; background-color:#e0e0e0;">
      <form action="#" id="mainform" name="mainform">
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.touchpoint", LanguageID))%>">
          <tr>
            <td colspan="2">
              <h1>
                <% Sendb(Copient.PhraseLib.Lookup("term.touchpoint", LanguageID))%>
              </h1>
            </td>
          </tr>
          <tr>
            <td>
              <label for="tpName"><% Sendb(Copient.PhraseLib.Lookup("term.areaname", LanguageID))%></label>
            </td>
            <td>
              <label for="tpX"><% Sendb(Copient.PhraseLib.Lookup("term.xpos", LanguageID))%></label>
            </td>
            <td>
              <label for="tpY"><% Sendb(Copient.PhraseLib.Lookup("term.ypos", LanguageID))%></label>
            </td>
            <td>
              <label for="tpWidth"><% Sendb(Copient.PhraseLib.Lookup("term.width", LanguageID))%></label>
            </td>
            <td>
              <label for="tpHeight"><% Sendb(Copient.PhraseLib.Lookup("term.height", LanguageID))%></label>
            </td>
            <td>
            </td>
          </tr>
          <tr>
            <td>
              <input type="text" id="tpName" name="tpName" value="" tabindex="1" maxlength="100" />
            </td>
            <td>
              <input type="text" id="tpX" name="tpX" value="" size="5" tabindex="2" maxlength="5" onchange="redraw();" />
            </td>
            <td>
              <input type="text" id="tpY" name="tpY" value="" size="5" tabindex="3" maxlength="5" onchange="redraw();" />
            </td>
            <td>
              <input type="text" id="tpWidth" name="tpWidth" value="" size="5" tabindex="4" maxlength="5" onchange="redraw();" />
            </td>
            <td>
              <input type="text" id="tpHeight" name="tpHeight" value="" size="5" tabindex="5" maxlength="5" onchange="redraw();" />
            </td>
            <td>
              <input type="button" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID))%>" onclick="addTouchpoint();" tabindex="6" />
            </td>
          </tr>
        </table>
      </form>
    </div>
<%
done:
  Send_BodyEnd("mainform", "tpName")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
