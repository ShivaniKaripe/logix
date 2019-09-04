<%@ Control Language="C#" AutoEventWireup="true" Debug="true" CodeFile="ProductAttributeFilter.ascx.cs"
  Inherits="logix_UserControls_ProductAttributeFilter" EnableViewState="true" %>
<script type="text/javascript" src="/javascript/jquery.min.js"></script>
<script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
<script type="text/javascript" src="/javascript/notify.min.js"></script>
<link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" />
<link type="text/css" href="/css/chosen.css" rel="stylesheet" />
<script type="text/javascript">
  var GlobalObject = {
    LookingFor: '<%=PhraseLib.Lookup("term.lookingFor", LanguageID,"Looking For")%>',
    NoResultsMatch: '<%=PhraseLib.Lookup("error.noResultsMatch", LanguageID,"No Results Match")%>',
    SelectAValue: '<%=PhraseLib.Lookup("term.selectvalue", LanguageID,"Select a value....")%>',
    KeepTyping: '<%=PhraseLib.Lookup("term.keepTyping", LanguageID,"Keep typing...")%>'
  };
</script>
<script type="text/javascript" src="/javascript/chosen.jquery.js"></script>
<script type="text/javascript" src="/javascript/ajax-chosen.js"></script>
<script type="text/javascript" src="/javascript/spin.min.js"></script>
<style type="text/css">
  .hideShaded
  {
    color: #f0f0f0;
  }
  .hide
  {
    color: white;
  }
    .panel_header {
    line-height: 16px;
    margin-left: 5px;
  }
  
    .ui-dialog-titlebar {
    background-color: #808080;
    background-image: none;
    color: #FFFFFF;
  }
  
    .ui-icon-closethick {
    background-image: url('/images/close_red.png') !important;
    background-repeat: no-repeat;
    background-position: center center;
  }
  
    .ui-dialog.ui-widget-content {
    background: #F5F5F5;
  }
  
    .ui-dialog .ui-dialog-buttonpane {
    border: 0;
    background: #F5F5F5;
  }
  .setIDContainer
  {
    width: 3%;
    float: left;
    display: block;
    margin-top: 12px;
    margin-left: 5px;
  }
  .excludeLabelContainer
  {
    width: 10%;
    float: left;
    display: block;
    margin-top: 12px;
  }
  .excludesetContainer
  {
    width: 90%;
    float: left;
    display: block;
    margin-top: 2px;
    margin-bottom: 4px;
  }
  .setContainer
  {
    width: 96%;
    float: left;
    display: block;
    margin-top: 2px;
    margin-bottom: 2px;
  }
  .divORContainer
  {
    width: 3%;
  }
  .setIDLabel
  {
    width: 5%;
    display: block;
    float: left;
    padding-top: 5px;
    padding-left: 10px;
    padding-right: 10px;
  }
  .setContainer
  {
    width: 95%;
    display: block;
  }
    .filterButton1
  {
    color: black;
    background-color: #E5E4E2;
    overflow: auto;
    border-style: solid;
    border-width: thin;
    border-color: Silver;
    height: 18px;
    width: auto;
    display: table;
    padding: 1px;
    line-height: 18px;
    margin-right: 10px;
    float: left;
    margin-top: 6px;
  }
  
    .filterButton1:hover {
       text-decoration: none;
    }

    .filterButton1:visited {
    color: black;
    text-decoration: none;
  }
  .filterButton
  {
    color: black;
    background-color: #E5E4E2;
    overflow: auto;
    border-style: solid;
    border-width: thin;
    border-color: Silver;
    height: 18px;
    width: auto;
    display: table;
    padding: 1px;
    line-height: 18px;
    margin-right: 10px;
    float: left;
    margin-top: 6px;
  }
  
        .filterButton:hover {
            background-color: #95b0e5;
            text-decoration: none;
        }

        .filterButton:visited {
    color: black;
    text-decoration: none;
  }
  
    table.fixHeader {
    table-layout: fixed;
  }
  
        table.fixHeader th {
    overflow: hidden;
    width: 120px;
  }
  
        table.fixHeader td {
    overflow: hidden;
    width: 120px;
  }
  
        table.fixHeader th:first-child {
    overflow: hidden;
    width: 40px;
  }
  
        table.fixHeader td:first-child {
    overflow: hidden;
    width: 40px;
  }
  
 .gvGroupLvl{ overflow: hidden; } 
 .gvGroupLvl td { min-width: 100px; width:100px}
 .gvGroupLvl th { min-width: 104px; width:104px}
 .fixHeaderLevel th { min-width: 104px; width:104px}
 .gvGroupLvl td:first-child
 {
   overflow-x: hidden;
   min-width: 23px;
   width:23px;
 }
 
 .gvGroupLvl td:nth-child(2) 
 {
   overflow-x: hidden;
   min-width: 23px;
   width:23px;
 }
 
  .gvGroupLvl th:first-child 
 {
   overflow-x: hidden;
   min-width: 26px;
   width:23px;
 }
 .gvGroupLvl th:nth-child(2) 
 {
   overflow-x: hidden;
   min-width: 26px;
   width:23px;
 }
.fixHeaderLevel th:nth-child(2) 
 {
   overflow-x: hidden;
   min-width: 26px;
   width:23px;
 }
 
 .fixHeaderLevel th:first-child
 {
   overflow-x: hidden;
   min-width: 26px;
   width:23px;
 }
 
 .gvGroupLvl td:last-child 
 {
   overflow-x: hidden;
   min-width: auto;
   width:auto;
 }
.gvGroupLvl table td:first-child
 {
   overflow-x: hidden;
   min-width: 26px;
   width:26px;
 }
 
 .gvGroupLvl table td:nth-child(2) 
 {
   overflow-x: hidden;
   min-width: 100px;
   width:100px;
 }
 .gvGroupLvl table td:last-child 
 {
   overflow-x: hidden;
   min-width: auto;
   width:auto;
 }
</style>
<script type="text/javascript" language="javascript">
    var pageIndex = 1;
    var scrollheight = 0;
    var LoadAttributesOnScroll = true;

    function GetDropDownVal() {
        document.getElementById('hdndropdownID').value = $('#ddlUsrCtrl').val();
        var s = "";
        $('#ddlUsrCtrl option:selected').each(function (index) {

            if (index == 0)
                s = $(this).text();
            else
                s = s + "|" + $(this).text();

        });
        document.getElementById('hdndropdownVal').value = s;
    }
    function ShoworHideDivs(fromContinuewithAttributes) {
        var selectedNodes = '';
        if (fromContinuewithAttributes) {
            if (selectednodenamesList == '') {
                $('#hdnSelctedNodeIDs').val("")
                alert('<%=Copient.PhraseLib.Lookup("error.select_nodes", LanguageID) %>');
                return;
            } else {
                $("#hdnPABStage").val("2");
                $("#PABScreen2").show()
                $("#PABScreen1").hide()
                DisplaySelectedHierarchy();
                $("#DeailedProducts").show();
                return;
            }
        }

        if ($("#hdnPABStage").val() == 2) {
            $("#PABScreen2").show()
            $("#PABScreen1").hide()

            if ($('#hdnSelctedNodes').val() != '' && $('#hdndivphselectedtree').val() != '') {
                $("#divphselectedtree").html(unescape($('#hdndivphselectedtree').val()));
            }

        } else {
            //switching from stage 2 to 1-->CHeck the selcted nodes by default
            var selectedNodes = $('#hdnSelctedNodes').val();
            var selectedNodeIDs = $('#hdnSelctedNodeIDs').val();
            $("#PABScreen1").show()
            $("#PABScreen2").hide()

            if (selectedNodes != '' && selectedNodeIDs != '' && selectedNodes != null && selectedNodeIDs != null) {

                var nodename = selectedNodes.split(',');
                var nodeid = selectedNodeIDs.split(',');
                for (var i = 0; i < nodename.length; i++) {
                    if (nodename[i] != '') {
                        var elem = document.getElementsByName("'chk" + nodename[i] + "'");
                        if (elem != null) {
                            elem.value = "N" + nodeid[i].toString();
                            elem.checked = true;
                            updateIdList(elem, nodeid[i], nodename[i]);
                        }
                    }
                }
            }
        }
    }
    function updateControls() {
        $(".chosen-select").chosen();
        if ($("#ddlAttributeValue_chosen").length > 0) {
            if ($("#ddlAttributeValue_chosen").width() == 0)
                $('#ddlAttributeValue_chosen').css({ "width": '250px' })
            if ($('#ddlAttributeValue_chosen ul.chosen-choices').css("max-height") != '100px' && $('#ddlAttributeValue_chosen ul.chosen-choices').css("overflow") != 'auto') {
                $('#ddlAttributeValue_chosen ul.chosen-choices').css({
                    'max-height': '60px',
                    overflow: 'auto',
                    'overflow-x': 'hidden'
                });
            }
        }
        if ($("#ddlAttributeValue_chosen input.default[type=text]").length > 0 && $('#ddlAttributeValue_chosen input.default[type=text]').width() < 50)
            $('#ddlAttributeValue_chosen input.default[type=text]').css({ "width": '121px' });
        if (document.getElementById("dvGridLocation") != null) {
            var gridPosition = document.getElementById("dvGridLocation").value;
            if (gridPosition != "") {
                document.getElementById("dvGrid").scrollTop = gridPosition;
                document.getElementById("dvGridLocation").value = "";
            }
        }
        var attributetype = $('#<%= ddlAttributeType.ClientID %>').val();
        var attribValues = [];
        $('#ddlAttributeValue :selected').each(function (i, selectedElement) {
            attribValues.push($(selectedElement).val());
        });
        attribValues = attribValues.toString();
        var keyValue = $('#hdnKeyValue').val();
        var params = '"attributetype" : ' + attributetype + ', "pageindex" : 0, "nodeIdList" : "' + $('#hdnSelctedNodeIDs').val() + '", "excludeattr" : "' + attribValues + '", "keyValue" : "' + $('#hdnKeyValue').val() + '"';
        $("#ddlAttributeValue").ajaxChosen({
            type: 'POST',
            url: "<%=Request.Url.AbsolutePath%>/GetAttributes",
            additionalparams: params,
            contentType: "application/json; charset=utf-8",
            dataType: 'json',

            failure: function (response) {
                //alert(response.d + " failure");
            },
            error: function (response) {
                //alert(response.d + " error");
            }
        },
          function (data) {
              var terms = {};
              if(data.d !=null && data.d != undefined)
              {
                  if (data.d.length < 100) {
                      LoadAttributesOnScroll = false;
                      $('#loadmoreattribvalues').hide();
                  }
                  else {
                      LoadAttributesOnScroll = true;
                      $('#loadmoreattribvalues').show();
                  }
                  pageIndex = 1;
                  
                  $.each(data, function (i, val) {
                      terms[i] = val;
                  });

              }
              return terms;
          });

        //$('div.chosen-drop').append('<a id="loadmoreattribvalues" href="javascript:LoadMoreAttribValues();">Load More Records...</a>');
        if ($('#ddlAttributeValue option').length < 100) {
            $('#loadmoreattribvalues').hide();
        }


        $('ul.chosen-results').bind('scroll', function () {
            if ($(this).scrollTop() + $(this).innerHeight() >= this.scrollHeight && LoadAttributesOnScroll == true) {
                LoadMoreAttribValues();
            }
        })
        $("#DeailedProducts").hide();
    } // end updateControls function

    function LoadMoreAttribValues() {
        var attributetype = $('#<%= ddlAttributeType.ClientID %>').val();
        var searchtext = $('.chosen-container').find(".search-field > input, .chosen-search > input").val();
        scrollheight = $('ul.chosen-results').scrollTop();
        var attribValues = [];
        var keyValue = $('#hdnKeyValue').val();
        $('#ddlAttributeValue :selected').each(function (i, selectedElement) {
            attribValues.push($(selectedElement).val());
        });
        //AL-9258 Removed parameter nodeIdList: $('#hdnSelctedNodeIDs').val()
        attribValues = attribValues.toString();
        $.ajax({
            type: "POST",
            url: "<%=Request.Url.AbsolutePath%>/GetAttributes",
            data: JSON.stringify({ term: searchtext, attributetype: attributetype, pageindex: pageIndex, excludeattr: attribValues, keyValue: $('#hdnKeyValue').val() }),
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        })
        .done(function (response) {
            OnSuccess(response);
        })
        .fail(function (response) {
                
        });
    }

    function OnSuccess(response) {
        pageIndex = pageIndex + 1;
        var lstattribval = response.d;
        if (lstattribval.length < 100) {
            $('#loadmoreattribvalues').hide();
            LoadAttributesOnScroll = false;
        }
        else {
            $('#loadmoreattribvalues').show();
        }
        $.each(lstattribval, function (i, re) {
            $('#ddlAttributeValue').append('<option value="' + re.value + '">' + re.text + '</option>');
        });
        $('.chosen-select').trigger("chosen:updated");
        $('ul.chosen-results').scrollTop(scrollheight);
    }

    function getClear() {
        $('.chosen-results').remove();
        // $('.chosen-choices').remove('li');
        $('.search-choice').remove();
        $('.search-field > input').val("Select a value....");

        return true;
    }

    var spinner ;
                    var opts = {
                        lines: 13, // The number of lines to draw
                        length: 20, // The length of each line
                        width: 10, // The line thickness
                        radius: 25, // The radius of the inner circle
                        corners: 1, // Corner roundness (0..1)
                        rotate: 0, // The rotation offset
                        direction: 1, // 1: clockwise, -1: counterclockwise
                        color: '#000', // #rgb or #rrggbb or array of colors
                        speed: 1, // Rounds per second
                        trail: 60, // Afterglow percentage
                        shadow: false, // Whether to render a shadow
                        hwaccel: false, // Whether to use hardware acceleration
                        className: 'spinner', // The CSS class to assign to the spinner
                        zIndex: 2e9 // The z-index (defaults to 2000000000)
                    };
    //This method will freeze the header by scrolling the header table along with the content table and
    //Loads the data by making ajax calls..
    var lastpos;
    function gridScroll(obj) {        
        var div = document.getElementById('dvGrid'); 
        var div2 = document.getElementById('dvGridHeader'); 
        //****** Scrolling HeaderDiv along with DataDiv ******
        div2.scrollLeft = div.scrollLeft;
        var hasVerticalScrollbar = obj.scrollHeight > obj.clientHeight;
        if (!hasVerticalScrollbar)  
            return;
        if(lastpos!=obj.scrollLeft){//Do not Process on Horizontal Scroll
            lastpos=obj.scrollLeft
            return;
        }
        if ($(obj).scrollTop() + $(obj).innerHeight() >= $(obj)[0].scrollHeight) {
            if($('#GridDataLoading').val()=="False" || $('#GridDataLoading').val()=="")
            {
                $('#GridDataLoading').val("True");
                 //If normal grid perform a button click to load products
                 if(!Groupgrid){
                    if ($('#lbNeedReload').text() != "True") {
                        return;
                    }
                    if ($('div.spinner').length == 0) {
                        $('#gridHeader tbody tr:first-child a').each(function () { $(this).replaceWith($(this).html()); });
                        $('#<%= gvData.ClientID %>').closest('div').css('opacity', 0.5);
                    }
                    $('#btnTemp').click();
                }
                //If group grid, make ajax calls to load products..
                else{
                  var target = document.getElementById('dvGrid');
                  spinner = new Spinner(opts).spin(target);
                  populateLevelGrid(); // should not process on horizontal scroll bar
                }
            }
          }
        }
    function SetDivPosition() {
        gridPosition = document.getElementById("dvGrid").scrollTop;
        document.getElementById("dvGridLocation").value = gridPosition;
    }
    function LoadEjectedValues() {

        if (document.getElementById("hdnEjectedValues").value != "") {
            var res = document.getElementById("hdnEjectedValues").value.split(",");
            $elemtns = $('#ddlAttributeValue').find('option');
            for (i = 0; i < $elemtns.length; i++) {
                $ele = $.grep($elemtns, function (n, i) { return (n.innerHTML == res[0]) });
                if ($ele[0] != undefined)
                    $ele[0].setAttribute("selected", "");
                res.splice(0, 1);
            }
        }
    }
    function GetSelection() {
        document.getElementById('hdnSelectedIDs').value = $('#ddlAttributeValue').val();
        var s = "";
        $('#ddlAttributeValue option:selected').each(function (index) {
            if (index == 0)
                s = $(this).text();
            else
                s = s + "," + $(this).text();

        });
        document.getElementById('hdnSelectedValues').value = s;
        if (s == "") {
            return false;
        }
        return true;
    }
    //When the master check box is clicked, it will select or deselect all the child check boxes 
    //and set the count accordingly..
    function ToggleCheckbox(isCheckAllChecked) {
        var Count = 0;
        var ActualProductCount = <%=TotalCount%>;
        if(!Groupgrid){
            $('#<%=gvData.ClientID %>').find("input:checkbox").each(function () {
                     this.checked = isCheckAllChecked;
             });
         }
         else
         {
             $('#<%=gvGroupLevel.ClientID %>').find("input:checkbox").each(function () {
                     this.checked = isCheckAllChecked;
             });
         }

        if (isCheckAllChecked)
        {
            SetCountText(0);
            $('#<%=hdnIsMasterCBChecked.ClientID %>').val(1);
        }
        else 
        {
            SetCountText(ActualProductCount);
            $('#<%=hdnIsMasterCBChecked.ClientID %>').val(0);
        }
     }
     //If any particualr product is selected or deslected using checkbox this method will update the count
    function ToggleMasterCheckbox(thiscntrl) {
         var IsChecked = false;
         var prodCount = 0;
         if (thiscntrl.type && thiscntrl.type === 'checkbox') {
            updateCount(thiscntrl);
          }
          //Update the Master check box, based on all the child checkboxes..
          $('#<%=gvData.ClientID %>').find("input:checkbox").each(function () {
             if ($(this).attr("id") != "checkall") {
                IsChecked = this.checked;
                if (IsChecked == false) {
                    return false;
                }
            }
        });
        $('#checkall').attr('checked', IsChecked);
    }

   function updateCount(control){
          var ActualProductCount = parseInt($('#<%=hdnInclProducts.ClientID %>').val());
          if (ActualProductCount == 'NaN')
            ActualProductCount = 0;
          if (control.id != "checkall") {
            if (control.checked) {
                ActualProductCount = ActualProductCount - 1;
            }
            else {
                ActualProductCount = ActualProductCount + 1;
            }
          }
      $('#<%=lbTotalProducts.ClientID %>').text(ActualProductCount);
      $('#<%=hdnInclProducts.ClientID %>').val(ActualProductCount);
      ChangeColor();
   }


function ChangeColor() {

    if ($('#<%=lbTotalProducts.ClientID %>').text() == "0") {
            $('#<%=spanTotalProd.ClientID%>').children().each(function () {
                $(this).css('color', 'red');
                $('#<%=hlShowDetails.ClientID %>').hide();
            });
            if($("#Groupcheckall").length)
               $("#Groupcheckall")[0].checked=true;
        }
        else {
            if ($('#<%=spanTotalProd.ClientID%>') != undefined)
                $('#<%=spanTotalProd.ClientID%>').children().each(function () {
                    $(this).css('color', 'black');
                });
        }
		//DOnt display warning label, if product count is not zero.
        var lbWarning = document.getElementById("lblWarningText");
        if($('#<%=lbTotalProducts.ClientID %>').text()== "0")
            lbWarning.style.display = "inline-block";
        else 
            lbWarning.style.display = "none";
    
    }
    function highLightInEditMode()
    {
            //var lbl = document.getElementById("lblDebugInfo");
            var hdnEditedElement = document.getElementById("hdnEditedElementID");
            var editedElement = document.getElementById(hdnEditedElement.value);

            if(editedElement)
            {
                editedElement.onmouseover = null;
                editedElement.onmouseout = null;
                highlight(editedElement, true);
            }
    }
    function highlight(obj, flag) {
        if ($('#disableHierarchyTree').val() == undefined || ($('#disableHierarchyTree').val() != undefined && $('#disableHierarchyTree').val() != "true")) {
            var id = document.getElementById(obj.id).id;
            if (flag)
                $('#' + id + '> a > div').each(function () {
                    this.style.backgroundColor = '#95B0E5';

                });
            else
                $('#' + id + '> a > div').each(function () {
                    this.style.backgroundColor = '#E5E4E2';

                });
            }
    }
    function editAttributeSetDialog(message) {
        $('<div align="left" style="white-space: normal; word-wrap: break-word;-moz-hyphens:auto;-webkit-hyphens:auto;-o-hyphens:auto;hyphens:auto;">' + message + '</div>').dialog({
            height: 200,
            width: 400,
            title: '<%=PhraseLib.Lookup("term.appliedAttrSet", LanguageID)%>',
            modal: true,
            buttons: {
                '<%=PhraseLib.Lookup("term.editset", LanguageID)%>': function () {
                    $(this).dialog("close");
                    $('#editTemp').click();
                    $(this).dialog("destroy");
                }
            }
        });
    }
    function RaiseSeverEvent() {
        if ($('#hdnSelctedNodeIDs').val().trim() != ''){
		      //Dont allow user to click on detailed list as we dont have the hierarchy information in code behind by this time.
          $('#hlShowDetails').replaceWith($('#hlShowDetails').text());
          //Update count text to default during the postback
          if ($('#<%=spanTotalProd.ClientID%>') != undefined)
              $('#<%=spanTotalProd.ClientID%>').children().each(function () {
                  $(this).css('color', 'black');
              });
          $('#<%=lbTotalProducts.ClientID %>').text('');
          var imgLoader = document.getElementById("imgLoader");
            imgLoader.style.display = "inline-block";
          var lbWarning = document.getElementById("lblWarningText");
              lbWarning.style.display = "none";

          $('#btnDummyForcatchingEvent').click();
          }
    }

    function DisplaySelectedHierarchy() {
        var selectedNodes = '';
        var selectedNodeIDs = '';

        if (typeof selectednodenamesList !== "undefined" && selectedNodes == '') {
            selectedNodes = selectednodenamesList;
        }
        if (selectedNodes == '') {
            selectedNodes = $('#hdnSelctedNodes').val();
        }

        if (typeof nodeidlist !== "undefined" && selectedNodeIDs == '') {
            selectedNodeIDs = nodeidlist;
        }

        if (selectedNodeIDs == '') {
            selectedNodeIDs = $('#hdnSelctedNodeIDs').val();
        }

        if (selectedNodes != '') {
            $("#hdnSelctedNodes").val(selectedNodes);
            $("#hdnSelctedNodeIDs").val(selectedNodeIDs);
            var html = '';
            var tooltip = '<%=PhraseLib.Lookup("term.edithierarchyselection", LanguageID)%>';

            //Code to generate the selcted nodes div in-order to display along with attribute filters
            var id = selectedNodes.split(',');
            var newLeft = ((levelSel + 1) * 17);

            for (var i = 0; i < id.length; i++) {
                if (id[i] != '') {
                    html = html + "<img src=\"/images/clear.png \" style=\"height:1px;width:" + newLeft.toString() + "px; \" />";
                    html = html + "<span onclick=\"return WarnUser();\" class=\"hrow\" title=\"" + tooltip + "\" onmouseover=\"highlightdiv(true)\" onmouseout=\"highlightdiv(false)\" ><img  border=\"0\" src=\"/images/folder.png\" \/><span style=\"left: 5px;\">" + "&nbsp;" + id[i] + "</span></span><br />";
                }
            }

            $("#divphselectedtree").html("<br />" + $("#leftpane").html() + html);

            $("#divphselectedtree  span").each(function (index, elem) {
                $(this).css({ "color": "", "backgroundColor": "" });
            });

            $("#divphselectedtree  span[id^='indent']").each(function (index, elem) {
                $(this).attr("title", tooltip);
                $(this).attr("onclick", "WarnUser();");
                $(this).attr("onmouseover", "highlightdiv(true);");
                $(this).attr("onmouseout", "highlightdiv(false);");
            });

            $("#hdndivphselectedtree").val(escape($("#divphselectedtree").html()));
        }
    }

    function WarnUser() {
        if ($('#disableHierarchyTree').val() != "true") {
            if (confirm('<%=PhraseLib.Lookup("Warning.hierarchyEdit", LanguageID)%>')) {
                $("#hdnPABStage").val("1");
                $("#hdnIsEditHierarchyInProgress").val("1");
                $("#btnReloadHierarchytree").click();
                return true;
            }
        }
        else {
            return false;
        }
    }

    function highlightdiv(AllowHighlight) {
        if (($('#disableHierarchyTree').val() == undefined) || ($('#disableHierarchyTree').val() != undefined && $('#disableHierarchyTree').val() != "true")) {
            if (AllowHighlight) {
                $("#divphselectedtree").css({ "backgroundColor": "#95B0E5" });
            } else {
                $("#divphselectedtree").css({ "backgroundColor": "" });
            }
        }
    }
    function EnableAttributeSetDDL()
    {
        var ddlAttributeSet = document.getElementById("ddlAttributeSet");
        ddlAttributeSet.disabled = false;

        EnableDisableAVDropDowns(ddlAttributeSet);
    }
    function DisableAttributeSetDDL()
    {
        var radExcludedSet = document.getElementById("radExcludedSet");
        //alert(radExcludedSet);
        if(radExcludedSet.checked == false)
            document.getElementById("ddlAttributeSet").disabled = true;
    }
    function EnableDisableAVDropDowns(ddlAttributeSet)
    {
        if(ddlAttributeSet && ddlAttributeSet.selectedIndex == 0)
        {
            $('#<%= ddlAttributeType.ClientID %>').attr("disabled", true);
            $('#<%= ddlAttributeValue.ClientID %>').attr("disabled", true);
            $('#<%= btnAddFilter.ClientID %>').attr("disabled", true);      
            $("#ddlAttributeValue_chosen").attr("disabled", true);       
            //return false; //Prevent postback
        }
        else
        {
            $('#<%= ddlAttributeType.ClientID %>').attr("disabled", false);
            $('#<%= ddlAttributeValue.ClientID %>').attr("disabled", false);
            $('#<%= btnAddFilter.ClientID %>').attr("disabled", false);
            $("#ddlAttributeValue_chosen").attr("disabled", false);       
            //return true;
        }        
    }
    function WarnUserForExcludeSetInvalidation()
    {
        var hdnExcludeSetExists = document.getElementById("hdnExcludeSetExists");

        if(hdnExcludeSetExists.value == "True")
        {
            var warnMsg = "<%=Copient.PhraseLib.Lookup(8327, LanguageID)%>";
            warnMsg = warnMsg.replace("&#39;", "'");
            var result = confirm(warnMsg);
            if(result == true)
                return true;
            else
                return false;
        }
    }
</script>
<asp:HiddenField ID="hdnIsEditHierarchyInProgress" runat="server" Value="0" ClientIDMode="Static" />
<asp:HiddenField ID="hdnProductGroupID" runat="server" />
<asp:HiddenField ID="hdnTotalRecords" runat="server" />
<asp:HiddenField ID="hdnSelectedAttributeValues" runat="server" ClientIDMode="Static" />
<asp:HiddenField ID="hdnKeyValue" runat="server" ClientIDMode="Static" />
<asp:HiddenField ID="hdnEjectedValues" runat="server" ClientIDMode="Static" />
<asp:Button runat="server" ID="btnDummyForcatchingEvent" OnClick="btnDummyForcatchingEvent_Click"
  ClientIDMode="Static" Height="0px" Width="0px" BorderWidth="0px" Style="padding: 0px" />
<input type="hidden" id="hdndivphselectedtree" runat="server" clientidmode="Static" />
<input type="hidden" id="hdnLocateHierarchyURL" runat="server" clientidmode="Static" />
<input type="hidden" id="hdndivphtree" runat="server" clientidmode="Static" />
<input type="hidden" id="hdnPABStage" runat="server" clientidmode="Static" />
<input type="hidden" id="hdnSelctedNodes" runat="server" clientidmode="Static" />
<input type="hidden" id="hdnSelctedNodeIDs" runat="server" clientidmode="Static" />
<asp:HiddenField ID="disableHierarchyTree" runat="server" ClientIDMode="Static" />
<asp:HiddenField ID="EditBoxMessageString" runat="server" ClientIDMode="Static" />
<asp:HiddenField ID="hdnEditedElementID" runat="server" ClientIDMode="Static" Value="" />
<asp:HiddenField ID="hdnExcludeSetExists" runat="server" ClientIDMode="Static" Value="" />
<asp:HiddenField ID="hdnExludedProductsCount" runat="server" ClientIDMode="Static" Value="" />
<asp:HiddenField ID="hdnPABAVPairsJson" runat="server" ClientIDMode="Static" Value="" />
<asp:Button ID="btnReloadHierarchytree" runat="server" Text="btnReloadHierarchytree"
  ClientIDMode="Static" OnClick="btnReloadHierarchytree_Click" Style="display: none;" />
<asp:Button ID="btnLocateHierarchyTree" runat="server" Text="btnReloadHierarchytree"
  ClientIDMode="Static" OnClick="btnLocateHierarchyTree_Click" Style="display: none;" />
<asp:Panel runat="server" ID="panelPAB" ClientIDMode="Static">
    <div id="PABScreen1" style="float: left; position: relative; width: 100%; height: auto; display: none;">
        <div id="producthierarchy" clientidmode="Static" style="float: left; position: relative; width: 100%; height: 300px;"
            runat="server">
    </div>
    <div style="float: left; position: relative;" id="throbber">
      <span id="pcount" style="display: none">
                <label id="warning" style="display: none"><%=Copient.PhraseLib.Lookup("term.warning", LanguageID) %> : </label>
                <label id="contains" style="display: inline-block"><%=Copient.PhraseLib.Lookup("term.contains",LanguageID) %></label>
        <label id="lblProductsCount" style=""></label>
                <img id="Img1" src="../../images/loader.gif" style="display: none" height="10px" width="40px" />
           <label id="products" style="display: inline-block"><%=Copient.PhraseLib.Lookup("term.products", LanguageID).ToLower()%></label>
      </span>
    </div>
  </div>
    <div id="PABScreen2" style="float: left; position: relative; width: 100%; height: auto; display: none">
        <asp:Label ID="lblDebug" runat="server" ClientIDMode="Static"></asp:Label>
        <div id="attributes" class="greybox" style="width: 100%; height: auto; margin-top: 10px; padding-bottom: 15px;" runat="server" clientidmode="Static">
      <div class="greyboxwrap panel_header">
        <h3 style="padding-top: 0;">
                    <span style="color: White;"><%=Copient.PhraseLib.Lookup("term.create_attr_set",LanguageID) %></span>
        </h3>
      </div>
      <div>
        <div style="padding-top: 15px; padding-left: 5px;">
                    <asp:RadioButton ID="radIncludedSet" runat="server" ClientIDMode="Static" GroupName="RadGroupIncludeExcludeSet" Checked="true" 
                        onclick='DisableAttributeSetDDL();EnableDisableAVDropDowns();' AutoPostBack="true" OnCheckedChanged="radIncludedSet_CheckedChanged"/>
          <br />
                    <asp:RadioButton ID="radExcludedSet" runat="server" ClientIDMode="Static" GroupName="RadGroupIncludeExcludeSet" Enabled="false" onclick='EnableAttributeSetDDL();'/>
          <asp:DropDownList ID="ddlAttributeSet" runat="server" ClientIDMode="Static" onchange="EnableDisableAVDropDowns(this);"
            OnSelectedIndexChanged="ddlAttributeSet_SelectedIndexChanged" AutoPostBack="true">
          </asp:DropDownList>
        </div>
        <div style="padding-top: 15px; padding-left: 5px;">
          <asp:Label ID="lblAttributeType" runat="server" Text="Attribute Type: "></asp:Label>
          <asp:DropDownList ID="ddlAttributeType" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAttributeType_SelectedIndexChanged"
            Width="100px" Height="25px" onchange="javascript:return getClear()" Enabled="false">
          </asp:DropDownList>
          <asp:Label ID="lblAttributeValue" runat="server" Text="Value: "></asp:Label>
          <asp:DropDownList ID="ddlAttributeValue" ClientIDMode="Static" runat="server" data-placeholder="<%# PhraseLib.Lookup(7431, LanguageID) %>"
            CssClass="chosen-select" Style="width: 250px; height: 25px;" Enabled="false">
          </asp:DropDownList>
                    <asp:Button ID="btnAddFilter" runat="server" Text="Add attribute to set" OnClientClick="return GetSelection();" Style="float: right; margin-right: 8px"
                        OnClick="btnAddFilter_Click" Enabled="false" />
          <label id="lblnodeid">
          </label>
        </div>
        <div style="display: block; width: 100%; height: auto; margin-bottom: 5px; padding: 5px;">
          <div style="display: block; width: 14%; float: left;">
            <h3 style="margin-top: 7px;">
              <asp:Label ID="lblFilter" runat="server" Text='<%# Copient.PhraseLib.Lookup("pab.setcontains", LanguageID) %>'></asp:Label>
            </h3>
          </div>
          <div id="forscroll" runat="server" style="display: block; width: 85%; float: left;"
            clientidmode="Static">
            <asp:Repeater ID="repFilter" runat="server" OnItemCommand="repFilter_ItemCommand">
              <ItemTemplate>
                <div id="tempfilter" align="left" title='<%#DataBinder.Eval(Container.DataItem, "ToolTip")%>'
                                    style="overflow: auto; border-style: solid; margin-top: 3px; border-width: thin; border-color: Silver; background-color: #E5E4E2; height: 18px; line-height: 20px; width: auto; display: table; margin-right: 8px;">
                  <b>
                    <%# (DataBinder.Eval(Container.DataItem, "AttributeTitle").ToString())%>:</b>
                  <span style="padding: 5px; padding-left: 7px;">
                    <%# DataBinder.Eval(Container.DataItem, "AttributeValue").ToString().Trim()%></span>
                  <asp:ImageButton ID="ImageButton1" align="right" CommandArgument='<%#DataBinder.Eval(Container.DataItem, "AttributeTypeID")%> '
                    ToolTip='<%#PhraseLib.Lookup(138, LanguageID)%>' ImageUrl="~/images/up_arrow.png"
                    Enabled='<%# ButtonStatus%>' CommandName="FilterUpdate" runat="server" title='Delete'
                    Text='<%# (DataBinder.Eval(Container.DataItem, "AttributeValue").ToString())%>'
                    Height="9pt" Style="margin-top: 3px;" />
                </div>
              </ItemTemplate>
            </asp:Repeater>
          </div>
        </div>
        <div style="margin-top: 10px; padding-top: 20px;">
                    <asp:Button ID="btnApplyFilter" runat="server" Text="" Style="margin-left: 5px;" OnClick="btnApplyFilter_Click" Enabled="false" OnClientClick="return WarnUserForExcludeSetInvalidation();"/>
          <asp:Button ID="btnClear" runat="server" Text="" Style="float: right; margin-right: 8px"
            OnClick="btnClear_Click" Enabled="false" />
        </div>
      </div>
    </div>
    <br />
        <div class="greybox" style="float: left; position: relative; width: 100%; height: auto; padding-bottom: 15px;">
      <div class="greyboxwrap">

                <asp:LinkButton ID="hlBack" runat="server" Text="" Style="float: right; color: White; font-size: 13px; font-weight: bold; text-decoration: none; line-height: 22px; margin-right: 5px;"
                    OnClick="hlBack_Click" OnClientClick="javascript:UpdateProductChanges()"></asp:LinkButton>
                <asp:ImageButton ID="backImg" runat="server" ClientIDMode="Static"
                    OnClick="hlBack_Click" Style="float: right; cursor: pointer; margin-top: 3px; margin-right: 5px;" Height="18" Width="30" ImageUrl="../../images/ncr/arrow_left.png" />
        <h3 class="panel_header" style="padding-top: 0;">
                    <span style="color: White;"><%=Copient.PhraseLib.Lookup("term.productselection", LanguageID, "Phrase Not Found")%></span>
        </h3>
      </div>
            <div id="divphselectedtree" clientidmode="Static" runat="server" style="float: left; position: relative; width: 100%; height: 50%; overflow: auto; padding-left: 5px;">
      </div>
      <div id="attibuteFilters" runat="server" style="margin-top: 5px;">
        <h3 style="margin-left: 5px; margin-bottom: 10px;">
          <asp:Label ID="lblFilter0" runat="server" Text=""></asp:Label>
        </h3>
        <asp:Repeater ID="repWrapMultiSet" runat="server">
          <ItemTemplate>
            <div id="divIncludeSet" runat="server">
              <div class="setIDContainer">
                <asp:Label ID="lblSetID" runat="server"><%# DataBinder.Eval(Container.DataItem, "AttributeSetID") %>:</asp:Label>
              </div>
                        <div id="divIncludeSetContainer" class="setContainer" runat="server" onmouseover="javascript:highlight(this,true);" onmouseout="javascript:highlight(this,false);">
                        <asp:Repeater ID="repIncludeSet" runat="server" DataSource='<%#((System.Data.DataRowView)Container.DataItem)[1] %>' OnItemCommand="repWrapMultiSet_ItemCommand">
                  <ItemTemplate>
                                <asp:LinkButton ID="lnkBtnIncludeSet" ClientIDMode="Static" runat="server" CssClass="filterButton" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "AttributeSetID") %>'> 
                                    <div id="divIncludeEditSet" runat="server" align="left" title='<%#PhraseLib.Lookup(8163, LanguageID)%>' >
                                    <b>
                                        <%# (DataBinder.Eval(Container.DataItem, "AttributeTitle").ToString())%>
                                        :</b> <span style="padding: 5px">
                                        <%# (DataBinder.Eval(Container.DataItem, "AttributeValue").ToString())%></span>
                                    </div>
                    </asp:LinkButton>
                                <asp:Button ID="editTemp" runat="server" ClientIDMode="Static" Style="Visibility:hidden; display:none;"
                                    Text="0" />                                    
                  </ItemTemplate>
                </asp:Repeater>
              </div>
            </div>
            <br />
            <div id="divExcludeSet" runat="server" style='<%# DataBinder.Eval(Container.DataItem, "ExcludeSetStyle") %>'>
              <div class="excludeLabelContainer">
                <asp:Label ID="lblExcludeSet" runat="server"><%# DataBinder.Eval(Container.DataItem, "ExcludeString") %></asp:Label>
              </div>
                        <div id="divExcludeSetContainer" class="excludeSetContainer" runat="server" onmouseover="javascript:highlight(this,true);" onmouseout="javascript:highlight(this,false);">
                        <asp:Repeater ID="repExcludeSet" runat="server" DataSource='<%#((System.Data.DataRowView)Container.DataItem)[2] %>' OnItemCommand="repWrapMultiSet_ItemCommand">
                  <ItemTemplate>
                                <asp:LinkButton ID="lnkBtnExcludeSet" ClientIDMode="Static" runat="server" CssClass="filterButton" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "AttributeSetID") %>'> 
                                    <div id="divExcludeEditSet" runat="server" align="left" title='<%#PhraseLib.Lookup(8163, LanguageID)%>' >
                                    <b>
                                        <%# (DataBinder.Eval(Container.DataItem, "AttributeTitle").ToString())%>
                                        :</b> <span style="padding: 5px">
                                        <%# (DataBinder.Eval(Container.DataItem, "AttributeValue").ToString())%></span>
                                    </div>
                    </asp:LinkButton>
                                <asp:Button ID="editTemp" runat="server" ClientIDMode="Static" Style="Visibility:hidden; display:none;"
                                    Text="0" />                                    
                  </ItemTemplate>
                </asp:Repeater>
              </div>
            </div>
            <br />
                        <div id="divOR" class="divORContainer" style='<%# DataBinder.Eval(Container.DataItem, "JoinStyle") %>'><%# DataBinder.Eval(Container.DataItem, "JoiningString") %></div>
          </ItemTemplate>
        </asp:Repeater>
      </div>
      <div style="padding: 5px">
      </div>
      <br />
      <div class="gvGroupLeveldiv" id="DetailedProductList" runat="server" clientidmode="Static"
        style="padding-top: 10px" visible="false">
        <div id="dvGridHeader" style="overflow: hidden">
                    <table id="gridHeader" class="fixHeader" style="word-break:break-all; word-wrap:break-word">
          </table>
        </div>
        <div id="dvGrid" style="height: 250px; overflow: auto;" onscroll="javascript:gridScroll(this);">
          <asp:HiddenField ID="excludeditems" runat="server" />
          <asp:HiddenField ID="rowCount" runat="server" />
          <asp:HiddenField ID="pageNum" runat="server" />
          <AMSControls:AMSGridView ID="gvData" runat="server" PageSize="50" ShowHeader="true"
            Height="30px" GridLines="None" CellSpacing="2" AllowSorting="True" OnRowDataBound="gvData_RowDataBound"
                        DataKeyNames="ProductID" ShowHeaderWhenEmpty="true" OnSorting="gvData_Sorting" CssClass="fixHeader">
            <RowStyle CssClass="shaded" />
            <AlternatingRowStyle CssClass="" />
            <Columns>
            </Columns>
          </AMSControls:AMSGridView>
                    <asp:Button ID="btnTemp" runat="server" ClientIDMode="Static" Style="visibility: hidden; display: none;"
                        OnClientClick="javascript:SetDivPosition()" OnClick="btnTemp_Click"
            Text="0" />
          <asp:HiddenField ID="hfPageIndex" runat="server" Value="1" ClientIDMode="Static" />
          <asp:HiddenField ID="hfGroupgrid" runat="server" Value="" ClientIDMode="Static" />
          <asp:HiddenField ID="hfGroupedOn" runat="server" Value="" ClientIDMode="Static" />
          <asp:HiddenField ID="hfItemUPC" runat="server" Value="" ClientIDMode="Static" />
          <asp:HiddenField ID="hfTotalPage" runat="server" Value="0" ClientIDMode="Static" />
          <asp:HiddenField ID="hfExcludedItems" runat="server" Value="" ClientIDMode="Static" />
          <asp:HiddenField ID="hfIncludedItems" runat="server" Value="" ClientIDMode="Static" />
          <asp:HiddenField ID="hfIsGridUpdated" runat="server" Value="" ClientIDMode="Static" />
          <asp:HiddenField ID="hfLevelInfoTable" runat="server" Value="" ClientIDMode="Static" />
          <%--varma--%>
          <AMSControls:AMSGridView ID="gvGroupLevel" runat="server" PageSize="100" ShowHeader="true"
            ClientIDMode="Static" Height="30px" GridLines="None" CellSpacing="2" AllowSorting="True"
            OnRowDataBound="GroupGrid_RowDataBound" DataKeyNames="ProductID" ShowHeaderWhenEmpty="true"
            OnSorting="GroupGrid_Sorting">
            <RowStyle CssClass="shaded" />
            <AlternatingRowStyle CssClass="" />
            <Columns>
              <asp:TemplateField>
                <HeaderTemplate>
                  <input type="checkbox" id="Groupcheckall" onclick="javascript:ToggleGroupMasterCB(this);"/>
                  <th scope="col" style="overflow:hidden;">
                    <img src="/images/Information.png" alt='(i)' title='<%#Copient.PhraseLib.Lookup("term.excludedprodcheckbox",LanguageID, "No phrase")%>' onload="javascript:updateGroupMasterCB()"/>
                  </th>
                </HeaderTemplate>
                <ItemTemplate>
                  <span class="<%#Eval("RowNum")%>">
                    <input type="checkbox" id="<%#Eval("RowNum")%>" <%#ProcessExcludeDataItem(Eval("DisplayLevel"),Eval("Excluded"))%>
                      name="<%#Eval("RowNum")%>" 
                      onclick="<%# String.Format("ToggleGroupCheckbox(this,\'{0}\')", HttpUtility.HtmlEncode(Eval("DisplayLevel").ToString())) %>" />
                    <td style="width: 20px">
                      <span class='<%#Eval("RowNum")%>'>
                        <img id="imgProductsShow" alt="+" src="~/images/plus2.png" runat="server" onclick='<%# String.Format("Show_Hide_ProductsGrid(this,\"{0}\")", HttpUtility.HtmlEncode(Eval("DisplayLevel").ToString())) %>' />
                        <asp:Panel ID="pnlProducts" runat="server" Visible="true" Style="position: relative"
                          ClientIDMode="Static" CssClass='<%#Eval("RowNum")%>'>
                        </asp:Panel>
                      </span>
                    </td>
                  </span>
                </ItemTemplate>
              </asp:TemplateField>
            </Columns>
          </AMSControls:AMSGridView>
          <asp:Label ID="lbNeedReload" runat="server" ClientIDMode="Static" Text="True" Style="visibility: hidden;
            display: none;"></asp:Label>
          <asp:HiddenField ID="dvGridLocation" runat="server" ClientIDMode="Static" />
          <asp:HiddenField ID="GridDataLoading" runat="server" ClientIDMode="Static" />
        </div>
        <asp:Label ID="lblexcludedcheckboxdetail" runat="server" Text="">
        <img src="../../images/arrow.png" alt="__/|\" style="float:left;padding-left:12px;"/>
          <span style="float:left;padding-top:14px;"><%=PhraseLib.Lookup("term.excludedprodcheckbox", LanguageID)  %></span>
          <div  id="notifier" style="float:left;margin-left:30%;"></div>
          <br clear="all"/>
        </asp:Label>
      </div>
    </div>
    <div style="float: left; position: relative;">
      <br />
      <span id="spanTotalProd" runat="server">
        <label ID="lblWarningText" style="display:none"><%=Copient.PhraseLib.Lookup("term.warning", LanguageID) %> : </label>
        <asp:Label ID="lblContainsText" ClientIDMode="Static" runat="server" Text=""><%=PhraseLib.Lookup("term.contains", LanguageID) %> </asp:Label>
        <asp:Label ID="lbTotalProducts" runat="server" Text=""> </asp:Label>
        <img id="imgLoader" src="../../images/loader.gif" style="display: none" height="10px" width="40px" />
        <asp:Label ID="lblProductText" ClientIDMode="Static" runat="server" Text=""><%=(lbTotalProducts.Text == "1" ? PhraseLib.Lookup("term.product", LanguageID).ToLower() : PhraseLib.Lookup("term.products", LanguageID).ToLower())%>    </asp:Label>
      </span>
      <asp:LinkButton ID="hlShowDetails" runat="server" Text="" OnClick="hlShowDetails_Click" ClientIDMode="Static"></asp:LinkButton>
    </div>
  </div>
    </label>
    </label>
</asp:Panel>
<asp:HiddenField ID="hdnInclProducts" ClientIDMode="Static" runat="server" />
<asp:HiddenField ID="hdnConsiderExclusions" runat="server" ClientIDMode="Static" />
<asp:HiddenField ID="hdnIsMasterCBChecked" runat="server" ClientIDMode="Static" />
<asp:HiddenField ID="hdnTotalProducts" ClientIDMode="Static" runat="server" />
<asp:HiddenField ID="hdnSelectedIDs" ClientIDMode="Static" runat="server" />
<asp:HiddenField ID="hdnSelectedValues" ClientIDMode="Static" runat="server" />
<asp:HiddenField ID="hidReadOnly" ClientIDMode="Static" runat="server" />
<br class="half" />
<script type="text/javascript">
    //To eliminate the horizontal alignment issue, because of the vertical scroll bar, adding new header column.
    $(window).load(function () {
        gridHeaderMove();
        if (document.getElementById("hdnEjectedValues").value != "")
            $('#ddlAttributeValue_chosen ul.chosen-choices').focus().click();

        if ($("#hdnPABStage").val() == 2) {
            updateControls();
        }
    })
  var ajaxRequestForCount;
  var Groupgrid=false;
  var gvGroupLevelheader, gvGroupLevelRow1;
  var LevelInfoTable;
  var ExcludedItems = [];
    var IncludedItems = [];
    var btnClicked = '';
    $(document).ready(function () {
        LoadEjectedValues();        
        highLightInEditMode();     
        GetProductCount();  
        //ChangeColor(); 
      Groupgrid=$('#<%=hfGroupgrid.ClientID%>').val()=="true"?true:false;
    if (Groupgrid && gvGroupLevelheader == undefined) {
        gvGroupLevelheader = $('#<%=gvGroupLevel.ClientID%> tr').eq(0).clone(true);
        gvGroupLevelRow1 = $('#<%=gvGroupLevel.ClientID%> tr').eq(1).clone(true);
        if($('#<%=hfLevelInfoTable.ClientID %>').val()!="")
          LevelInfoTable=JSON.parse($('#<%=hfLevelInfoTable.ClientID %>').val());
        if($('#<%=hfIncludedItems.ClientID %>').val()!="")
          IncludedItems= JSON.parse($('#<%=hfIncludedItems.ClientID %>').val());
        if($('#<%=hfExcludedItems.ClientID %>').val()!="")
          ExcludedItems= JSON.parse($('#<%=hfExcludedItems.ClientID %>').val());
      }

      $("form input[type=submit]").click(function () {
          $("input[type=submit]", $(this).parents("form")).removeAttr("clicked");
          $(this).attr("clicked", "true");
      });
//     if(Groupgrid)
        //updateIndeterminateState();
    });

    function AbortAsyncCountRequest()
    {
        btnClicked = $("input[type=submit][clicked=true]").val();
       if(ajaxRequestForCount)
       {
           ajaxRequestForCount.abort();
           if (btnClicked != 'Download') {
               var imgLoader = document.getElementById("imgLoader");
               imgLoader.style.display = "inline-block";
               $('#<%=lbTotalProducts.ClientID %>').text("");
           }
           else {
               toggleDropdown1();
           }
       }
    }
    function GetProductCount()
    {
        var hdnPABAVPairsJson = document.getElementById("hdnPABAVPairsJson");
        //var lbl = document.getElementById("lblDebug");
        //lbl.innerHTML = "GetProductCount Called.";
        if(hdnPABAVPairsJson.value)
        {
            //Set the count label to empty to clear off any previous result
            $('#<%=lbTotalProducts.ClientID %>').text("");
            var imgLoader = document.getElementById("imgLoader");
            imgLoader.style.display = "inline-block";
            //lbl.innerHTML += "Before ajax call." + hdnPABCacheTableName.value;
            var dataParam = ""; 
            var hdnStage = document.getElementById("hdnPABStage");            

            if(hdnPABAVPairsJson.value == "FetchNodesProductCount")  //Condition for getting count based on nodes
                dataParam = JSON.stringify({'FetchProductCountInNodesFlag': 1, 'AVPairs': hdnPABAVPairsJson.value});
            else 
                dataParam = JSON.stringify({'FetchProductCountInNodesFlag': 0, 'AVPairs': hdnPABAVPairsJson.value});
            ajaxRequestForCount = $.ajax({
                type: "POST",
                url: "/logix/pgroup-edit.aspx/GetProductCount",
                //url: "/Connectors/AjaxProcessingFunctions.asmx/GetProductCount",
                data: dataParam,
                contentType: "application/json; charset=utf-8",
                dataType: 'json'
            })
            .done(function (response) {
                OnCountFetchSuccess(response);
            })
            .fail(function (response) {
                OnCountFetchError(response);
            });
                  
        }
        else
        {
            //The count elements lbTotalProducts and hdnInclProducts already have values
            var count = $('#<%=hdnInclProducts.ClientID %>').val();
            //lbl.innerHTML += ":Else:" + parseInt(count);
            if(parseInt(count))
            {
                //lbl.innerHTML += "Finding count.";
                var lblExcludeCount = document.getElementById("hdnExludedProductsCount");
                //var lblHierarchyProductCount = document.getElementById("lblProductsCount");
                if(parseInt(lblExcludeCount.value))
                {
                    count = parseInt(count) - parseInt(lblExcludeCount.value);
                    lblExcludeCount.value = 0;
                }
            }
            SetCountText(count);  
        }
    }
    function OnCountFetchSuccess(http)
    {
        //var lbl = document.getElementById("lblDebug");
        //lbl.innerHTML = http.d;
        var hdnPABAVPairsJson = document.getElementById("hdnPABAVPairsJson");
        //lbl.innerHTML += "Success";        
        //TODO: equivalent of what is there in server side variable HeirarchyProductCount 1. set PCount, 2. remove excluded items count 3....
        var imgLoader = document.getElementById("imgLoader");
        imgLoader.style.display = "none";

        var count = JSON.parse(http.d); 
        
        var lblExcludeCount = document.getElementById("hdnExludedProductsCount");
        if(parseInt(lblExcludeCount.value))
        {
            count = parseInt(count) - parseInt(lblExcludeCount.value);
            lblExcludeCount.value = 0;
        }

        SetCountText(count);  
        //lbl.innerHTML += ":Count:" + count.toString() + (count.toString() == "270");
        //Clear these hiddent values to prevent repeated ajax calls in case of other operations in view details screen that reloads the page
        if(hdnPABAVPairsJson.value)
        {
            hdnPABAVPairsJson.value = "";
            //lbl.innerHTML = "Emptied hidden fields.";
        }
    }
    function SetCountText(count)
    {
        var lbWarning = document.getElementById("lblWarningText");
        if(parseInt(count) <= 0)
        {
            lbWarning.style.display = "inline-block";
            count = 0;
        }
        else 
            lbWarning.style.display = "none";
        
        $('#<%=lbTotalProducts.ClientID %>').text(count);
        $('#<%=hdnInclProducts.ClientID %>').val(count);                
        ChangeColor();     
    }
    function OnCountFetchError(http)
    {
        var imgLoader = document.getElementById("imgLoader");
        imgLoader.style.display = "none";
        var count = $('#<%=hdnInclProducts.ClientID %>').val();
        SetCountText(count);
    }

    updateTemplateProperties();
    if ($("#hdnPABStage").val() == 2) {
        if (document.getElementById("dvGrid") != null) {
            var HasVerticalScrollbar = document.getElementById("dvGrid").scrollHeight > document.getElementById("dvGrid").clientHeight;
        }
    }


    function updateTemplateProperties() {
        if ($('#disableHierarchyTree').val() == "true") {
            $('#PABScreen1').fadeTo('slow', .6);
            $('#PABScreen1').append('<div id="mask" style="position: absolute;top:0;left:0;width: 100%;height:100%;z-index:2;opacity:0.4;filter: alpha(opacity = 50)"></div>');
        }
    }
    function gridHeaderMove() {
        if (document.getElementById("dvGrid") != null) {
            if ($('#<%= gvData.ClientID %> tbody tr').length > 1) {
                ToggleMasterCheckbox(this);
                var extracol = "";
                //If it has Vertical scroll bar
                if ($('#dvGrid').hasVerticalScrollBar())
                    //If it has horizotal scroll bar
                    if ($('#dvGrid').hasHorizontalScrollBar())
                        extracol = '<th style="width:17px;">&nbsp;</th></tr>';
                    else
                        extracol = '<th style="width:8px;">&nbsp;</th></tr>';
                $('#dvGridHeader table').html('<tr>' + $('#<%= gvData.ClientID %> tbody tr:first-child').html() + extracol);
                $('#<%= gvData.ClientID %> tbody tr:nth-child(1)').css('visibility', 'collapse');
                $('#dvGridHeader table tbody tr input').attr('onclick', 'javascript: ToggleCheckbox(this.checked);');
            }
        }
        if (document.getElementById("gvGroupLevel") != null) {
            if ($('#<%= gvGroupLevel.ClientID %> tbody tr').length > 1) {
                var extracol = "";
                //Remove the fixHeader class in case of group level grid, as we set the width using jquery
                $('#gridHeader').removeClass("fixHeader");
                $('#gridHeader').addClass("fixHeaderLevel");
                $('#gvGroupLevel').addClass("gvGroupLvl");

                var cols = $("#gvGroupLevel tr:first > th").length
                var HasHorizontalScrollBar=document.getElementById("dvGrid").scrollWidth > document.getElementById("dvGrid").clientWidth;
                //If it has Vertical scroll bar
                var HasVerticalScrollbar = document.getElementById("dvGrid").scrollHeight > document.getElementById("dvGrid").clientHeight;
                if (HasVerticalScrollbar){
                    //If it has horizotal scroll bar
                    if (cols>8)
                        extracol = '<th style="min-width:17px; width:17px;">&nbsp;</th></tr>';
                    else
                        extracol = '<th style="min-width:8px; width:8px;">&nbsp;</th></tr>';
                }
                $('#dvGridHeader table').html('<tr>' + $('#<%= gvGroupLevel.ClientID %> tbody tr:first-child').html() + extracol);
                $('#<%= gvGroupLevel.ClientID %> tbody tr:nth-child(1)').css('visibility', 'collapse');

                if(cols<8){
                  $('#gridHeader').removeClass("fixHeaderLevel");
                  $('#gvGroupLevel').removeClass("gvGroupLvl");
                  FormatTable('#<%=gvGroupLevel.ClientID %>',"#gridHeader",true);
                }
            }
        }
     }

    (function ($) {
        $.fn.hasHorizontalScrollBar = function () {
            return this.get(0) ? this.get(0).scrollWidth > this.innerWidth() : false;
        }
    })(jQuery);
    (function ($) {
        $.fn.hasVerticalScrollBar = function () {
            return this.get(0) ? this.get(0).scrollHeight > this.get(0).clientHeight : false;
        }
        })(jQuery);
  // populate the grouped level data from database
  function populateLevelGrid() {
    var lgpageIndex = parseInt($('#<%=hfPageIndex.ClientID%>').val());
    var totalPage = parseInt($('#<%=hfTotalPage.ClientID%>').val());
    var lvls = LevelInfoTable.map(function(d) { if(d["Excluded"]==1)return d['DisplayLevel']; });
    var prods=ExcludedItems.map(function(d) { return d['ProductID']; });
    lvls = lvls.join(',');
    prods=prods.join(',');
    if(lgpageIndex < totalPage )
    { 
        var warnMessage = '<%=PhraseLib.Lookup("term.loadingmorelevels", LanguageID)%>';
        warnMessage = warnMessage.replace("&#39;", "'");
        $('#notifier').notify(warnMessage ,"info",{ position:"right center" });
        $.ajax({
            type: "POST",
            url: "/logix/pgroup-edit.aspx/LoadLevelsGV",
            data: JSON.stringify({ PageIndex: lgpageIndex, PageSize: 100, _sortBy: '<%=_sortBy%>', _sortOrder: '<%=_sortOrder%>', ProductGroupID: <%=ProductGroupID%>, strExcludedProductIDs: prods, strExcludedLevels:lvls }),
            contentType: "application/json; charset=utf-8",
            dataType: 'json'
        })
        .done(function (response) {
            OnLoadlevelGridSuccess(response);
        })
        .fail(function (response) {
            OnLoadlevelGridonError(response);
        });
    }
    else
    {
        var warnMessage = '<%=PhraseLib.Lookup("term.NoMoreLevelGrps", LanguageID)%>';
        warnMessage = warnMessage.replace("&#39;", "'");
        $('#notifier').notify(warnMessage ,"info",{ position:"right center" });
        spinner.stop();
        $('#GridDataLoading').val("False");
        return;
    }
  }
  function OnLoadlevelGridSuccess(response) {

    var lgpageIndex = parseInt($('#<%=hfPageIndex.ClientID%>').val());
    var totalPage = parseInt($('#<%=hfTotalPage.ClientID%>').val());
    if (lgpageIndex < totalPage) {
      lgpageIndex = lgpageIndex + 1;
      $('#<%=hfPageIndex.ClientID%>').val(lgpageIndex.toString());
    }
    //Handle case where response is null
    var obj = JSON.parse(response.d);
    if(obj==undefined ){
        $('#notifier').notify('<%=PhraseLib.Lookup("term.NolevelGrps", LanguageID)%>',"info",{ position:"right center" });
        spinner.stop();
        $('#GridDataLoading').val("False");
        return;
    }

    var Levels = obj.Level
    if(Levels.length==0)
    {
      var warnMessage = '<%=PhraseLib.Lookup("term.NoMoreLevelGrps", LanguageID)%>';
      warnMessage = warnMessage.replace("&#39;", "'");
      $('#notifier').notify(warnMessage ,"info",{ position:"right center" });
      spinner.stop();
      $('#GridDataLoading').val("False");
      return;
    }
    var flag = true;
    var header = gvGroupLevelheader;
    //Loop through each result
    for (var i = 0; i < Levels.length; i++) {
      var level = Levels[i]

      var row = gvGroupLevelRow1.clone();
      //Use the flag to switch between the shaded row and normal row
      if (flag) {
        flag = false;
        row.toggleClass("shaded");
      }
      else {
        flag = true;
      }
      var index = 1;
      //Go through the each cloned row and update the cells of the row with new data and append it to the Grid
      header.find('th').each(function () {
        var headertext = this.textContent;
        //If it is check box cell update the checkbox property id, with current product id using regex 
        if (headertext.trim() == '') {

		      var prodid = $("td:nth-child(" + index + ")", row).find("span").attr('class');
          var target = level["RowNum"];
          var cb =$("td:nth-child(" + index + ")", row).html();
          //If the Master checkbox is checked, load all the child check boxes by checked..
          if($(cb).children().is("input")){
            $("td:nth-child(" + index + ")", row).find('span').attr('class',target);
            $("td:nth-child(" + index + ")", row).find('input').attr('id',target);
            $("td:nth-child(" + index + ")", row).find('input').attr('name',target);
            $("td:nth-child(" + index + ")", row).find('input').attr("onclick","ToggleGroupCheckbox(this, \""+escape(level.DisplayLevel)+"\")")
            cb= $("td:nth-child(" + index + ")", row).html();
            cb= cb.replace(/checked=/g, '');//Remove the checked attribute if it has any..
            var checked=false;
            checked=level.Excluded;
            //If any modifications happend in the grid, consider the current state          
            if( $('#<%=hfIsGridUpdated.ClientID%>').val()=="1"){
              var lvlInfo;
              var _pos= LevelInfoTable.map(function(d) { return d['DisplayLevel']; }).indexOf(unescape(level.DisplayLevel));
              if(_pos!=-1)
              lvlInfo = LevelInfoTable[_pos];
              //If the level is excluded, mark the checkbox.
              if(lvlInfo != undefined && lvlInfo.ConsiderLastState==0 ){
                if(lvlInfo.Excluded==1)
                  checked = true;
               }
            }
            //If the product inside the level is included after excluding level then the level should not be excluded.
            _pos=IncludedItems.map(function(d) { return d['LevelID']; }).indexOf(level["DisplayLevel"]);
            if(checked && _pos!=-1)
              checked=false;

             if(checked ){
               $(cb).children().attr("checked",true);
            }
          }
          else{
                $("td:nth-child(" + index + ")", row).find('span,div').attr('class',target);
                $("td:nth-child(" + index + ")", row).find('img').attr("onclick","Show_Hide_ProductsGrid(this, \""+escape(level.DisplayLevel)+"\")")
                cb= $("td:nth-child(" + index + ")", row).html();
          }
          $("td:nth-child(" + index + ")", row).html(cb);
        }
        //If it is content cell update content  directly
        else {
        if(headertext==$('#<%=hfGroupedOn.ClientID%>').val())//for group grid, the header varies, consider Display level..
          $("td:nth-child(" + index + ")", row).html(level["DisplayLevelName"]);
        else if(headertext==$('#<%=hfItemUPC.ClientID%>').val())//for group grid, the header varies. For ExtProductid take it as UPC..
          $("td:nth-child(" + index + ")", row).html(level["ExtProductID"]);
        else
          $("td:nth-child(" + index + ")", row).html(level[headertext]);
        }
        index = index + 1;
      });
      $('#<%=gvGroupLevel.ClientID%>').append(row);
    }
    updateIndeterminateState();
    spinner.stop();
    $('#GridDataLoading').val("False");
  }
  function OnLoadlevelGridonError(response) {
    $('#notifier').notify('<%=PhraseLib.Lookup("term.someerror", LanguageID)%>',"error",{ position:"right center" });
    $('#GridDataLoading').val("False");
    spinner.stop();
  }
  //Called when Expand/Collapse button is clicked.
  function Show_Hide_ProductsGrid(control,levelId) {
    levelId=unescape(levelId);
    //Get the level id and create a table by appending levelid to table name and populate that particular data.
    if($('#GridDataLoading').val()=="False" || $('#GridDataLoading').val()==""){
         var levelRowId = $(control).parent().attr('class');
        if ($(control).attr("src").indexOf('plus') != -1) {
          var target = document.getElementById('dvGrid');
          $('#GridDataLoading').val("True");
          spinner = new Spinner(opts).spin(target);
          $(control).attr("src", "../../images/minus2.png")
          var i=1;
          $(control).closest('tr').children("td").each(function(){
            if(i>3){
              //Set the font color to back-ground color to hide the text
              if($(control).closest('tr').attr("class") == "shaded")
                $(this).addClass("hideShaded");
              else
                $(this).addClass("hide");
            }
            i++;
          });

          populateProducts(levelRowId,levelId, control);
        }
        else {
          $(control).attr("src", "../../images/plus2.png")
          document.getElementById("divProducts" + levelRowId).style.display = "none";
          var i=1;
          $(control).closest('tr').children("td").each(function(){
            if(i>3){
            //Set the font-color back to previous color
            if($(control).closest('tr').attr("class") == "shaded")
                $(this).removeClass("hideShaded");
              else
                $(this).removeClass("hide");
            }
            i++;
          });
        }
    }
    else{
    $('#notifier').notify('<%=PhraseLib.Lookup("term.waitdataloading", LanguageID)%>',"info",{ position:"right center" });
    }
  }
  //For the first time we need pass control, to create the table, from 2nd call its not required..
    function populateProducts(levelRowId, levelId, control) {

        // populate data from database
        var ProdspageIndex = 0;
        var TotalProdPages = 0;
        var PageIndexattr = "#tblProducts" + levelRowId + "PageInd"; ;
        var TotalProdPagesattr = "#tblProducts" + levelRowId + "TotalPage"; ;
        var table = $("#tblProducts" + levelRowId + "").html();
        var warnMessage1;
        //Store the page index of products group page respective to each level in body using 'data' function
        if (table != undefined) {
            ProdspageIndex = parseInt($.data(document.body, PageIndexattr));
            TotalProdPages = parseInt($.data(document.body, TotalProdPagesattr));
            if(ProdspageIndex >=TotalProdPages )
            {
                if(document.getElementById("divProducts" + levelRowId)!=undefined){
                    document.getElementById("divProducts" + levelRowId).style.display = "block";
                }
                if(control==undefined){
                    warnMessage1 = '<%=PhraseLib.Lookup("term.nomoreproduct", LanguageID)%>';
                    warnMessage1.replace("&#39;", "'");
                    $('#notifier').notify(warnMessage1 ,"info",{ position:"right center" });
                }
                spinner.stop();
                $('#GridDataLoading').val("False");
                return;
            }
        }
        if (ProdspageIndex < TotalProdPages) {
            ProdspageIndex = ProdspageIndex + 1;
            $.data(document.body, PageIndexattr, ProdspageIndex);
        }
        var warnMessage1;        
        var prods = ExcludedItems.map(function(d) { if(d['LevelID']==levelId)return d['ProductID']; });
        prods = prods.join(',');
        warnMessage1 = '<%=PhraseLib.Lookup("term.loadingmoreproduct", LanguageID)%>';
        warnMessage1.replace("&#39;", "'");
        $('#notifier').notify(warnMessage1 ,"info",{ position:"right center" });
        $.ajax({
            type: "POST",
            url: "/logix/pgroup-edit.aspx/LoadProductsByLevel",
            data: JSON.stringify({ Level : levelId,PageIndex: ProdspageIndex, PageSize: 100, _sortBy: '<%=_sortBy%>', _sortOrder: '<%=_sortOrder%>', ProductGroupID: <%=ProductGroupID%>, strExcludedProductIDs: prods}),
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        })
        .done(function (response) {
            OnLoadProductGridSuccess(response, levelRowId,levelId, control) ;
        })
        .fail(function (response) {
            OnLoadProductGridonError(response);
        });
    }
  function OnLoadProductGridSuccess(response, levelRowId, levelId, control) {
    var PageIndexattr = "#tblProducts" + levelRowId + "PageInd"; ;
    var TotalProdPagesattr = "#tblProducts" + levelRowId+ "TotalPage"; ;
    var obj = JSON.parse(response.d);
    if(obj!= ''){
        var Products = obj.Product;
        var TotalPages = obj.TotalPages[0].TotalPages;
        if ($("#tblProducts" + levelRowId + "").html() == undefined) {
          CreateProductsTable(levelRowId, levelId, Products, control);
          $.data(document.body, TotalProdPagesattr, TotalPages);
          $.data(document.body, PageIndexattr, 0);
        }
        else {
          updateProductsTable(levelRowId,levelId, Products);
        }
    }
    spinner.stop();
    $('#GridDataLoading').val("False");
  }
  function CreateProductsTable(levelRowId, levelId, Products, control) {
    $("div." + levelRowId + "").html("<div id='divProducts" + levelRowId + "' style='max-height: 180px;overflow-y: auto;' onscroll='levelScroll(this,\""+levelRowId+"\",\""+escape(levelId)+"\");'><table id='tblProducts"+ levelRowId + "' class='tablescroll'><tbody style=''></tbody></table></div>");
    FormProductsTable(levelRowId , levelId, Products)
    showProducts(control);
    FormatTable('#<%=gvGroupLevel.ClientID %>','#tblProducts'+ levelRowId ,false);
  }
  //This method will format the target table based on source table format
  //When IsHeader is true, then format the header table ths width with the content table th width.
  //When IsHeader is false, then Format the inner table tds and align them with parent table.
  function FormatTable(src,target,IsHeader){
    var count=0;
    var flag=false
    var max;
    if(IsHeader){
        max= $(target+' tr:first-child').eq(0).children("th").length-1;
        $(src+' tr:first-child').eq(0).children("th").each(function(){
          if(count<max){//Set the width for all the ths, except the last th, as we dont have scroll bar on the top..
                 if($(target+' tr:first-child').eq(0).children("th")[count]!=undefined){
                     $(target+' tr:first-child').eq(0).children("th")[count].width=$(this).css("width");
                     }
                 count=count+1;
             }
        });
    }
    else{
        max= $(target+' tr:first-child').eq(0).children("td").length-1;
        //Loop through 2nd tr,(1st is th) and get the width and set the same to corresponding new table tr.
        $(src+' tr:nth-child(2)').eq(0).children("td").each(function(){
          if(count<max){
              if(flag){
                if($(target+' tr:first-child').eq(0).children("td")[count]!=undefined)
                  $(target+' tr:first-child').eq(0).children("td")[count].width=$(this).css("width");//this.offsetWidth+'px';
                  //$(target+' tr:first-child').eq(0).children("td")[count].minWidth=$(this).css("width");//this.offsetWidth+'px';
                  count=count+1;
               }
              else//For inner table we dont have image row, so dont set the width.
                flag=true;
           }
           else//For last row, set the width less, as we may have scrollbar.
           {
              var wid=this.offsetWidth-20;
              if($(target+' tr:first-child').eq(0).children("td")[count]!=undefined)
                 $(target+' tr:first-child').eq(0).children("td")[count].width=wid+'px';
             count=count+1;
           }
        });
      }
   }
  function updateProductsTable(levelRowId, levelId, Levels) {
    document.getElementById("divProducts" + levelRowId).style.display = "block";
    FormProductsTable(levelRowId, levelId, Levels)
  }
  function FormProductsTable(levelRowId,levelId, Products) {
    var flag = true;
    //Get the header rows.
    var header = gvGroupLevelheader;
    for (var i = 0; i < Products.length; i++) {
      var Product = Products[i]
      var row = gvGroupLevelRow1.clone();
      if (flag) {
        row.toggleClass("shaded");
        flag = false;
      }
      else {
        flag = true;
      }
      var index = 1;
      var generateColumn = false; 
      header.find('th').each(function () {
        //Get the header text of each column
        var headertext = this.textContent;
        headertext= headertext.replace('▲','').trim();
        headertext= headertext.replace('▼','').trim();

        if (headertext.trim() != '') {
            if(headertext==$('#<%=hfGroupedOn.ClientID%>').val())//Dont display Level info inner table
              $("td:nth-child(" + index + ")", row).html("");
            else if(headertext==$('#<%=hfItemUPC.ClientID%>').val())//for group grid, the header varies. For ExtProductid take it as UPC..
              $("td:nth-child(" + index + ")", row).html(Product["ExtProductID"]);
            else
              $("td:nth-child(" + index + ")", row).html(Product[headertext]);
        }
        else {
          //Generate check box column
          if (generateColumn) {
            var prodid = Product["ProductID"].toString();
            var checked ="";
            //If any modifications happend in grid and went back to PAB screen, use the excluded itms to display exclusions.
            if( $('#<%=hfIsGridUpdated.ClientID%>').val()=="1"){
                var lvlInfo;
                var _pos= LevelInfoTable.map(function(d) { return d['DisplayLevel']; }).indexOf(unescape(levelId));
                if(_pos!=-1)
                lvlInfo = LevelInfoTable[_pos];
                //If the group is modified, dont consider previous state
                if(lvlInfo!=undefined && lvlInfo.ConsiderLastState==0 ){
                  //If the product level is excluded, then exclude the product.
                  if(lvlInfo.Excluded==1)
                   checked = "checked";
                  //Check whether the product is in excluded list
                  if(checked ==""){
                    var pos = ExcludedItems.map(function(d) { return d['ProductID']; }).indexOf(prodid);
                    if(pos!=-1)
                      checked = "checked";
                  }
                  if(checked !=""){
		        	      //If the product is in included list remove checked attribute 
                    var pos = IncludedItems.map(function(d) { return d['ProductID']; }).indexOf(prodid);
                    if(pos!=-1)
                      checked = "";
                   }
                }
                //If group is not modified, then consider the product state
                else
                {
                   //Check whether the product is in excluded list
                   if(checked ==""){
                      var pos = ExcludedItems.map(function(d) { return d['ProductID']; }).indexOf(prodid);
                      if(pos!=-1)
                        checked = "checked";
                   }
                   if(checked ==""){
                     checked = $("#"+levelRowId).is(':checked')? "checked" : "" ;//Check whether the group checkbox is checked..
                   }
                   if(checked ==""){
                     checked=Product.Excluded.toLowerCase()=="true"? "checked" : "" ;//Check whether the product is excluded
                   }
                   if(checked !=""){
		        	      //If the product is in included list remove checked attribute 
                    var pos = IncludedItems.map(function(d) { return d['ProductID']; }).indexOf(prodid);
                    if(pos!=-1)
                      checked = "";
                   }
                }
             }
             //If grid is not updated, use the database state
             else{
                 if(checked ==""){
                   checked=Product.Excluded.toLowerCase()=="true"? "checked" : "" ;//Check whether the product is excluded
                 }
             }
             //Insert in to exclude list
              
             if(checked!=""){
               //var pos = ExcludedItems.map(function(d) { return d['ProductID']; }).indexOf(prodid);
               //if(pos==-1)
                //ExcludedItems.push({"LevelID": levelId ,"ProductID": prodid });
             }
            //$("td:nth-child(" + index + ")", row).html("<input type='checkbox' id='" + prodid + "' "+ checked +" onclick=\"javascript:ToggleProductCB(this,\'"+levelId+"\',"+levelRowId+");\"/>");
            var _html="<input type='checkbox' id='" + prodid + "' "+ checked +" onclick=\"javascript:ToggleProductCB(this,\'"+escape(levelId)+"\',"+levelRowId+");\"/>";
            $("td:nth-child(" + index + ")", row).html(_html);
            generateColumn = false;
          }
          //Dont generate the column for Expand/Collapse image in the row.
          else {
            $("td:nth-child(" + index + ")", row).remove();
            index = index - 1;
            generateColumn = true;
          }
        }
        index = index + 1;
      });
      $("#tblProducts" + levelRowId + "").append(row);
    }
  }
  function OnLoadProductGridonError(response) {

    $('#notifier').notify('<%=PhraseLib.Lookup("term.waitdataloading", LanguageID)%>',"info",{ position:"right center"  });
    spinner.stop();
    $('#GridDataLoading').val("False");
  }
  //Add the newly generated grid in the downside..
  function showProducts(control) {
    if ($(control)[0].src.indexOf("minus") != -1 && $(control).next().html() != undefined) {
      $(control).closest("tr").after("<tr><td></td><td colspan = '999'>" + $(control).next().html() + "</td></tr>");
      $(control).next().remove();
    }
  }
  function levelScroll(con,levelRowId, levelId){
    levelId= unescape(levelId);//To handle single quotes in the string
    if($(con)[0].scrollWidth >= $(con)[0].innerWidth){
     return;
    }
    if ($(con).scrollTop() + $(con).innerHeight() >= $(con)[0].scrollHeight) {
        if($('#GridDataLoading').val()=="False" || $('#GridDataLoading').val()=="")
        {
          //Get the levelId and populate the products..
          //var levelId = con.id.substring(11);//con is div and its name starts with "divProducts", so take from 11.
          var target = document.getElementById('dvGrid');
          $('#GridDataLoading').val("True"); // GridDataLoading value will be made false, in the end of loading.
          spinner = new Spinner(opts).spin(target); 
          populateProducts(levelRowId,levelId, null);
       }
     }
   }
  
  //If master CB is checked
  function ToggleGroupMasterCB(control){
    $('#<%=hfIsGridUpdated.ClientID%>').val(1);
    ToggleCheckbox(control.checked);
      $(LevelInfoTable).each(function(){
      var ele= document.getElementsByName(this.ID)[0];
      if(ele!=undefined){
        ele.indeterminate = false;
	  	  ele.className = "";
        }
        if(control.checked){
          this.ExcludedCount=this.TotalCount;
        this.ConsiderLastState=0;
        this.Excluded=1;
      }
      else{
          this.ExcludedCount=0;
        this.ConsiderLastState=0;
        this.Excluded=0;
        }
      });
     ExcludedItems=[];
     IncludedItems=[];
  }
  
   //If any group check box is checked or unchecked, update the count and hidden field..
  function ToggleGroupCheckbox(control, id){
    id=unescape(id);//To handle single quotes in the string
    var isMasterCBChecked=$('#<%=hdnIsMasterCBChecked.ClientID %>').val();
    var obj="#tblProducts"+control.id;
    $('#<%=hfIsGridUpdated.ClientID%>').val(1);

      //Incase of indeterminate state set this to check instead of uncheck.
      if($(control).hasClass("indeterminate")){
        control.checked=true; 
        $(control).removeClass("indeterminate");
      }
    //Upadte the check boxes
    $(obj).find("input:checkbox").each(function () {
          this.checked = control.checked;
      });
      if(isMasterCBChecked){
        }
      $(LevelInfoTable).each(function(){
        var level=this;
        if(level.DisplayLevel==id){
            var ActualProductCount = parseInt($('#<%=hdnInclProducts.ClientID %>').val());
            if (ActualProductCount == 'NaN')
                ActualProductCount = 0;
            if (control.checked) {//If whole group is Excluded
              if(level.ExcludedCount != level.TotalCount){
                  ActualProductCount = ActualProductCount - (level.TotalCount-level.ExcludedCount);
                 }
                else{
                  ActualProductCount = ActualProductCount - level.TotalCount;
                }
               level.ConsiderLastState=0; 
               level.ExcludedCount = level.TotalCount;
               level.Excluded=1;
               $('#notifier').notify('<%=PhraseLib.Lookup("term.excluded", LanguageID)%>' + " "+level.TotalCount +" "+'<%=PhraseLib.Lookup("term.prodsinlevel", LanguageID) %>',"info",{ position:"right center" });
            }
            else {//If group is included
              $("#Groupcheckall")[0].checked=false;
              if(level.ExcludedCount == 0 || level.ExcludedCount == level.TotalCount){
                ActualProductCount = ActualProductCount + level.TotalCount;
                }
                else{
                  ActualProductCount = ActualProductCount + (level.TotalCount-level.ExcludedCount );
                }
                level.ConsiderLastState=0;
                level.ExcludedCount = 0;
                level.Excluded=0;
               $('#notifier').notify('<%=PhraseLib.Lookup("term.included", LanguageID)%>' + " " +level.TotalCount +" " + '<%=PhraseLib.Lookup("term.prodsinlevel", LanguageID) %>',"info",{ position:"right center" });
            }
            //if whole group is included/excluded, delete their entries from exlude lis and include list
            updateIncludeProductsInLevel(id);
            updateExcludeProductsInLevel(id);

            $('#<%=lbTotalProducts.ClientID %>').text(ActualProductCount);
            $('#<%=hdnInclProducts.ClientID %>').val(ActualProductCount);
            ChangeColor();
          }
      });
  }

  //Removes all the products in the exclude list for the given level 
  function updateExcludeProductsInLevel(lvlid){
      if(lvlid!=undefined){
        var pos = ExcludedItems.map(function(d) { return d['LevelID']; }).indexOf(lvlid);
        if(pos!=-1){
          ExcludedItems.splice(pos,1);
          updateExcludeProductsInLevel(lvlid);
          }
          else
            return;
      }
    }
	//Removes all the products in the include list for the given level
    function updateIncludeProductsInLevel(lvlid){
      if(lvlid!=undefined){
          var pos = IncludedItems.map(function(d) { return d['LevelID']; }).indexOf(lvlid);
          if(pos!=-1){
            IncludedItems.splice(pos,1);
            updateIncludeProductsInLevel(lvlid);
          }
          else
            return;
       }
   }

  //Handles Check/Uncheck of product
  function ToggleProductCB(control,id,LevelRowId){
      id=unescape(id);//To handle single quotes in the string
      $('#<%=hfIsGridUpdated.ClientID %>').val(1);
      updateCount(control);
      var level;
      if(Groupgrid){
           if (control.checked) {
            $(LevelInfoTable).each(function(){
                        level=this;
                        if(level.DisplayLevel==id){
                            level.ExcludedCount=level.ExcludedCount +1;
                            return false;
                            }
                   });
              var pos = IncludedItems.map(function(d) { return d['ProductID']; }).indexOf(control.id)
              if(pos!=-1)
                IncludedItems.splice(pos,1);
              pos = ExcludedItems.map(function(d) { return d['ProductID']; }).indexOf(control.id)
              if(pos==-1)
                ExcludedItems.push({"LevelID": id ,"ProductID": control.id});
            }
            else {
            $("#Groupcheckall")[0].checked=false;
            $(LevelInfoTable).each(function(){
                        level=this;
                        if(level.DisplayLevel==id){
                            if(level.ExcludedCount > 0)
                            level.ExcludedCount=level.ExcludedCount-1;
                            return false;
                            }
                   });
              var pos = ExcludedItems.map(function(d) { return d['ProductID']; }).indexOf(control.id)
              if(pos!=-1)
                 ExcludedItems.splice(pos,1);
                
              pos = IncludedItems.map(function(d) { return d['ProductID']; }).indexOf(control.id)
              if(pos == -1)
                 IncludedItems.push({"LevelID": id,"ProductID": control.id });
            }
            //Update the level checkbox
            if($('#'+LevelRowId)!=undefined){
              var ele=document.getElementsByName(LevelRowId)[0];
              if(level.ExcludedCount==level.TotalCount){
                ele.checked=true;
                ele.indeterminate = false;
                ele.className ='';
               }
              else if(level.ExcludedCount==0){
                ele.checked=false;
                ele.indeterminate = false;
                ele.className ='';
                }
              else{
                ele.indeterminate = true;
                ele.className = " indeterminate ";
              }
            }
        }
    }
    
  //Sets the checkbox state to indeterminate(partially checked) if only few proucts are excluded inside the level
  //This needs to be set only using the jquery as this is not a HTML property.
  function updateIndeterminateState(){
  if(Groupgrid){
      $(LevelInfoTable).each(function(){
         var level=this;
         var LevelRowId=level.ID;
         var ele=document.getElementsByName(LevelRowId)[0];
         if(ele!=undefined){
              if(level.ExcludedCount==level.TotalCount){
              ele.checked=true;
              ele.indeterminate = false;
              ele.class = "";
              }
            else if(level.ExcludedCount==0){
              ele.checked=false;
              ele.indeterminate = false;
              ele.class = "";
              }
            else{
              ele.checked=false;
              ele.indeterminate = true;
              ele.className =" indeterminate ";
            }
        }});
      }
    }


    function updateGroupMasterCB(){
      if(Groupgrid){
         var lvls = $.grep(LevelInfoTable, function(d) { if(d["Excluded"]==1)return d['DisplayLevel']; });
         if(lvls != undefined && lvls.length==LevelInfoTable.length && $("#Groupcheckall").length)
           $("#Groupcheckall")[0].checked=true;
      }
    }

	//Updates the json object changes to the hidden field
    function UpdateProductChanges(){
      if(Groupgrid && typeof ExcludedItems !== "undefined"){
        $('#<%=hfExcludedItems.ClientID %>').val(JSON.stringify(ExcludedItems));
        $('#<%=hfIncludedItems.ClientID %>').val(JSON.stringify(IncludedItems));
        $('#<%=hfLevelInfoTable.ClientID %>').val(JSON.stringify(LevelInfoTable));
      }
    }

    function spin(target){
      var target = document.getElementById(target);
      if(target !=undefined)
         spinner = new Spinner(opts).spin(target);
    }
</script>
