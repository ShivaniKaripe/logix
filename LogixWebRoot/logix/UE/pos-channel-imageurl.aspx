<%@ Page Language="C#" AutoEventWireup="true" CodeFile="pos-channel-imageurl.aspx.cs" Inherits="logix_UE_pos_channel_imageurl" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title></title>

    <style>
        #main {
            padding: 20px !important;
            padding-left: 50px !important;
        }

        #formDiv {
            padding-bottom: 20px;
        }

        #ptext {
            padding-bottom: 3px;
        }

        #clearFix {
            padding-bottom: 5px;
        }
    </style>

    <script>

        function imgError(me) {
            var AlterNativeImg = "../../images/PreviewNotAvailable.jpg";
            me.src = AlterNativeImg;
        }

        function buttonClick() {
            window.onunload = null;
            return true;
        }

        function refreshParent() {
            if (window.opener != null) {
                window.opener.location.reload();
            }
        }

    </script>

    <script>
        function handleSaveButton() {
            var _txtImageUrl = document.getElementById("txtImageUrl");
            _txtImageUrl.value = _txtImageUrl.value.trim();
            if (_txtImageUrl.value != "") {
                document.getElementById('btnSave').disabled = false;
                document.getElementById('btnPreview').disabled = false;
            } else {
                document.getElementById('btnSave').disabled = true;
                document.getElementById('btnPreview').disabled = true;
            }
        }
    </script>
</head>
<body class="popup" onload="handleSaveButton()" onunload="refreshParent()">

    <form id="mainform" runat="server">

        <div id="intro">
            <h1 id="title" runat="server"></h1>
        </div>

        <div id="main">

            <div id="infobar" runat="server" clientidmode="Static" style="display: none; width: 400px;" />
            <div id="formDiv">
                <br />
                <asp:Label ID="ptext" runat="server" ClientIDMode="Static"></asp:Label>
                <br />
                <asp:TextBox ID="txtImageUrl" runat="server" MaxLength="500" Width="300px" onkeyup="handleSaveButton()"></asp:TextBox>
                <br />

                <div class="clearFix">&nbsp;</div>

                <asp:Button ID="btnPreview" runat="server" OnClientClick="return buttonClick();" OnClick="btnPreview_Click" />

                &nbsp;&nbsp;&nbsp;&nbsp;

                <asp:Button ID="btnSave" runat="server" OnClientClick="return buttonClick();" OnClick="btnSave_Click" />

            </div>
            <div class="clearFix"></div>

            <asp:Image ID="PreviewImg" runat="server" Visible="false" onerror="imgError(this)" Height="200" Width="200" />

        </div>

    </form>

</body>
</html>
