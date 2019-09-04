<%@ Page Language="C#" AutoEventWireup="true" CodeFile="UEoffer-rew-proximitymsgpreview.aspx.cs" Inherits="logix_configurator_UEoffer_rew_proximitymsgpreview" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript" src="/javascript/logix.js"></script>
    <script type="text/javascript" src="/javascript/jquery.min.js"></script>
    <script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
    <link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" /> 
     <style type="text/css">
       .wrapword{
            white-space: -moz-pre-wrap !important;  /* Mozilla, since 1999 */
            white-space: -pre-wrap;      /* Opera 4-6 */
            white-space: -o-pre-wrap;    /* Opera 7 */
            white-space: pre-wrap;       /* css-3 */
            word-wrap: break-word;       /* Internet Explorer 5.5+ */
            word-break: break-all;
            white-space: normal;
            }

    </style>   
</head>
<body class="popup">

    <div id="custom1">
    </div>
    <div id="wrap">
        <div id="custom2"></div>
        <a id="top" name="top"></a>'

        <div id="intro">
            <h1 id="title">Proximity message preview</h1>
        </div>
        <div id="main">
            <div id="column2x">
                <div id="pmsgpreview">
                    <div id="pmsgpreviewbody" class="wrapword">
                        <label runat="server" id="proximityMessage"  ></label>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
