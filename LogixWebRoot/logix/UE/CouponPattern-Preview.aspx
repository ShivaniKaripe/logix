<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CouponPattern-Preview.aspx.cs" Inherits="logix_UE_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
   <style>
       #patternprev {
           background-color: #ffd060;
           border: 1px solid #666666;
           border-top: 1px dotted #666666;
           border-bottom: 1px dotted #666666;
           margin: 6px 0 1px 0;
           padding: 19px;
           text-align: center;
           width:500px;
       }
       #lblpatternprev {
           display: block;
           word-wrap: break-word;
           margin: 1em 0px 1em;
       }
       #ptext{
           white-space: nowrap;
       }
   </style>
</head>
<body class="popup">
    
    <form id="mainform" runat="server">
    
        
    <div id="intro">
        <h1 id="title" runat="server"></h1>
     
    </div>
        <div id="main">
          
            <div id="column1">
              <br /><b> <asp:Label ID="ptext" runat="server" ClientIDMode="Static"></asp:Label></b>
                    <div id="patternprev">
                       <b> <asp:Label ID="lblpatternprev" runat="server" ClientIDMode="Static"></asp:Label></b>
                    </div>
                   
                </div>
           
        </div>
    </form>
</body>
</html>
