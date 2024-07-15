﻿<%@ Page Language="C#" Title="BE_CoreDigitalSplitReport" AutoEventWireup="true" CodeBehind="BE_CoreDigitalSplitReport.aspx.cs" Inherits="BECodeProd.BE_CoreDigitalSplitReport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>

    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <script src="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>


   
    <style type="text/css">
            .button
        {
            border-style: solid;
            background-color: #f8da92;
            padding: 1px 0px;
            border-color: red;
            border-width: 1px;
            cursor: pointer;
            cursor: hand;
            font-family: Calibri;
            font-size: 9pt;
        }
        .button:hover
        {
            border-style: solid;
            background-color: #c41502;
            border-color: Black;
            color: White;
            border-width: 1px;
            padding: 1px 0px;
            cursor: pointer;
            cursor: hand;
            font-family: Calibri;
            font-size: 9pt;
        }
        
                .FormLabel
        {
            
            font-family: Verdana;
            color: #000000;
            font-size: 10px;
            font-weight: normal;
        }
        
        .form-control {
     display: block;
    width: 100%;
     background: url('Images/down.png') no-repeat right;
    padding: 5px 5px 5px 5px;
    font-size: 10px;
    line-height: 1.4285;
    color: #555;
    
   
    border: 1px solid #ccc;
    border-radius: 4px;
    -webkit-box-shadow: inset 0 1px 1px rgba(0, 0, 0, .075);
    box-shadow: inset 0px 1px 1px rgba(0,0,0,0.075);
    -webkit-transition: border-color ease-in-out .15s, -webkit-box-shadow ease-in-out .15s;
    -o-transition: border-color ease-in-out .15s, box-shadow ease-in-out .15s;
    transition: border-color ease-in-out .15s, box-shadow ease-in-out .15s;
}

 select::-ms-expand {
    display: none;
}
    </style>
     <script type="text/javascript">
         var myVar;
         function isvalidupload() {
             var btnreport = document.getElementById('btnreport');
             btnreport.style.visibility = 'hidden';
             document.getElementById('loading').style.visibility = 'visible';
             var a = document.getElementById('lbl').innerHTML;
             var d = $('#dots');
             if (a == "Downloaded" || a == "") {
                 document.getElementById('lbl').innerHTML = 'Downloading';
             }
             (function loading() {

                 myVar = setTimeout(function () {

                     draw = d.text().length >= 5 ? d.text('') : d.append('.');
                     loading();
                 }, 300);

             })();
         }

         function myStopFunction() {
             clearTimeout(myVar);

         }

         function isvaliduploadClose() {
             var btnreport = document.getElementById('btnreport');
             btnreport.style.visibility = 'visible';

         }
    </script>

</head>
<body>
    <form id="form1" runat="server">

    
    <asp:ScriptManager ID="sm" runat ="server">
    </asp:ScriptManager>
    <asp:UpdatePanel ID="up" runat="server" UpdateMode="Conditional">
    <ContentTemplate>
    <div align="center" style="margin-top:5px">
       <table class="style2">
           <tr>
          
           </td>
            <td class="FormLabel" style="padding-left:5px">
               <asp:Label ID="lblQtr" runat="server" Text="Select Quarter: "></asp:Label></td>
           <td >
               <asp:DropDownList ID="ddlQuarter" runat="server" Font-Names="Calibri" Font-Size="11px"
              Height="25" class="form-control" width="78" style="margin-left:5px"
              >
               </asp:DropDownList>
           </td>
            <td class="FormLabel" style="padding-left:5px">
                             Select ServiceLine: &nbsp;
                        </td>
                      
                                    <td>
                                        <asp:DropDownList ID="ddlServiceLine" runat="server" Font-Names="Calibri" Font-Size="11px"
                                            AutoPostBack="True" Height="25" class="form-control" width="78" style="margin-left:5px">
                                        </asp:DropDownList>
                                    </td>

           </tr>

           <tr>
              <td colspan="3" align="left">
               

              </td>

           </tr>
           
            </table>
            </div>
            <div style ="width:130px;margin:0px auto;padding-top:20px">
           
            <div style="float:left;margin-right:10px">   
            <asp:Button ID="btnreport" runat="server" Text="Generate Report" 
                        OnClick="btnreport_Click" OnClientClick="return isvalidupload()" CssClass="btn btn-success"
                         /></div>
          
             
              <div id="loading" align="left" runat="server" style="font-size: medium; font-weight: bold;
                                     visibility: hidden; color: #FF0000; font-family: Calibri;float:left;margin-left:17px">
                                 <asp:Label ID="lbl" runat="server" Text="Downloading"></asp:Label>   <span style="width: 50px" id="dots"></span></div>
             </div>
                                  
                                  <iframe id="iframe" runat ="server" style="display:none"></iframe>
               <asp:Label ID="lblError" runat="server" 
                   Text="No data available for this selection" Visible="False" ForeColor="Red"></asp:Label>
    
   </ContentTemplate>
    <Triggers>
    <asp:PostBackTrigger ControlID="btnreport" />
    </Triggers>
    </asp:UpdatePanel>
    </div>

   </form>
   
</body>
</html>
