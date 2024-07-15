<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Error.aspx.cs" Title="Error-Page"
    Inherits="iSHARE.Error" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    
    <link href="SiteError.css" rel="stylesheet" />
    <title></title>
    <link rel="icon" href="Images/favicon.ico" type="image/vnd.microsoft.icon" />
    <script type="text/javascript">
        $(document).ready(function () {
            var div = document.getElementById('divgrid');
            if (div != null) {
                div.style.width = (window.screen.width - 65) + 'px';
                div.style.height = (window.screen.height - 280) + 'px'; // address bar, favorites, tool bar , status bar 
            }
        });
            
    </script>
    <style type="text/css">
        .GridDock
        {
            overflow-x: auto;
            overflow-y: auto;
            padding: 0 0 0 0;
        }
    </style>
</head>
<body style="background-color: white">
    <form id="form1" runat="server" style="background-color: white">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" style="padding-right: 0px;
        padding-left: 0px; padding-bottom: 0px; padding-top: 0px; margin: 0px; background-color: #FFFFFF;">
         
        <tr>
            <td colspan="3" align="center">
                <div id="divgrid" class="GridDock" runat="server">
                    <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" style="height: 424px">
                        <tr>
                            <td style="width: 0px; height: 415px;">
                            </td>
                            <td style="width: 100%; height: 415px; background-color: White" valign="middle">
                                <table width="100%">
                                    <tr>
                                        <td align="center">
                                            <div style="background-color: #f9fcff; border: 3px solid #e6eaf1; height: 200px;
                                                width: 550px">
                                                <table width="100%">
                                                    <tr>
                                                        <td style="width: 55px">
                                                        </td>
                                                        <td>
                                                            <table width="100%">
                                                                <tr>
                                                                    <td align="left">
                                                                        &nbsp;
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td style="width: 55px">
                                                           
                                                            <img src="error1.png" alt="Alternate Text" />
                                                        </td>
                                                        <td>
                                                            <table width="100%">
                                                                <tr>
                                                                    <td align="left">
                                                                        &nbsp;
                                                                        <asp:Label ID="lbl1" Text="Error" Font-Size="15px" runat="server"> </asp:Label>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table width="100%">
                                                                <tr>
                                                                    <td align="left">
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td style="width: 55px">
                                                        </td>
                                                        <td>
                                                            <table width="100%">
                                                                <tr>
                                                                    <td align="left">
                                                                        <asp:Label ID="lblMessage" Text="An Unexpected Error Occured while performing this operation.  
Please Try again later. If problem persists Please contact administrator  " Font-Size="12px" runat="server"></asp:Label>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td style="width: 55px">
                                                        </td>
                                                        <td>
                                                            <table width="100%">
                                                                <tr>
                                                                    <td align="left">
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                     
                                                    <tr>
                                                        <td style="width: 55px">
                                                        </td>
                                                        <td>
                                                            <table width="100%">
                                                                <tr>
                                                                    <td align="left">
                                                                        &nbsp; &nbsp;
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                     
                                                     
                                                </table>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </div>
            </td>
        </tr>
         
    </table>
    </form>
</body>
</html>
