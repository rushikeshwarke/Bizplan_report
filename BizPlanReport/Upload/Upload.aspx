<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Upload.aspx.cs" Inherits="BEBIZ.Upload.Upload" %>



<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>

    <style>
        /*Modal*/
        /* The Modal (background) */
.modal {
  display: none; /* Hidden by default */
  position: fixed; /* Stay in place */
  z-index: 1; /* Sit on top */
  padding-top: 100px; /* Location of the box */
  left: 0;
  top: 0;
  width: 100%; /* Full width */
  height: 100%; /* Full height */
  overflow: auto; /* Enable scroll if needed */
  background-color: rgb(0,0,0); /* Fallback color */
  background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
}

/* Modal Content */
.modal-content {
  position: relative;
  background-color: #fefefe;
  margin: auto;
  padding: 0;
  border: 1px solid #888;
  width: 700px;
  box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2),0 6px 20px 0 rgba(0,0,0,0.19);
  -webkit-animation-name: animatetop;
  -webkit-animation-duration: 0.4s;
  animation-name: animatetop;
  animation-duration: 0.4s
}

/* Add Animation */
@-webkit-keyframes animatetop {
  from {top:-300px; opacity:0} 
  to {top:0; opacity:1}
}

@keyframes animatetop {
  from {top:-300px; opacity:0}
  to {top:0; opacity:1}
}

/* The Close Button */
.close {
  color: white;
  float: right;
  font-size: 28px;
  font-weight: bold;
}

.close:hover,
.close:focus {
  color: #000;
  text-decoration: none;
  cursor: pointer;
}

.modal-header {
  padding: 2px 16px;
  background-color: #4F617A;
  color: white;
}

.modal-body {padding: 2px 16px;}

.modal-footer {
  padding: 2px 16px;
  background-color: #5cb85c;
  color: white;
}
    </style>
    
    
   
    
     <script>

         

         function gif() {
             setTimeout(function () {

                 document.getElementById('gif').style.display = "block";
             }, 50);




         }

         function gif1() {
             setTimeout(function () {

                 document.getElementById('gif1').style.display = "block";
             }, 50);

             


         }

           function gif2() {
             setTimeout(function () {

                 document.getElementById('gif2').style.display = "block";
             }, 50);




         }

     </script>

    <style>
        .btn {
            color: white;
            background-color: #5cb85c;
            border: 0px;
            padding: 5px;
            cursor: pointer;
            outline:none!important;
            outline-width:0px!important;
        }

       span.select2-selection.select2-selection--single {
        outline: none;
    }

        .txt {
            width: 50px;
        }
    </style>
    <script>
        function funvalidate() {
            var sl = document.getElementById('ddlSL').value;
            if (sl == '') {
                alert("Please select the service line");
                return false;
            }
            var file = document.getElementById('fuUploader').value;
            if (file == '') {
                alert("Please select the file");
                return false;
            }
            return true;
        }

        function showdownload(obj) {
            var tr = document.getElementById('trdwonload');
            var tr1 = document.getElementById('trdwonload1');
            var tr2 = document.getElementById('trdwonload2');

            var l2 = document.getElementById('link2');
            var l3 = document.getElementById('link3');

            tr.style.display = (tr.style.display == 'none') ? 'block' : 'none';

            tr1.style.display =  'none' ;
            tr2.style.display = 'none';

            l2.textContent = 'Show Download SPO EAS Report (Internal)'
            l3.textContent = 'Show Download Service Line BPlan (CDO)'

            if (tr.style.display == 'none')
                obj.textContent = 'Show Download'
            else
                obj.textContent = 'Hide Download'

        }

        function showdownload1(obj) {
            var tr = document.getElementById('trdwonload');
            var tr1 = document.getElementById('trdwonload1');
            var tr2 = document.getElementById('trdwonload2');

            var l1 = document.getElementById('link1');
            var l3 = document.getElementById('link3');

            tr1.style.display = (tr1.style.display == 'none') ? 'block' : 'none';
             
            tr.style.display = 'none';
            tr2.style.display = 'none';

            l1.textContent = 'Show Download'
            l3.textContent = 'Show Download Service Line BPlan (CDO)'

            if (tr1.style.display == 'none')
                obj.textContent = 'Show Download SPO EAS Report (Internal)'
            else
                obj.textContent = 'Hide Download SPO EAS Report (Internal)'

        }

        function showdownload2(obj) {
            var tr = document.getElementById('trdwonload');
            var tr1 = document.getElementById('trdwonload1');
            var tr2 = document.getElementById('trdwonload2');

            var l1 = document.getElementById('link1');
            var l2 = document.getElementById('link2');

            tr2.style.display = (tr2.style.display == 'none') ? 'block' : 'none';

            tr.style.display = 'none';
            tr1.style.display = 'none';

            l1.textContent = 'Show Download'
            l2.textContent = 'Show Download SPO EAS Report (Internal)'
            
            if (tr2.style.display == 'none')
                obj.textContent = 'Show Download Service Line BPlan (CDO)'
            else
                obj.textContent = 'Hide Download Service Line BPlan (CDO)'

       

        }

        function changemodel(obj) {
            debugger;
            var value = obj.value;
            var div = document.getElementById('divforModerateType');
            div.style.display = (value == 'Moderate') ? 'block' : 'none';

            

            var div1 = document.getElementById('divforModerate');
            div1.style.display = ($('#ddlModerateType').val() == 'Dynamic') ? 'block' : 'none';
        }

        function changeModerateType(obj) {
            var value = obj.value;
            var div = document.getElementById('divforModerate');
            div.style.display = (value == 'Dynamic') ? 'block' : 'none';
        }

        var myVar;
        function isvalidupload() {
            var btnDownloadReport = document.getElementById('btnDownload');
            btnDownloadReport.style.visibility = 'hidden';

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
    </script>

      <script src="../bootstrap-3.3.6-dist/jquery-1.10.2.js"></script>
    <link href="bootstrap-3.3.6-dist/css/bootstrap.css" rel="stylesheet" />
    <script src="../bootstrap-3.3.6-dist/js/bootstrap.min.js"></script>
    <link href="../bootstrap-3.3.6-dist/css/bootstrap.css" rel="stylesheet" />


     


    <script src="../Select2/JScriptSelect2.js"></script>
    <link href="../Select2/select2.css" rel="stylesheet" />
<script>

    $(function () {
        initdropdown();

          var tab = document.getElementById('<%= hidTAB.ClientID%>').value;
        $('#myTabs a[href="' + tab + '"]').tab('show');
        
        var role = document.getElementById('<%= hdnRole.ClientID%>').value;   
        if (role == 'SPOC') {
            
                $('#myTabs a[href="#home"]').hide();
                $('#myTabs a[href="#menu3"]').hide();
            $('#myTabs a[href="#menu1"]').hide();
            $('#myTabs a[href="#menu4"]').hide();
            
        }
        })

        function initdropdown() {
            
            $("#ddlSL,#ddl_SL_1,#ddl_Internal,#ddl_Type,#ddlVertical,#ddlInput,#ddlSLConsolidated").select2({
                selectOnClose: true,
                minimumResultsForSearch: -1

            })
            //.on("change", function (e) { angular.element(document.getElementById("myctrl")).scope().fnstatus(); });

            
    }

    function download1() {

        //document.getElementsByTagName("body").innerHTML+='<iframe class="iframe" style="display:none"></iframe>'
        document.getElementById('iframe').src = 'Download.aspx';
    }

</script>

    <style>
        a {
            color: darkgray!important;
            text-decoration: none;
        }
        td {
            border: 1px solid black;
            font-family: Calibri;
            background-color:ActiveCaption;
            color:Black;
        }

        th {
            background-color: Black;
            font-family: Calibri;
            color: white;
            border: 1px solid black;
            text-align: center;
        }

        a:focus, button:focus {
            outline: 0px !important;
        }

        .well {
            min-height: 20px;
            padding: 19px;
            margin-bottom: 20px;
            border: 1px solid #e3e3e3;
            background-color: white !important;
            border-radius: 4px;
            -webkit-box-shadow: inset 0 1px 1px rgba(0, 0, 0, .05);
            box-shadow: inset 0px 1px 1px rgba(0,0,0,0.05);
        }

       .nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover {
    color: black!important;
    cursor: default;
     font-family:Calibri;
    background-color: white !important;
    height:27px!important;
    padding-top:2px!important;
}
.nav-tabs > li > a {
    margin-right: 2px;
    line-height: 1.4285;
  font-family:Calibri;
    border-radius: 4px 4px 0 0;
   color:white!important;
    background-color:  #4F617A!important;
    height:27px!important;
     padding-top:2px!important;
     font-size:11pt;
}
    </style>
  
 
</head>
    
<body style="background-color: white">
    <iframe id="iframe" style="display:none"></iframe>
    <form id="form1" runat="server">
         <asp:HiddenField ID="hidTAB" runat="server" Value="#menu1" />
          <asp:HiddenField ID="hdnRole" runat="server" Value="" />
          <div class="container" style="margin-top:5px;margin-bottom: 30px;"  >

             
                    <div class="well" >
                        <div style="margin-top:5px;">
                              
                            <ul class="nav nav-tabs" id="myTabs">
                                <li ><a data-toggle="tab" href="#home">Upload </a></li>
                                <li ><a data-toggle="tab" href="#menu3">Input Template </a></li>
                                <li><a data-toggle="tab" href="#menu1">CDO</a></li>
                                 <li class="active"><a data-toggle="tab" href="#menu2">Internal</a></li>
                                <li class="active"><a data-toggle="tab" href="#menu4">Consolidated</a></li>
                            </ul>
                            <div class="tab-content">
                              
                                <div id="home" class="tab-pane fade">
                                    <div style="width:950px;margin:0px auto;padding-top:40px;height:450px">
                                        
                                     <div style="float:left"> Service Line:</div>   <div  style="float:left;margin-left:20px"> <asp:DropDownList ID="ddlSL" runat="server" Width="130" Style="font-family: Calibri; font-size: small">
                               <asp:ListItem Value="SAP" Text="SAP" />
                              <asp:ListItem Value="ORC" Text="ORC" />
                            <asp:ListItem Value="ECAS" Text="ECAS" />
                            <asp:ListItem Value="EAIS" Text="EAIS" />
                           
                          
                            <%--<asp:ListItem Value="Fluido" Text="Fluido" />--%>
                        </asp:DropDownList></div>
                                           

                                     <div  style="float:left;margin-top:20px; clear:both"> <asp:FileUpload ID="fuUploader" runat="server" /></div>   
                                        <div  style="float:left;margin-left:20px;margin-top:10px"> <asp:Button ID="btnUpload" runat="server" Text="Upload" CssClass="btn" OnClientClick="return funvalidate();" OnClick="btnUpload_Click" /></div>
                                   
                                        </div>
                                     
                                    
                                </div>
                                <div id="menu1" class="tab-pane fade">
                                        <div style="width:950px;margin:0px auto;padding-top:40px;height:450px">
                                              <div style="float:left;margin-top:5px"> Service Line:</div>
                                            <div  style="float:left;margin-left:20px"> <asp:DropDownList ID="ddl_SL_1" runat="server" Width="130" Style="font-family: Calibri; font-size: small">
                               <asp:ListItem Value="EAS" Text="EAS" />
                              <asp:ListItem Value="SAP" Text="SAP" />
                              <asp:ListItem Value="ORC" Text="ORC" />
                            <asp:ListItem Value="ECAS" Text="ECAS" />
                            <asp:ListItem Value="EAIS" Text="EAIS" />
                             <%--<asp:ListItem Value="Fluido" Text="Fluido" />--%>  
                        </asp:DropDownList></div>
                                                  <div style="float:left;margin-left:20px">
    <asp:Button  ID="btnSLDownload" runat="server" Text="Download" onClientClick="gif1()"  CssClass="btn" OnClick="btn_ServiceLineBPlan_Download_Click" /> </div>
                           <div style="float:left;margin-left:20px;margin-top:5px">   <div id="gif1" style="float:left;display:none">
                                  <img src="../loadimg/ajax-loader.gif" /></div>   
                                     <asp:Label ID="lblError" runat="server" Text=""></asp:Label>
                                                                                             </div>
                                </div>
                                   
                                    </div>
                                 <div id="menu2" class="tab-pane fade in active">
                                     <div style="width:950px;margin:0px auto;padding-top:40px;height:300px">
                                         <div style="float:left">
                                              <div style="float:left;margin-top:5px"> Service Line:</div>
                                            <div  style="float:left;margin-left:20px"> <asp:DropDownList ID="ddl_Internal" AutoPostBack="true" runat="server" Width="130" Style="font-family: Calibri; font-size: small" OnSelectedIndexChanged="ddl_Internal_SelectedIndexChanged">
                               
                             <%--<asp:ListItem Value="Fluido" Text="Fluido" />--%>  
                        </asp:DropDownList></div>
                                                   
<div id="div_vertical" runat="server" style="float:left;margin-left:20px"> Vertical :
<asp:DropDownList ID="ddlVertical" runat="server" Width="110" Style="font-family: Calibri; font-size: small">                   
                                      </asp:DropDownList> </div>
<div style="float:left;margin-left:20px">                     Type :
                                 <asp:DropDownList ID="ddl_Type" runat="server" Width="100" Style="font-family: Calibri; font-size: small">                                        
                                    <%-- <asp:ListItem>Overall</asp:ListItem> 
                                     <asp:ListItem >Digital</asp:ListItem>--%>    
                                      <asp:ListItem>Standalone</asp:ListItem> 
                                     <asp:ListItem >Consolidated</asp:ListItem>
                                      </asp:DropDownList> </div>  
                  </div>
                                         <div style="float:left;clear:both;margin-top:50px;padding-left:180px">
                                         <div style="float:left">
                                <asp:Button ID="Button2" runat="server" Text="Download" onClientClick="gif()"  CssClass="btn" OnClick="lnkSPOReport_Click" /></div>
                         <div style="float:left;margin-top:5px;margin-left:20px">  <div id="gif" style="float:left;display:none">
                                  <img src="../loadimg/ajax-loader.gif" />
                             </div>
                             </div>
                                                                                             </div>
                                            
                                </div>

                                      <btn class="btn" id="myBtn"   onclick="OpenModal(); return false;"> Download Core Digital Split</btn>
                               
                            </div>
                                <div id="menu3" class="tab-pane fade">
                                        <div style="width:950px;margin:0px auto;padding-top:40px;height:450px">
                                              <div style="float:left;margin-top:5px"> Service Line:</div>
                                            <div  style="float:left;margin-left:20px"> <asp:DropDownList ID="ddlInput" runat="server" Width="130" Style="font-family: Calibri; font-size: small">
                             
                              <asp:ListItem Value="SAP" Text="SAP" />
                              <asp:ListItem Value="ORC" Text="ORC" />
                            <asp:ListItem Value="ECAS" Text="ECAS" />
                            <asp:ListItem Value="EAIS" Text="EAIS" />
                             <%--<asp:ListItem Value="Fluido" Text="Fluido" />--%>  
                        </asp:DropDownList></div>
                                                  <div style="float:left;margin-left:20px">
    <asp:Button  ID="Button1" runat="server" Text="Download" onClientClick="gif2()"  CssClass="btn" OnClick="btn_ServiceLineBPlan_Input_Click" /> </div>
                           <div style="float:left;margin-left:20px;margin-top:5px">   <div id="gif2" style="float:left;display:none">
                                  <img src="../loadimg/ajax-loader.gif" /></div>   
                                     <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
                                                                                             </div>
                                </div>


                                  

                                    
                                    </div>

                                  <div id="menu4" class="tab-pane fade">
                                      <br />
                                      <asp:DropDownList runat="server" ID="ddlSLConsolidated">
                                          <asp:ListItem Text="EAS"  />
                                          <asp:ListItem Text="SAP" />
                                          <asp:ListItem Text="ORC" />
                                          <asp:ListItem Text="ECAS" />
                                          <asp:ListItem Text="EAIS" />
                                      </asp:DropDownList>
                                       <br /> 
                                      
                                        <br />
            <asp:Button CssClass="btn" ID="btnConsolidatedDownload" runat="server" Text="Download" OnClick="btnConsolidatedDownload_Click" />
                                  </div>
                        </div>
                    </div>
                   
        </div>

              
     </div>
    </form>


    <!-- The Modal -->
<div id="myModal" class="modal">

  <!-- Modal content -->
  <div class="modal-content">
    <div class="modal-header">
      <span class="close" onclick="closeModal();">&times;</span>
      <h2 style="margin:0px 0px 0px 0px;">Digital Core Split</h2>
    </div>
    <div class="modal-body">
       
        <iframe src="BE_CoreDigitalSplitReport.aspx" style="border:none"></iframe>
     
    </div>
    
  </div>

</div>

    <script>
        //MOdal

        function OpenModal() 
        {
                  
// Get the modal
            var modal = document.getElementById("myModal");
             modal.style.display = "block";

        }

        function closeModal() {
            var modal = document.getElementById("myModal");
             modal.style.display = "none";
        }

  

 

// When the user clicks anywhere outside of the modal, close it
        window.onclick = function (event) {
     var modal = document.getElementById("myModal");
  if (event.target == modal) {
    modal.style.display = "none";
  }
}
</script>

</body>
</html>

