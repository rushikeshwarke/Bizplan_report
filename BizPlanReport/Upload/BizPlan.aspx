<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="BizPlan.aspx.cs" Inherits="BEBIZ4.Upload.BizPlan" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
    <script src="../bootstrap-3.3.6-dist/jquery-1.10.2.js"></script>
    <link href="bootstrap-3.3.6-dist/css/bootstrap.css" rel="stylesheet" />
    <script src="../bootstrap-3.3.6-dist/js/bootstrap.min.js"></script>
    <link href="../bootstrap-3.3.6-dist/css/bootstrap.css" rel="stylesheet" />
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
}
    </style>

 
</head>
<body>
   <form name="form1" id="form1">
        <div class="container" style="margin-top: 30px;margin-bottom: 30px;"  >
                    <div class="well" >
                        <div style="margin-top:10px;">
                            <ul class="nav nav-tabs" id="myTabs">
                                <li class="active"><a data-toggle="tab" href="#home">Upload </a></li>
                                <li><a data-toggle="tab" href="#menu1">Biz Plan CDO</a></li>
                                 <li><a data-toggle="tab" href="#menu2">Biz Plan Internal</a></li>
                            </ul>
                            <div class="tab-content">
                                <div id="home" class="tab-pane fade in active">
                                    <div style="width:950px;margin:0px auto;padding-top:40px">
                                        
                                            
                                        </div>
                                     
                                    
                                </div>
                                <div id="menu1" class="tab-pane fade">
                                     
                                </div>
                                 <div id="menu2" class="tab-pane fade">
                                     
                                </div>
                            </div>
                        </div>
                    </div>
                   
        </div>
       
    </form>
    
</body>
</html>
