<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Expires=-1000%>
<%'------------------------ Include File Start Here -------------------------%>
<!-- #include file="../include/Authentication.asp"	-->
<!-- #include file="../common/json_new.asp"	-->
<!-- #include file="../common/smtp_open.asp"	-->
<%'------------------------ Include Files end Here ---------------------------%>
<%
Dim TaskTitle,TaskDesc,TaskExpiryDate, TaskHandleBy, TaskCreatedBy,TaskStatus, TaskCreatedDate
Dim objHttp, url, jsonData, strResponse, JSON,i, MsgString, apiResponse ,tasks, TaskID
Dim TaskCount, PendingTaskCount
%>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <meta name="description" content="">
    <meta name="author" content="">
    <title>Home</title>
  
    <!-- Bootstrap core CSS -->
    <link href="../vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">
    <!-- Custom styles for this template -->
    <link href="../css/simple-sidebar.css" rel="stylesheet">
    <!-- Custom styles for app -->
    <link href="../css/style.css" rel="stylesheet">
	 <!-- Vendor CSS Files -->
  <link href="../assets/vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">
  <link href="../assets/vendor/bootstrap-icons/bootstrap-icons.css" rel="stylesheet">
  <link href="../assets/vendor/boxicons/css/boxicons.min.css" rel="stylesheet">
  <link href="../assets/vendor/quill/quill.snow.css" rel="stylesheet">
  <link href="../assets/vendor/quill/quill.bubble.css" rel="stylesheet">
  <link href="../assets/vendor/remixicon/remixicon.css" rel="stylesheet">
  <link href="../assets/vendor/simple-datatables/style.css" rel="stylesheet">

  <!-- Template Main CSS File -->
  <link href="../assets/css/style.css" rel="stylesheet">

<style>
    body {
      font-family: Arial, sans-serif;
    }
    .submit-form-button {
      padding: 10px 20px;
      background-color: #4CAF50;
      color: #fff;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }
    .delete-form-button {
      padding: 10px 20px;
      background-color: red;
      color: #fff;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }
  </style>
<script language="JavaScript">
 
 
</script>
</head>
<body>
<form name="frmasp" action="" method="POST">
<input type="hidden" name="myCommand" value>   

    <!-- Wrapper -->
    <div class="d-flex" id="wrapper">
        <!-- #include file="../include/sidebar.asp"			-->

        <div id="page-content-wrapper">
            <!-- #include file="../include/topbar.asp"			-->

            <!-- Page Content -->
            <table width="100%" CELLPADDING="1" CELLSPACING="1" BORDER="0">
		<tr style="FONT-SIZE: 25px; FONT-WEIGHT: bolder; BACKGROUND-COLOR: #1b2021; COLOR: white;" align="left" height="50">
			<td align="center">HOME</td>
		</tr>
            </table>
            <br>

            <!-- Page Content -->


    <div class="pagetitle" style="margin:14px">
      <h1>Dashboard</h1>
      <nav>
        <ol class="breadcrumb">
          <li class="breadcrumb-item">Home</li>
          <li class="breadcrumb-item active">Dashboard</li>
        </ol>
      </nav>
    </div><!-- End Page Title -->

	<section class="section dashboard">
	<div class="row" style="margin:2px">

        <!-- Left side columns -->
        <div class="col-lg-12">
          <div class="row">
	<%
TaskID = "SUMTask"
jsonData = ""
jsonData = jsonData & "{"
jsonData = jsonData & """TaskID"": """ & TaskID & """"
jsonData = jsonData & "}"

url = Session("URL") & "/App_Management/API/viewtask.asp"
' Create an instance of the HTTP object
Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Open a connection to the API
objHttp.Open "POST", url, False

' Set the request headers
objHttp.setRequestHeader "Content-Type", "application/json"
objHttp.SetRequestHeader "x-api-key", Session("APIKEY")

' Send the JSON data to the API
objHttp.Send jsonData

' Get the response from the API
strResponse = objHttp.responseText

Set JSON = New aspJSON

' Parse the JSON string
Call JSON.loadJSON(strResponse)

' Get the array of objects from the parsed JSON

Set tasks = JSON.getData() 

For i = 0 To tasks.count - 1

PendingTaskCount = tasks(i)("PendingTaskCount")

next

%>

            <div class="col-xxl-4 col-md-6">
              <div class="card info-card ">
  		<a href="viewtask.asp"><div class="card-body">
                  <h5 class="card-title">Outstanding Task</span></h5>

                  <div class="d-flex align-items-center">
                    <div class="card-icon rounded-circle d-flex align-items-center justify-content-center" style = " background-color: #f6f6fe;color: red;">
                      <i class="bi bi-exclamation-square-fill"></i>
                    </div>
                    <div class="ps-3">

                     <h6><%=PendingTaskCount%> Task</h6>
                     <a href="viewtask.asp"> <span class="text-success small pt-1 fw-bold">Click here to view details</span> </a>

                    </div>
                  </div>
                </div>

              </div> 
            </div> </a>

      </div>
    </section>
        </div>		
    </div>
    <!-- Wrapper -->

    <br><br><br><br><br>	


    <!-- Bootstrap core JavaScript -->
    <script src="../vendor/jquery/jquery.min.js"></script>
    <script src="../vendor/bootstrap/js/bootstrap.bundle.min.js"></script>

    <!-- Menu Toggle Script -->
    <script>
        $("#menu-toggle").click(function(e) {
            e.preventDefault();
            $("#wrapper").toggleClass("toggled");
        });
    </script>
</form>
</body>
</html>
<%'------------------------ Close Connection Here ---------------------------%>

<%'------------------------ Include File end Here ---------------------------%>


