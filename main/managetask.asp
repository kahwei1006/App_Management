<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Expires=-1000%>
<% Response.Charset = "UTF-8" %>
<!-- #include file="../include/Authentication.asp"	-->
<!-- #include file="../common/json_new.asp"	-->
<!-- #include file="../common/smtp_open.asp"	-->
<%'------------------------ Include Files end Here ---------------------------%>

<%

Dim TaskTitle,TaskDesc,TaskExpiryDate, TaskHandleBy, TaskCreatedBy,TaskStatus, TaskCreatedDate,TaskCompletedDate
Dim objHttp, url, jsonData, strResponse, JSON,i, MsgString, apiResponse ,tasks, TaskID

TaskID = Request.QueryString("tid")

if Request.Form("myCommand") = "Save" then
TaskStatus = "Completed"
TaskID = Request.Form("TaskID")
TaskTitle = Request.Form("tasktitle")
TaskDesc = Request.Form("taskDesc")
TaskExpiryDate = Request.Form("taskExpiryDate")
TaskCreatedBy = Request.Form("CreatedBy")

Set objHttp = Nothing
Set JSON = Nothing
 jsonData = jsonData & "{" & vbCrLf
	jsonData = jsonData & """TaskID"": """ & TaskID & """," & vbCrLf
        jsonData = jsonData & """TaskStatus"": """ & TaskStatus & """" & vbCrLf
        jsonData = jsonData & "}"


url = Session("URL") & "/App_Management/API/updatetask.asp"
' Create an instance of the HTTP object
Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Open a connection to the API
objHttp.Open "POST", url, False

' Set the request headers
objHttp.setRequestHeader "Content-Type", "application/json"
objHttp.SetRequestHeader "x-api-key", Session("APIKEY")

' Send the JSON data to the API
objHttp.Send jsonData

strResponse = objHttp.responseText

Set JSON = New aspJSON

' Parse the JSON string
Call JSON.loadJSON(strResponse)

' Get the array of objects from the parsed JSON

Set tasks = JSON.getData() 

For i = 0 To tasks.count - 1
apiResponse = tasks(i)("response")

Next

%>

 <script type="text/javascript">
        alert("Your task is update successful!");
        window.location.href = "home.asp";
    </script>

<%

end if

%>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <meta name="description" content="">
    <meta name="author" content="">
    <title>Manage Task</title>
  
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
</head>
<body>
<form name="frmasp" action="" method="POST">
 <input type="hidden" name="myCommand" value="">
 <input type="hidden" name="TaskID" value="<%=TaskID%>">
    <!-- Wrapper -->
    <div class="d-flex" id="wrapper">
        <!-- #include file="../include/sidebar.asp"			-->
       
        <div id="page-content-wrapper">
            <!-- #include file="../include/topbar.asp"			-->
 <table width="100%" CELLPADDING="1" CELLSPACING="1" BORDER="0">
		<tr style="FONT-SIZE: 25px; FONT-WEIGHT: bolder; BACKGROUND-COLOR: #1b2021; COLOR: white;" align="left" height="50">
			<td align="center">Manage Your Task</td>
		</tr>
            </table>
            <br>
<%

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

TaskTitle = tasks(i)("TaskTitle")
TaskDesc = tasks(i)("TaskDesc")
TaskCreatedDate = tasks(i)("TaskCreatedDate")
TaskExpiryDate = tasks(i)("TaskExpiryDate")
TaskHandleBy = tasks(i)("TaskHandleBy")
TaskCreatedBy = tasks(i)("TaskCreatedBy")

Next



%>
            <!-- Page Content -->
            <div class="container mt-5">
                <div class="card p-4">
                    <h2 class="text-center mb-4">Task Form</h2>
                    <form method="post" action="your_action_page.asp">
                        <div class="mb-3">
                            <label for="taskDesc" class="form-label">Task Title</label>
                            <input type="text" class="form-control" id="taskTitle" name="taskTitle" value = "<%=TaskTitle%>" disabled>
                        </div>
                        <div class="mb-3">
                            <label for="taskType" class="form-label">Task Description</label>
                            <input type="text" class="form-control" id="taskDesc" name="taskDesc" value="<%=TaskDesc%>" disabled>
                        </div>
			<div class="mb-3">
                            <label for="taskType" class="form-label">Created By</label>
                            <input type="text" class="form-control" id="createdby" name="createdby" value= "<%=TaskCreatedBy%>" disabled>
                        </div>
                        <div class="mb-3">
                            <label for="taskExpiryDate" class="form-label">Task Expiry Date</label>
                            <input type="date" class="form-control" id="taskExpiryDate" name="taskExpiryDate" value = "<%=TaskExpiryDate%>" disabled>
                        </div>
                        <div class="text-center">
                            <button type="submit" class="btn btn-primary" onclick="update_onclick()">Complete Task</button>
			   <a <button type="submit" class="btn btn-primary" href="javascript:goBack()">Back</button> </a>
                        </div>
                    </form>
                </div>
            </div>
        </div>

    </div>
    <!-- Wrapper -->

    <!-- Bootstrap core JavaScript -->
    <script src="../vendor/jquery/jquery.min.js"></script>
    <script src="../vendor/bootstrap/js/bootstrap.bundle.min.js"></script>

    <!-- Menu Toggle Script -->
    <script>
        $("#menu-toggle").click(function(e) {
            e.preventDefault();
            $("#wrapper").toggleClass("toggled");
        });

	function update_onclick() {	
	    			event.preventDefault();
				document.frmasp.myCommand.value = "Save"
				document.frmasp.action = "managetask.asp"
            			document.frmasp.submit()
        		}
	function back_onclick() {	
	    			
				document.frmasp.action = "home.asp"
            			document.frmasp.submit()
        		}
	function goBack() {
        if (document.referrer && document.referrer !== window.location.href) {
            window.location.href = document.referrer;
        } else {
            // Optionally, use history.go(-2) if the last history entry is problematic
            history.go(-2);
        }
    }
    </script>
</form>
</body>
</html>
<%'------------------------ Close Connection Here ---------------------------%>

<%'------------------------ Include File end Here ---------------------------%>
