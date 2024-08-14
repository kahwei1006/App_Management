<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Expires=-1000%>
<% Response.Charset = "UTF-8" %>
<!-- #include file="../include/Authentication.asp"	-->
<!-- #include file="../common/json_new.asp"	-->
<!-- #include file="../common/smtp_open.asp"	-->
<%'------------------------ Include Files end Here ---------------------------%>

<%
Dim TaskTitle,TaskDesc,TaskExpiryDate, TaskID,AllocateTo, CreatedBy
Dim objHttp, url, jsonData, strResponse, JSON,i, MsgString, apiResponse ,tasks

if Request.Form("myCommand") = "Save" then

TaskTitle = Request.Form("tasktitle")
TaskDesc = Request.Form("taskDesc")
TaskExpiryDate = Request.Form("taskExpiryDate")
AllocateTo = Request.Form("allocateto")
CreatedBy = Request.Form("CreatedBy")

Set objHttp = Nothing
Set JSON = Nothing
 jsonData = jsonData & "{" & vbCrLf
        jsonData = jsonData & """TaskTitle"": """ & Tasktitle & """," & vbCrLf
	jsonData = jsonData & """TaskDesc"": """ & TaskDesc & """," & vbCrLf
	jsonData = jsonData & """CreatedBy"": """ & CreatedBy & """," & vbCrLf
	jsonData = jsonData & """AllocateTo"": """ & AllocateTo & """," & vbCrLf
        jsonData = jsonData & """TaskExpiryDate"": """ & TaskExpiryDate & """" & vbCrLf
        jsonData = jsonData & "}"


url = Session("URL") & "/App_Management/API/createtask.asp"
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
TaskID = tasks(i)("TaskID")

Next

if apiResponse = "200" then

MsgString = "Hi ," & "<br><br>"
MsgString = MsgString & "There are a new task has been created and assigned to you." & "<br><br>"
MsgString = MsgString & "You can view your task by clicking the link below:" & "<br>"
MsgString = MsgString & "<a href=""" & Session("URL") & "/App_Management/main/managetask.asp?tid=" & Server.URLEncode(TaskID) & """>Click here to view your task</a>" & "<br><br>"

MsgString = MsgString & "Best Regards," & "<br>"
MsgString = MsgString & "Admin"

objEmail.From = "testwechat0001@gmail.com"
objEmail.To = AllocateTo
objEmail.Subject = "[Action Required] : New Task created by " & CreatedBy
objEmail.HTMLBody = MsgString

' Send the email
On Error Resume Next
objEmail.Send

end if
%>

 <script type="text/javascript">
        alert("Your task is created successfully!");
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
    <title>New Task</title>
  
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
    <!-- Wrapper -->
    <div class="d-flex" id="wrapper">
        <!-- #include file="../include/sidebar.asp"			-->

        <div id="page-content-wrapper">
            <!-- #include file="../include/topbar.asp"			-->
 <table width="100%" CELLPADDING="1" CELLSPACING="1" BORDER="0">
		<tr style="FONT-SIZE: 25px; FONT-WEIGHT: bolder; BACKGROUND-COLOR: #1b2021; COLOR: white;" align="left" height="50">
			<td align="center">Create New Task</td>
		</tr>
            </table>
            <br>

            <!-- Page Content -->
            <div class="container mt-5">
                <div class="card p-4">
                    <h2 class="text-center mb-4">Task Form</h2>
                    <form method="post" action="your_action_page.asp">
                        <div class="mb-3">
                            <label for="taskDesc" class="form-label">Task Title</label>
                            <input type="text" class="form-control" id="taskTitle" name="taskTitle" placeholder="Title" value>
                        </div>
                        <div class="mb-3">
                            <label for="taskType" class="form-label">Task Description</label>
                            <input type="text" class="form-control" id="taskDesc" name="taskDesc" placeholder="Description" value>
                        </div>
			<div class="mb-3">
                            <label for="taskType" class="form-label">Created By</label>
                            <input type="text" class="form-control" id="createdby" name="createdby" placeholder="Created By" value>
                        </div>
			<div class="mb-3">
                            <label for="taskType" class="form-label">Allocate To</label>
                            <input type="email" class="form-control" id="allocateto" name="allocateto" placeholder="Email address" value>
                        </div>
                        
                        <div class="mb-3">
                            <label for="taskExpiryDate" class="form-label">Completed By</label>
                            <input type="date" class="form-control" id="taskExpiryDate" name="taskExpiryDate" value>
                        </div>
                        <div class="text-center">
                            <button type="submit" class="btn btn-primary" onclick="update_onclick()">Submit</button>
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
	    			var isError = false
				//document.frmasp.taskTitle.value = trimSpace(document.frmasp.taskTitle.value)
				if (!isError) {
				 if (document.frmasp.taskTitle.value == "") {
            				isError = true
            				alert("This field is required!")
            				document.frmasp.taskTitle.focus();
        			 }
				}
				if (!isError) {
				 if (document.frmasp.taskDesc.value == "") {
            				isError = true
            				alert("This field is required!")
            				document.frmasp.taskDesc.focus();
        			}
				}
				if (!isError) {
				 if (document.frmasp.createdby.value == "") {
            				isError = true
            				alert("This field is required!")
            				document.frmasp.createdby.focus();
        			}
				}
				if (!isError) {
				 if (document.frmasp.allocateto.value == "") {
            				isError = true
            				alert("This field is required!")
            				document.frmasp.allocateto.focus();
        			}
				}
				if (!isError) {
				 if (document.frmasp.taskExpiryDate.value == "") {
            				isError = true
            				alert("This field is required!")
            				document.frmasp.taskExpiryDate.focus();
        			}
				}
				if (!isError) {
				document.frmasp.myCommand.value = "Save"
				document.frmasp.action = "newtask.asp"
            			document.frmasp.submit()
				}
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
