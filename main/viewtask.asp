<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Expires=-1000%>
<%'------------------------ Include File Start Here -------------------------%>
<!-- #include file="../include/Authentication.asp"	-->
<!-- #include file="../common/LongDate.asp"			-->
<!-- #include file="../common/ShortDate.asp"			-->
<!-- #include file="../common/json_new.asp"	-->
<!-- #include file="../common/smtp_open.asp"	-->
<%'------------------------ Include Files end Here ---------------------------%>
<%
Dim TaskTitle,TaskDesc,TaskExpiryDate, TaskHandleBy, TaskCreatedBy,TaskStatus, TaskCreatedDate
Dim objHttp, url, jsonData, strResponse, JSON,i, MsgString, apiResponse ,tasks, TaskID
Dim TaskCount
Dim PeriodStartDate, PeriodEndDate, StartDate, EndDate, TmpDate, KeyWords
Dim Command
PeriodStartDate = Request.Form("PeriodStartDate")
PeriodEndDate = Request.Form("PeriodEndDate")
if PeriodStartDate = "" then
    TmpDate = Cdate("1 " & MonthName(Month(Date),true) & " " & Year(Date))
    PeriodStartDate = ShortDate(TmpDate)
    StartDate = LongDate(TmpDate)
else
    StartDate = LongDate(PeriodStartDate)
end if
if PeriodEndDate = "" then
    PeriodEndDate = LongDate(Date)
    EndDate = LongDate(Date+1)
else
    EndDate = LongDate(Cdate(LongDate(PeriodEndDate)) + 1)
end if


if Request.Form("myCommand") = "Search" then
Command = "Search"
KeyWords = Request.Form("keywords")
PeriodStartDate = Request.Form("PeriodStartDate")
PeriodEndDate = Request.Form("PeriodEndDate")
if PeriodStartDate = "" then
    TmpDate = Cdate("1 " & MonthName(Month(Date),true) & " " & Year(Date))
    PeriodStartDate = ShortDate(TmpDate)
    StartDate = LongDate(TmpDate)
else
    StartDate = LongDate(PeriodStartDate)
end if
if PeriodEndDate = "" then
    PeriodEndDate = LongDate(Date)
    EndDate = LongDate(Date+1)
else
    EndDate = LongDate(Cdate(LongDate(PeriodEndDate)) + 1)
end if

end if


' Assuming StartDate is in a DateTime format
' Convert StartDate to yyyy-mm-dd format
Dim formattedStartDate , formattedEndDate
formattedStartDate = Year(StartDate) & "-" & Right("0" & Month(StartDate), 2) & "-" & Right("0" & Day(StartDate), 2)
formattedEndDate = Year(EndDate) & "-" & Right("0" & Month(EndDate), 2) & "-" & Right("0" & Day(EndDate), 2)





%>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <meta name="description" content="">
    <meta name="author" content="">
    <title>View Task</title>
  
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
 <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
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
 .datepicker-container {
            display: flex;
            justify-content: flex-end;
 /* Aligns the date picker to the right */
            margin: 10px;
        }
        .datepicker {
            padding: 1px;
            font-size: 16px;
        }
 .search-icon {
            cursor: pointer;
            font-size: 25px;
            color: white;
	    margin-right:8px;
            vertical-align: middle;	
        }
        
     
#search-container {
            display: none;
 /* Initially hidden */
            margin-top: 8px;
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
<table width="100%" CELLPADDING="1" CELLSPACING="1" BORDER="0">
		<tr style="FONT-SIZE: 25px; FONT-WEIGHT: bolder; BACKGROUND-COLOR: #1b2021; COLOR: white;" align="left" height="50">
			<td align="center">Task Listing</td>
			<td align="right" width="40px">

            			<i class="fas fa-search search-icon" onclick="toggleSearch()"></i>

		        </td>
		</tr>
            </table>
</br>

<div class="container-fluid"  id="search-container">
    <div class="row align-items-center" style="padding:4px;">
      
        <div class="col-12 col-md-1 text-center">
            Keywords
        </div>
        
        <!-- Machine ID Input  -->
        <div class="col-12 col-md-2 text-center">
            <input name="keywords" value="" type="text" class="form-control text-center" onfocus="select()">
        </div>

        <!-- Period Label -->
        <div class="col-12 col-md-1 text-center" style="padding:10px;">
            Period
        </div>

        <!-- Period Start and End Date Inputs -->
        <div class="col-12 col-md-4">
            <div class="row">
                <div class="col-5">
                    <input type="date" name="PeriodStartDate" id="datepicker1" class="form-control text-center" value="<%=formattedStartDate%>">
                </div>
                <div class="col-1 text-center d-flex align-items-center justify-content-center">
                    to
                </div>
                <div class="col-5">
                    <input type="date" name="PeriodEnddate" id="datepicker2" class="form-control text-center" value="<%=formattedEndDate%>">
                </div>
            </div>
        </div>

        <!-- Search Button -->
        <div class="col-12 col-md-2 text-center" style="padding:10px;">
            <input class="btn btn-primary" onclick="search_onclick()" name="btnSearch" type="button" value="Search">
        </div>
    </div>
</div>
            <!-- Page Content -->
<div style="overflow-x: auto; width: 100%;">
        <table class="table">
  <thead class="table-dark">
    <tr>
      <th scope="col">No</th>
      <th scope="col">TaskTitle</th>
      <th scope="col">TaskDesc</th>
      <th scope="col">TaskStatus</th>
      <th scope="col">Task HandleBy</th>
      <th scope="col">CompletedBy</th>
    </tr>
  </thead>

<%
if Command <> "" then

TaskID = "SEARCHDATA"
jsonData = ""
jsonData = jsonData & "{"
jsonData = jsonData & """TaskID"": """ & TaskID & ""","
jsonData = jsonData & """KeyWords"": """ & KeyWords & ""","
jsonData = jsonData & """StartDate"": """ & StartDate & ""","
jsonData = jsonData & """EndDate"": """ & EndDate & """"
jsonData = jsonData & "}"

else
TaskID = "ALLDATA"
jsonData = ""
jsonData = jsonData & "{"
jsonData = jsonData & """TaskID"": """ & TaskID & ""","
jsonData = jsonData & """StartDate"": """ & StartDate & ""","
jsonData = jsonData & """EndDate"": """ & EndDate & """"
jsonData = jsonData & "}"

end if

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
TaskStatus = tasks(i)("TaskStatus")
TaskID = tasks(i)("TaskID")
TaskCount = TaskCount + 1


%>

  <tbody>
     <tr onclick="View('<%=TaskID%>')">
      <th scope="row"><%=Taskcount%></th>
      <td><%=TaskTitle%></td>
      <td><%=TaskDesc%></td>
      <td><%=TaskStatus%></td>
      <td><%=TaskHandleBy%></td>
      <td><%=TaskExpiryDate%></td>
    </tr>
  </tbody>
<%Next%>

</table>
</div>
<div class="text-center">
           <a <button type="submit" class="btn btn-primary" href="javascript:goBack()">Back</button> </a>
        </div>
        </div>		
    </div>
    <!-- Wrapper -->
	
    <br><br>	


    <!-- Bootstrap core JavaScript -->
    <script src="../vendor/jquery/jquery.min.js"></script>
    <script src="../vendor/bootstrap/js/bootstrap.bundle.min.js"></script>

    <!-- Menu Toggle Script -->
    <script>
        $("#menu-toggle").click(function(e) {
            e.preventDefault();
            $("#wrapper").toggleClass("toggled");
        });

	function View(TaskID) {
            document.frmasp.target = "_self"
            document.frmasp.action = "managetask.asp?tid=" + TaskID
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
	function toggleSearch() {

        var searchContainer = document.getElementById('search-container');

        if (searchContainer.style.display === "none" || searchContainer.style.display === "") {

            searchContainer.style.display = "block";

        } else {

            searchContainer.style.display = "none";

        }

    }

   function search_onclick() {
	
        	document.frmasp.myCommand.value = "Search";
        	document.frmasp.target = "_self"
        	document.frmasp.action = "viewtask.asp"
        	document.frmasp.submit()
	
    }
    </script>
</form>
</body>
</html>
<%'------------------------ Close Connection Here ---------------------------%>

<%'------------------------ Include File end Here ---------------------------%>


