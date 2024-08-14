<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Expires=-1000%>
<%'------------------------ Include File Start Here -------------------------%>
<!-- #include file="../include/Authentication.asp"			-->
<!-- #include file="../include/token.asp"	-->
<!-- #include file="../common/db_openconn.asp"	-->
<!-- #include file="../common/LongDate.asp"			-->
<!-- #include file="../common/ShortDate.asp"			-->
<!-- #include file="../common/aspJSON1.19.asp"			-->
<!-- #include file="../common/JSON_Helper.asp"			-->
<%'------------------------ Include Files end Here ---------------------------%>
<%
Dim ASPFile,TmpASPFileArray
TmpASPFileArray = split(Request.ServerVariables("URL"),"/")
ASPFile = Left(TmpASPFileArray(ubound(TmpASPFileArray)),len(TmpASPFileArray(ubound(TmpASPFileArray)))-4)

Dim RS, RS1, RS2, RS3, sqlstr, i

Dim ResponseString
Dim TaskID
Dim lngBytesCount, jsonString
Dim oJSON
Dim RoomType, RoomDesc,RoomPrice, BookingID, KeyWords, StartDate, EndDate

Function FormatDateToISO(inputDate)
    Dim year, month, day

    ' Ensure the input is a date
    If IsDate(inputDate) Then
        ' Extract year, month, and day
        year = DatePart("yyyy", inputDate)
        month = Right("0" & DatePart("m", inputDate), 2) ' Pad month with leading zero if needed
        day = Right("0" & DatePart("d", inputDate), 2) ' Pad day with leading zero if needed

        ' Return formatted date as yyyy-mm-dd
        FormatDateToISO = year & "-" & month & "-" & day
    Else
        ' Return an empty string or handle the error as needed
        FormatDateToISO = ""
    End If
End Function

If Request.TotalBytes > 0 Then
	lngBytesCount = Request.TotalBytes
	jsonString = BytesToStr(Request.BinaryRead(lngBytesCount))
	Set oJSON = New aspJSON
	oJSON.loadJSON(jsonstring)

	on error resume next
	TaskID = oJSON.data("TaskID")
	KeyWords = oJSON.data("KeyWords")
	StartDate = oJSON.data("StartDate")
	EndDate = oJSON.data("EndDate")
end if

if TaskID <> "" then

	if TaskID = "SUMTask" then

		sqlstr = "SELECT COUNT(*) AS PendingTaskCount " & _
			 "FROM TaskInfo " & _
			 "WHERE TaskStatus = 'pending' "

	else

		sqlstr = "SELECT TaskID,TaskDesc, TaskStatus, TaskCreatedDate, TaskCreatedBy, " & _
	 		"TaskCompletedDate,TaskCompletedBy, TaskExpiryDate, TaskTitle, TaskHandleBy " & _
	 		"FROM TaskInfo " 

		if TaskID = "ALLDATA" then
			sqlstr = sqlstr & "WHERE TaskCreatedDate >= '" & (Replace(StartDate,"'","''")) & "' " 	
			sqlstr = sqlstr & "AND TaskCreatedDate < '" & (Replace(EndDate,"'","''")) & "' " 
			sqlstr = sqlstr & "Order By TaskStatus desc, TaskID desc"
		elseif TaskID = "SEARCHDATA" then

			sqlstr = sqlstr & "WHERE TaskCreatedDate >= '" & (Replace(StartDate,"'","''")) & "' " 	
			sqlstr = sqlstr & "AND TaskCreatedDate < '" & (Replace(EndDate,"'","''")) & "' " 
			sqlstr = sqlstr & "AND TaskTitle LIKE '%" & Replace(KeyWords,"'","''") & "%' "	
			sqlstr = sqlstr & "OR TaskDesc LIKE '%" & Replace(KeyWords,"'","''") & "%' "
			sqlstr = sqlstr & "Order By TaskStatus desc"
		else
			sqlstr = sqlstr & "WHERE TaskID = '" & Replace(TaskID,"'","''") & "' " 
			
		end if

	end if

SET RS = Conn.Execute(sqlstr)

if RS.EOF then
		ResponseString = "[{" & vbCrLf
		ResponseString = ResponseString & """response"": ""403""," & vbCrLf
		ResponseString = ResponseString & """message"": ""Record Not Found!""" & vbCrLf	 
		ResponseString = ResponseString & "}]" 
	else


' Start the JSON array
    ResponseString = "["

    ' Initialize a flag to track if the current record is not the first one
    Dim FirstRecord
    FirstRecord = True

    Do While Not RS.EOF
       
        ' Add a comma if this is not the first record
        If Not FirstRecord Then
            ResponseString = ResponseString & ","
        End If
        FirstRecord = False

        ' Append the JSON object

	if TaskID = "SUMTask" then

		ResponseString = ResponseString & "{" & vbCrLf
		ResponseString = ResponseString & """response"": ""200""," & vbCrLf
		ResponseString = ResponseString & """PendingTaskCount"": """ & RS("PendingTaskCount") & """" & vbCrLf	 
	 
		ResponseString = ResponseString & "}" 
	else
       		ResponseString = ResponseString & "{" & vbCrLf
		ResponseString = ResponseString & """response"": ""200""," & vbCrLf
		ResponseString = ResponseString & """message"": ""Success""," & vbCrLf	
		ResponseString = ResponseString & """TaskID"": """ & RS("TaskID") & """," & vbCrLf 
		ResponseString = ResponseString & """TaskDesc"": """ & RS("TaskDesc") & """," & vbCrLf
		ResponseString = ResponseString & """TaskStatus"": """ & RS("TaskStatus") & """," & vbCrLf
		ResponseString = ResponseString & """TaskCreatedDate"": """ & RS("TaskCreatedDate") & """," & vbCrLf	 
		ResponseString = ResponseString & """TaskCreatedBy"": """ & RS("TaskCreatedBy") & """," & vbCrLf
		ResponseString = ResponseString & """TaskCompletedDate"": """ & RS("TaskCompletedDate") & """," & vbCrLf
		ResponseString = ResponseString & """TaskExpiryDate"": """ & FormatDateToISO(RS("TaskExpiryDate")) & """," & vbCrLf
		ResponseString = ResponseString & """TaskTitle"": """ & RS("TaskTitle") & """," & vbCrLf
		ResponseString = ResponseString & """TaskHandleBy"": """ & RS("TaskHandleBy") & """" & vbCrLf	 
	 
		ResponseString = ResponseString & "}" 
	end if
        ' Move to the next record
        RS.MoveNext
    Loop

    ' End the JSON array
    ResponseString = ResponseString & "]"



	end if
else

responseString = "[{" & vbCrLf
		ResponseString = ResponseString & """response"": ""403""," & vbCrLf
		ResponseString = ResponseString & """message"": ""Record Not Found!""" & vbCrLf	 
		ResponseString = ResponseString & "}]" 

end if


response.write ResponseString



%>	 
 
<%'------------------------ Close Connection Here ---------------------------%>
<!-- #include file="../common/db_closeconn.asp" -->	
<%'------------------------ Include File end Here ---------------------------%>
