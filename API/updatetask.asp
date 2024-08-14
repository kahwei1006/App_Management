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
Dim TaskStatus,TaskID
Dim ResponseString
Dim lngBytesCount, jsonString
Dim oJSON
Dim RoomType,RoomPrice, GuessName, GuessEmail, GuessContact, CheckInDate, CheckOutDate, BookingID , CreatedDate
If Request.TotalBytes > 0 Then
	lngBytesCount = Request.TotalBytes
	jsonString = BytesToStr(Request.BinaryRead(lngBytesCount))
	Set oJSON = New aspJSON
	oJSON.loadJSON(jsonstring)

	on error resume next
	TaskID = oJSON.data("TaskID")
	TaskStatus = oJSON.data("TaskStatus")
end if

CreatedDate = now

if TaskID <> "" then

sqlstr = "UPDATE TaskInfo " & _
	 "SET TaskStatus = '" & Replace(TaskStatus,"'","''") & "' " & _
	 "WHERE TaskID = '" & Replace(TaskID,"'","''") & "' "

Conn.Execute(sqlstr)
		
		ResponseString = "["
		ResponseString = ResponseString & "{" & vbCrLf
		ResponseString = ResponseString & """response"": ""200""," & vbCrLf
		ResponseString = ResponseString & """message"": ""Success""" & vbCrLf	 
		ResponseString = ResponseString & "}" 
  		ResponseString = ResponseString & "]"
else

responseString = "[{" & vbCrLf
		ResponseString = ResponseString & """response"": ""403""," & vbCrLf
		ResponseString = ResponseString & """message"": ""Update Not Successful!""" & vbCrLf	 
		ResponseString = ResponseString & "}]" 


end if


response.write ResponseString



%>	 
 
<%'------------------------ Close Connection Here ---------------------------%>
<!-- #include file="../common/db_closeconn.asp" -->	
<%'------------------------ Include File end Here ---------------------------%>
