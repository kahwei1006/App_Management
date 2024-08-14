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
Dim MemberName,MemberEmail
Dim lngBytesCount, jsonString
Dim oJSON
Dim TaskTitle,TaskDesc,TaskExpiryDate,CreatedDate, TaskID, AllocateTo,CreatedBy,TaskCompletedDate, TaskCompletedBy

If Request.TotalBytes > 0 Then
	lngBytesCount = Request.TotalBytes
	jsonString = BytesToStr(Request.BinaryRead(lngBytesCount))
	Set oJSON = New aspJSON
	oJSON.loadJSON(jsonstring)

	on error resume next
	TaskTitle = oJSON.data("TaskTitle")
	TaskDesc = oJSON.data("TaskDesc")
	TaskExpiryDate = oJSON.data("TaskExpiryDate")
	AllocateTo = oJSON.data("AllocateTo")
	CreatedBy = oJSON.data("CreatedBy")

end if

CreatedDate = now

if TaskTitle <> "" then

sqlstr = "Insert Into BookingInfo (GuessName,GuessContact, GuessEmailAddress,BookingDate, " & _
	 "CheckInDate, CheckOutdate,CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, RoomType, RoomPrice) " & _
	 "VALUES ('" & Replace(GuessName,"'","''") & "', '" & Replace(GuessContact,"'","''") & "' , '" & Replace(GuessEmail,"'","''") & "', '" & now & "' , " & _  
         "'"& CheckInDate &"', '"& CheckOutDate &"' , 'SYSTEM', '"&CreatedDate&"', 'SYSTEM' , '"&CreatedDate&"' , '" & Replace(RoomType,"'","''") & "' , '" & Replace(RoomPrice,"'","''") & "') " 


sqlstr = "INSERT INTO TaskInfo (TaskDesc, TaskStatus, TaskCreatedDate, TaskCreatedBy, TaskCompletedDate, TaskCompletedBy, TaskExpiryDate, TaskTitle, TaskHandleBy) " & _
	 "VALUES ('" & Replace(TaskDesc,"'","''") & "', 'Pending' , '" & Replace(CreatedDate,"'","''") & "', '" & Replace(CreatedBy,"'","''") & "', '" & Replace(TaskCompletedDate,"'","''") & "','" & Replace(TaskCompletedBy,"'","''") & "','" & Replace(TaskExpiryDate,"'","''") & "', '" & Replace(TaskTitle,"'","''") & "', '" & Replace(AllocateTo,"'","''") & "' )"
'response.write sqlstr

Conn.Execute(sqlstr)
		

Set RS = Conn.Execute("SELECT SCOPE_IDENTITY() AS TaskID")
    If Not RS.EOF Then
        TaskID = RS("TaskID")
        
    End If
		ResponseString = "["
		ResponseString = ResponseString & "{" & vbCrLf
		ResponseString = ResponseString & """response"": ""200""," & vbCrLf
		ResponseString = ResponseString & """message"": ""Success""," & vbCrLf	 
		ResponseString = ResponseString & """TaskID"": """ & TaskID & """" & vbCrLf
		ResponseString = ResponseString & "}" 
  		ResponseString = ResponseString & "]"

else

responseString = "[{" & vbCrLf
		ResponseString = ResponseString & """response"": ""403""," & vbCrLf
		ResponseString = ResponseString & """message"": ""Create Task Not Successul!""" & vbCrLf	 
		ResponseString = ResponseString & "}]" 

end if


response.write ResponseString



%>	 
 
<%'------------------------ Close Connection Here ---------------------------%>
<!-- #include file="../common/db_closeconn.asp" -->	
<%'------------------------ Include File end Here ---------------------------%>
