<%
Dim RSCommand
Set RSCommand=Server.CreateObject("ADODB.Recordset")
sqlstr = "SELECT CommandName, CommandName_New FROM SMV_Command WHERE (MachineUniqueID = '000000') AND (CommandName_New <> '')"
Set RSCommand = optObj.Execute(sqlstr)

Function Get_CommandName_New(CommandName)
	Dim CommandName_New
	CommandName_New = CommandName

	if RSCommand.EOF = False then
		RSCommand.Find "CommandName = '" & Replace(CommandName,"'","''") & "'"
		if RSCommand.EOF = False then
			CommandName_New = RSCommand("CommandName_New")
		end if

		RSCommand.MoveFirst
	end if

	Get_CommandName_New = CommandName_New
End Function

Sub UpdateCommand(ByVal MachineUniqueID, ByVal DestinationID)
    Dim NotificationResponse

    sqlstr = "DELETE FROM _NotificationCommands " & _
	     "WHERE (ClientProfileID = " & ClientProfileID & ") " & _
             "AND (DestinationID = '" & Replace(DestinationID,"'","''") & "') " & _
             "AND (CommandName = '" & Replace(CommandName,"'","''") & "') " & _
             "AND (CommandStatus = 'Pending')"
    optObj.Execute(sqlstr)

    if PayloadType <> "" then
	NotificationResponse = SendNotification(DestinationID, PayloadType, Payload)
	
	sqlstr = "INSERT INTO _NotificationCommands (" & _
                 "ClientProfileID, SourceID, MachineUniqueID, DestinationID, PayloadType, Payload, CommandName, CommandStartDate, CommandCompletedDate, CommandStatus, UpdatedBy, UpdatedDate, CreatedBy, CreatedDate" & _
                 ") VALUES (" & _
                 "" & ClientProfileID & ", 'Vengo', '" & Replace(MachineUniqueID,"'","''") & "', '" & Replace(DestinationID,"'","''") & "', '" & Replace(PayloadType,"'","''") & "', '" & Replace(Payload,"'","''") & "', '" & Replace(CommandName,"'","''") & "', '" & now & "', NULL, 'Sent', '', NULL, 'Vengo', '" & now & "'" & _
		 ")"
	optObj.Execute(sqlstr)
    end if
End Sub

%>