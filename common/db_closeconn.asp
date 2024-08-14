<%
	' Close the recordset
	RS.Close
	Set RS = Nothing

	' Close the connection
	conn.Close
	Set conn = Nothing
%>