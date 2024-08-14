<%
	Dim conn, connString

	' Set up the connection string
	connString = "Provider=SQLOLEDB;Data Source=EC2AMAZ-VS3MRJF;Initial Catalog=TechKW;User Id=sa;Password=s@db123;"

	' Create and open the database connection
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open connString


	
%>
