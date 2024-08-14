<%
if Request.ServerVariables("HTTP_X_API_KEY") = "" then
	response.ContentType = "text/plain; charset=utf-8"
	response.write "API Key is not provided"
	response.end
elseif Request.ServerVariables("HTTP_X_API_KEY") <> Session("APIKEY") then
	response.ContentType = "text/plain; charset=utf-8"
	response.write "Unauthorized client"
	response.end
end if
%>
 