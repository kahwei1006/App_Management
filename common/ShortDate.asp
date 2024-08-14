<%
Function ShortDate(byval InDate)
''*********************************************************************************	
''Assume: InDate is date type, return dd/mm/yyyy
''*********************************************************************************	
	ShortDate=""
	if IsDate(InDate)=false then
		ShortDate = ""
	else
		if Day(InDate)<10 then ShortDate = ShortDate & "0"
		ShortDate = ShortDate & Day(InDate) & "/"
		if Month(InDate)<10 then ShortDate = ShortDate & "0"
		ShortDate = ShortDate & Month(InDate) & "/" & Year(InDate)				
	end if
End Function
%>