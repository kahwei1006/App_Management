<%
Function TranslatedDate(byval InDate)
	TranslatedDate = ReturnLongDate(InDate,true)
End Function
Function LongDate (byval InDate)
	LongDate = ReturnLongDate(InDate,false)
End Function
Function ReturnLongDate (byval InDate,byval Translated)
''*********************************************************************************	
''Assume: InDate is in 'dd/mm/yyyy' format or date type
''*********************************************************************************	
	dim dateArr
	
	ReturnLongDate=""
	if IsNull(InDate) then
		ReturnLongDate = ""
	elseif InDate="" then
		ReturnLongDate = ""
	else
		if TypeName(InDate)="Date"  or TypeName(InDate)="Field" then
			dateArr = Array (Day(InDate),Month(InDate),Year(InDate))
		else
			dateArr = Split(InDate,"/")
		end if
		
		if UBound(dateArr)=2 then
			ReturnLongDate = String(2-len(dateArr(0)),"0") & dateArr(0) & " " & ReturnMonthName(dateArr(1),Translated) & " " & dateArr(2)
		else
			ReturnLongDate=InDate
		end if
	end if
End Function
Function ReturnMonthName(byval InMonth,byval Translated)
	Dim MonthArr
	MonthArr = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
	if Cint(InMonth)>12 then
		ReturnMonthName = MonthArr(0)
	else
		ReturnMonthName = MonthArr(Cint(InMonth)-1)
	end if
	if Translated then ReturnMonthName=Translate(ReturnMonthName)
End Function
Function CShortDate(byval InDate)
''*********************************************************************************	
''Assume: InDate is in 'dd/mm/yyyy' format
''*********************************************************************************	
	CShortDate=LongDate(InDate)
	if IsDate(CShortDate) then
		CShortDate = CDate(CShortDate)
	else
		CShortDate = null
	end if
End Function
Function MonthLastDay(byval InDate,byval ReturnDayOnly)
	Dim NextMonth
	NextMonth = DateAdd("m",1,InDate)
	NextMonth = CDate("01 " & MonthName(Month(NextMonth),true) & " " & Year(NextMonth))
	MonthLastDay = DateAdd("d",-1,NextMonth)
	if ReturnDayOnly then MonthLastDay=Day(MonthLastDay)
End Function
Function sqlDate(byval InputDate)
	Dim TmpDate
	TmpDate = LongDate(InputDate)
	
	'mySql
	'sqlDate = "'" & Year(Cdate(TmpDate)) & "/" & Month(Cdate(TmpDate)) & "/" & Day(Cdate(TmpDate)) & "'"
	
	'MSSQL
	sqlDate = "'" & TmpDate & "'"

	'Access
	'sqlDate = "#" & TmpDate & "#"
End Function
Function Get_DateTimeString(InputDateTime)
    Get_DateTimeString = ""
    if NOT ISNULL(InputDateTime) then
        Get_DateTimeString = LongDate(InputDateTime) & " "
        if Hour(InputDateTime) < 10 then
            Get_DateTimeString = Get_DateTimeString & "0"
        end if
        Get_DateTimeString = Get_DateTimeString & Hour(InputDateTime) & ":"
        if Minute(InputDateTime) < 10 then
            Get_DateTimeString = Get_DateTimeString & "0"
        end if
        Get_DateTimeString = Get_DateTimeString & Minute(InputDateTime) & " "
        if Hour(InputDateTime) < 12 then
            Get_DateTimeString = Get_DateTimeString & "AM"
        else
            Get_DateTimeString = Get_DateTimeString & "PM"
        end if
    end if
End Function
%>