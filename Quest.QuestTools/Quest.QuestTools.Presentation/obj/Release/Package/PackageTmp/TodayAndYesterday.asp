<%
' This script is created to correct the Today and Yesterday function
' This code is an include to collect all the date and time variables to include both today and yesterday into a program.
' incorporates day, month, year, (checking for end of month/year, leap years, and end of week)
' Originally coded into each page in January 2014,  Turned into an Include in May 2014
' See TodayandYesterday.inc for the initial change and the old incorrect code it replaced
 
 
STAMPVAR = year(now) & " " & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now)
ccTime = hour(now) & ":" & minute(now) & ":" & second(now)
cDay = day(now)
cYesterday = cDay - 1
cMonth = month(now)
cMonthy = cMonth
cYear = year(now)
cYeary = cYear
chour = hour(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)
weekNumbery = weekNumber
lastweek = WeekNumber -1


'Go back a week if today is day 1 (Sunday)
if Weekday(now) = 1 then
weekNumbery = Weeknumber - 1
end if

' Go back a Month (or year) if today is the first of the month
If cDay = 1 then
	if cMonth = 2 OR cMonth = 4 OR cMonth = 6 OR cMonth = 8 OR cMonth = 9 OR cMonth = 11 OR cMonth = 1 then
	cYesterday = 31
	end if
	if cMonth = 5 OR cMonth = 7 OR cMonth = 10 OR cMonth = 12 then
	cYesterday = 30
	end if
	if cMonth = 3 then
		if cyear = 2016 OR cyear = 2020 OR cyear = 2024 OR cyear = 2028 OR cyear = 2032  then 
			cYesterday = 29
		else
			cYesterday = 28
		end if
	end if
		
	cMonthy = cMonth - 1	
		
	if cMonth = 1 then
	cMonthy = 12
	cYeary = cYear - 1
	end if
	
	

end if

'Call SetTestDate(cDay, cMonth, cYear)

%>
