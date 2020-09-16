<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>

  </head>
  <body>

  <!--#include file="dbpath.asp"-->
 <% 
 
 ForelCheck = 0
 WillianCheck = 0
 ForelStatus = "Forel Scan Over an Hour "
 WillianStatus = "Willian Scan Over an Hour "
 
 timeFifteen = DateAdd("n",-15, now)
 timeHalfHour = DateAdd("n",-30, now)
 timeHour = DateAdd("h",-1, now)
 
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Top 1 * FROM X_BARCODEGA WHERE DEPT = 'Forel' ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

if timeHour > rs("Datetime") then
else
	if timeHalfHour > rs("DateTime") then
		ForelCheck = 1
		ForelStatus = "Forel Scan Over an Half an Hour "
	else
		if timeFifteen > rs("DateTime") then
			ForelCheck = 2
			ForelStatus = "Forel Scan Over Fifteen Minutes "
		else
			ForelCheck = 3
			ForelStatus = "Forel Scan Within Fifteen Minutes - - GOOD "
		end if
	end if

end if

rs.close
set rs=nothing


Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Top 1 * FROM X_BARCODEGA WHERE DEPT = 'Willian' ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


if timeHour > rs("Datetime") then
else
	if timeHalfHour > rs("DateTime") then
		WillianCheck = 1
		WillianStatus = "Willian Scan Over an Half an Hour "
	else
		if timeFifteen > rs("DateTime") then
			WillianCheck = 2
			WillianStatus = "Willian Scan Over Fifteen Minutes "
		else
			WillianCheck = 3
			WillianStatus = "Willian Scan Within Fifteen Minutes - - GOOD"
		end if
	end if

end if

rs.close
set rs=nothing


if (ForelCheck + WillianCheck) < 6 and (ForelCheck + WillianCheck) > 0 then
' Send the email
'*******************

Set Mail = CreateObject("CDO.Message")

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="smtp.gmail.com"
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 300

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="quest.windows@gmail.com" 'You can also use you email address that's setup through google apps.
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="gNpXe9fewe9N"

Mail.Configuration.Fields.Update

if (WillianCheck + ForelCheck) = 0 then
	Mail.Subject="Glass Line Notification - Both Machines OVER 1 HOUR"
else
	Mail.Subject="Glass Line Notification!"
end if



Mail.From="quest.windows@gmail.com" 'This has to be an actual email address or an alias that's setup on the gmail account you used above
Mail.To="mbernholtz@questwindows.com, mdungo@questwindows.com, aramirez@questwindows.com" 'TEST
'Mail.Bcc="someoneelse@somedomain.com" 'Carbon Copy
'Mail.Cc="someoneelse2@somedomain.com" 'Blind Carbon Copy


Mail.HTMLBody="Notification about Glass Lines:" & " <br> " & "Forel: " & ForelStatus & " <br> " & "Willian: " & WillianStatus 


Mail.Send
Set Mail = Nothing
'***********************


%>     
<p> Notification Sent out (Forel/Willian) line Notification of over 15 Minutes </p>

<%
end if

DBConnection.close
Set DBConnection = nothing
%>


<p> End Program </p>
<p> 
<br>
<% response.write ForelStatus%>
</p>
<p> 
<br>
<% response.write WillianStatus%>
</p>
</body>
</html>
