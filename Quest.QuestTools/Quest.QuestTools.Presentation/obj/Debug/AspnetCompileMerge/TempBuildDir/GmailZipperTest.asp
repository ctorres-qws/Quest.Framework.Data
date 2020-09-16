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
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Top 1 * FROM PROZipperShearTest ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

LastWeek = DateAdd("d",-7, now)
LastMonth = DateAdd("m",-1, now)

if LastWeek > rs("DateTime") then
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

Mail.Subject="Zipper Report"
Mail.From="quest.windows@gmail.com" 'This has to be an actual email address or an alias that's setup on the gmail account you used above
Mail.To= "mbernholtz@questwindows.com, mdungo@questwindows.com, aramirez@questwindows.com" 'TEST
'Mail.Bcc="someoneelse@somedomain.com" 'Carbon Copy
'Mail.Cc="someoneelse2@somedomain.com" 'Blind Carbon Copy

if LastMonth > rs("DateTime") then
Mail.TextBody=" Last Shear Test:" & RS("DATETIME") & " - Over a Month old"
else
Mail.TextBody=" Last Shear Test:" & RS("DATETIME") & " - Over a Week old"
End if

Mail.Send
Set Mail = Nothing
'***********************
%>     
<p> Glass Tools Report (Optima and Completed) ~ E-mail Sent </p>
<%
end if

rs.close
set rs = nothing
DBConnection.close
Set DBConnection = nothing
%>
</body>
</html>
