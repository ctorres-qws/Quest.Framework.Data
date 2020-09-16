



 <% 
 
TruckChild = rs("job")
TruckParent =""
ScanChild = jobname
ScanParent =""
 
Set rsParent = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Job, Parent FROM Z_jobs ORDER BY JOB ASC"
'rsParent.Cursortype = 2
'rsParent.Locktype = 3
'rsParent.Open strSQL, DBConnection
Set rsParent = GetDisconnectedRS(strSQL, DBConnection)

rsParent.filter = "JOB = '" & TruckChild & "'"
	if rsParent.eof then 
		TruckParent = TruckChild
	else
		TruckParent = rsParent("Parent")
	end if
rsParent.filter =""
rsParent.filter = "JOB = '" & ScanChild & "'"
	if rsParent.eof then 
		ScanParent = ScanChild
	else
		ScanParent = rsParent("Parent")
	end if
 
rsParent.close
set rsParent = nothing

if TruckParent = ScanParent then
else

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

Mail.Subject="Incorrect Load - Check if Correct"
Mail.From="quest.windows@gmail.com" 'This has to be an actual email address or an alias that's setup on the gmail account you used above
Mail.To= "mbernholtz@questwindows.com, mdungo@questwindows.com, aramirez@questwindows.com" 'TEST
'Mail.Bcc="someoneelse@somedomain.com" 'Carbon Copy
'Mail.Cc="someoneelse2@somedomain.com" 'Blind Carbon Copy


BodyCode = BodyCode &  windowValue & " : (" & jobname & floorname & "-" & tagname & ")"
BodyCode = BodyCode & " Scanned onto Truck: " & truck & " - " & RS("JOB") & " " & RS("FLOOR")
BodyCode = BodyCode & " Please Confirm!"

Mail.TextBody= BodyCode

Mail.Send
Set Mail = Nothing
'***********************

end if



%>
