            
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>

  </head>
  <body>

<%


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

Mail.Subject="Canada Closed Truck: " & TruckName
Mail.From="quest.windows@gmail.com" 'This has to be an actual email address or an alias that's setup on the gmail account you used above
'Mail.To="jodycash@gmail.com, mbernholtz@questwindows.com, mmohanlall@questwindows.com, kcosgrove@questwindows.com, shipping@questwindows.com , AlexS@questwindows.com, agjondedaj@questwindows.com"
Mail.To="mdungo@questwindows.com" ' for testing

'Mail.Bcc="mdungo@questwindows.com, aramirez@questwindows.com" 'Carbon Copy
'Mail.Cc="someoneelse2@somedomain.com" 'Blind Carbon Copy

'**Below are different options for the Body of an email. *Only one of the below body types can be sent.
'sMail.TextBody="Report"


'Mail.HTMLBody="This is an email message that accepts HTML tags"
Mail.CreateMHTMLBody "http://172.18.13.31:8081/ShipTruckViewerGmailMichelle.asp?Truck=" & truck  'Sends an email which has a body of a specified webpage
'Mail.CreateMHTMLBody "file://c:/mydocuments/email.htm" 'Sends an email which has a body of an html file that's stored on your computer. This MUST be on the server that this script is being served from.

' How to add an attachment
'myMail.AddAttachment "c:\mydocuments\test.txt" 'Again this must be on the server that is serving this script.

Mail.Send
Set Mail = Nothing
%>     
<p> Window Production ~ E-mail Sent </p>
</body>
</html>
