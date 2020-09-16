<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!--Window Remake List Entry Confirmation-->
<!--When new glass is called in, this form will be used by all so there is central storage of all information and no chasing -->
<!-- Collects from WindowRemakeEnter -->
<!-- Designed August 2014, by Michael Bernholtz at request of Jody Cash-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>



<% 


JOB = UCASE(REQUEST.QueryString("JOB"))
FLOOR = REQUEST.QueryString("FLOOR")
TAG = REQUEST.QueryString("TAG")
BREAKDATE = REQUEST.QueryString("BREAKDATE")
if isdate(BREAKDATE) = false then
	BREAKDATE = DATE()
end if
BREAKCAUSE= REQUEST.QueryString("BREAKCAUSE")
REQUIREDDATE = REQUEST.QueryString("REQUIREDDATE")
if isdate(REQUIREDDATE) = false then
	REQUIREDDATE = DateAdd("d",10,Date()) 
end if
READY = REQUEST.QueryString("Ready")
If READY = "on" then
	READY = TRUE
Else
	READY = FALSE
End If
SENDTO  = REQUEST.QueryString("Sendto")
NOTES = REQUEST.QueryString("Notes")

response.write BREAKDATE & REQUIREDDATE

''Create a Query
	strSQL = "INSERT INTO Window_Remakes ([JOB], [FLOOR], [TAG], [BREAKDATE], [BREAKCAUSE], [REQUIREDDATE], [READY], [SENDTO], [NOTES]) VALUES( '" & JOB & "', '" & FLOOR &  "', '" & TAG & "', '" & BREAKDATE & "', '" & BREAKCAUSE & "', '" & REQUIREDDATE & "', " & READY & ", '" & SENDTO & "', '" & NOTES & "')"
''Get a Record Set
    Set RS = DBConnection.Execute(strSQL)
	
%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="WindowRemakeENTER.asp" target="_self">Remake</a>

    </div>


    
<ul id="Report" title="Added" selected="true">
	
	<li> Window Added to the Remake list:</li>
    <li><% response.write "JOB: " & JOB %></li>
	<li><% response.write "FLOOR: " & FLOOR %></li>
    <li><% response.write "TAG: " & TAG %></li>
	<li><% response.write "BREAK DATE: " & BREAKDATE %></li>
    <li><% response.write "BREAK CAUSE: " & BREAKCAUSE %></li>
    <li><% response.write "REQUIRED DATE: " & REQUIREDDATE %></li>
    <li><% response.write "READY: " & READY %></li>
	<li><% response.write "SEND WINDOW TO: " & SENDTO %></li>
	<li><% response.write "Additional Notes: " & NOTES %></li>
	<br>
	<li> Please do not forget to update information during the replacement process or mark it completed when the window is shipped back out</li>
<li><a class="button" type="cancel" href="WindowRemakeReport.asp" target="_self">Home</a></li>

</ul>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection = nothing
%>

</body>
</html>



