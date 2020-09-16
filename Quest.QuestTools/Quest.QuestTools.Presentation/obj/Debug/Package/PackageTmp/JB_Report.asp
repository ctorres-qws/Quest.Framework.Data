<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Jamb Receptor Report (Left and Right), with given Job and Floor - Checks JB_XXX##-->
<!-- Basic Entry form gets Job And Floor JB_ReportEnter.asp sends to JB_Report.asp-->
<!-- JB_Report Designed by Ariel Aziza, Coded by Michael Bernholtz, October 2018 -->
<!-- Left and Right Jamb Receptor -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Jamb Receptor Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="JB_ReportEnter.asp" target="_self">Back</a>
        </div>
  
        <ul id="Profiles" title="Jamb Receptor L/R" selected="true">
<%

JOB = request.querystring("JOB")
FLOOR = request.querystring("FLOOR")


On Error Resume Next	
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM [JR_" & JOB & FLOOR & "] ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

IF Err.Number <>0 then

Response.write "<li> Jamb Receptor Table was not processed. </li>"
Response.write "<li> Please Process JB_" & JOB & FLOOR & " before continuing</li>"

Else

	
	%>
	
	<li class='group'>Left and Right Jamb Receptors for <%Response.write JOB%>   <%Response.write FLOOR%></li>
	<li><table border='1' class='sortable' align ='center'>
		<tr><th align='center'>LEFT</th><th align='center'>RIGHT</th></tr>
	<TR>
		<td align='center'><img src="\partpic\LeftJR.jpg" alt="Left Jamb Receptor Image" width="100%" height="100%" %></td>
		<td align='center'><img src="\partpic\RightJR.jpg" alt="Right Jamb Receptor Image"  width="100%" height="100%"></td>
	</TR>
	<TR>
	<TD valign ='top'><table border='1' class='sortable'>
		<TR><TH>Job</TH><TH>Floor</TH><TH>Tag</TH><TH>Length</TH><TH>Colour Code</TH><TH>Bundle Label</TH></TR>
	<%
	rs.filter = "Extrusion = 'LEFT JAMB-RECEPTOR'"
	Do While not rs.eof
		response.write "<TR>"
		Response.write "<TD>" & rs("Job") & "</TD>"
		Response.write "<TD>" & rs("Floor") & "</TD>"
		Response.write "<TD>" & rs("Tag") & "</TD>"
		Response.write "<TD>" & rs("Length") & "</TD>"
		Response.write "<TD>" & rs("ColorCode") & "</TD>"
		Response.write "<TD>" & rs("BundleLabel") & "</TD>"
		response.write "</TR>"
	rs.movenext 
	loop
	%>
	
	
	</table>
	</TD>
	<TD valign ='top'><table border='1' class='sortable' >
	<TR><TH>Job</TH><TH>Floor</TH><TH>Tag</TH><TH>Length</TH><TH>Colour Code</TH><TH>Bundle Label</TH></TR>
	<%
	rs.filter = "Extrusion = 'RIGHT JAMB-RECEPTOR'"
	Do While not rs.eof
		response.write "<TR>"
		Response.write "<TD>" & rs("Job") & "</TD>"
		Response.write "<TD>" & rs("Floor") & "</TD>"
		Response.write "<TD>" & rs("Tag") & "</TD>"
		Response.write "<TD>" & rs("Length") & "</TD>"
		Response.write "<TD>" & rs("ColorCode") & "</TD>"
		Response.write "<TD>" & rs("BundleLabel") & "</TD>"
		response.write "</TR>"
	rs.movenext 
	loop
	%>
	</table>
	</TD>
		
	
	</TR>
	</TABLE></li>
	

	<%	
	rs.close
	set rs = nothing
	DBConnection.close 
	set DBConnection = nothing
	%>

    </ul>   

<%

End if 
%>	
                           
</body>
</html>
