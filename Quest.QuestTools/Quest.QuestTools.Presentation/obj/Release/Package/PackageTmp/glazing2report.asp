<!--#include file="dbpath.asp"-->
<!--#include file="dbpath-QUESTSQL.asp "--> 


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
		 <!-- Glazing 2 Report - Table in SQL Server, Michael Bernholtz July 2014 -->
		 <!-- Information collected about Glazing 2 for explanation -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glazing 2</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1000" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript"> iui.animOn = true; </script>
  <script src="sorttable.js"></script>
  

<% 

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * From X_BARCODE_LINEitem WHERE DEPT = 'GLAZING2' ORDER BY ID DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection	



currentDate = Date()
cYear = year(now)
weekNumber = DatePart("ww", currentDate)
sixweeks = weekNumber - 6

%>

</head>

<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

	
	  <form id="reason" title="Glazing 2" class="panel" name="reason" action="glazing2reportform.asp" method="GET" target="_self" selected= "true">
<ul>


	              <li class="group">LAST 6 WEEK'S ACTIVITY</li>
        
<%

	


	Response.write "<li> <table border ='1' class='sortable' cellpadding='3' >"
	Response.write "<tr title = 'Click on a Header to Sort by that Column' ><th>Job</th><th>Floor</th><th>Glazing2</th><th>G2 Today</th><th>Date</th><th>Reason</th><th>Add Reason</th>"
	
	rs.filter = "WEEK > " & sixweeks & " AND YEAR = " & cYear

	do while not rs.eof
	
	response.write "<tr>"	
	response.write " <td> " & rs("job") & " </td><td> " & rs("floor") & " </td><td> " & rs("tag") & " </td><td> " & rs("last") & " </td>"
	
	RecentDate = DateValue(rs("month") & "-" & rs("day") & "-" & rs("year"))
		
	response.write "<td>" & RecentDate &"</td>"
	
	
	response.write " <td> "& rs("G2Reason") & "</td>"
	
	response.write "<td><a href='glazing2reportform.asp?g2id=" & rs("id")& "' target='_self' />Add Reason </td>"
	
	response.write "</tr> "
	

	rs.movenext
	loop
	
	
	
	Response.write "</table></li>"

	

%>

<li>	<a class="whiteButton" href="javascript:reason.submit()">Submit</a> </li>
        </ul>
       
		

		
        
     <% 


rs.close
set rs=nothing
DBConnection.close
set DBConnection=nothing
DBConnection2.close
set DBConnection2=nothing
%> 





</body>
</html>



