                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<!--WindowRemake Format designed as Quick Summary of all Jobs for reference. -->
<!-- Designed August 2014, by Michael Bernholtz -->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Window Remake Report</title>
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
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Window_Remakes WHERE COMPLETED = FALSE ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>
 <!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Ship" target="_self">Shipping</a>
        </div>
      
       
        <ul id="Profiles" title=" Remake Window - All Active" selected="true">
        <li class='group'>Window Remake Summary </li>
         <a class="whiteButton" href="WindowRemakeENTER.asp" target='_Self'>Add Window Remake</a>
<% 

response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='sortable'><tr><th>Job</th><th>Floor</th><th>Tag</th><th>Break Cause</th><th>Re Order By</th><th>SendTo</th><th>Notes</th><th>Ready</th><th>Break Date</th><th>Re-order Date</th><th>Required Date</th><th>Received Date</th></tr>"

do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("JOB") & "</td>"
	response.write "<td>" & RS("FLOOR") &"</td>"
	response.write "<td>" & RS("TAG") & "</td>"
	response.write "<td>" & RS("BreakCause") & "</td>"
	response.write "<td>" & RS("ReOrderBY") & "</td>"
	response.write "<td>" & RS("Sendto") & "</td>"
	response.write "<td>" & RS("notes") & "</td>"
	
	response.write "<td>" & RS("Ready") & "</td>"
	
	response.write "<td>" & RS("BreakDate") & "</td>"
	response.write "<td>" & RS("ReOrderDate") & "</td>"
	response.write "<td>" & RS("RequiredDate") & "</td>"
	response.write "<td>" & RS("ReceivedDate") & "</td>"

	response.write "<td> <a class='lightblueButton' href='WindowRemakeEditForm.asp?WRID=" & RS("ID") & "' target='_Self'>Update Window</a></td>"
	response.write " </tr>"

	rs.movenext
loop





rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing



%>
      </ul>                 
            
     
               
</body>
</html>
