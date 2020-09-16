                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Form Created December 31, 2014, Michael Bernholtz at request of Slava Kotek, Lev Bedoev, Jody Cash-->
<!-- View form to show items on Zipper-->
<!-- Zipper will be changed to automatic only-->
<!-- Entry from Quest Dashboard ZipperEnter.asp -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Zipper Report</title>
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
strSQL = "SELECT * FROM Roll_Table WHERE cStatus < Qty ORDER BY ID ASC"
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
        <a class="button leftButton" type="cancel" href="index.html#_Zipper" target="_self">Zipper</a>
        </div>
   
   
         
       
        <ul id="Profiles" title=" Glass Report - All Active" selected="true">
        
        

<li class='group'>All Active Zipper REPORT </li>
<li> Click on the Headers of each column to sort Ascending/Descending</li>  
<li><table border='1' class='sortable'>
<tr><th>ID</th><th>Job</th><th>Floor</th><th>Profile</th><th>Status</th><th>Qty</th><th>Length</th><th>Entry Date</th></tr>
<% 
do while not rs.eof
	response.write "<tr><td>" & RS("ID") & "</td><td>" & RS("Job") & "</td><td>" & RS("Floor") &"</td><td>" & RS("Profile") & "</td>"
	response.write "<td>" & RS("cStatus") & "</td><td>" & RS("Qty") &  "</td><td>" & RS("Length") &  "</td><td>" & RS("EnterDate") & "</td>"
	response.write "</tr>"

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
