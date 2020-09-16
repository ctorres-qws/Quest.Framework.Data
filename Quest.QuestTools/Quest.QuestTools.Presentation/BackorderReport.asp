<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Back Order Report - Showing all Back Order items as scanned -->
<!-- Michael Bernholtz, April 2015, Developed at Request of Jody Cash - Adapted mainly from Forel scan and XA9Backorder code-->
<%
' Use same page for email report - hide style and javascript
	Dim b_Email: b_Email = False
	If Request("Email") = "T" Then
		b_Email = True
	End If
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Back Order Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
<% If b_Email = False Then %>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    <!-- DataTables CSS -->
<link rel="stylesheet" type="text/css" href="../DataTables-1.10.2/media/css/jquery.dataTables.css">
  
<!-- jQuery -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.js"></script>
  
<!-- DataTables -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/media/js/jquery.dataTables.js"></script>
<!-- Fixed Headers -->
<script type="text/javascript" charset="utf8" src="../DataTables-1.10.2/extensions/FixedHeader/js/dataTables.fixedHeader.js"></script>
 
  <script type="text/javascript">
  $(document).ready( function () {
    $('#back').DataTable();
} );
  
  </script>
<% End if %>
    <%
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = FixSQL("SELECT X_BACKORDER.ID AS ID2, * FROM X_BACKORDER INNER JOIN X_BACKORDER_REASON ON X_BACKORDER.REASONID = X_BACKORDER_REASON.ID WHERE ACTIVE = TRUE ORDER BY X_BACKORDER.ID ASC")
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



'afilter = request.QueryString("aisle")


%>
<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
    </head>
<body>
<% If b_Email = False Then %>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target='_self' >Stock</a>
        </div>
<% Else %>
<style>
	.group { color: black !important; }
</style>
<% End If %>
         
       
        <ul id="Profiles" title="Glass Report - Service" selected="true">
        
        
<% 
response.write "<li class='group'>BACK ORDER REPORT </li>"
response.write "<li> Click on the Headers of each column to sort Ascending/Descending</li>  "
response.write "<li><table border='1' class='back' id ='back' ><thead><tr><th>ID</th><th>Backorder</th><th>Job</th><th>Floor</th><th>Tag</th><th>Section</th><th>Reason</th><th>Location</th><th>Backorder Date</th><th>Notes</th><th></th></tr></thead><tbody>"
do while not rs.eof
	response.write "<tr>"
	response.write "<td>" & RS("ID2") & "</td>"
	response.write "<td>" & RS("Barcode") & "</td>"
	response.write "<td>" & RS("Job") & "</td>"
	response.write "<td>" & RS("Floor") & "</td>"
	response.write "<td>" & RS("Tag") & "</td>"
	response.write "<td>" & RS("Section") & "</td>"
	response.write "<td>" & RS("Reason") & "</td>"
	response.write "<td>" & RS("Location") & "</td>"
	response.write "<td>" & RS("BackOrderDate") & "</td>"
	response.write "<td>" & RS("Note") & "</td>"
	If b_Email Then
		response.write "<td>&nbsp;</td>"
	Else
		response.write "<td><a class='greenButton' href='BackOrderC.asp?Returnsite=backorderreport.asp&bocid=" & RS("ID2") & "'  target='_self'> Reordered</a></td>"
	End If
	response.write " </tr>"
	rs.movenext
loop
response.write "</tbody></table></li></ul>"



rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>
        
               
</body>
</html>
