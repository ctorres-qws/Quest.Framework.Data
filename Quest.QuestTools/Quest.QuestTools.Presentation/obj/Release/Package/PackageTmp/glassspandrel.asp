<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted July 31st, 2014 - by Michael Bernholtz at request of Jody Cash --> 
<!-- Submits to page GLASSspandrelCONF.asp -->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Enter Spandrel Colour</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job/Colour</a>
        </div>

              <form id="enter" title="Enter Spandrel Color" class="panel" name="managetypes" action="glassSpandrelconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter Spandrel Colour</h2>


<fieldset>

<!--Entry form to add new Type and Description and then submit it to the database-->

    <div class="row">
        <label> Code: </label>
        <input type="text" name='CODE' id='CODE' >
    </div>
	
    <div class="row">
        <label> Description: </label>
        <input type="text" name='DESCRIPTION' id='DESCRIPTION' >
    </div>

	<div class="row">
        <label> JOB: </label>
        <input type="text" name='JOB' id='JOB' >
    </div>
	
	<div class="row">
        <label> NOTES: </label>
        <input type="text" name='NOTES' id='NOTES' >
    </div>
	       
	<div class="row">
        <label>Active</label>
        <input type="checkbox" name='Active' id='Active' checked>
    </div> 
        
		
	<a class="whiteButton" href="javascript:managetypes.submit()">Submit</a>
</fieldset>


<% 
Set rs5 = Server.CreateObject("adodb.recordset")
strSQ5L = "Select * FROM Y_COLOR_SPANDREL ORDER BY CODE ASC"
rs5.Cursortype = GetDBCursorType
rs5.Locktype = GetDBLockType
rs5.Open strSQ5L, DBConnection
%>


<ul id="Profiles" title="Glass" selected="true">

<%




response.write "<li class = 'group'> Current Spandrel Colours:</li>" 
do while not rs5.eof
	response.write "<li>Glass Type: " & rs5("Code") & " - " & rs5("Description") & " - " & rs5("Active") & "</li>"
rs5.movenext
loop
	
%>
	</ul>
	</form>
	
                
<%             
rs5.close
set rs5=nothing

DBConnection.close
set DBConnection=nothing           
%>
</body>
</html>
