<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 4th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!-- Submits to page GLASStypeCONF.asp -->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Enter Glass Types</title>
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
        <a class="button leftButton" type="cancel" href="index.html#_Job" target="_self">Job / Colour</a>
        </div>
              <form id="enter" title="Enter Glass Types" class="panel" name="managetypes" action="glasstypeconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter Glass Types</h2>

<fieldset>
            
<!--Entry form to add new Type and Description and then submit it to the database-->

    <div class="row">
        <label> Type: </label>
        <input type="text" name='GLASSTYPE' id='GLASSTYPE' >
    </div>
	
    <div class="row">
        <label> Description: </label>
        <input type="text" name='DESCRIPTION' id='DESCRIPTION' >
    </div>
	
	<div class="row">
        <label> Shop Code: </label>
        <input type="text" name='ShopCode' id='ShopCode' >
    </div>
		<div class="row">
		<label>Normal</label>
        <input type="radio" name='status' id='status' value="" checked />  
	</div>
	<div class="row">
		<label>Tempered</label>
        <input type="radio" name='status' id='status' value="TMP" />  
	</div>
	<div class="row">
		<label>Heat Strengthened</label>
		<input type="radio" name='status' id='status' value="HS" /> 
	</div>
	<div class="row">
        <label> JOB: </label>
        <input type="text" name='Job' id='Job' >
    </div>

	<a class="whiteButton" href="javascript:managetypes.submit()">Submit</a>
</fieldset>

<% 
Set rs5 = Server.CreateObject("adodb.recordset")
strSQ5L = "Select * FROM XQSU_GlassTypes ORDER BY TYPE ASC"
rs5.Cursortype = 2
rs5.Locktype = 3
rs5.Open strSQ5L, DBConnection
%>


<ul id="Profiles" title="Glass" selected="true">
<li class = 'group'> Current Glass Types:</li>
<li><table border ='1'>
<tr><th>Optima Glass Type</th><th>Description</th><th>ShopCode</th><th>Status</th><th>Job</th></tr>

<%
do while not rs5.eof
	response.write "<tr>"
	response.write"<td>" & rs5("Type") & "</td>"
	response.write"<td>" & rs5("Description") & "</td>"
	response.write"<td>" & rs5("ShopCode") & "</td>"
	response.write"<td>" & rs5("Status") & "</td>"
	response.write"<td>" & rs5("Job") & "</td>"
	response.write "</tr>"
rs5.movenext
loop
	
%>
	</table></li>
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
