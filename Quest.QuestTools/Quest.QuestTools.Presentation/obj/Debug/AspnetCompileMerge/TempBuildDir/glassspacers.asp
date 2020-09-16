<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Drafted December 4th, 2013 - by Michael Bernholtz at request of Jody Cash --> 
<!-- Submits to page GLASSSpacersCONF.asp -->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Enter Spacers</title>
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

              <form id="enter" title="Enter Spacers" class="panel" name="ManageSpacers" action="glassspacersconf.asp" method="GET" target="_self" selected="true">

        <h2>Enter Spacers </h2>

<fieldset>

    <div class="row">
        <label> Spacer: </label>
        <input type="number" name='SPACER' id='SPACER' >
		<ul>
        <li>Note: Spacer must be a number</li>
		</ul>
    </div>

    <div class="row">
        <label> OT: </label>
        <input type="text" name='OT' id='OT' >
		<ul>
        <li>Note: Please List OT in format "Overall Inches (Xmm / X / Xmm) </li>
		</ul>
    </div>

	<a class="whiteButton" href="javascript:ManageSpacers.submit()">Submit</a>
</fieldset>

<% 

Set rs6 = Server.CreateObject("adodb.recordset")
strSQL6 = "Select * FROM XQSU_OTSpacer ORDER BY SPACER ASC"
rs6.Cursortype = 2
rs6.Locktype = 3
rs6.Open strSQL6, DBConnection
%>
 
<ul id="Profiles" title="Glass" selected="true">

<%

response.write "<li class = 'group'>Current Glass Spacers</li>" 
do while not rs6.eof
	response.write "<li class = 'group'>Spacer: " & rs6("SPACER") & " - " & rs6("OT") & "</li>"
rs6.movenext
loop
	
%>
	</ul>



            
            </form>
                
<%             
rs6.close
set rs6=nothing

DBConnection.close
set DBConnection=nothing           
%>
</body>
</html>
