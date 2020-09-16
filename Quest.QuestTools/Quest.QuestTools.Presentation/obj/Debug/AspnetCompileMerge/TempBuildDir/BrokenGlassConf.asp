<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 11th, by Michael Bernholtz - Conf Page: Marks new Item as broken-->
<!-- Form created at Request of Ariel Aziza Implemented by Michael Bernholtz--> 
<!-- Using Tables: X_Broken -->
<!-- Inputs from BrokenGlassAdd.asp -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Broken Glass</title>
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
job = trim(UCASE(Request.Form("job")))
floor = trim(UCASE(Request.Form("floor")))
tag = trim(UCASE(Request.Form("tag")))
opening = trim(UCASE(Request.Form("opening")))
width = trim(Request.Form("width"))
height = trim(Request.Form("height"))
addby = trim(UCASE(Request.Form("addby")))
reason = trim(UCASE(Request.Form("reason")))
notes = trim(UCASE(Request.Form("notes")))
addDate = FormatDateTime(Date, 0)

		'Set Glass Input Statement
			StrSQL = "INSERT INTO X_Broken (job, floor, tag, opening, width, height, addby, reason, notes, addDate) VALUES ('" & job & "', '" & floor & "', '" & tag & "', '" & opening & "', '" & width & "', '" & height & "', '" & addby & "', '" & reason & "', '" & notes & "', '" & addDate & "')"
			
		'Get a Record Set
			Set RS = DBConnection.Execute(StrSQL)

%>  
  
  
  
    </head>

<body >

    <div class="toolbar">
        <h1 id="pageTitle">Add New Broken Glass</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
	</div>
   
       
    
<ul id="Report" title="Added" selected="true">
	
<%		
		
		Response.Write "<li>Broken Glass Values Entered:</li>"
		Response.Write "<li> Job: " & job & "</li>"
		Response.Write "<li> Floor: " & floor & "</li>"
		Response.Write "<li> Tag: " & tag & "</li>"
		Response.Write "<li> Opening: " & opening & "</li>"
		Response.Write "<li> Width: " & width & "''</li>"
		Response.Write "<li> Height: " & height& "''</li>"
		Response.Write "<li> Added By: " & addby & "</li>"
		Response.Write "<li> Added Date: " & addDate & "</li>"
		Response.Write "<li> Reason: " & reason & "</li>"
		Response.Write "<li> Notes: " & notes & "</li>"


DBConnection.close
set DBConnection = nothing
%>

            
</body>
</html>
