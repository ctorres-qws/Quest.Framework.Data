<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created April 11th, by Michael Bernholtz - Order Page for all items that are marked Broken-->
<!-- Form created at Request of Ariel Aziza Implemented by Michael Bernholtz--> 
<!-- Using Tables: X_Broken -->
<!--Inputs from BrokenGlassOrder.asp-->
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
ordernum = Request.Form("ordernum")
orderDate = FormatDateTime(Date, 0)
count = 0
If NOT Request.Form("glass")= "" Then
	count = Request.Form("glass").count	
end if
%>




  
  
    </head>

<body >

    <div class="toolbar">
        <h1 id="pageTitle">Add New Broken Glass</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
	</div>
 
<ul id="Report" title="Added" selected="true">

<%


if count = 0 or ordernum = "" then
	Response.Write "<li>Please enter a Valid Order Number and select at least Broken item that was reordered.</li>"
else
	Response.Write "<li>" & " Items Added to order:" & ordernum & "</li>"
For Each item In Request.Form("glass")

currentid= Int(item)
	'Set Glass Order Statement
		StrSQL = "Update X_Broken Set ordernum = '" & ordernum & "', orderDate = '" & orderDate & "' Where id = " & currentid
			
	'Get a Record Set
		Set RS = DBConnection.Execute(StrSQL)
Next
 Response.Write "<li>Glass now listed as ordered</li>"
end if
%>

</ul>
<%
DBConnection.close
set DBConnection = nothing
%>

            
</body>
</html>
