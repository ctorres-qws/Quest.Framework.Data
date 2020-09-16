<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->
		 
<!--Window Remake List Update Form-->
<!--When new glass is called in, this form will be used by all so there is central storage of all information and no chasing -->
<!-- Collects from to WindowRemakeEditForm -->
<!-- Designed August 2014, by Michael Bernholtz at request of Jody Cash-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Summary Edited </title>
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

WRID = request.querystring("WRID")
%>

	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
                <a class="button leftButton" type="cancel" href="WindowRemakeEditForm.asp?WRID=<% response.write WRid %>" target="_self">Remake</a>
    </div>
    
      
    
<form id="conf" title="Window Remake Updated" class="panel" name="conf" action="WindowRemakeReport.asp" method="GET" target="_self" selected="true" >              

  
   
        <h2>Window Remake Updated</h2>
  
<%       

JOB = UCASE(REQUEST.QueryString("JOB"))
FLOOR = REQUEST.QueryString("FLOOR")
TAG = REQUEST.QueryString("TAG")
BREAKDATE = REQUEST.QueryString("BREAKDATE")
if isdate(BREAKDATE) = false then
	BREAKDATE = DATE()
end if
BREAKCAUSE= REQUEST.QueryString("BREAKCAUSE")
REQUIREDDATE = REQUEST.QueryString("REQUIREDDATE")
if isdate(REQUIREDDATE) = false then
	REQUIREDDATE = DateAdd("d",10,Date()) 
end if
READY = REQUEST.QueryString("Ready")
If READY = "on" then
	READY = TRUE
Else
	READY = FALSE
End If
SENDTO  = REQUEST.QueryString("Sendto")
COMPLETED = REQUEST.QueryString("Completed")
If COMPLETED = "on" then
	COMPLETED = TRUE
Else
	COMPLETED = FALSE
End If
REORDERDATE = REQUEST.QueryString("REORDERDATE")
RECEIVEDDATE = REQUEST.QueryString("RECEIVEDDATE")
REORDERBY = REQUEST.QueryString("REORDERBY")
NOTES = REQUEST.QueryString("NOTES")

		   

	
			'Set Glass Inventory Update Statement
				StrSQL = "UPDATE Window_Remakes  SET [JOB]='"& JOB & "', [FLOOR]='" & FLOOR & "', [TAG]='" & TAG & "', BREAKDATE= '" & BREAKDATE & "', [BREAKCAUSE]='" & BREAKCAUSE & "', [REQUIREDDATE]='" & REQUIREDDATE & "', [READY]= " & READY & ", [COMPLETED]= " & COMPLETED & ", [REORDERBY]= " & REORDERBY & ", [NOTES]= " & NOTES & " WHERE ID = " & WRID
				
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	
		if isdate(REORDERDATE) then
				
			'Set DATE Update Statement only if ISDATE
				StrSQL2 = "UPDATE Window_Remakes  SET [REORDERDATE]='"& REORDERDATE & "' WHERE ID = " & WRID
		else
			'Set DATE Update Statement only if ISDATE
				StrSQL2 = "UPDATE Window_Remakes  SET [REORDERDATE]= NULL WHERE ID = " & WRID
		end if
			'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
		
		
		if isdate(RECEIVEDDATE) then
		
			'Set DATE Update Statement only if ISDATE
				StrSQL3 = "UPDATE Window_Remakes  SET [RECEIVEDDATE]='"& RECEIVEDDATE & "' WHERE ID = " & WRID
		else
			'Set DATE Update Statement only if ISDATE
				StrSQL3 = "UPDATE Window_Remakes  SET [RECEIVEDDATE]= NULL WHERE ID = " & WRID
				
		end if		
			'Get a Record Set
				Set RS3 = DBConnection.Execute(strSQL3)
		
		
	
	

Set RS = Nothing
DBConnection.close
set DBConnection=nothing


%>

    
<ul id="Report" title="Added" selected="true">

	<li> Window Remake UPDATED:</li>
    <li><% response.write "JOB: " & JOB %></li>
	<li><% response.write "FLOOR: " & FLOOR %></li>
    <li><% response.write "TAG: " & TAG %></li>
	<li><% response.write "BREAK DATE: " & BREAKDATE %></li>
    <li><% response.write "BREAK CAUSE: " & BREAKCAUSE %></li>
    <li><% response.write "REQUIRED DATE: " & REQUIREDDATE %></li>
    <li><% response.write "READY: " & READY %></li>
	<li><% response.write "SEND WINDOW TO: " & SENDTO %></li>
	<br>
	<li><% response.write "RE ORDER DATE: " & REORDERDATE %></li>
	<li><% response.write "RECEIVED DATE: " & RECEIVEDDATE %></li>
	<li><% response.write "COMPLETED: " & COMPLETED %></li>
	<br>
	
	<li> Please do not forget to update information during the replacement process or mark it completed when the window is shipped back out</li>


</ul>
        <BR>
       
         <a class="whiteButton" href="javascript:conf.submit()">Home</a>
            
            </form>

            
    
</body>
</html>


