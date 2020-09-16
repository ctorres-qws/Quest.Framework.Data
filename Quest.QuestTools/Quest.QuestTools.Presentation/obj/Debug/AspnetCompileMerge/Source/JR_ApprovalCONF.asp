<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
			<!--#include file="dbpath.asp"-->
	<!-- Approval Form for JR PROCESSING -->
	<!-- Given Job and Floor, Updates to Z_Floors Approval -->
	<!-- Sent as an Email in 8080 JR_PROCESSING sends to JR_ApprovalCONF-->
	<!-- Michael Bernholtz, October 2018 - for Ariel Aziza, David Ofir and Jody Cash -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Approval Confirmation</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>
  
  
</head>
<body>
<%
JOB = Request.QueryString("Job")
Floor = Request.QueryString("Floor")
Approval = Request.QueryString("Approval")
ApprovalManager = Request.QueryString("ApprovalManager")
%>
    <div class="toolbar">
        <h1 id="pageTitle">Approval Confirmation</h1>
               
    </div>
    
  <%
  
  
	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Z_Floors Where JOB = '" & JOB & "' AND Floor = '" & Floor & "' ORDER BY Floor ASC"
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection  
	
	if rs.eof then
	
	else
		if APPROVAL = "TRUE" THEN
			rs("Approval") = "TRUE"
			rs("ApprovalManager") = ApprovalManager
			rs("ApprovalTime") = Now
			rs.update
		else
			rs("Approval") = "FALSE"
			rs("ApprovalManager") = ApprovalManager
			rs("ApprovalTime") = Now
			rs.update
		end if
	end if
	
	rs.close
	set rs = nothing
	
%>   
    
              <form id="edit" title="Approval Form" class="panel" name="edit" action="addFloor3Conf.asp" method="GET" target="_self" selected="true" >
        <h2>Approval Condition</h2>

<fieldset>
     <div class="row">
                
	<% 
	if APPROVAL = "TRUE" THEN
		Response.write "<P><center>Approval Given to " & JOB & " / " & FLOOR & "</center></P>"
		Response.write "<P><center>Please inform Data Entry Operator that they can now proceed.</center></P>"
	else
		Response.write "<P><center>Approval NOT Given to " & JOB & " / " & FLOOR & "</center></P>"
	end if

	
%>
	
            </div>
            
</fieldset>
                       
            </form>
            
    
</body>
</html>

<% 

DBConnection.close
set DBConnection=nothing
%>

