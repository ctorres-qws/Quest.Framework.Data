<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
	 "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	<!-- Approval Form for JR PROCESSING -->
	<!-- Given Job and Floor, Updates to Z_Floors Approval -->
	<!-- Sent as an Email in 8080 JR_PROCESSING sends to JR_ApprovalCONF-->
	<!-- Michael Bernholtz, October 2018 - for Ariel Aziza, David Ofir and Jody Cash -->
			
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Approve Jamb Processing</title>
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
Job = Request.queryString("JOB")
Floor = Request.queryString("FLOOR")
%>


    <div class="toolbar">
        <h1 id="pageTitle">Approval for <%Response.write JOB%> : <%Response.write Floor%> </h1>

    </div>
    
              <form id="add" title="Approval" class="panel" name="add" action="JR_ApprovalCONF.asp" method="GET" target="_self" selected="true" > 
        <h2>Approval for <%Response.write JOB%> : <%Response.write Floor%> </h2>
  
   

<fieldset>
<input type="hidden" name='Job' id='Job' value ='<% response.write JOB%>' >
<input type="hidden" name='Floor' id='Floor' value ='<% response.write Floor%>' >
     <div class="row">
		<label> Approval </label>
        <select name ='Approval'>
			<option value = "TRUE">Approved</option>
			<option value = "FALSE">Not Approved</option>
		</select>
     </div>
	<div class="row">
		<label>Approved By?</Label>
		<input type="text" name="ApprovalManager" required /></TD>
	</div>
            
</fieldset>


        <BR>
		
	<%
	if JOB="" or Floor = "" then
	else
	%>
        <a class="whiteButton" href="javascript:add.submit()">Give Approval</a><BR>
   	<%
	end if
	%>
         
            
            </form> 

            
    
</body>
</html>


