<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 <!--#include file="dbpath.asp"-->
		 
<!--Window Remake List Update Form-->
<!--When new glass is called in, this form will be used by all so there is central storage of all information and no chasing -->
<!-- Sends to WindowRemakeEditConf -->
<!-- Designed August 2014, by Michael Bernholtz at request of Jody Cash-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Remake Summary</title>
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
WRID = Request.QueryString("WRid")


		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Window_Remakes WHERE ID = " & WRID
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle">Update Window Remake Information</h1>
                <a class="button leftButton" type="cancel" href="WindowRemakeReport.asp" target="_self">All Remakes</a>

    </div>			
    
    
    <form id="WindowRemakeEdit" title="Update Window Remake" class="panel" action="WindowRemakeEditConf.asp" name="WindowRemakeEdit"  method="GET" target="_self" selected="true" > 
  
	<fieldset>
	
	<div class="row">
		<label>Job </label>
		<input type="text" name='JOB' id='JOB' value ='<% response.write Trim(rs.fields("JOB")) %>' >
	</div>


	<div class="row">
		<label>Floor </label>
		<input type="text" name='Floor' id='Floor' value ='<% response.write Trim(rs.fields("FLOOR")) %>' >
	</div>
	<div class="row">
		<label>Tag</label>
		<input type="text" name='Tag' id='Tag' value ='<% response.write Trim(rs.fields("TAG")) %>' >
	</div>
	<div class="row">
        <label>Break Date</label>
		<input type="text" name='BREAKDATE' id='BREAKDATE' size='8' value ='<% response.write Trim(rs.fields("BREAKDATE")) %>' >
	</div>  
	<div class="row">
		<label>Break Cause</label>
		<input type="text" name='BreakCause' id='BreakCause' value ='<% response.write Trim(rs.fields("BREAKCAUSE")) %>' >
	</div>	
	<div class="row">
        <label>Required Date</label>
		<input type="text" name='RequiredDATE' id='RequiredDATE' size='8' value ='<% response.write Trim(rs.fields("REQUIREDDATE")) %>'  >
	</div>  
	<div class="row">
        <label>Send to </label>
		<input type="text" name='Sendto' id='Sendto' value ='<% response.write Trim(rs.fields("SENDTO")) %>' >
	</div>  
    <div class="row">
		<label>Ready</label>
        <input type="checkbox" name='Ready' id='Ready' <% if rs.fields("READY") = TRUE THEN response.write "checked" END IF%>>
    </div>   
	<div class="row">
        <label>Notes</label>
		<input type="text" name='Notes' id='Notes' value ='<% response.write Trim(rs.fields("Notes")) %>' >
	</div>  	
	</fieldset>
	<fieldset>
	<div class="row">
        <label>Re Order By</label>
		<input type="text" name='ReOrderBy' id='ReOrderBy' value ='<% response.write Trim(rs.fields("ReOrderBy")) %>' >
	</div>  
	<div class="row">
        <label>ReOrder Date</label>
		<input type="text" name='REORDERDATE' id='REORDERDATE' size='8' value ='<% response.write Trim(rs.fields("REORDERDATE")) %>' >
	</div>  
		<div class="row">
        <label>Received Date</label>
		<input type="text" name='RECEIVEDDATE' id='RECEIVEDDATE' size='8' value ='<% response.write Trim(rs.fields("RECEIVEDDATE")) %>' >
	</div>  
	    <div class="row">
		<label>Completed</label>
        <input type="checkbox" name='COMPLETED' id='COMPLETED' <% if rs.fields("COMPLETED") = TRUE THEN response.write "checked" END IF%>>
    </div> 
	
	
						<input type="hidden" name='WRID' id='WRID' value="<%response.write WRID %>" />
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="WindowRemakeEdit.action='WindowRemakeEditConf.asp'; WindowRemakeEdit.submit()">Update Remake Information</a><BR>
		
            
            </form>
                        
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

