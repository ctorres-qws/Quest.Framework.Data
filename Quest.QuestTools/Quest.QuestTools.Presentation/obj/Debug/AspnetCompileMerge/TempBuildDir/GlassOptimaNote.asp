<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 
<!--Add Notes Form GlassOptimaNote->
<!-- Submits to page GlassOptimaNoteConf.asp -->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Add Note</title>
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
  


<% 
GID = Request.QueryString("Gid")


		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Z_GLASSDB"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	
		rs.filter = "ID = " & GID
		

ticket = Request.Querystring("ticket")


%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
               
<%
	Select Case ticket
		case "select"
		%>
			   <a class="button leftButton" type="cancel" href="GlassExportSelect.asp" target="_self">Select Glass</a>
		<%
		case "active"
		%>
			   <a class="button leftButton" type="cancel" href="GlassReportActive.asp" target="_self">Active Glass</a>
		<%
	End Select
		%>
		
		
    </div>			
    
    
    <form id="GlassEdit" title="Edit Glass" class="panel" action="GlassOptimaNoteConf.asp" name="GlassEdit"  method="GET" target="_self" selected="true" > 
  <H2> <% response.write Trim(rs.fields("JOB")) & Trim(rs.fields("FLOOR")) & Trim(rs.fields("TAG")) %> </h2>
	<fieldset>


		<div class="row">
                <label>Add Notes</label>
                <input type="text" name='NOTES' id='NOTES' value="<% response.write Trim(rs.fields("NOTES")) %>">
        </div>
                    
						<input type="hidden" name='GID' id='GID' value="<%response.write GID %>" />
						<input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>" />
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="GlassEdit.action='GlassOptimaNoteConf.asp'; GlassEdit.submit()">ADD NOTE</a><BR>
		
            
            </form>
                        
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

