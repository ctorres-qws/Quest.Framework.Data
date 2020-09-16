<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created August 8th, by Michael Bernholtz - Edit and Delete Form for items in Optimization Log-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Optimization Log Edit Form</title>
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

OPID = Request.QueryString("OPid")

Dim Identifier, IdentifierID
'Identifier changes the entry between Serial Number for Glass, Box Number for Spacer, Lot Number for Sealant



		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM OptimizeLog"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	
		rs.filter = "ID = " & OPID
		

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="OptimizationLogManage.asp" target="_self">Opt Log </a>
				
    </div>			
    
    
    <form id="OptLog" title="Edit Stock" class="panel" action="OptimizationLogEditConf.asp" name="OptLog"  method="GET" target="_self" selected="true" > 
  
	<fieldset>
	
	  <div class="row">
                <label>Job</label>
                <select name="Job">
					<% ActiveOnly = True %>
			<option value="SRV">SRV</option>
			<option value="EB">EB</option>
			<!--#include file="JobsList.inc"-->
			<% 
			' Coded this 3 times - to show Description again, despite collected value is TYPE ( USER CHOOSES DESCRIPTION, but SYSTEM NEEDS TYPE)
			rsJob.filter = "Job = '" & rs("Job") & "'"
			if rsJob.eof then
			if rs("JOB") = "SRV" or rs("JOB") = "EB" or rs("JOB") = "RM" then
				%>
			<option value = "<%response.write rs("Job") %>" selected><%response.write rs("Job") %></option>
				<%
			else
				%><option value = "" selected>-</option><%
			end if
			else
			%>
			<option value = "<% response.write rsJob("Job") %>" selected><% response.write rsJob("Job") %></option> 

			<%
			end if
			%>
			</select>
            </div>

            <div class="row">
                <label>Floor</label>
                <input type="text" name='floor' id='floor' value="<% response.write Trim(rs.fields("Floor")) %>">
            </div>
			
			 <div class="row">
                <label>Glass</label>
                <select name="Glass">
				
<% mat = "Glass" %>
<% entertype = "Code" %>
<option name="Glass" option value = '<% response.write Trim(rs.fields("Glass")) %>'> <% response.write Trim(rs.fields("Glass")) %></option>"
<!--#include file="QSU.inc"-->
</select>
</div>
	  
			<div class="row">
                <label>Type</label>
                <select name="inventorytype">
					<option value="<% response.write Trim(rs.fields("Type")) %>"><% response.write Trim(rs.fields("Type")) %></option>
					<option value="SU">SU</option>
					<option value="OV">OV</option>
					<option value="SP">SP</option>
					<option value="SU/OV">SU/OV</option>
					<option value="Door">Door</option>
					<option value="Other">Other</option>
				</select>
            </div>	  
           	<div class="row">
                <label>Opt File</label>
                <input type="text" name='opfile' id='opfile' value="<% response.write Trim(rs.fields("opfile")) %>">
            </div>
			
			<div class="row">
                <label># of Lites</label>
                <input type="number" name='Lites' id='Lites'  value="<% response.write Trim(rs.fields("LITES")) %>">
            </div>

			<div class="row">
                <label>Bending File</label>
                <input type="text" name='bendfile' id='bendfile'  value="<% response.write Trim(rs.fields("bendfile")) %>">
            </div>
			<div class="row">
                <label>Opt Date</label>
												<%
				InputDate = rs.fields("OpDate")
     mm = Month(InputDate)
     dd = Day(InputDate)
     yy = Year(InputDate)
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
     DateInput = yy & "-" & mm & "-" & dd 
				%>
                <input type="date" name='opdate' id='opdate' value='<%= DateInput %>'>
            </div>
			<div class="row">
                <label>Shift</label>
                <input type="text" name='shift' id='shift'  value="<% response.write Trim(rs.fields("shift")) %>">
            </div>
			<div class="row">
                <label>Employee</label>
                <input type="text" name='employee' id='employee' value="<% response.write Trim(rs.fields("employee")) %>">
            </div>

			<div class="row">
                <label>Glass Cut Date</label>
												<%
				InputDate = rs.fields("GlassCutDate")
     mm = Month(InputDate)
     dd = Day(InputDate)
     yy = Year(InputDate)
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
     DateInput = yy & "-" & mm & "-" & dd 
				%>
                <input type="date" name='glasscutdate' id='glasscutdate' value='<%= DateInput %>'>
            </div>

			<div class="row">
                <label>Pack Date</label>
												<%
				InputDate = rs.fields("PackDate")
     mm = Month(InputDate)
     dd = Day(InputDate)
     yy = Year(InputDate)
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
     DateInput = yy & "-" & mm & "-" & dd 
				%>
                <input type="date" name='PackDate' id='PackDate' value='<%= DateInput %>'>
            </div>
			
			<div class="row">
                <label>Skid Number</label>
                <input type="text" name='Skid' id='Skid'  value="<% response.write Trim(rs.fields("Skid")) %>">
            </div>
			
			<div class="row">
                <label>Ship Date</label>
								<%
				InputDate = rs.fields("ShipDate")
     mm = Month(InputDate)
     dd = Day(InputDate)
     yy = Year(InputDate)
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
     DateInput = yy & "-" & mm & "-" & dd 
				%>
                <input type="date" name='ShipDate' id='ShipDate' value='<%= DateInput %>'>
            </div>
						
			<div class="row">
                <label>Received Date</label>
				<%
	InputDate = rs.fields("ReceivedDate")
     mm = Month(InputDate)
     dd = Day(InputDate)
     yy = Year(InputDate)
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
     DateInput = yy & "-" & mm & "-" & dd 
				%>
               <input type="date" name='ReceivedDate' id='ReceivedDate' value='<%= DateInput %>'>
            </div>
			
			<div class="row">
                <label>Back-order Lites</label>
                <input type="number" name='BackOrder' id='BackOrder'  value="<% response.write Trim(rs.fields("BACKORDER")) %>">
            </div>
			<div class="row">
                <label>Back-order Names</label>
                <input type="text" name='backordertext' id='backordertext'  value="<% response.write Trim(rs.fields("backordertext")) %>">
            </div>
            
						<input type="hidden" name='OPID' id='OPID' value="<%response.write OPID %>" />
						
						
						
						
						
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="OptLog.action='OptimizationLogEditConf.asp'; OptLog.submit()">Submit Changes</a><BR>
		'<a class="redButton" onClick="OptLog.action='OptimizationLogDelConf.asp'; OptLog.submit()">Delete Entry</a><BR>
	

            
            </form>
                        
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

