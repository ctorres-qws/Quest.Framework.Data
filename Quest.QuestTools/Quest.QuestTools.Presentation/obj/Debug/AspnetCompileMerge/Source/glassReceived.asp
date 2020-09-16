<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--Optima Selection Page for Received Date, Joe wants to Add by PO-->
		<!--Created MArch 2015, at Request of Joe for adding a note to multiple items at once-->
		<!-- Sends to glassReceivedConf.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    <style type="text/css">
	ul{
    margin: 0;
    padding: 0;
   }
   </style>
       <style type="text/css">
   #Recived
    padding-left: 300px;
   }
   </style>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       <form id="Optima" action="glassReceived.asp" name="Optima"  method="GET" target="_self" selected="true" >  
        
		<h2><center>Enter PO or Cardinal Work Order and the Actual Date Received<center></h2>

		
		<fieldset>
		<!--#include file="dbpath.asp"-->

<%
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT PO, EXTORDERNUM, INTORDERNUM, COMPLETEDDATE FROM Z_GLASSDB ORDER BY PO,EXTORDERNUM, INTORDERNUM DESC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection
rs.filter = "COMPLETEDDATE = NULL OR COMPLETEDDATE = ''"
%>
 

		<div class="row">
			<label>Work Order</label>
			
			<Select name='WorkOrder' id='WorkOrder' >
			<%
			
			DO while not rs.eof
			if rs("PO") <> "" then
				if rs("PO")<> lastpo then
					response.write "<option value = '" & rs("PO") & "'> " & rs("PO") & "</option>"
				end if
			lastpo= rs("po")
			end if
			if rs("ExtOrderNum") <> "" then
				if rs("ExtOrderNum")<> lastEON then
					response.write "<option value = '" & rs("ExtorderNum") & "'> " & rs("ExtorderNum") & "</option>"
				end if
			lastEON = rs("ExtOrderNum")	
			end if
			if rs("IntOrderNum") <> "" then
				if rs("INTORDERNUM")<> lastION then
					response.write "<option value = '" & rs("IntOrderNum") & "'> " & rs("IntOrderNum") & "</option>"
				end if
			lastION = rs("IntOrderNum")	
			end if
			rs.movenext
			loop
			
			rs.close
			set rs = nothing
			DBConnection.close
			SET DBConnection = nothing
			%>
			
			</select>
			
		</div>
	
		<div class="row">
		<% 
		PreSetTime = Date()
		%>		
                <label>Receive Date</label>
                <input type="text" name='Received' id='Received' value='<% response.write PreSetTime %>' />
        </div>
		    
		<a class="whiteButton" onClick="Optima.action='GlassReceivedConf.asp'; Optima.submit()">Add Received Dates</a><BR>
		</fieldset>
        <ul id="Profiles" title=" Optima Report" selected="true">
	

	</table>
	
      </ul>    
		</form>     
             
               
</body>
</html>
