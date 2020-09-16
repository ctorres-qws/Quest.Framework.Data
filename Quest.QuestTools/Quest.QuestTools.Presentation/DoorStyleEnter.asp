<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- May 2019 -->
<!-- DoorStyle pages collect information about Door types to get Machining Data for Emmegi Saws -->
<!-- Programmed by Michelle Dungo - At request of Ariel Aziza, using PanelStyle Pages as a template -->
<!-- DoorStyle.asp (General View) -- DoorStyleEditForm.asp (Manage Form) -- DoorStyleEditConf.asp (Manage Submit) -- DoorStyleEnter.asp (Enter Form)-- DoorStyleConf.asp (Enter Submit)--DoorStyleByJob.asp (view By Job filter) -->
<!-- SQL Table StylesDoor - NOT IN ACCESS -->
<!-- Date: June 14, 2019
	 Modified By: Michelle Dungo
	 Changes: Dynamic drop down of job names which only shows parent jobs taken from Quest DB			  
			  Add script to refresh parameters when job dropdown value changes
			  Added logic to recalculate next available door style name based on refreshed job name
-->
<%
JOB = REQUEST.QueryString("Job")

Set DBConnectionJob = Server.CreateObject("ADODB.Connection")
Set rsJob = Server.CreateObject("ADODB.Recordset")
DSN = GetConnectionStr(False) ' connect to access for job list
DBConnectionJob.Open DSN
strSQL = "SELECT DISTINCT Parent "
strSQL = strSQL & "FROM Z_Jobs Where Parent <> '' and Completed = False "
'strSQL = strSQL & " and Parent NOT LIKE 'AA%' " 'exclude test job
rsJob.Cursortype = 2
rsJob.Locktype = 3
rsJob.Open strSQL, DBConnectionJob
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Panel Styles</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>
  <script>
	function periodChange() {
		enter.action = "DoorStyleEnter.asp"
		enter.qwsAction.value = 'JOB_CHANGE';
		enter.submit();
	}
	function addStyle() {		
			enter.action = "DoorStyleConf.asp";
			enter.submit();
	}
  </script>	
</head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="DoorStyle.asp" target="_self">Styles</a>
    </div>
   
	<form id="enter" title="Enter New Door Style" class="panel" name="enter" action="DoorStylerEnter.asp" method="GET" target="_self" selected="true">
	<input name="qwsAction" value="" type="hidden">
	<h2>Enter New Door Style:</h2>
		<fieldset>
			<div class="row">
				<label>Job</label>
    			<Select name='Job' id='Job' onchange="periodChange()">
				<% 
				if JOB = "" then
					' do nothing
				else
				%>
					<option value = "<%response.write JOB%>" ><%response.write JOB%></option>
				<%
				end if
				%>
<%
While Not rsJob.EOF
	response.write "<option value=" & rsJob("Parent") & ">" & rsJob("Parent") & "</option>"
rsJob.MoveNext
Wend

rsJob.Close
DBConnectionJob.Close
set rsJob = nothing
set DBConnectionJob = nothing
%>				
				</Select>	
			</div>
			<div class="row">
				<label>Name</label>
<%
Set DBConnectionName = Server.CreateObject("ADODB.Connection")
Set rsName = Server.CreateObject("ADODB.Recordset")
DSN = GetConnectionStr(True) ' connect to sql server for style list
DBConnectionName.Open DSN

strSQL = "select CONCAT('D', ISNULL(MAX(CAST((SUBSTRING(name, PATINDEX('%[0-9]%', name), LEN(name))) as INT)),0) + 1) as max FROM  StylesDoor "
strSQL = strSQL & "Where Job ='" & JOB &"'"
'strSQL = strSQL & " and Parent NOT LIKE 'AA%' " 'exclude test job
rsName.Cursortype = 2
rsName.Locktype = 3
rsName.Open strSQL, DBConnectionName

response.write "<input readonly type='text' name='Name' id='Name' maxlength='4' title='Maximum 3 character Door Style Name' value='" &rsName("max")&"'>"

rsName.Close
DBConnectionName.Close
set rsName = nothing
set DBConnectionName = nothing
%>						
			</div>
			
			<div class="row">
				<label>Int</label>
				<Select name='IntDoorType'>
					<option value="Fapim">Fapim</option>
					<option value="Metra">Metra</option>
				</Select>				
			</div>			

			<div class="row">
				<label>Ext</label>
				<Select name='ExtDoorType'>
					<option value="Fapim">Fapim</option>
					<option value="Hopi">Hopi</option>					
					<option value="Metra">Metra</option>
					<option value="None">None</option>					
				</Select>				
			</div>
			
            <a class="whiteButton" href="javascript:addStyle();">Submit</a>
		</fieldset>
  
    </form>
</body>
</html>
