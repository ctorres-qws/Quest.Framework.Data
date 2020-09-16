                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Cut File Progress Table - Searches for a Job and Floor and then finds the Cut / Not Cut Percentage-->
<!-- Created March 9, 2015 by Michael Bernholtz - Requested by Michael Angel-->

<!-- Updated at Shaun Levy request to Remove Nothing DH from HCUT, truing up the numbers, Michael Bernholtz March 2017-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Cut File Progress</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>

<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
 <script>
 function example() {
    alert('test');
}
function disable(){
  document.getElementById('CutFile').setAttribute("disabled", "disabled"); 
}
</script>



    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
        </div>
<form id="CutFile" title="Cut File Progress" class="panel" name="CutFile" action="cutFileProgress.asp" method="GET" selected="true">
<fieldset>
<%
Job = UCASE(Request.querystring("Job"))
FLoor = UCase(Request.querystring("Floor"))
%>
       

         <div class="row">   
            <label>Job </label>
            <input type="text" name='Job' id='Job' value = "<%response.write Job%>" >
		</div>
		<div class="row">     
            <label>Floor </label>
            <input type="text" name='Floor' id='Floor' value = "<%response.write Floor%>">
		</div>
		<a id ='buttonname' class="whiteButton" onclick="disable(); buttonname.value = 'PleaseWait'; CutFile.submit();" return true;> View Cut File Progress</a><BR>
</fieldset>
       <div>
        <h1>Progress of <% Response.write JOB & " " & Floor%></h1>
		<h2><table border='1' class='sortable'><tr><th>Type</th><th>Cut Status</th><th>Total Items</th><th>Last Cut Date</th></tr>
		
				
		
<% 


if JOB <> "" and FLOOR <>"" then 
	i= 0
	cycle = "c0"

		CutData = 0
		CutDataTotal = 0
	Do until i = 8
	Select Case cycle
			Case "c0"
				cycle = "c1"
			Case "c1"
				cycle = "c2"
			Case "c2"
				cycle = "c3"
			Case "c3"
				cycle = "c4"
			Case "c4"
				cycle = "c5"
			Case "c5"
				cycle = "c6"
			Case "c6"
				cycle = "c7"
			Case "c7"
				cycle = "c8"
		End Select

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT cstatus, cdate FROM CUT_" & JOB & FLOOR & Cycle
	On Error Resume Next  
	rs.Open strSQL, DBConnection
	On Error GoTo 0
	If rs.State = 1 Then 
		CutDate = "01/01/1999"
		do while not rs.eof
		CutDataTotal = CutDataTotal + 1
			if rs("cstatus") = -1 then
				CutData = CutData + 1
		end if
			if rs("cDate") > CutDate then
				CutDate = rs("cDate")
			end if
			
			rs.movenext
		loop
		 if CutDate = "01/01/1999" then
		 CutDate = ""
		 end if
		rs.close
		set rs = nothing
	end if

	i= i+1
	loop
	%>
	<% 
	i= 0
	cycle = "c0"
	HCutData = 0
	HCutDataTotal = 0
	Do until i = 8

	Select Case cycle
			Case "c0"
				cycle = "c1"
			Case "c1"
				cycle = "c2"
			Case "c2"
				cycle = "c3"
			Case "c3"
				cycle = "c4"
			Case "c4"
				cycle = "c5"
			Case "c5"
				cycle = "c6"
			Case "c6"
				cycle = "c7"
			Case "c7"
				cycle = "c8"
		End Select

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT cstatus, cdate,pName FROM HCUT_" & JOB & FLOOR & Cycle
	On Error Resume Next  
	rs.Open strSQL, DBConnection
	On Error GoTo 0
	If rs.State = 1 Then 
		HCutDate = "01/01/1999"
		do while not rs.eof
		if rs("pName") = "Nothing DH" then
		else
			HCutDataTotal = HCutDataTotal + 1
			if rs("cstatus") = -1 then
				HCutData = HCutData + 1
			end if
			if rs("cDate") > HCutDate then
				HCutDate = rs("cDate")
			end if
		end if
			
			rs.movenext
		loop
		 if HCutDate = "01/01/1999" then
		 HCutDate = ""
		 end if
		rs.close
		set rs = nothing
	end if

	i= i+1
	loop
	%>
	<% 
	i= 0
	cycle = "c0"
	STOPDataTotal = 0
	STOPData = 0
	Do until i = 8
		
	Select Case cycle
			Case "c0"
				cycle = "c1"
			Case "c1"
				cycle = "c2"
			Case "c2"
				cycle = "c3"
			Case "c3"
				cycle = "c4"
			Case "c4"
				cycle = "c5"
			Case "c5"
				cycle = "c6"
			Case "c6"
				cycle = "c7"
			Case "c7"
				cycle = "c8"
		End Select

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT cStatus, ddate FROM STOP_" & JOB & FLOOR & Cycle
	On Error Resume Next  
	rs.Open strSQL, DBConnection
	On Error GoTo 0
	If rs.State = 1 Then 

		STOPDate = "01/01/1999"
		do while not rs.eof
			StopDataTotal = StopDataTotal + 1
			if rs("cStatus") = -1 then
				StopData = StopData + 1
			end if
			if rs("dDate") > StopDate then
				StopDate = rs("dDate")
			end if
			
			rs.movenext
		loop
		 if StopDate = "01/01/1999" then
		 StopDate = ""
		 end if
		rs.close
		set rs = nothing
	end if

	i= i+1
	loop
	%>
	<% 
	i= 0
	cycle = "c0"
	DMSAWData = 0
	DMSAWDataTotal = 0
	

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT cstatus, cdate FROM DMSAW_" & JOB & FLOOR
	On Error Resume Next  
	rs.Open strSQL, DBConnection
	On Error GoTo 0
	If rs.State = 1 Then 
		DMSAWDate = "01/01/1999"
		do while not rs.eof
		DMSAWDataTotal = DMSAWDataTotal +1
			if rs("cstatus") = -1 then
				DMSAWData = DMSAWData + 1
		end if
			if rs("cDate") > DMSAWDate then
				DMSAWDate = rs("cDate")
			end if
			
			rs.movenext
		loop
		 if DMSAWDate = "01/01/1999" then
		 DMSAWDate = ""
		 end if
		rs.close
		set rs = nothing
	end if

end if
%>
		
		<tr><td>CUT</td><td><% response.write CutData %></td><td><% response.write CutDataTotal %></td><td><% response.write CutDate %></td></tr>
		<tr><td>HCUT</td><td><% response.write HCutData %></td><td><% response.write HCutDataTotal %></td><td><% response.write HCutDate %></td></tr>
		<tr><td>STOP</td><td><% response.write StopData %></td><td><% response.write StopDataTotal %></td><td><% response.write StopDate %></td></tr>
		<tr><td>DMSAW</td><td><% response.write DMsawData %></td><td><% response.write DMSAWDataTotal %></td><td><% response.write DMSAWDate %></td></tr>
		</table></h2>
		         
</div>      

<%
' Close the Connection to Database
DBConnection.close 
set DBConnection = nothing
%>
               
    </ul>    
</form>	
             
</body>
</html>
