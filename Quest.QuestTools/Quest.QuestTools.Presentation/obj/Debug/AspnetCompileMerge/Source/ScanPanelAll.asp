<!--#include file="dbpath.asp"-->                    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            
			<!-- Panel Scan form based on Panel Scan program with 4 distinct Scan options - Allows the scanning of Panel items by -->
			<!-- Scan Cut / Bend / Shipped / Received  -->
			<!-- Changed February 2018  to include all Machines not just Cut/Bend-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Panel Scan</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>

     <!--#include file="TodayAndYesterday.asp"-->

	 <% 

EMPLOYEE = request.querystring("EMPLOYEEID")
bc = UCASE(request.querystring("window"))

	 Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODEP WHERE BARCODE='" & bc & "'"
'if Len(bc) >2 then
rs.Cursortype = 2
rs.Locktype = 3
'else
'rs.Cursortype = GetDBCursorType
'rs.Locktype = GetDBLockType	
'end if
rs.Open strSQL, DBConnection

ScanType = REQUEST.QueryString("ScanType")



DEPTVAR = ScanType





OUTDATE = DATE
'The date to add on the item marked completed
IsError = False
' Reset the Variable for locating an Error
IDFound = False
' Flag to see if the Scanned Item is found on the Z_GLASSDB
DetailsID =""
' Details to show for item successfully marked complete
RecordLocated=0

if Len(bc) >2 then
	jobname = Left(bc, 3)
	Marker = inStr(1, bc, "-", 0)
	floor = Mid(bc, 4, Marker - 4)
	Tag = Right(bc, Len(bc) - Marker  )

	sizecheckid = 0
end if
	
	
	Do while not rs.eof
	

		' Check to see if the item has already been scanned
		if rs("BARCODE") = bc AND rs("DEPT") = DEPTVAR then
			sizecheckid = rs("ID")
			error = bc & ": Already Scanned"
			IsError = True
		end if
		'Check to see if the item has already been UNSCANNED and rescan the same one, 
		if rs("BARCODE") = bc AND rs("DEPT") = "UNSCAN" then
			sizecheckid = rs("ID")
			RecordLocated = rs("ID")
			'rs.Move RecordLocated
			rs.fields("DEPT") = DEPTVAR
			rs.UPDATE
			DetailsID = "JOB: " &jobname & " FLOOR: " & floor & " TAG: " & tag
		end if
		
		
	rs.movenext
	loop
  'Create new if did not get caught above
	if sizecheckid = 0 then

		if Len(bc) > 3 AND Len(EMPLOYEE) > 3 then

			rs.addnew 
			rs.fields("BARCODE") = bc
			rs.fields("JOB") = jobname
			rs.fields("FLOOR") = floor
			rs.fields("TAG") = tag
			rs.fields("DEPT") = DEPTVAR
			rs.fields("DATETIME") = STAMPVAR
			rs.fields("TYPE") = glasstype
			rs.fields("Employee") = EMPLOYEE
			rs.fields("TIME") = cctime
			if hour(now) <= 6 then  ' Changed to 6am from 3 by Michael Bernholtz February 2018
				rs.fields("DAY") = cYesterday
				rs.fields("MONTH") = cMonthy
				rs.fields("YEAR") = cYeary
				rs.fields("WEEK") = weekNumbery
			else
				rs.fields("DAY") = cDay
				rs.fields("MONTH") = cMonth
				rs.fields("YEAR") = cYear
				rs.fields("WEEK") = weekNumber
			end if	
			
			
			rs.UPDATE
			DetailsID = "JOB: " &jobname & " FLOOR: " & floor & " TAG: " & tag

		else 
			if Len(EMPLOYEE) <4 then
				error = "Not a Valid Employee ID, Try Again"
			else
				error = bc & ": Not Valid, Try Again"
			end if
			IsError = True
		end if

	end if

  
 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
     <a class="button leftButton" type="cancel" href="ScanHome.HTML" target="_self">Scan Tools</a>
        </div>
   
   
   
    <form id="igline" title="Panel Scan" class="panel" name="igline" action="ScanPanelAll.asp" method="GET" selected="true">
         <h2>Panel Scan</h2>
		 

		 
		 
		<div class="row">
			<table>
			<tr>
				<td><a class="whiteButton" href="ScanPanelCut.asp?Machine=1" target = "_self" >Euromac Alum</a> </td>
				<td><a class="whiteButton" href="ScanPanelBend.asp?Machine=1" target = "_self" >All Steel</a> </td>
			</tr>
			<tr>
				<td><a class="whiteButton" href="ScanPanelCut.asp?Machine=2" target = "_self" >HK Laser</a> </td>
				<td><a class="whiteButton" href="ScanPanelBend.asp?Machine=2" target = "_self" >Toskar</a> </td>
			</tr>
			<tr>
				<td><a class="whiteButton" href="ScanPanelCut.asp?Machine=3" target = "_self" >Durma</a> </td>
				<td><a class="whiteButton" href="ScanPanelBend.asp?Machine=3" target = "_self" >Schechtl</a> </td>
			</tr>
			
			<!--Old Options-->
			<tr>
				<td><a class="whiteButton"" href="ScanPanelCut.asp" target = "_self" >Sheet Cut</a> </td>
				<td><a class="whiteButton"" href="ScanPanelBend.asp" target = "_self" >Panel Bend </a> </td>
			</tr>
			<tr>
				<td><a class="whiteButton"" href="ScanPanelShip.asp" target = "_self" >Panel Sent </a> </td>
				<td><a class="whiteButton"" href="ScanPanelReceive.asp" target = "_self" > Panel Received </a> </td>
			</tr>
			
			</table>
		</div>
		

			<BR>       
				<a class="lightblueButton" href="PanelReport.asp?RangeView=Today" target = "_self" >Today's Scans</a>
				<a class="lightblueButton" href="PanelReportJobFloor.asp?JOB=<%response.write jobname %>& Floor=<%response.write Floor %>" target = "_self" >Job Floor Report</a>
            </form>
			
			
	
		
<%
On Error Resume Next
rs.close
set rs=nothing
rs2.close
set rs2=nothing

DBConnection.close
set DBConnection=nothing
%>	

</body>
</html>