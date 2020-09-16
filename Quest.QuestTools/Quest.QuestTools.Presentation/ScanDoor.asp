<!--#include file="dbpath.asp"-->                    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            
			<!-- Door Scan based on Awning Scan -->
			<!-- Update May 2014 - Scan before 3:00am count as yesterday, Today and Yesterday Include -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Door Scan</title>
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
error = ""
IDEntry = 0
EMPLOYEE = request.querystring("EMPLOYEEID")
bc = UCASE(request.querystring("window"))
bc = Replace(bc, "+", "-")
ScanType = REQUEST.QueryString("ScanType")
if ScanType = "" then
	ScanType= "SlidingFrame"
end if


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

	if Left(bc,2) = "GT" then
		GlassID = ""
		GlassID = Right(bc, Len(bc)-2)
	
		Set rs2 = Server.CreateObject("adodb.recordset")
		strSQL2 = "SELECT top 10 * FROM Z_GLASSDB WHERE ID=" & GlassID & " ORDER BY ID ASC"
		rs2.Cursortype = 2
		rs2.Locktype = 3
		rs2.Open strSQL2, DBConnection
		
		if rs2.eof then
		else
			jobname = rs2("Job")
			floor = rs2("Floor")
			Tag = rs2("Tag")
			
			rs2("GlazedDate") = DATE
			rs2.update
		end if
		
		
		rs2.close
		set rs2 = nothing
		
	

		sizecheckid = 0
	
	else
	
		jobname = Left(bc, 3)
		Marker = inStr(1, bc, "-", 0)
		floor = Mid(bc, 4, Marker - 4)
		Tag = Right(bc, Len(bc) - Marker  )
		sizecheckid = 0
	end if
end if




If Len(bc) > 0 Then
	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select
End If

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODESW WHERE Barcode='" & bc & "'"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

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
			if isSQLServer then
				rs.fields("ID") = IDEntry
			end if
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
			if not isSQLServer then
				rs.movelast
				IDEntry = rs("ID")
			end if
		else 
			if Len(EMPLOYEE) <4 then
				error = "Not a Valid Employee ID, Try Again"
			else
				error = bc & ": Not Valid, Try Again"
			end if
			IsError = True
		end if

	end if

DbCloseAll

End Function

%>



</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
     <a class="button leftButton" type="cancel" href="ScanHome.HTML" target="_self">Scan Tools</a>
        </div>
   
   
   
    <form id="igline" title="Door Scan" class="panel" name="igline" action="ScanDoor.asp" method="GET" selected="true">
         <h2>Door Scan: <%Response.write ScanType%></h2>
		 
		 <%
		 
		 
		 SlidingFrameColour ="whiteButton"
		 SlidingGlassColour ="whiteButton"
		 SwingDoorColour ="whiteButton"
		 SunviewDoorColour ="whiteButton"
		 
		 Select Case ScanType
			Case "SlidingFrame"
				SlidingFrameColour ="greenButton"
			Case "SlidingGlass"
				SlidingGlassColour ="greenButton"
			Case "SwingDoor"
				SwingDoorColour ="greenButton"
			Case "SunviewDoor"
				SunviewDoorColour ="greenButton"
		 End Select
		 %>
		 
		 
		<div class="row">
			<table align = 'center'>
			<tr>
				<td><button class="<%response.write SlidingFrameColour%>"><a href="ScanDoor.asp?ScanType=SlidingFrame" target = "_self" >Sliding Frame</a></button> </td>
				<td><button class="<%response.write SlidingGlassColour%>"><a href="ScanDoor.asp?ScanType=SlidingGlass" target = "_self" >Sliding Glass</a></button></td>
			</tr>
			<tr>
				<td><button class="<%response.write SwingDoorColour%>"><a href="ScanDoor.asp?ScanType=SwingDoor" target = "_self" >Swing Door</a></button></td>
				<td><button class="<%response.write SunviewDoorColour%>"><a href="ScanDoor.asp?ScanType=SunviewDoor" target = "_self" >Sunview Door</a></button></td>
			</tr>
			</table>
		</div>
		
		 <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if IsError = False then
				response.write bc & " - Sent"
				' & DetailsID
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
			<BR>
        <fieldset>
       
	   
	        <div class="row">
                <label>Employee#</label>
                <input type="text" name='employeeID' id='inputbce' >
            </div>
	   
            
			<div class="row">
                <label>Door Glass</label>
                <input type="text" name='window' id='inputbcw' >
				<input type="hidden" name='ScanType' id='ScanType' value='<%response.write ScanType%>' />
            </div>
            </fieldset>
			 <BR>
				<a class="whiteButton" href="javascript:igline.submit()">Submit</a>				
            </form>
			
<script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("inputbcw");
    
    if ( barcode.length == 4 ) {
        textbox = document.getElementById("inputbce");
    }
	if ( barcode.length == 5 ) {
        textbox = document.getElementById("inputbce");
    }
    
        textbox.value = barcode;

}
    
        </script>
		


</body>
</html>