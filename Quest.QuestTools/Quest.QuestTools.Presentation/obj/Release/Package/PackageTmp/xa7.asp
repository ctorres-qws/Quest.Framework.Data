<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
             <!--#include file="dbpath.asp"-->
			<!-- Update May 2014 - Scan before 3:00am count as yesterday, Today and Yesterday Include -->
			
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>CDN Assembly</title>
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
	window = request.QueryString("Window")      
	employeeID = 0
	employeeID = request.QueryString("EmployeeID")
	STAMP = REQUEST.QueryString("STAMP")
	if bc = "&TA" then
		bc ="STA"
	end if
	EMPLOYEE = request.querystring("EMPLOYEEID")
	DEPTVAR = "ASSEMBLY"
	ERROR = "Already Scanned - Not Sent"
	

	bc = request.querystring("window")
	
	if not bc ="" then
		Label = bc
		jobname = Left(bc, 3)
		Label = Right(Label, Len(Label)-3)
		floor = Left(Label,  inStr(1, label, "-", 0) - 1)
		tag = Right(Label, Len(label)- inStr(1, Label, "-", 0))
	end if
if floor >= "0" then

	sizecheckid = 0

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			If gstr_ErrMsg="" Then Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

end if

Function Process(isSQLServer)

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_BARCODE WHERE Barcode='" & window & "'"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.filter = "BARCODE = '" & bc & "' AND DEPT = 'ASSEMBLY'"
  
if RS.EOF then
sizeCheckID= 1
	if (Len(employee) = 4 or Len(employee)= 5) AND Len(bc) > 5 then

		rs.addnew 
		rs.fields("BARCODE") = bc
		rs.fields("JOB") = jobname
		rs.fields("FLOOR") = floor
		rs.fields("TAG") = tag
		rs.fields("DEPT") = DEPTVAR
		rs.fields("EMPLOYEE") = EMPLOYEE
		rs.fields("DATETIME") = STAMPVAR
		rs.fields("FIRST") = FIRST
		rs.fields("LAST") = LAST
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
				
		RS.UPDATE

	else 
		sizeCheckID= 2
		error = "Wrong Barcode or Invalid Employee #, Try Again"
		gstr_ErrMsg = "Err"
	end if

Else
sizeCheckID= 2
	error = "Barcode Already Scanned"
	gstr_ErrMsg = "Err"
End if

DbCloseAll

End Function

 %>

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ScanHome.HTML" target="_self">Scan Tools</a>
        </div>
   
   
   
    <form id="assy1" title="Assembly Scan" class="panel" name="assy1" action="xa7.asp" method="GET" selected="true">
         <% if employeeID = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if sizecheckid = 1 AND (len(employee) = 4 or len(employee) = 5) then
				response.write EMPLOYEE & " - " & window & " - Sent <BR>" 
				else
				response.write error & sizeCheckID
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Employee # Scan</h2>
        <fieldset>
       
            
         <div class="row">
            <label>Employee#</label>
            <input type="text" name='employeeID' id='inputbce' >
         </div>
            
        <div class="row">
            <label>Window</label>
            <input type="text" name='window' id='inputbcw' >
        </div>

 
        </fieldset>
		
		<script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
		function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
			var textbox = document.getElementById("inputbcw");
			if ( barcode.length == 4) {
				textbox = document.getElementById("inputbce");
			}
			if ( barcode.length == 5) {
				textbox = document.getElementById("inputbce");
			}
			textbox.value = barcode;
    
}
			
        </script>
        <BR>
        <a class="whiteButton" href="javascript:assy1.submit()">Submit</a>
            
            
            
            </form>
</body>
</html>
