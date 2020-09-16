                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodesw.asp"-->
			<!-- Update May 2014 - Scan before 3:00am count as yesterday, Today and Yesterday Include -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>iUI Theaters</title>
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
	 
	 
	 employeeID = 0
 employeeID = request.QueryString("EmployeeID")
  window = request.QueryString("Window") 
  

STAMP = REQUEST.QueryString("STAMP")
		%>
     <!--#include file="TodayAndYesterday.asp"-->
     <%

bc = request.querystring("barcode")

EMPLOYEE = request.querystring("EMPLOYEEID")
DEPTVAR = "SWING"
ERROR = "Already Scanned - Not Sent"

bc = request.querystring("window")
jobname = Left(bc, 3)
if inStr(1, bc, "-", 0) = 5 then
floor = Mid(bc, 4, 1)
tag = Mid(bc, 5, 5)
'code to scrub ! from tag is required here !!!!! 
else
floor = Mid(bc, 4, 2)
tag = Mid(bc, 6, 5)
end if
glasstype = right(bc,2)

if floor > "0" then

sizecheckid = 0

  Do while not rs.eof
  if rs("BARCODE") = bc AND rs("DEPT") = "SWING" then
  sizecheckid = rs("ID")
  end if
  rs.movenext
  loop
  
if sizecheckid = 0 then

'rs2.filter = "NUMBER = " & EMPLOYEE
'IF RS2("NUMBER") = "" THEN
'ELSE
'FIRST = RS2("FIRST")
'LAST = RS2("LAST")
'END IF

if Len(bc) > 5 then

rs.addnew 
rs.fields("BARCODE") = bc
rs.fields("JOB") = jobname
rs.fields("FLOOR") = floor
rs.fields("TAG") = tag
rs.fields("DEPT") = DEPTVAR
rs.fields("DATETIME") = STAMPVAR
rs.fields("TYPE") = glasstype
rs.fields("TIME") = cctime
	if hour(now) <= 2 then
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
error = "Wrong Barcode, Try Again"
end if

end if

end if
  
 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Scan" target="_webapp">Scan Tools</a>
        </div>
   
   
   
       <form id="scansw" title="Swing Door Scan" class="panel" name="scansw" action="sw1.asp" method="GET" selected="true">
         <% if employeeID = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if sizecheckid = 0 AND len(employee) = 4 then
				response.write EMPLOYEE & " - " & window & " - Sent <BR>" 
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Swing Door Scan</h2>
        <fieldset>
       
            
         <div class="row">
                <label>Employee#</label>
                <input type="text" name='employeeID' id='inputbce' >
            </div>
            
                        <div class="row">
                <label>Window</label>
                <input type="text" name='window' id='inputbcw' >
            </div>
            
                              	
                <% response.write window & "<Br>" & employeeID %>
            
            
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
    
        textbox.value = barcode;
    
}
			
			

    

        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:scansw.submit()">Submit</a>
            
            
            
            </form>
</body>
</html>
