                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodega.asp"-->

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

STAMPVAR = year(now) & " " & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

bc = request.querystring("barcode")

DEPTVAR = "GLASSLINE"
ERROR = "Already Scanned - Not Sent"



bc = request.querystring("window")

if Left(bc,2) = "GT" then
gtactive = "HEY YO"
else


jobname = Left(bc, 3)
if inStr(1, bc, "-", 0) = 5 then
floor = Mid(bc, 4, 1)
tag = Mid(bc, 5, 5)
END IF

if inStr(1, bc, "-", 0) = 6 then
floor = Mid(bc, 4, 2)
tag = Mid(bc, 6, 5)
end if

if inStr(1, bc, "-", 0) = 7 then
floor = Mid(bc, 4, 3)
tag = Mid(bc, 7, 5)
end if

end if

glasstype = right(bc,2)

'if floor > "0" then

sizecheckid = 0

  Do while not rs.eof
  if rs("BARCODE") = bc AND rs("DEPT") = "GLASSLINE" then
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

if Len(bc) > 3 then

rs.addnew 
rs.fields("BARCODE") = bc
rs.fields("JOB") = jobname
rs.fields("FLOOR") = floor
rs.fields("TAG") = tag
rs.fields("DEPT") = DEPTVAR
rs.fields("DATETIME") = STAMPVAR
rs.fields("TYPE") = glasstype
rs.fields("DAY") = cday
rs.fields("MONTH") = cmonth
rs.fields("YEAR") = cyear
rs.fields("WEEK") = weeknumber
rs.fields("TIME") = cctime
RS.UPDATE

else 
error = "Wrong Barcode, Try Again"
end if

end if

'end if
  
 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Scan" target="_webapp">Scan Tools</a>
        </div>
   
   
   
    <form id="igline" title="Glass Line Scan" class="panel" name="igline" action="ga1.asp" method="GET" selected="true">
         <% if bc = "" then
		 response.write ""
		 else %>
			<div class="row">
                <label><% if sizecheckid = 0 then
				response.write window & " - Sent <BR>"  & gtactive
				else
				response.write error
				end if				
				%></label>
              
            </div>
            <% 			
			end if %>
        <h2>Scan Glass</h2>
        <fieldset>
       
            
         <div class="row">

                        <div class="row">
                <label>Window</label>
                <input type="text" name='window' id='inputbcw' >
            </div>
            
                              	
                <% 'response.write window & "<Br>" & employeeID %>
            
            
                  <script type="text/javascript">
				  
				 		  
            function callback1(barcode) {
                var barcodeText = "BARCODE:" + barcode;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
        
            }
            
	function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
    var textbox = document.getElementById("inputbcw");
    
    
        textbox.value = barcode;
		igline.submit();
    
}
			
			

    

        </script>
        
        
           
         
 
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:igline.submit()">Submit</a>
            
            
            
            </form>
</body>
</html>
