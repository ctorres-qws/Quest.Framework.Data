                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


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

EMPLOYEE = request.querystring("EMPLOYEEID")
DEPTVAR = "ASSEMBLY"
ERROR = "Already Scanned - Not Sent"

bc = request.querystring("window")
jobname = Left(bc, 3)
if inStr(1, bc, "-", 0) = 5 then
floor = Mid(bc, 4, 1)
tag = Mid(bc, 5, 5)
END IF
'
if inStr(1, bc, "-", 0) = 6 then
floor = Mid(bc, 4, 2)
tag = Mid(bc, 6, 5)
end if

if inStr(1, bc, "-", 0) = 7 then
floor = Mid(bc, 4, 3)
tag = Mid(bc, 7, 5)
end if

RESPONSE.WRITE "JOB: " & JOBNAME & " <BR>"
RESPONSE.WRITE "FLOOR: " & FLOOR & " <BR>"
RESPONSE.WRITE "WINDOW: " & TAG

%>

</body>
</html>
