                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <!--#include file="connect_barcodeqc.asp"-->
			<!-- Update May 2014 - Scan before 3:00am count as yesterday, Today and Yesterday Include -->
			<!-- Update June 2014 - Glazing 2 for different Employee Number  & remove ability to type --> 

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
	 
	 
	 
employeeID = request.QueryString("EmployeeID")
EMPLOYEE = request.querystring("EMPLOYEEID")
DEPTVAR = "GLAZING"
bc = request.querystring("window")

 %>

     

</head>

<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Scan" target="_webapp">Scan Tools</a>
        </div>
   
   
   
    <form id="Backorder" title="Backorder" class="panel" name="BackOrder" action="xa9BackOrder.asp" method="GET" selected="true">

                <H2><%Response.write bc%></H2>

        <fieldset>
       
            
         <div class="row">
                <label>Backorder Reason</label>
                <select name="backorder">
				<option value = "0"> Not Backorder</Option>
<%				
Set rsReason = Server.CreateObject("adodb.recordset")
strSQL = "Select * FROM X_Backorder_Reason"
rsReason.Cursortype = 2
rsReason.Locktype = 3
rsReason.Open strSQL, DBConnection
if rsReason.eof then 
else
rsReason.movefirst
end if
do while not rsReason.eof

response.write "<option value='" & rsReason("ID") & "' >"
response.write rsReason("Reason")
response.write "</option>"
rsReason.movenext
loop

%>
            </div>

		<div class="row">
            <label>Section</label>
            <select name= 'Section' id = 'Section'>
				<option value="Multiple"><Multiple</option>
				<option value="1">1</option>
				<option value="2">2</option>
				<option value="3">3</option>
				<option value="4">4</option>
				<option value="5">5</option>
				<option value="6">6</option>
				<option value="7">7</option>
				<option value="8">8</option>
				
			</select>
        </div>

            
            	<input type="hidden" name='employeeID'  value ="<% response.write EMPLOYEEID %>" >
				<input type="hidden" name='window' value ="<% response.write bc %>" >
              
        </fieldset>
        <BR>
        <a class="whiteButton" href="javascript:Backorder.submit()">Submit</a>
            
            
            
            </form>
</body>
</html>
