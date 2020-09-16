<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Bill of Material Summary Page by Job and Floor-->
<!-- Gets Job and Floor from ROD_BOM_JF.asp and uses ROD_BOM_FINDER.asp in 172.18.13.31\quest to determine totals-->
<!-- Created September 27th, by Michael Bernholtz - For Jody Cash and Daniel Zalcman-->
<!-- Report to be used by Kirk Campbell and Carlos to get correct Hardware -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Service Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="ROD_BOM_JF.asp" target="_self">Back</a>
        </div>
   
   
         

        <ul id="Profiles" title="Glass Report - Recut" selected="true">
<%

BOM_JOB = request.querystring("JOB")
BOM_FLOOR = request.querystring("FLOOR")
'Commented in for Testing Purposes.
'BOM_JOB = "AAA"
'BOM_FLOOR = "98"

On Error Resume Next	
	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM [ROD_" & BOM_JOB & BOM_FLOOR & "] WHERE LEFT(RCODE,1) = 'H' ORDER BY BOM ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

IF Err.Number <>0 then

Response.write "<li> ROD Table was not processed. </li>"
Response.write "<li> Please Process DMSAW and then ROD before continuing</li>"
Response.write "<li>" & BOM_JOB & BOM_FLOOR & "</li>"

Else

	BOM1 = "NA"
	BOM2 = "NA"
	TKeeper = 0
	TPin = 0
	TOpeningMech = 0
	TTransmission = 0
	TShavedTransmission = 0
	TotalBOMCOUNTER = 0
	TotalKeeper = 0
	TotalPin = 0
	TotalOpeningMech = 0
	TotalTransmission = 0
	TotalShavedTransmission = 0


	response.write "<li class='group'>Bill Of Materials Summary Sheet for " & BOM_JOB & " " & BOM_FLOOR & "</li>"
	response.write "<li><table border='1' class='sortable'><tr><th>BOM</th><th>COUNT</th><th>Keeper</th><th>Pin</th><th>Opening Mech</th><th>Transmission</th><th>Shaved Transmission</th></tr>"


	if rs.eof then
	else

		Do while not rs.eof
			BOM2 = BOM1
			BOM1 = rs("BOM")
			IF BOM1 = BOM2 then
			BOMCOUNTER = BOMCOUNTER + 1
			TotalBOMCOUNTER = TotalBOMCOUNTER + 1
			BOM = BOM1
			%>
			<!--#include file="ROD_BOM_FINDER.asp"-->
			<%
			TKeeper = TKeeper + AKeeper
			TPin = TPin + APin
			TOpeningMech = TOpeningMech + AOpeningMech
			TTransmission = TTransmission + ATransmission
			TShavedTransmission = TShavedTransmission + AShavedTransmission
				TotalKeeper = TotalKeeper + AKeeper
				TotalPin = TotalPin + APin
				TotalOpeningMech = TotalOpeningMech + AOpeningMech
				TotalTransmission = TotalTransmission + ATransmission
				TotalShavedTransmission = TotalShavedTransmission + AShavedTransmission
			
			else
				if BOM2 = "NA" then
				else
					Response.write "<TR>"
					Response.write "<TD>" & BOM & "</TD>"
					Response.write "<TD>" & BOMCOUNTER & "</TD>"
					Response.write "<TD>" & TKeeper & "</TD>"
					Response.write "<TD>" & TPin & "</TD>"
					Response.write "<TD>" & TOpeningMech & "</TD>"
					Response.write "<TD>" & TTransmission & "</TD>"
					Response.write "<TD>" & TShavedTransmission & "</TD>"
					Response.write "</TR>"
					
				End if
			
			BOM = BOM1
			%>
			<!--#include file="ROD_BOM_FINDER.asp"-->
			<%
			BOMCOUNTER = 1
			TotalBOMCOUNTER = TotalBOMCOUNTER + 1
			TKeeper = AKeeper
			TPin = APin
			TOpeningMech = AOpeningMech
			TTransmission = ATransmission
			TShavedTransmission = AShavedTransmission
				TotalKeeper = TotalKeeper + AKeeper
				TotalPin = TotalPin + APin
				TotalOpeningMech = TotalOpeningMech + AOpeningMech
				TotalTransmission = TotalTransmission + ATransmission
				TotalShavedTransmission = TotalShavedTransmission + AShavedTransmission
			
			end if 
		
		rs.movenext
		loop
			Response.write "<TR>"
			Response.write "<TD>" & BOM1 & "</TD>"
			Response.write "<TD>" & BOMCOUNTER & "</TD>"
			Response.write "<TD>" & TKeeper & "</TD>"
			Response.write "<TD>" & TPin & "</TD>"
			Response.write "<TD>" & TOpeningMech & "</TD>"
			Response.write "<TD>" & TTransmission & "</TD>"
			Response.write "<TD>" & TShavedTransmission & "</TD>"
			Response.write "</TR>"
			Response.write "<TR>"
			Response.write "<TD>Total</TD>"
			Response.write "<TD><B>" & TotalBOMCOUNTER & "</B></TD>"
			Response.write "<TD><B>" & TotalKeeper & "</B></TD>"
			Response.write "<TD><B>" & TotalPin & "</B></TD>"
			Response.write "<TD><B>" & TotalOpeningMech & "</B></TD>"
			Response.write "<TD><B>" & TotalTransmission & "</B></TD>"
			Response.write "<TD><B>" & TotalShavedTransmission & "</B></TD>"
			Response.write "</TR>"
		

	end if
	%>    


			
	<% 
		
	response.write "</table></li>"

	


	response.write "<li class='group'>Individual BOM per Awning/Casement</li>"
	response.write "<li><table border='1' class='sortable'><tr><TH>OV Tag</TH><th>BOM</th><th>Keeper</th><th>Pin</th><th>Opening Mech</th><th>Transmission</th><th>Shaved Transmission</th></tr>"
	
	rs.close
	set rs = nothing
		
	Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM [ROD_" & BOM_JOB & BOM_FLOOR & "] WHERE LEFT(RCODE,1) = 'H' ORDER BY TAG ASC"
	rs2.Cursortype = 2
	rs2.Locktype = 3
	rs2.Open strSQL2, DBConnection
	
	do while not rs2.eof
		BOM = rs2("BOM")
		%>
		<!--#include file="ROD_BOM_FINDER.asp"-->
		<%
			Response.write "<TR>"
			Response.write "<TD>" & RS2("TAG") & "</TD>"
			Response.write "<TD><B>" & BOM & "</B></TD>"
			Response.write "<TD>" & AKeeper & "</TD>"
			Response.write "<TD>" & APin & "</TD>"
			Response.write "<TD>" & AOpeningMech & "</TD>"
			Response.write "<TD>" & ATransmission & "</TD>"
			Response.write "<TD>" & AShavedTransmission & "</TD>"
			Response.write "</TR>"
	rs2.movenext
	loop
	
	rs2.close
	set rs2 = nothing
	
	response.write "</table></li>"
	
	
	

		
	DBConnection.close 
	set DBConnection = nothing

	
END IF ' No Table	
	%>

    </ul>        
            
            
       
            
              
               
                
             
               
</body>
</html>
