<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Update QC Inventory</title>
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
function startTime()
{
var today=new Date();
var h=today.getHours();
var m=today.getMinutes();
var s=today.getSeconds();
// add a zero in front of numbers<10
m=checkTime(m);
s=checkTime(s);
document.getElementById('clock').innerHTML=h+":"+m+":"+s;
t=setTimeout(function(){startTime()},500);
}

function checkTime(i)
{
if (i<10)
  {
  i="0" + i;
  }
return i;
}
</script>

<%
InventoryType = REQUEST.QueryString("InventoryType")
Identifier = request.querystring("Identifier")
%>

	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="QCInventoryUpdate.asp" target="_self">Update Stock</a>
    </div>
    
      
    
<form id="conf" title="Edit Stock" class="panel" name="conf" action="index.html#_QC" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>
  
<%                  

Dim Identifier, IdentifierID
'Identifier changes the entry between Serial Number for Glass, Box Number for Spacer, Lot Number for Sealant

	
		Select Case InventoryType
	Case "QCGlass"

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_Glass"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	
		rs.filter = "SerialNumber = '" & Identifier & "'"
		
		IdentifierID = "SerialNumber"
		IdentifierTitle = "Serial Number"
		
	Case "QCSpacer"
	
	
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_Spacer"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "BoxNumber = '" & Identifier & "'"
		
		IdentifierID = "BoxNumber"
		IdentifierTitle = "Box Number"
		
	Case "QCSealant"
	
	
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_Sealant"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "Lot Number = '" & Identifier &"'"
		
		IdentifierID = "LotNumber"
		IdentifierTitle = "Lot Number"
	
	End Select


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)




Quantity = Trim(REQUEST.QueryString("Quantity"))

IF not isNumeric(QUANTITY) then
	Quantity = 0
	Response.Write "<h2> " & Quantity & ": Is not a valid Number, Update was Cancelled </h2>" 
End if
UpdateDate = Date()

if not rs.eof then 
	rs.Fields("UpdateDate") = UpdateDate
	rs.Fields("Quantity") = rs.Fields("Quantity") + CINT(Quantity)
	rs.update

	
%>

    
<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	
Select Case InventoryType
	Case "QCGlass"
		Response.Write "<li>QC GLASS Quantity Updated:</li>"
		Response.Write "<li> Serial #: " & Identifier & "</li>"  	' Serial Number for Glass
		Response.Write "<li> Added: " & Quantity & "</li>"
		Response.Write "<li> New Quantity: " & rs.Fields("Quantity") & "</li>"
		Response.Write "<li> Date of Update " & UpdateDate & "</li>"
		
	Case "QCSpacer"
		Response.Write "<li>QC Spacer Quantity Updated:</li>"
		Response.Write "<li> Box #: " & Identifier & "</li>"		' Box Number for Spacer
		Response.Write "<li> Added: " & Quantity & "</li>"
		Response.Write "<li> New Quantity: " & rs.Fields("Quantity") & "</li>"
		Response.Write "<li> Date of Update " & UpdateDate & "</li>"
	
	Case "QCSealant"
		Response.Write "<li>QC Sealant Quantity Updated:</li>"
		Response.Write "<li> Lot #: " & Identifier & "</li>"		' Lot Number for Glass
		Response.Write "<li> Added: " & Quantity & "</li>"
		Response.Write "<li> New Quantity: " & rs.Fields("Quantity") & "</li>"
		Response.Write "<li> Date of Update " & UpdateDate & "</li>"
		
End Select	

else
Response.Write "<h2>" & InventoryType & " Item does not Exist: " & Identifier & " Please Retry</h2>"

end if
%>

        <BR>
       
         <a class="whiteButton" href="javascript:conf.submit()">Home</a>
            
            </form>

            
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

