<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit and Delete Form for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant, Miscellaneous go to QC_Misc-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
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

Dim InventoryType, ItemID
InventoryType = Request.Querystring("InventoryType")
QCID = Request.QueryString("QCid")

Dim Identifier, IdentifierID
'Identifier changes the entry between Serial Number for Glass, Box Number for Spacer, Lot Number for Sealant

	
		Select Case InventoryType
	Case "QCGlass"

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_MASTER_Glass"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	
		rs.filter = "ID = " & QCID
		
		
	Case "QCSpacer"
	
	
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_MASTER_Spacer"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "ID = " & QCID
		
		
	Case "QCSealant"
	
	
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_MASTER_Sealant"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "ID = " & QCID
		
	Case "QCMisc"
	
	
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM QC_MASTER_Misc"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection
		rs.filter = "ID = " & QCID	
	End Select


STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="QCMasterEdit.asp?InventoryType=<% response.write InventoryType %>" target="_self">Edit Master</a>

    </div>			
    
    
    <form id="QCedit" title="Edit Master" class="panel" action="QCMasterEditConf.asp" name="QCedit"  method="GET" target="_self" selected="true" > 
  
	<fieldset>
	<div class="row">
        <label>Item Name</label>
        <input type="text" name='ItemName' id='ItemName' value="<%response.write Trim(rs.fields("ItemName")) %>" >
    </div>

    <div class="row">
        <label>Manufacturer</label>
        <input type="text" name='MANUFACTURER' id='MANUFACTURER' value="<%response.write Trim(rs.fields("Manufacturer")) %>" >
    </div>
               
    <div class="row">
        <label>Entry Date</label>
        <input type="text" name='EntryDate' id='EntryDate' value="<%response.write Trim(rs.fields("EntryDate")) %>" >
    </div>        
	<%
	If InventoryType = "QCGlass" Then
	%>	
	
	<div class="row">
        <label>Code</label>
        <input type="text" name='CODE' id='CODE' value="<%response.write Trim(rs.fields("CODE")) %>" >
    </div>
	
	<div class="row">
        <label>Pack Count</label>
        <input type="Number" name='Pieces' id='Pieces' value="<% response.write Trim(rs.fields("Pieces")) %>" >
    </div>

    <div class="row">
        <label>Width</label>
        <input type="Number" name='Width' id='Width' value="<%response.write Trim(rs.fields("Width")) %>" >
    </div>
               
    <div class="row">
        <label>Height</label>
        <input type="Height" name='Height' id='Height' value="<%response.write Trim(rs.fields("Height")) %>" >
    </div> 
	               
    <div class="row">
        <label>$ per Sqft</label>
        <input type="Number" name='Price' id='Price' value="<%response.write Trim(rs.fields("Price")) %>" >
    </div> 	
		
	<div class="row">
        <label>Extra Lites</label>
        <input type="Number" name='Lites' id='Lites' value="<%response.write Trim(rs.fields("Lites")) %>" >
    </div> 	
	<%	
	End if
     %>       
                  
                        <input type="hidden" name='qcid' id='qcid' value="<%response.write Trim(rs.fields("id")) %>" />
						<input type="hidden" name='InventoryType' id='InventoryType' value="<%response.write InventoryType %>" />
            
</fieldset>


        <BR>
        
		
		<a class="whiteButton" onClick="QCedit.action='QCMasterEditConf.asp'; QCedit.submit()">Submit Changes</a><BR>
       <!-- <a class="redButton" onClick="QCedit.action='QCMasterDelConf.asp'; QCedit.submit()">Delete Master</a><BR> -->

            
            </form>
                        
    
</body>
</html>

<% 

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

