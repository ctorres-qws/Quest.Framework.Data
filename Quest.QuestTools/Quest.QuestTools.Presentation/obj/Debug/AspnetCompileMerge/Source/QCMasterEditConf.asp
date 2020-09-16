<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant, Miscellaneous go to QC_Misc-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Edit QC Inventory</title>
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
qcid = request.querystring("qcid")
%>

	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="QCMasterEditForm.asp?InventoryType=<% response.write InventoryType %>&QCID=<% response.write qcid %>" target="_self">Edit Stock</a>
    </div>
    
      
    
<form id="conf" title="Edit Stock" class="panel" name="conf" action="index.html#_QC" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>
  
<%                  

ItemName = Trim(Request.Querystring("ItemName"))
if ItemName <> "" then
	ItemName = replace (ItemName,"'","")
end if

Manufacturer = Trim(Request.Querystring("Manufacturer"))
EntryDate = Trim(Request.Querystring("EntryDate"))
Pieces = Trim(Request.Querystring("Pieces"))
Code = Trim(Request.Querystring("Code"))
if Pieces = "" then
Pieces = 0
End If
Width = Trim(Request.Querystring("Width"))
if Width = "" then
Width = 1
End If
Height = Trim(Request.Querystring("Height"))
if Height = "" then
Height = 1
End If
Price = Trim(Request.Querystring("Price"))
if Price = "" then
Price = 0
End If	
Lites = Trim(Request.Querystring("Lites"))
if Lites = "" then
Lites = 0
End If

Select Case(gi_Mode)
	Case c_MODE_ACCESS
		Process(false)
	Case c_MODE_HYBRID
		Process(false)
		Process(true)
	Case c_MODE_SQL_SERVER
		Process(true)
End Select

Function Process(isSQLServer)

DBOpenQC DBConnection, isSQLServer

		Select Case InventoryType
	Case "QCGlass"
	
			'Set Glass Master Update Statement
				StrSQL = "UPDATE QC_MASTER_GLASS  SET ItemName= '" & ItemName & "', Manufacturer='"& Manufacturer & "', Code='"& Code & "', EntryDate='" & EntryDate & "', Pieces='"& Pieces & "', Width='"& Width & "', Height='"& Height & "', Price='"& Price & "', Lites='"& Lites & "'  WHERE ID = " & QCID
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		
	Case "QCSpacer"
	
			'Set Spacer Master Update Statement
				StrSQL = "UPDATE QC_MASTER_SPACER  SET ItemName= '" & ItemName & "', Manufacturer='"& Manufacturer & "', EntryDate='" & EntryDate & "'  WHERE ID = " & QCID
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)

	Case "QCSealant"
	
			'Set Sealant Master Update Statement
				StrSQL = "UPDATE QC_MASTER_SEALANT  SET ItemName= '" & ItemName & "', Manufacturer='"& Manufacturer & "', EntryDate='" & EntryDate & "'  WHERE ID = " & QCID
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	
	Case "QCMisc"
	
			'Set Sealant Master Update Statement
				StrSQL = "UPDATE QC_MASTER_MISC  SET ItemName= '" & ItemName & "', Manufacturer='"& Manufacturer & "', EntryDate='" & EntryDate & "'  WHERE ID = " & QCID
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
				
	End Select

DbCloseAll

End Function

STAMPVAR = month(now) & "/" & day(now) & "/" & year(now)
ccTime = hour(now) & ":" & minute(now)
cDay = day(now)
cMonth = month(now)
cYear = year(now)
currentDate = Date
weekNumber = DatePart("ww", currentDate)


%>

    
<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	
Select Case InventoryType
	Case "QCGlass"
		Response.Write "<li>QC MASTER GLASS Edited:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Code: " & Code & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		Response.Write "<li> Pack Count: " & Pieces & "</li>"
		Response.Write "<li> Dimensions: " & Width & " X " & Height & "</li>"
		Response.Write "<li> Square Foot " & Width*Height*Pieces & "</li>"
		Response.Write "<li> Price Per Square Foot: " & Price & "</li>"
		
		
	Case "QCSpacer"
		Response.Write "<li>QC MASTER SPACER Edited:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"	
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
	
	Case "QCSealant"
		Response.Write "<li>QC MASTER SEALANT Edited:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		
	Case "QCMisc"
		Response.Write "<li>QC MASTER MISCELLANEOUS Edited:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		
End Select	
%>

        <BR>
       
         <a class="whiteButton" href="javascript:conf.submit()">Home</a>
            
            </form>

            
    
</body>
</html>

<% 

'DBConnection.close
'set DBConnection=nothing
%>

