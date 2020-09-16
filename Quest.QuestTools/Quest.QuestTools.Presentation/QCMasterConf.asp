<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 6th, by Michael Bernholtz - Confirmation  of input items to QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>QC Inventory</title>
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



ItemName = REQUEST.QueryString("ItemName")
Manufacturer = REQUEST.QueryString("Manufacturer")
EntryDate = Date()

InventoryType = REQUEST.QueryString("InventoryType")

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

'Entry will be 1 of 3 types split up by Case
Select Case InventoryType
	Case "QCGlass"
		Code = REQUEST.QueryString("Code")
		Pieces = Trim(Request.Querystring("Pieces"))
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

		Set RS = DBOpenRS(DBConnection, "SELECT * FROM QC_MASTER_GLASS WHERE ID=-1", GetDBCursorTypeInsert, GetDBLockTypeInsert)
		RS.AddNew
		RS.fields("ItemName") = ItemName
		RS.fields("Manufacturer") = Manufacturer
		RS.fields("Code") = Code
		RS.fields("EntryDate") = EntryDate
		RS.fields("Pieces") = Pieces
		RS.fields("Width") = Width
		RS.fields("Height") = Height
		RS.fields("Price") = Price
		RS.fields("Lites") = 0
		If GetID(isSQLServer,1) <> "" Then RS.Fields("ID") = GetID(isSQLServer,1)
		RS.Update
		Call StoreID1(isSQLServer, RS.Fields("ID"))

	Case "QCSpacer"

		Set RS = DBOpenRS(DBConnection, "SELECT * FROM QC_MASTER_SPACER WHERE ID=-1", GetDBCursorTypeInsert, GetDBLockTypeInsert)
		RS.AddNew
		RS.fields("ItemName") = ItemName
		RS.fields("Manufacturer") = Manufacturer
		RS.fields("EntryDate") = EntryDate
		If GetID(isSQLServer,1) <> "" Then RS.Fields("ID") = GetID(isSQLServer,1)
		RS.Update
		Call StoreID1(isSQLServer, RS.Fields("ID"))

	Case "QCSealant"

		Set RS = DBOpenRS(DBConnection, "SELECT * FROM QC_MASTER_SEALANT WHERE ID=-1", GetDBCursorTypeInsert, GetDBLockTypeInsert)
		RS.AddNew
		RS.fields("ItemName") = ItemName
		RS.fields("Manufacturer") = Manufacturer
		RS.fields("EntryDate") = EntryDate
		If GetID(isSQLServer,1) <> "" Then RS.Fields("ID") = GetID(isSQLServer,1)
		RS.Update
		Call StoreID1(isSQLServer, RS.Fields("ID"))

	Case "QCMisc"

		Set RS = DBOpenRS(DBConnection, "SELECT * FROM QC_MASTER_MISC WHERE ID=-1", GetDBCursorTypeInsert, GetDBLockTypeInsert)
		RS.AddNew
		RS.fields("ItemName") = ItemName
		RS.fields("Manufacturer") = Manufacturer
		RS.fields("EntryDate") = EntryDate
		If GetID(isSQLServer,1) <> "" Then RS.Fields("ID") = GetID(isSQLServer,1)
		RS.Update
		Call StoreID1(isSQLServer, RS.Fields("ID"))

	Case Else 
		Response.Write "<h2>Invalid Input</h2>"

End Select

DbCloseAll

End Function

%>
	</head>
<body onload="startTime()" >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="QCMasterAdd.asp" target="_self">Add QC Item</a>
        <a class="button" href="#searchForm" id="clock"></a>
    </div>


    
<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	
Select Case InventoryType
	Case "QCGlass"
		Response.Write "<li>QC MASTER GLASS Entered:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Code: " & Code & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		Response.Write "<li> Pack Count: " & Pieces & "</li>"
		Response.Write "<li> Dimensions: " & Width & " X " & Height & "</li>"
		Response.Write "<li> Square Foot " & Width*Height*Pieces & "</li>"
		Response.Write "<li> Price Per Square Foot: " & Price & "</li>"
		
		
	Case "QCSpacer"
		Response.Write "<li>QC MASTER SPACER Entered:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"	
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
	
	Case "QCSealant"
		Response.Write "<li>QC MASTER SEALANT Entered:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		
	Case "QCMisc"
		Response.Write "<li>QC MASTER MISCELLANEOUS Entered:</li>"
		Response.Write "<li> Item Name: " & ItemName & "</li>"
		Response.Write "<li> Manufacturer: " & Manufacturer & "</li>"
		Response.Write "<li> Date of Entry " & EntryDate & "</li>"
		
	Case Else
		Response.Write("Invalid Entry type")						' No Input for Invalid
End Select	
%>	

  

</ul>



</body>
</html>

<% 

RS.close
set RS=nothing

DBConnection.close
set DBConnection=nothing
%>

