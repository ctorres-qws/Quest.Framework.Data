<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Delete Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Delete QC Inventory</title>
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
                <a <a class="button leftButton" type="cancel" href="QCMasterEditForm.asp?InventoryType=<% response.write InventoryType %>&QCID=<% response.write qcid %>" target="_self">Edit Stock</a>
    
    </div>
    
      <%       
	
		Select Case InventoryType
	Case "QCGlass"

		'Set Glass Master Delete Statement
			StrSQL = "DELETE FROM QC_MASTER_GLASS WHERE ID = " & QCID
		'Get a Record Set
			Set RS = DBConnection.Execute(strSQL)
		
	Case "QCSpacer"
	
		'Set Spacer Master Delete Statement
			StrSQL = "DELETE FROM QC_MASTER_SPACER WHERE ID = " & QCID
		'Get a Record Set
			Set RS = DBConnection.Execute(strSQL)

	Case "QCSealant"
		
		'Set Sealant Master Delete Statement
			StrSQL = "DELETE FROM QC_MASTER_SEALANT WHERE ID = " & QCID
		'Get a Record Set
			Set RS = DBConnection.Execute(strSQL)
	
	Case "QCMisc"
		
		'Set Miscellaneous Master Delete Statement
			StrSQL = "DELETE FROM QC_MASTER_MISC WHERE ID = " & QCID
		'Get a Record Set
			Set RS = DBConnection.Execute(strSQL)
		
	End Select



%>
    
<form id="conf" title="Delete Stock" class="panel" name="conf" action="index.html#_QC" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Deleted</h2>
		<div class="row">

		</div>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>
            
            </form>

            
    
</body>
</html>

<% 


DBConnection.close
set DBConnection=nothing
%>

