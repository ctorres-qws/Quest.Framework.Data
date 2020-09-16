<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="QCdbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Delete Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant, Miscellaneous go to QC_Misc-->
<!-- February 2019 - USA Tables added -->

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
Passkey = "DLL"
Passkey2 = "JODY"
Password = UCASE(TRIM(Request.Form("pwd")))
back = request.querystring("back")

InventoryType = REQUEST.QueryString("InventoryType")
QCID = request.querystring("QCID")

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
                <a <a class="button leftButton" type="cancel" href="QCInventoryEditForm.asp?InventoryType=<% response.write InventoryType %>&QCID=<% response.write qcid %>" target="_self">Edit Stock</a>

    </div>
<%
If (Password = Passkey) or (Password = Passkey2) then

	Dim IdentifierTitle, IdentifierID
'Identifier changes the entry between Serial Number for Glass, Box Number for Spacer, Lot Number for Sealant

	Select Case(gi_Mode)
		Case c_MODE_ACCESS
			Process(false)
		Case c_MODE_HYBRID
			Process(false)
			Process(true)
		Case c_MODE_SQL_SERVER
			Process(true)
	End Select

Else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="QCInventoryDelConf.asp?InventoryType=<%response.write InventoryType%>&QCID=<%response.write QCID%>" method="post" target="_self" selected="True">
<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>

<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>
	
<%
End If
Function Process(isSQLServer)

	DbOpenQC DBConnection, isSQLServer

	Select Case InventoryType
		Case "QCGlass"
			'Set Glass Master Delete Statement
			if CountryLocation = "USA" then
				StrSQL = "DELETE FROM QC_GLASS_USA WHERE ID = " & QCID
			else
				StrSQL = "DELETE FROM QC_GLASS WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		Case "QCSpacer"
			'Set Spacer Master Delete Statement
			if CountryLocation = "USA" then
				StrSQL = "DELETE FROM QC_SPACER_USA WHERE ID = " & QCID
			else
				StrSQL = "DELETE FROM QC_SPACER WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		Case "QCSealant"
			'Set Sealant Master Delete Statement
			if CountryLocation = "USA" then
				StrSQL = "DELETE FROM QC_SEALANT_USA WHERE ID = " & QCID
			else
				StrSQL = "DELETE FROM QC_SEALANT WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
		Case "QCMisc"
			'Set Sealant Master Delete Statement
			if CountryLocation = "USA" then
				StrSQL = "DELETE FROM QC_MISC_USA WHERE ID = " & QCID
			else
				StrSQL = "DELETE FROM QC_MISC WHERE ID = " & QCID
			end if
			'Get a Record Set
				Set RS = DBConnection.Execute(strSQL)
	End Select

	DbCloseAll

End Function

If (Password = Passkey) or (Password = Passkey2) Then

		if CountryLocation = "USA" then 
			HomeSite = "indexTexas.html"
			HomeSiteSuffix = "-USA"
		else
			HomeSite = "index.html"
			HomeSiteSuffix = ""
		end if 
		%>

<form id="conf" title="Delete Stock" class="panel" name="conf" action="<%response.write Homesite%>#_QC" method="GET" target="_self" selected="true" >                

        <h2>Stock Deleted</h2>
		<div class="row">

		</div>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Back to Stock</a>

            </form>
<%
End If
%>
<%
On Error Resume Next
DBConnection.close
set DBConnection=nothing
%>
</body>
</html>



