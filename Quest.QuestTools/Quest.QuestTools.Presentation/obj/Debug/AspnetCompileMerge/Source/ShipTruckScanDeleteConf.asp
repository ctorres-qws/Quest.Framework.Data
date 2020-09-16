<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created July 2019 - by Michael Bernholtz --> 
<!--Delete Confirmation Page for Scan Items-->
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Delete Ship Scan</title>
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
DeleteBarcode = request.querystring("DeleteID")
%>
	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle">Delete Confirmation</h1>
        <a class="button leftButton" type="cancel" href="ShipHomeManager.HTML" target="_self">Scan</a>
     </div>

<%


	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "Select ID, BARCODE, DELETEDATE, DELETED from X_SHIP WHERE [DELETED] = FALSE AND Barcode = '" & DeleteBarcode & "' "
	rs.Cursortype = 2
	rs.Locktype = 3
	rs.Open strSQL, DBConnection
	if not RS.eof then
	'Set Glass Master Delete Statement
		DeleteID = rs("ID")
		RS("DeleteDate") = Now
		RS("Deleted") = True
		RS.UPDATE
		
		rs.close
		set rs =nothing
		
	Else
	DeleteID = "ID Not Found"
	End If
	DbCloseAll
	

%>

<form id="conf" title="Delete Stock" class="panel" name="conf" action="ShipTruckScanDelete.asp" method="GET" target="_self" selected="true" >

        <h2>Scan Deleted <%response.write DeleteID &" :**"& DeleteBarcode%></h2>
		<div class="row">

		</div>

        <BR>

         <a class="whiteButton" href="javascript:conf.submit()">Back to Ship Reports</a>
            
            </form>

</body>
</html>

<%
'DBConnection.close
'set DBConnection=nothing
%>

