<!--#include file="connect_barcodeqc.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Quest Dashboard</title>
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
'Create a Query
    SQL = "Select * FROM X_BARCODE ORDER BY DATETIME DESC"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
	
bcid = request.querystring("bcid")
If bcid = "" Then bcid=330525
rs.filter = "ID = " & bcid
if not rs.eof then

bc = rs("barcode")
bctarget = "anything"
%> <!--#include file="bcgenerate.asp"--> <%
 'rs.movenext
 
 bc = rs("barcode")
bctarget = "barcodetarget1"
%> <!--#include file="bcgenerate.asp"--> <%


end if

rs.close
 set rs=nothing
 rs2.close
 set rs2=nothing
' rs3.close
' set rs3=nothing
DBConnection.close
set DBConnection=nothing
%>
</head>

<body>
<p>Test</P>
</body>
</html>



