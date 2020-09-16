<!--#include file="dbpath.asp"-->
    <!-- Updated May 9th to include length in feet, Michael Bernholtz -->
	<!-- Special Code to Overwrite Metra Door Material which ALWAYS comes in at 21.33 feet May 2018-->
	<!-- USA Seperation added February 2019-->
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
Server.ScriptTimeout=400
fail = "0"
counter = 0
'gi_mode= c_MODE_SQL_SERVER
gi_mode =c_MODE_SQL_SERVER
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

DBOpen DBConnection, isSQLServer

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Warehouse,Bundle,Supplier,[Note 9],[Note 8] FROM Y_INV where warehouse IN ('METRA','SAPA','HYDRO')"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

rs.movefirst
if rs.eof then
	fail = "1"	
else
	Do While Not rs.eof
		Response.write rs.Fields("Warehouse")
		Warehouse = rs.Fields("Warehouse")
		rs.Fields("Supplier") = Warehouse
		rs.Fields("Note 8") = "Changed from " & Warehouse
		rs.Fields("Note 9") = "Supplier Updated April 22,2019"
		rs.update
		' bundle = rs("Bundle")	
		' oldSupplier = rs("Supplier")
		
		' if oldSupplier = "KEYMARK" then
		'	do nothing
		' else
			' rs.Fields("Supplier") = GetSupplier(bundle)
		' end if
		' rs.Fields("Note 8") = "Changed from " & oldSupplier
		' rs.Fields("Note 9") = "Supplier Updated April 22,2019"
		' rs.update
		' counter = counter + 1
		rs.movenext
	Loop	
	
end if

DbCloseAll

End Function

%>
	</head>
<body>
    
<ul id="Report" title="Stock Entered" selected="true">
<% if fail = "1" then
response.write "<li>Supplier for Inventory Item Not Updated</li>"
else
%>
	<li><% response.write "Items Updated " & counter %></li>
<%
end if
%>
</ul>



</body>
</html>

<%

Function GetSupplier(bundle)

		Dim supplier

		'extract first bundle number found
		Dim a_Bundles:
		a_Bundles = Replace(bundle & "",", ",",") ' replace ', ' with ','		
		a_Bundles = Replace(a_Bundles & "",",","/") ' replace ',' with '/'		
		a_Bundles=Replace(a_Bundles & ""," / ","/") ' replace ' / ' with '/'		
		a_Bundles = Split(a_Bundles, "/")
		Dim str_Bundle: str_bundle = bundle
		Dim origBundle : origBundle = bundle

		If UBound(a_Bundles) >= 0 Then
			str_Bundle = Trim(a_Bundles(0)&"")
		End If

		If str_Bundle <> "" Then
			'get Supplier from Warehouse field then check bundle numbers to confirm
			If IsNumeric(trim(str_Bundle)) Then
				If Left(str_Bundle,3) = "113" Then
					supplier = "HYDRO"
				ElseIf CDbl(str_Bundle) > (958153 - 300000) AND CDbl(str_Bundle) < (958153 + 100000) Then
					supplier = "HYDRO"			
				ElseIf Left(str_Bundle,3) = "109" AND Len(str_Bundle) = 7 Then
					supplier = "HYDRO"
				ElseIf Len(trim(str_Bundle)) = 7 Then
						supplier = "CANART"
				ElseIf CDbl(str_Bundle) > (560442 - 100000) AND CDbl(str_Bundle) < (560442 + 100000) Then
						supplier = "APEL"
				ElseIf InStr(UCase(origBundle),"MCA") > 0 Then
					supplier = "METRA"											
				Else
					supplier = ""
				End If
			Else
				If UCase(Left(str_Bundle,1)) = "A" Then
					supplier = "EXTAL"
				ElseIf UCase(Left(Trim(str_Bundle),5)) = "METRA" Then
					supplier = "METRA"
				ElseIf UCase(Left(Trim(str_Bundle),3)) = "MCA" Then
					supplier = "METRA"
				ElseIf InStr(UCase(origBundle),"MCA") > 0 Then
					supplier = "METRA"					
				Else
					supplier = ""			
				End If
			End If
		End If

		
		GetSupplier = supplier
End Function

%>
