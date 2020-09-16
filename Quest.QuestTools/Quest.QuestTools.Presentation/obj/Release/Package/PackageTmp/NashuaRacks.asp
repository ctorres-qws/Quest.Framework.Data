<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Nashua Empty Racks - Report for Shaun and Lev,  April 2017-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Nashua Inventory</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
<% Server.ScriptTimeout = 500 %> 
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT Distinct Aisle, Rack, Shelf FROM Y_INV WHERE WAREHOUSE = 'NASHUA' AND Len(Aisle) > 0 AND Len(Rack) > 0 AND Len(Shelf) > 0  ORDER BY AISLE ASC, RACK ASC, SHELF ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Dim str_Used: str_Used = ""
Do While Not rs.EOF
	str_Used = str_Used & "[" & rs("Aisle") & "|" & rs("Rack") & "|" & rs("Shelf") & "] <br/>" & vbCrLf
	rs.MoveNext
Loop

%>

</head>
<body>
<%

%>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Inv" target="_self">Stock</a>
        </div>

        <ul id="Profiles" title="Profiles" selected="true">
<%

SpaceCounter = 0
Dim aisle(19)
aisle(1) = "Aa"
aisle(2) = "Bb"
aisle(3) = "Cc"
aisle(4) = "Dd"
aisle(5) = "Ee"
aisle(6) = "Ff"
aisle(7) = "Gg"
aisle(8) = "Hh"
aisle(9) = "Ii"
aisle(10) = "Jj"
aisle(11) = "Kk"
aisle(12) = "Ll"
aisle(13) = "Mm"
aisle(14) = "Nn"
aisle(15) = "Oo"
aisle(16) = "Pp"
aisle(17) = "Qq"
aisle(18) = "Rr"

shelf= "0"
rack ="0"

CurrentAisle = 1
Do Until CurrentAisle > 18
	Rack = 1
	Shelf = 1
	ShelfFound = "0" 'False
	Do Until Rack >8

		

		Do Until Shelf > 8
			ShelfFound = "0"
			If (IsUsed(Aisle(CurrentAisle), Rack, Shelf)) Then
				ShelfFound = "1"
			End If

			If ShelfFound = "0" Then
				Response.Write "<li> Aisle: " & Aisle(CurrentAisle) & " Rack: " & rack & " Shelf: " & shelf & "</li>"
				SpaceCounter = SpaceCounter + 1
			End If

			Shelf = Shelf + 1
		Loop

		Shelf = 1
		Rack = Rack +1
	Loop 

	CurrentAisle = CurrentAisle +1
Loop

response.write "<li>Spaces Available: " & SpaceCounter & "</li>"


rs.close
set Rs = nothing

DBConnection.close
Set DBConnection = nothing


%>
      </ul>
</body>
</html>
<%
	Function IsUsed(str_Aisle, str_Rack, str_Shelf)
		Dim b_Ret: b_Ret = False
		Dim str_Token, str_TokenA, str_TokenB

		'Possible Rack Locations [Aa, 1, 1], [Aa, 1, 1a], [Aa, 1, 1c] etc
		str_Token = "[" & str_Aisle & "|" & str_Rack & "|" & str_Shelf & ""				'Don't check last character
		'str_TokenA = "[" & str_Aisle & "|" & str_Rack & "|" & str_Shelf & "A]"
		'str_TokenB = "[" & str_Aisle & "|" & str_Rack & "|" & str_Shelf & "B]"

		'If Instr(1, str_Used, str_Token, 1) > 0 Or Instr(1, UCase(str_Used), UCase(str_TokenA), 1) > 0 Or Instr(1, UCase(str_Used), UCase(str_TokenB), 1) > 0 Then
		If Instr(1, str_Used, str_Token, 1) > 0 Then
			b_Ret = True
		End If

		IsUsed = b_Ret
	End Function

%>