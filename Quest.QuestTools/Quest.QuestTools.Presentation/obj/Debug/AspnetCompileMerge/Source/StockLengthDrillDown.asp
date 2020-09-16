<!--#include file="dbpath.asp"-->
              <!-- Changed September 2014, to include Durapaint and Goreway as Stock and rest as Pending: (DuraPaint removed from Pending)-->
				<!-- Change requested by Shaun Levy, Approved by Jody Cash -->
				<!-- Updated  to Split up Texas from Canada October 2019 - Michael Bernholtz-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Stock Drill Down </title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />

  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  
  
  
  </script>

<!-- Added a script to include Sorttable.js to allow tables to be sorted on screen rather than by repeating SQL string  December 6th, Michael Bernholtz-->
 <script src="sorttable.js"></script>
	</head>
<body>

<%
ticket=Request.Querystring("ticket")
%>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
				
				
				
				<%	select case ticket
					case  "plastic" 
				%>
			   <a class="button leftButton" type="cancel" href="StocklevelsSummaryPL.asp" target="_self">Levels-PL</a>
				<% 
					case "sheet"
				%>
				<a class="button leftButton" type="cancel" href="StocklevelsSummarySH.asp" target="_self">Levels-SH</a>
				<% 
					case else
				%>
				<a class="button leftButton" type="cancel" href="StocklevelsSummaryEX.asp" target="_self">Levels-EX</a>
				<% 
				End Select
				
				%>
    </div>
    
      
  
<%
Part = Request.QueryString("Part")
SWidth = Request.QueryString("SWidth") + 0
SHeight = Request.QueryString("SHeight") + 0
Set rs = Server.CreateObject("adodb.recordset")


		if CountryLocation = "USA" then 
			if ticket = "sheet" then
				strSQL = "SELECT * FROM Y_INV WHERE (Warehouse = 'JUPITER') AND Part = '" & Part & "' AND Width = " & SWidth & " AND Height = " & SHeight & " order by Lft ASC, width ASC, Colour ASC"
			else
				strSQL = "SELECT * FROM Y_INV WHERE (Warehouse = 'JUPITER') AND Part = '" & Part & "' order by Lft ASC, width ASC, Colour ASC"
			end if

		else
			if ticket = "sheet" then
				strSQL = "SELECT * FROM Y_INV WHERE (Warehouse = 'GOREWAY' OR Warehouse = 'HORNER' OR Warehouse = 'NASHUA' OR Warehouse = 'DURAPAINT' OR Warehouse = 'DURAPAINT(WIP)' OR Warehouse = 'CAN-ART' OR Warehouse = 'MILVAN' OR Warehouse = 'EXTAL SEA' OR Warehouse = 'DEPENDABLE' OR Warehouse = 'SAPA' OR WAREHOUSE = 'HYDRO' OR WAREHOUSE = 'NPREP') AND Part = '" & Part & "' AND Width = " & SWidth & " AND Height = " & SHeight & " order by Lft ASC, width ASC, Colour ASC"
			else
				strSQL = "SELECT * FROM Y_INV WHERE (Warehouse = 'GOREWAY' OR Warehouse = 'HORNER' OR Warehouse = 'NASHUA' OR Warehouse = 'DURAPAINT' OR Warehouse = 'DURAPAINT(WIP)' OR Warehouse = 'CAN-ART' OR Warehouse = 'MILVAN' OR Warehouse = 'EXTAL SEA' OR Warehouse = 'DEPENDABLE' OR Warehouse = 'SAPA' OR WAREHOUSE = 'HYDRO' OR WAREHOUSE = 'NPREP') AND Part = '" & Part & "' order by Lft ASC, width ASC, Colour ASC"
			end if
		end if 




rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

Part = Request.QueryString("Part")
Set rs2 = Server.CreateObject("adodb.recordset")
	strSQL2 = "SELECT * FROM Y_Color Order by Project ASC"
rs2.Cursortype = 2
rs2.Locktype = 3
rs2.Open strSQL2, DBConnection	

Set rs3 = Server.CreateObject("adodb.recordset")
	strSQL3 = "SELECT * FROM Y_master"
rs3.Cursortype = 2
rs3.Locktype = 3
rs3.Open strSQL3, DBConnection	

%>  
<ul id="screen1" title="Stock Level <% response.write ": " & Part %>" selected="true">            

<%


response.write "<li class='group'>Length Per Die - " & Part & " </li>"

if rs2.eof then
	response.write "<li>No Colours in Colour Table - Something fishy is going on...</li>"
else
	do while not rs2.eof
		rs.filter ="Colour = '" & rs2("Project") & "'"

		
		if rs.eof then
		else
		response.write "<li class='group'>Length Per Colour - " & RS2("Project") & " </li>"

		select case ticket
			case "plastic"
		response.write "<li><table border='1' class='sortable' ><tr><th>Length</th><th>Count</th><th>Available</th><th>Min Level</th></tr>"
			case "sheet"
		response.write "<li><table border='1' class='sortable' ><tr><th>Size</th><th>Count</th><th>Available</th><th>Thickness</th></tr>"
			case else
		response.write "<li><table border='1' class='sortable' ><tr><th>Length</th><th>Count</th><th>Available</th></tr>"
		end select
			
			
			if ticket = "sheet" then
				Length1 = rs("Width") & " X " & rs("Height")
				Length2 = 0
			Length2 = 0
			
			
			else
			Length1 = CDBL(rs("lft"))
			Length2 = 0
			
			end if
			
			LengthCount = 0
			AvailableCount = 0

			Do while not rs.eof
			
			if ticket = "sheet" then
				Length2 = Length1
				Length1 = rs("Width") & " X " & rs("Height")
			
			
			else
				Length2 = Length1
				Length1 = CDBL(rs("Lft"))
			end if	
				
				
				if Length1 = Length2 then
				
				
					LengthCount = LengthCount + rs("qty")
					If RS("Warehouse") = "GOREWAY" or RS("Warehouse") = "DURAPAINT" or RS("Warehouse") = "HORNER" or RS("Warehouse") = "NASHUA" or RS("Warehouse") = "NPREP" or RS("Warehouse") = "MILVAN" then
					AvailableCount = AvailableCount + rs("qty")
					Thickness = rs("Thickness")
					end if
				else 
					response.write "<tr><td>" & Length2 & "</td><td> " & LengthCount & "</td><td> " & AvailableCount & "</td>"
					if ticket= "sheet" then
						response.write "<td> " & rs("Thickness") & "</td>"
					end if
	
	
					if ticket= "plastic" then
						rs3.filter ="Part = '" & Part & "'"
						if not rs3.eof then
							Select Case Length2
								Case 16
									response.write "<td> " & rs3("Min-16") & "</td>"
								Case 18
									response.write "<td> " & rs3("Min-18") & "</td>"
								Case 20
									response.write "<td> " & rs3("Min-20") & "</td>"
								Case 21
									response.write "<td> " & rs3("Min-21") & "</td>"
								Case 22
									response.write "<td> " & rs3("Min-22") & "</td>"
								Case Else
									response.write "<td>OTHER</td>"
							End Select
						Else
							response.write  "<td>Search Invalid</td>"
							response.write "</tr>"
						End If
					END IF
				LengthCount = rs("qty")
				AvailableCount = 0
				Thickness = rs("Thickness")
					If RS("Warehouse") = "GOREWAY" or RS("Warehouse") = "DURAPAINT" or RS("Warehouse") = "HORNER" or RS("Warehouse") = "NASHUA" or RS("Warehouse") = "NPREP" or RS("Warehouse") = "MILVAN" then
					AvailableCount = rs("qty")
					end if
				end if
				
			rs.movenext
			loop
			response.write "<tr><td>" & Length1 & "</td><td> " & LengthCount & "</td><td> " & AvailableCount & "</td>"
			if ticket= "sheet" then
				response.write "<td> " & Thickness & "</td>"
			end if
			if ticket= "plastic" then
				rs3.filter ="Part = '" & Part & "'"
				if not rs3.eof then
					Select Case Length2
						Case 16
							response.write "<td> " & rs3("Min-16") & "</td>"
						Case 18
							response.write "<td> " & rs3("Min-18") & "</td>"
						Case 20
							response.write "<td> " & rs3("Min-20") & "</td>"
						Case 21
							response.write "<td> " & rs3("Min-21") & "</td>"
						Case 22
							response.write "<td> " & rs3("Min-22") & "</td>"
						Case Else
							response.write "<td>OTHER</td>"
					End Select
				Else
					response.write  "<td>Search Invalid</td>"
					response.write "</tr>"
				End If
			END IF
			response.write"</tr>"




		end if
		
		response.write "</table></li>"

	rs2.movenext
	loop
end if


rs.close
set rs=nothing
rs2.close
set rs2=nothing
DBConnection.close
set DBConnection=nothing
%>

   
            
   </ul>
</body>
</html>



