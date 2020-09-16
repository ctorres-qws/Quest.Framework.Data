<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
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
	

	
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection



'afilter = request.QueryString("aisle")


%>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>
   
   
         
       
        <ul id="Profiles" title=" Glass Report - All Active" selected="true">
        
        
<% 

response.write "<li class='group'>ALL ACTIVE GLASS</li>"
' Sets an inital background colour, this will be changed back and forth in the loop to alternate glass items at two rows each
COLOUR = "#BBBBBB"

do while not rs.eof
	if not isdate(RS("COMPLETEDDATE")) then
	if isdate(RS("COMPLETEDDATE")) then
		response.write "<li style='background-color: "& COLOUR & "; border: 0px'>" & rs("ID") & " - " & rs("JOB") & " " & RS("FLOOR") & " " & RS("TAG") 
		response.write " - " & RS("DIM X") & "'' x " & RS("DIM Y") & "'' " 
		response.write " - " & RS("DEPARTMENT")
		response.write " -  " & " " & RS("1 MAT") & " " & RS("1 SPAC") & " " & RS("2 MAT") &  "</li>"
		response.write "<li style='background-color: "& COLOUR & "; border: 0px'>IN: " & RS("INPUTDATE") & " OPT: " & RS("OPTIMADATE") & " REQ: " & RS("REQUIREDDATE") & " OUT: " & RS("COMPLETEDDATE") & "</li>"
		' If statement to determine current background color of the glass and switch it back between BBBBBB (grey) and FFFFFF (white)
		if COLOUR = "#BBBBBB" then
			COLOUR = "#FFFFFF"
		else
			COLOUR = "#BBBBBB"
		end if
	end if
rs.movenext
loop




rs.close
Set rs = nothing
DBConnection.close
Set DBConnection = nothing

%>
      </ul>     
</body>
</HTML>	  
            
            
            
       
            
              
               
                
             
               
</body>
</html>
