<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- Service Glass Status Report for Mailer - GT Codes found in Forel and Willian Scans - With Completion of PO included -->
<!-- Union Z_GLASSDB and X_BARCODEGA to get the GT CODE and determine POs. Then count number of items in the PO -->
<!-- Shows PO percentage based on completed-->
<!-- Michael Bernholtz June 2018 -->


<!--

Union Z_GLASSDB and X_BARCODEGA on GT (BARCODE)
WHERE WEEKNUMBER AND YEAR ORDER by PO

PO / OLD PO Counter

Second Dataset of just Z_GLASSDB for PO totals


-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Service Glass</title>
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


</head>
<body>
<!--#include file="todayandyesterday.asp"-->

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Report" target="_self">Reports</a>
    </div>

<ul id="screen1" title="Service Glass Status" selected="true">

		<li class="group">Service Glass Number</li>
<%


Set rs = Server.CreateObject("adodb.recordset")

strSQL = "Select GA.BARCODE, GA.Year, GA.Month, GA.WEEK, GA.DAY, DB.Barcode, DB.PO, DB.Job, DB.FLoor, DB.Tag FROM X_BARCODEGA AS GA, Z_GLASSDB AS DB WHERE GA.Barcode = DB.Barcode AND ((GA.Month = " & cMonth & " or GA.WEEK = " & weekNumber & ") AND GA.YEAR = " & cYear & ") ORDER BY DB.PO ASC"
rs.Cursortype = GetDBCursorType
rs.Locktype = GetDBLockType
rs.Open strSQL, DBConnection

Set rs2 = Server.CreateObject("adodb.recordset")
strSQL2 = "Select [PO] FROM Z_GLASSDB ORDER BY PO ASC"
rs2.Cursortype = GetDBCursorType
rs2.Locktype = GetDBLockType
rs2.Open strSQL2, DBConnection
rs.filter = "DAY = " & cDay & " AND Month = " & cMonth & " AND YEAR = " & cYear		
		
Response.write "<li class='group'>Today</li>"	
if rs.eof then
	response.write "<li> There are currently no items of Activity Today, Please check later</li>"
else
	
	NewPO = "BLANK"
	OldPO = ""
	NewBarcode = ""
	OldBarcode = ""
	POCounter = 0
	TotalPOCounter = 0
	
	Do While not rs.eof
	OldPO = NewPO
	NewPO = rs("PO")
	OldBarcode = NewBarcode
	NewBarcode = rs("Barcode") & " - " & rs("Job") & rs("Floor") & "-" & rs("Tag")
	
	if OldPO = NewPO then
		POCounter = POCounter + 1
		Response.write "<li>" & OldBarcode & "</li>"
	else
		rs2.filter = " PO = '" & OldPO & "'"
		if not rs2.eof then
			TotalPOCounter = 0
			Do while not rs2.eof
				TotalPOCounter = TotalPOCounter + 1
				rs2.movenext
				loop
			else
				TotalPOCounter = 40000
		end if
		if OldPO = "BLANK" then
		else 
			Response.write "<li class = 'smallLi'>" & OldBarcode & "</li>"
			Response.write "<li>PO " & OldPO & " : " & POCounter & " / " &TotalPOCounter & "</li><hr>"
		end if	
		POCounter = 1
		rs2.filter = ""
	
	end if
	
	rs.movenext
	loop
	rs2.filter = " PO = '" & NewPO & "'"
		if not rs2.eof then
			TotalPOCounter = 0
			Do while not rs2.eof
				TotalPOCounter = TotalPOCounter + 1
				rs2.movenext
				loop
			else
				TotalPOCounter = 40000
		end if
		Response.write "<li class = 'smallLi'>" & NewBarcode & "</li>"
		Response.write "<li>PO " & NewPO & " : " & TotalPOCounter & "</li><hr>"
	
	

end if

 rs.filter = "Year = " & cYear & " AND Week = " & weekNumber 
	
 Response.write "<li class='group'>This Week</li>"	

if rs.eof then
	response.write "<li> There are currently no items of Activity This Week, Please check later</li>"
else
	
	NewPO = "BLANK"
	OldPO = ""
	NewBarcode = ""
	OldBarcode = ""
	POCounter = 0
	TotalPOCounter = 0
	
	Do While not rs.eof
	OldPO = NewPO
	NewPO = rs("PO")
	OldBarcode = NewBarcode
	NewBarcode = rs("Barcode") & " - " & rs("Job") & rs("Floor") & "-" & rs("Tag")
	
	if OldPO = NewPO then
		POCounter = POCounter + 1
		Response.write "<li>" & OldBarcode & "</li>"
	else
		rs2.filter = " PO = '" & OldPO & "'"
		if not rs2.eof then
			TotalPOCounter = 0
			Do while not rs2.eof
				TotalPOCounter = TotalPOCounter + 1
				rs2.movenext
				loop
			else
				TotalPOCounter = 40000
		end if
		if OldPO = "BLANK" then
		else 
			Response.write "<li class = 'smallLi'>" & OldBarcode & "</li>"
			Response.write "<li>PO " & OldPO & " : " & POCounter & " / " &TotalPOCounter & "</li><hr>"
		end if	
		POCounter = 1
		rs2.filter = ""
	
	end if
	
	rs.movenext
	loop
	rs2.filter = " PO = '" & NewPO & "'"
		if not rs2.eof then
			TotalPOCounter = 0
			Do while not rs2.eof
				TotalPOCounter = TotalPOCounter + 1
				rs2.movenext
				loop
			else
				TotalPOCounter = 40000
		end if
		Response.write "<li class = 'smallLi'>" & NewBarcode & "</li>"
		Response.write "<li>PO " & NewPO & " : " & TotalPOCounter & "</li><hr>"
	
end if


 rs.filter = "Year = " & cYear & " AND Month = " & cMonth
	
 Response.write "<li class='group'>This Month</li>"	

if rs.eof then
	response.write "<li> There are currently no items of Activity This Month, Please check later</li>"
else
	
	
	NewPO = "BLANK"
	OldPO = ""
	NewBarcode = ""
	OldBarcode = ""
	POCounter = 0
	TotalPOCounter = 0
	
	Do While not rs.eof
	OldPO = NewPO
	NewPO = rs("PO")
	OldBarcode = NewBarcode
	NewBarcode = rs("Barcode") & " - " & rs("Job") & rs("Floor") & "-" & rs("Tag")
	
	if OldPO = NewPO then
		POCounter = POCounter + 1
		Response.write "<li>" & OldBarcode & "</li>"
	else
		rs2.filter = " PO = '" & OldPO & "'"
		if not rs2.eof then
			TotalPOCounter = 0
			Do while not rs2.eof
				TotalPOCounter = TotalPOCounter + 1
				rs2.movenext
				loop
			else
				TotalPOCounter = 40000
		end if
		if OldPO = "BLANK" then
		else 
			Response.write "<li class = 'smallLi'>" & OldBarcode & "</li>"
			Response.write "<li>PO " & OldPO & " : " & POCounter & " / " &TotalPOCounter & "</li><hr>"
		end if	
		POCounter = 1
		rs2.filter = ""
	
	end if
	
	rs.movenext
	loop
	rs2.filter = " PO = '" & NewPO & "'"
		if not rs2.eof then
			TotalPOCounter = 0
			Do while not rs2.eof
				TotalPOCounter = TotalPOCounter + 1
				rs2.movenext
				loop
			else
				TotalPOCounter = 40000
		end if
		Response.write "<li class = 'smallLi'>" & NewBarcode & "</li>"
		Response.write "<li>PO " & NewPO & " : " & TotalPOCounter & "</li><hr>"
	
end if

%>

	</ul>
        
  
<% 
On Error Resume Next

rs.close
set rs=nothing
rs2.CLOSE
Set rs2= nothing
DBConnection.close
set DBConnection=nothing

%>


</body>
</html>
