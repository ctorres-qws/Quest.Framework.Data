<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created February 7th, by Michael Bernholtz - Edit Confirmation for items in QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  
<!-- Glass go to QC_GLASS, Spacer go to QC_Spacer, Sealant go to QC_Sealant-->
<!-- Updated February 26th to include Consumed and the ability to clear consumption and reactivate -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Edit QC Inventory</title>
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
Passkey = "DLL"
Password = UCASE(TRIM(Request.Form("pwd")))


OPID = request.querystring("OPID")
JOB = REQUEST.QueryString("Job")
FLOOR = REQUEST.QueryString("Floor")
GLASS = REQUEST.QueryString("Glass")
LITES = REQUEST.QueryString("Lites")
if LITES = "" then 
	LITES = 0
end if
BACKORDER = REQUEST.QueryString("Backorder")
if BACKORDER = "" then 
	BACKORDER = 0
end if
BackorderText = REQUEST.QueryString("BackOrderText")
GType = REQUEST.QueryString("InventoryType")
Bendfile = REQUEST.QueryString("BendFile")
Opfile = REQUEST.QueryString("OpFile")
OpDate = REQUEST.QueryString("OpDate")

GlassCutDate = REQUEST.QueryString("GlassCutDate")
PackDate = REQUEST.QueryString("PackDate")
ShipDate = REQUEST.QueryString("ShipDate")
ReceivedDate = REQUEST.QueryString("ReceivedDate")
Shift = REQUEST.QueryString("Shift")
Employee = REQUEST.QueryString("Employee")
Skid = REQUEST.QueryString("Skid")
if Skid ="" then
Skid = 0
end if




BarcodeSuffix = ""
if instr(1,Glass, "TMP") > 0 then
BarcodeSuffix = "_TMP"
end if
if instr(1,Glass, "HS") > 0 then
BarcodeSuffix = "_HS"
end if

Barcode = Skid & "_QT_" & JOB & FLOOR & BarcodeSuffix


%>

	</head>
<body>

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>
                <a class="button leftButton" type="cancel" href="OptimizationLogEditForm.asp?OPID=<% response.write OPID %>" target="_self">Edit Opt Log</a>
    </div>
    
 <% 
if Password = Passkey then
%>          
    
<form id="conf" title="Production" class="panel" name="conf" action="index.html#_GlassP" method="GET" target="_self" selected="true" >              

  
   
        <h2>Stock Edited</h2>
  
<%       
           
	
		'Set Sealant Inventory Update Statement
			StrSQL = "UPDATE OptimizeLog SET Job='"& JOB & "', Floor='" & Floor & "', Glass='" & Glass & "', Lites='" & Lites & "', Type='" & GType & "', BendFile='" & Bendfile & "', OpFile='" & Opfile & "', Shift='" & Shift & "', Employee='" & EMPLOYEE & "', Skid='" & Skid & "', Backorder='" & BACKORDER & "', Backordertext='" & BACKORDERTEXT & "'  WHERE ID = " & OPID
		'Get a Record Set
		'		Set RS = DBConnection.Execute(strSQL)
				
		' -------------------------------------------------------OP Date ---------------------------------------		
		if isDate(OpDate) then		
			
			StrSQL2 = "UPDATE OptimizeLog SET OpDate='" & OpDate & "' WHERE ID = " & OPID
		'Get a Record Set
				Set RS2 = DBConnection.Execute(strSQL2)
		end if		
		
		if OpDate = "" then		
			StrSQL2 = "UPDATE OptimizeLog SET OpDate= NULL WHERE ID = " & OPID
		'Get a Record Set
		'		Set RS2 = DBConnection.Execute(strSQL2)
			
		end if
		
		' -------------------------------------------------------Glass Cut Date ---------------------------------------	
		if isDate(GlassCutDate) then		
			
			StrSQL3 = "UPDATE OptimizeLog SET GlassCutDate='" & GlassCutDate & "' WHERE ID = " & OPID
		'Get a Record Set
		'		Set RS3 = DBConnection.Execute(strSQL3)
		end if
		
		if GlassCutDate = "" then		
			StrSQL3 = "UPDATE OptimizeLog SET GlassCutDate= NULL WHERE ID = " & OPID
		'Get a Record Set
			'	Set RS3 = DBConnection.Execute(strSQL3)
			
		end if
		' -------------------------------------------------------Received Date ---------------------------------------	
		if isDate(ReceivedDate) then		
			
			StrSQL4 = "UPDATE OptimizeLog SET ReceivedDate='" & ReceivedDate & "' WHERE ID = " & OPID
		'Get a Record Set
		'		Set RS4 = DBConnection.Execute(strSQL4)
		end if
		
		if ReceivedDate = "" then		
			StrSQL4 = "UPDATE OptimizeLog SET ReceivedDate= NULL WHERE ID = " & OPID
		'Get a Record Set
			'	Set RS34 = DBConnection.Execute(strSQL4)
			
		end if
				' -------------------------------------------------------Pack Date ---------------------------------------	
		if isDate(PackDate) then		
			
			StrSQL4 = "UPDATE OptimizeLog SET PackDate='" & PackDate & "' WHERE ID = " & OPID
		'Get a Record Set
				'Set RS4 = DBConnection.Execute(strSQL4)
		end if
		
		if PackDate = "" then		
			StrSQL4 = "UPDATE OptimizeLog SET packDate= NULL WHERE ID = " & OPID
		'Get a Record Set
'				Set RS34 = DBConnection.Execute(strSQL4)
			
		end if
				' -------------------------------------------------------Ship Date ---------------------------------------	
		if isDate(ShipDate) then		
			
			StrSQL4 = "UPDATE OptimizeLog SET ShipDate='" & ShipDate & "' WHERE ID = " & OPID
		'Get a Record Set
			'	Set RS4 = DBConnection.Execute(strSQL4)
		end if
		
		if ShipDate = "" then		
			StrSQL4 = "UPDATE OptimizeLog SET ShipDate= NULL WHERE ID = " & OPID
		'Get a Record Set
		'	Set RS34 = DBConnection.Execute(strSQL4)
			
		end if
else
%>
<form id="adminpass" title="Administrative Tools" class="panel" name="enter" action="OptimizationLogEditConftest.asp?OPID=<%response.write OPID%>&JOB=<%response.write JOB%>&Floor=<%response.write Floor%>&Glass=<%response.write Glass%>&Lites=<%response.write Lites%>&InventoryType=<%response.write GType%>&Bendfile=<%response.write Bendfile%>&Opfile=<%response.write Opfile%>&Shift=<%response.write Shift%>&Employee=<%response.write Employee%>&Opdate=<%response.write Opdate%>&GlassCutDate=<%response.write GlassCutDate%>&BackOrder=<%response.write BackOrder%>&ReceivedDate=<%response.write ReceivedDate%>&ShipDate=<%response.write ShipDate%>&PackDate=<%response.write PackDate%>&Skid=<%response.write Skid%>&Backordertext=<%response.write BackOrderText%>" method="post" target="_self" selected="True">



<fieldset>
			<div class="row" >
				<label>Password:</label>
				<input type="password" name='pwd' id='pwd' ></input>
			</div>
			
</fieldset>




<a class="whiteButton" href="javascript:adminpass.submit()">Enter password</a>
	</form>
	
<%
end if
	
if Password = Passkey then
%>

<ul id="Report" title="Added" selected="true">
	
	
	

	
<%	

		Response.Write "<li>Optimization GLASS Edited:</li>"
		Response.Write "<li> Job: " & JOB & "</li>"
		Response.Write "<li> Floor: " & Floor & "</li>"
		Response.Write "<li> Glass" & Glass & "</li>"
		Response.Write "<li> Glass Type: " & Gtype & "</li>"
		Response.Write "<li> Optimization File: " & OpFile & "</li>"
		Response.Write "<li> Number of Lites: " & Lites & "</li>"
		Response.Write "<li> Bending File: " & Bendfile & "</li>"
		Response.Write "<li> Optimization Date: " & OpDate & "</li>"
		Response.Write "<li> Shift: " & Shift & "</li>"
		Response.Write "<li> Employee: " & Employee & "</li>"
		Response.Write "<li> Glass Cut Date: " & GlassCutDate & "</li>"
		Response.Write "<li> Pack Date: " & PackDate & "</li>"
		Response.Write "<li> Skid Number: " & Skid & "</li>"
		Response.Write "<li> Ship Date: " & ShipDate & "</li>"
		Response.Write "<li> Received Date: " & ReceivedDate & "</li>"
		Response.Write "<li> Number of Back Order Lites: " & BackOrder & "</li>"
		Response.Write "<li> Barcode: " & Barcode & "</li>"
		
		
		'Added for Jody and Ruslan 
if isDate(ReceivedDate) then
Response.write "</ul><h2>Below Glass Items with Matching MO Code are being marked Received:</h2><ul>"
Set rsG = Server.CreateObject("adodb.recordset")
strSQLG = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
rsG.Cursortype = 2
rsG.Locktype = 3
rsG.Open strSQLG, DBConnection

rsG.filter = " extordernum = '" & opfILE & "' OR INTordernum = '" & opfILE & "'"

 if rsG.eof then

 else

 do while not rsG.eof 
	
	
	if rsG("ExtOrderNum") = opFile then
		rsG("ExtReceived") = Date()
		BackTag = "Ext"
	end if
	if rsG("IntOrderNum") = opFile then
		rsG("IntReceived") = Date()
		BackTag = "Int"
	end if
	rsG.update
	
	Response.write "<li>ID: " & RSG("ID") & " Window: " & RSG("Job") & RSG("Floor") & "-"& RSG("Tag") 
	if isnull(RSG("backorderflag")) then
	Response.write " : Mark as <a href='OptimizeBackorder.asp?opid=" & rsG("ID") & "&BackTag=" & BackTag & "' target='_blank' >Backorder</a>" 
	else
	Response.write " : On BackOrder"
	end if
	response.write "</li>"
 
 rsG.movenext
 loop
 end if
 
 
end if
%>

        <BR>
       
         <a class="whiteButton" href="OptimizationLogManage.asp" target="_self"> Back</a>

            </form>

     <%
	
end if
%>           
    
</body>
</html>

<% 

DBConnection.close
set DBConnection=nothing
%>

