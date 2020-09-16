<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--Page Created March 5th, 2014 - by Michael Bernholtz --> 
<!--Edit Form Page for Glass Items-->
<!-- Submits to page GlassManageCONF.asp -->
		 <!--#include file="dbpath.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Manage Glass</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
<meta http-equiv="refresh" content="1120" >
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
  </script>

<%
GID = Request.QueryString("Gid")
ticket = Request.QueryString("ticket")

select case ticket
	case "active"
		returnEmail = "GlassToolViewActive.asp"
	case "optima"
		returnEmail = "GlassToolViewOptima.asp"
	case "waiting"
		returnEmail = "GlassToolViewWait.asp"
	case "waitingcom"
		returnEmail = "GlassToolViewWaitCommercial.asp"
	case "waitingser"
		returnEmail = "GlassToolViewService.asp"
	case "received"
		returnEmail = "GlassToolViewReceived.asp"
	case "completed"
		returnEmail = "GlassToolViewCompleted.asp"
	case "shipped"
		returnEmail = "GlassToolViewShipped.asp"
	case Else
		returnEmail = "GlassManage.asp"
	End select

		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "SELECT * FROM Z_GLASSDB"
		rs.Cursortype = 2
		rs.Locktype = 3
		rs.Open strSQL, DBConnection	
		rs.filter = "ID = " & GID

%>
	</head>
<body >

    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a id="backButton" class="button" href="#"></a>

                <a class="button leftButton" type="cancel" href="<%response.write returnEmail%>" target="_self">Select Glass</a>

    </div>

    <form id="GlassEdit" title="Edit Glass" class="panel" action="GlassDelConf.asp" name="GlassEdit"  method="GET" target="_self" selected="true" > 

	<fieldset>

        <div class="row">
            <label>Job </label>
            <input type="text" name='PROJECT' id='PROJECT' value ="<% response.write Trim(rs.fields("JOB")) %>" >
		</div>

		<div class="row">
            <label>Floor</label>
            <input type="text" name='FLOOR' id='FLOOR' value ="<% response.write Trim(rs.fields("FLOOR")) %>" >
        </div>

        <div class="row">
                <label>Tag</label>
                <input type="text" name='TAG' id='TAG' value ="<% response.write Trim(rs.fields("TAG")) %>" >
        </div>
		<div class="row">
            <label>Department</label>
            <select name= 'DEPARTMENT' id = 'DEPARTMENT'>
				<option selected='selected' value="<% response.write Trim(rs.fields("DEPARTMENT")) %>"> <% response.write Trim(rs.fields("DEPARTMENT")) %> </option>
				<option value="Production">Production</option>
				<option value="Service">Service</option>
				<option value="Commercial">Commercial</option>
				<option value="Recut">Recut</option>
			</select>
        </div>

		<div class="row">
                <label>Notes</label>
                <input type="text" name='NOTES' id='NOTES' value="<% response.write Trim(rs.fields("NOTES")) %>">
        </div>

        <div class="row">
		   <label>Ext. Method</label>
			<select name ='EXTMethod'>
				<option value = '<%response.write TRIM(rs.fields("EXTMethod")) %>' selected ><%response.write TRIM(rs.fields("EXTMethod"))%></option>
				<option value = 'CUT' >CUT</option>
				<option value = 'ORDER'>ORDER</option>
				<option value = 'ALREADY-HAVE'>ALREADY-HAVE</option>
			</select>
		</div>
            <div class="row">
                <label>Ext.Glass</label>
                <select name="ONEMAT">
					<option selected='selected' option value = '<% response.write Trim(rs.fields("1 MAT")) %>'><% response.write Trim(rs.fields("1 MAT")) %></option>
<% mat = mat1 %>
<!--#include file="QSU.inc"-->
</select>
                 </div>
      
            
              <div class="row">
                <label>Spacer.</label>
                <select name="ONESPAC">
				<option selected='selected' option value = '<% response.write Trim(rs.fields("1 SPAC")) %>'> <% response.write Trim(rs.fields("1 SPAC")) %></option>
<% mat = spac1 %>
<!--#include file="QSU2.inc"-->
</select>
            </div>
        <div class="row">
		 <label>Int. Method</label>
			<select name ='INTMethod'>
				<option value = '<%response.write TRIM(rs.fields("INTMethod")) %>' selected ><%response.write TRIM(rs.fields("INTMethod"))%></option>
				<option value = 'CUT' >CUT</option>
				<option value = 'ORDER'>ORDER</option>
				<option value = 'ALREADY-HAVE'>ALREADY-HAVE</option>
            </select>
		</div>
            <div class="row">
                <label>Int.Glass</label>
                <select name="TWOMAT">
				<option selected='selected' option value = '<% response.write Trim(rs.fields("2 MAT")) %>'> <% response.write Trim(rs.fields("2 MAT")) %></option>
<% mat = mat1 %>
<!--#include file="QSU.inc"-->
</select>
                 </div>
      
        
             <div class="row">
                <label>AIR / ARGON.</label>
			<select name ='AIR'>
					<option selected='selected' value = '<% response.write Trim(rs.fields("AIR")) %>'> <% response.write Trim(rs.fields("AIR")) %></option>
					<option value = 'Argon' >Argon</option>
					<option value = 'Air'>Air</option>
					<option value = 'N/A'>N/A</option>
                </select>
            </div>
 
            
              <div class="row">
                <label>Width.</label>
                <input type="text" name='WIDTH' id='WIDTH' size='8' value = '<% response.write Trim(rs.fields("Dim X")) %>'>
                </div>
      
            
              <div class="row">
                <label>Height.</label>
                <input type="text" name='HEIGHT' id='HEIGHT' size='8' value = '<% response.write Trim(rs.fields("Dim Y")) %>'>
            </div>
			
			<div class="row">
                <label>Order By</label>
				<select name ='orderBy'>
					<option selected='selected' value = '<% response.write Trim(rs.fields("ORDERBY")) %>'> <% response.write Trim(rs.fields("ORDERBY")) %></option>
			
				<option value = 'Yegor'>Yegor</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'Joe'>Joe</option>
				<option value = 'Ariel'>Ariel</option>
				<option value = 'Tomas'>Tomas</option>
				<option value = 'Michael'>Michael</option>
				<option value = 'WIS'>WIS</option>
                </select>
				
             </div>
			 
		<div class="row">
                <label>Order For</label>
				<select name ='orderFor'>
					<option selected='selected' value = '<% response.write Trim(rs.fields("ORDERFor")) %>'> <% response.write Trim(rs.fields("ORDERfor")) %></option>
				<option value = 'Arten'>Artem</option>
				<option value = 'Daniel'>Daniel</option>
				<option value = 'Ellerton'>Ellerton</option>
				<option value = 'Eric'>Eric</option>
				<option value = 'George'>George</option>
				<option value = 'Gurveen'>Gurveen</option>
				<option value = 'Hamlet'>Hamlet</option>
				<option value = 'Ivan'>Ivan</option>
				<option value = 'John'>John</option>
				<option value = 'Kenny'>Kenny</option>
				<option value = 'Rob'>Rob</option>
				<option value = 'Roman'>Roman</option>
				<option value = 'Vince'>Vince</option>
				<option value = 'Yegor'>Yegor</option>
				<option value = 'WIS'>WIS</option>
			</select>

		</div>

			 <div class="row">
                <label>PO</label>
                <input type="text" name='PoNum' id='PoNum' value = '<% response.write Trim(rs.fields("PO")) %>' >
             </div>

			 <div class="row">
                <label>Ext Work #</label>
                <input type="text" name='ExtorderNum' id='ExtorderNum' value = '<% response.write Trim(rs.fields("ExtorderNum")) %>' >
             </div>

			 <div class="row">
                <label>Ext From</label>
                <select name ='ExtFrom'>
				<option selected='selected' value = '<% response.write Trim(rs.fields("ExtFrom")) %>'> <% response.write Trim(rs.fields("ExtFrom")) %></option>
				<option value = 'Quest'>Quest</option>
				<option value = 'Cardinal'>Cardinal</option>
				<option value = 'TruLite'>TruLite</option>
				<option value = 'Woodbridge'>Woodbridge</option>
				<option value = 'Other'>Other</option>
                </select>
             </div>

			 <div class="row">
                <label>Int Work #</label>
                <input type="text" name='IntorderNum' id='IntorderNum' value = '<% response.write Trim(rs.fields("IntorderNum")) %>' >
             </div>

			 <div class="row">
                <label>Int From</label>
				<select name ='IntFrom'>
				<option selected='selected' value = '<% response.write Trim(rs.fields("IntFrom")) %>'> <% response.write Trim(rs.fields("IntFrom")) %></option>
				<option value = 'Quest' >Quest</option>
				<option value = 'Cardinal'>Cardinal</option>
				<option value = 'TruLite'>TruLite</option>
				<option value = 'Woodbridge'>Woodbridge</option>
				<option value = 'Other'>Other</option>
                </select>
             </div>

			<div class="row">
                <label>QT File</label>
                <input type="text" name='QTFile' id='QTFile' value = '<% response.write Trim(rs.fields("QTFile")) %>' >
             </div>

						<input type="hidden" name='GID' id='GID' value="<%response.write GID %>" />
						<input type="hidden" name='ticket' id='ticket' value="<%response.write ticket %>" />
</fieldset>

        <BR>

		<a class="whiteButton" onClick="GlassEdit.action='GlassManageConf.asp'; GlassEdit.submit()">Submit Changes</a><BR>
		<input class="redButton" type='submit' value='Delete Glass' onclick='return confirm(" Delete This Piece of Glass from Inventory? \n Only Delete if certain it will not affect other items. ")'></td>

            </form>

</body>
</html>
<%

rs.close
set rs=nothing

DBConnection.close
set DBConnection=nothing
%>

