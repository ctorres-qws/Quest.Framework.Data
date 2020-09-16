<!--#include file="dbpath.asp"-->                     
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		<!--Optima Selection Page for adding notes, shows all items that do not have an  Completed Date and a checkbox-->
		<!--Created July 2014, at Request of Sasha for adding a note to multiple items at once-->
		<!--Receives note from GlassOptimaNoteMultiple.asp -->
		<!-- Sends to glassOptimaNoteSelectConf.asp-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Report</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />
	 <script src="sorttable.js"></script>
  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
    <style type="text/css">
	ul{
    margin: 0;
    padding: 0;
   }
   </style>
<%

AddNote = Request.form("Note")

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Z_GLASSDB ORDER BY ID ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

%>

    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass Tools</a>
        </div>

       <form id="Optima" action="glassOptimaNoteConf.asp" name="Optima"  method="GET" target="_self" selected="true" >  

		<h2><center>Add Key Details to many records all at once by selecting the checkbox<center></h2>
		<h2><center>For Quest Glass please put MO number into the Ext/Int Work # Field<center></h2>
		<h2><center>For Externally Ordered Glass please put Order Number into the Ext/Int Work # Field<center></h2>
		<fieldset>

			<div class="row">
					<label>Add Note</label>
					<input type="text" name='NOTES' id='NOTES' />
			</div>
			<div class="row">
               <label>Window PO</label>
               <input type="text" name='PoNum' id='PoNum' >
            </div>

			<div class="row">
               <label>Ext Work #</label>
               <input type="text" name='ExtorderNum' id='ExtorderNum' >
            </div>

			<div class="row">
               <label>Int Work #</label>
               <input type="text" name='IntorderNum' id='IntorderNum' >
            </div>

			<div class="row">
               <label>QT File</label>
               <input type="text" name='QTFile' id='QTFile' >
            </div>
               <input type="hidden" name='ticket' id='ticket' value = 'multiple' />     
		<a class="whiteButton" onClick="Optima.action='GlassOptimaNoteConf.asp'; Optima.submit()">ADD Key Details</a><BR>
		</fieldset>
        <ul id="Profiles" title=" Optima Report" selected="true">

<%

response.write "<li class='group'>Choose Records below to add the Note</li>"
response.write "<li><table border='1' class='sortable'><tr><th></th><th>ID</th><th>Job</th><th>Floor</th><th>Tag</th><th>Width</th><th>Height</th><th>1 Mat</th><th>1 SPAC</th><th>2 Mat</th><th>Type</th><th>Order</th><th>PO</th><th>QT File Name</th><th>Notes</th><th>TimeLine</th></tr>"

Do While not rs.eof
	If not isdate(RS("COMPLETEDDATE")) Then
		response.write "<tr><td><input type='checkbox' name='GID' value='" & RS("ID")& "'></td>"
		response.write" <td>" & RS("ID") & "</td><td>" & RS("JOB") & "</td><td>" & RS("FLOOR") &"</td><td>" & RS("TAG") & "</td><td>" & RS("DIM X") & "''</td><td>" & RS("DIM Y") & "''</td><td>" & RS("1 MAT") & "</td><td>" & RS("1 SPAC") & "</td><td>" & RS("2 MAT") & "</td>" 
		response.write "<td>" & RS("DEPARTMENT") & "</td><td>" & RS("ORDERBY") & "</td><td>" & RS("PO") & "</td><td>" & RS("QTFile") & "</td><td>" & RS("NOTES") & "</td>"
		response.write "<td><a class = 'greenButton' href='glassTimeLine.asp?gid="  & RS("ID") & "&ticket=production' target ='#_blank' >Time Line</a> </td>"
		response.write "</tr>"
	End If
	rs.movenext
Loop

rs.close
set rs = nothing
DBConnection.close 
set DBConnection = nothing

%>

	</table>

      </ul>
		</form>

</body>
</html>
