<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="dbpath.asp"-->
<!-- Created February 6th, by Michael Bernholtz - Entry Form to input items to QC Inventory Tables-->
<!-- QC_INVENTORY Tables created for Victor at Request of Jody Cash, Implemented by Michael Bernholtz-->  


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Glass Import</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle">Add Items to QC Inventory</h1>
        <a class="button leftButton" type="cancel" href="index.html#_Glass" target="_self">Glass</a>
        </div>

              <form id="enter" title="Add QC Master Item" class="panel" name="enter" action="GlassUploadConf.asp" method="GET" target="_self" selected="true">
              
			  <h2>Add QC Inventory Master Item</h2>
			  <h2>Please find the Excel Template in the UploadsRecords folder</h2> 
			  <a href="/UploadRecords/Template/Glass_Upload_Template.xls"  download="filename" target="_blank">Download link</a>
			  <h3>Excel Template must be filled in and stored in \\172.18.13.31\_Websites\Prod\QWS_Tools\UploadRecords</h3>
			  <h3>Template records must be saved as .xls (NOT .xlsx) in order to import</h3>
			   <h3>Do not put commas in the file name.</h3>

              <fieldset>

        <div class="row">
            <label>File Name</label>
            <input type="File" name='ItemName' id='ItemName' >
        </div>

        <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
		


</fieldset>

		<Table border = "1"><tr><th>CODE</th><th>Description</th></tr>
		<%
		Set rs = Server.CreateObject("adodb.recordset")
		strSQL = "Select * FROM XQSU_GLASSTYPES ORDER BY TYPE ASC"
		rs.Cursortype = GetDBCursorType
		rs.Locktype = GetDBLockType
		rs.Open strSQL, DBConnection
		
		Do While not rs.eof
		
		Response.write "<tr>"
		Response.write "<TD>" & RS("TYPE") & "</TD>"
		Response.write "<TD>" & RS("DESCRIPTION") & "</TD>"
		Response.write "</tr>"
		
		rs.movenext
		loop
		
		rs.close
		set rs = nothing
		DBConnection.close
		set DBConnection = nothing
		
		
		
		%>
		
		
		</Table>

            </form>
</body>
</html>
