<html>
<!--https://stackoverflow.com/questions/29698796/how-to-convert-html-table-to-excel-with-multiple-sheet -->

<!--#include file="dbpath.asp"-->
<!-- Import JUPITER Inventory into Excel for Inventory Counts Splitting up BY Aisle Rack-->
<!-- Table Values Organized by Aisle Rack Shelf -->
<!-- JUPITER Inventory Requested by David Jan 2020 for Texas -->
<!-- All Items at once was too big, so broken down into Aisle-->



<head>
<script type="text/javascript">
var tablesToExcel = (function() {
    var uri = 'data:application/vnd.ms-excel;base64,'
    , tmplWorkbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
      + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Axel Richter</Author><Created>{created}</Created></DocumentProperties>'
      + '<Styles>'
      + '<Style ss:ID="Currency"><NumberFormat ss:Format="Currency"></NumberFormat></Style>'
      + '<Style ss:ID="Date"><NumberFormat ss:Format="Medium Date"></NumberFormat></Style>'
      + '</Styles>' 
      + '{worksheets}</Workbook>'
    , tmplWorksheetXML = '<Worksheet ss:Name="{nameWS}"><Table>{rows}</Table></Worksheet>'
    , tmplCellXML = '<Cell{attributeStyleID}{attributeFormula}><Data ss:Type="{nameType}">{data}</Data></Cell>'
    , base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) }
    , format = function(s, c) { return s.replace(/{(\w+)}/g, function(m, p) { return c[p]; }) }
    return function(tables, wsnames, wbname, appname) {
      var ctx = "";
      var workbookXML = "";
      var worksheetsXML = "";
      var rowsXML = "";

      for (var i = 0; i < tables.length; i++) {
        if (!tables[i].nodeType) tables[i] = document.getElementById(tables[i]);
        for (var j = 0; j < tables[i].rows.length; j++) {
          rowsXML += '<Row>'
          for (var k = 0; k < tables[i].rows[j].cells.length; k++) {
            var dataType = tables[i].rows[j].cells[k].getAttribute("data-type");
            var dataStyle = tables[i].rows[j].cells[k].getAttribute("data-style");
            var dataValue = tables[i].rows[j].cells[k].getAttribute("data-value");
            dataValue = (dataValue)?dataValue:tables[i].rows[j].cells[k].innerHTML;
            var dataFormula = tables[i].rows[j].cells[k].getAttribute("data-formula");
            dataFormula = (dataFormula)?dataFormula:(appname=='Calc' && dataType=='DateTime')?dataValue:null;
            ctx = {  attributeStyleID: (dataStyle=='Currency' || dataStyle=='Date')?' ss:StyleID="'+dataStyle+'"':''
                   , nameType: (dataType=='Number' || dataType=='DateTime' || dataType=='Boolean' || dataType=='Error')?dataType:'String'
                   , data: (dataFormula)?'':dataValue
                   , attributeFormula: (dataFormula)?' ss:Formula="'+dataFormula+'"':''
                  };
            rowsXML += format(tmplCellXML, ctx);
          }
          rowsXML += '</Row>'
        }
        ctx = {rows: rowsXML, nameWS: wsnames[i] || 'Sheet' + i};
        worksheetsXML += format(tmplWorksheetXML, ctx);
        rowsXML = "";
      }

      ctx = {created: (new Date()).getTime(), worksheets: worksheetsXML};
      workbookXML = format(tmplWorkbookXML, ctx);



      var link = document.createElement("A");
      link.href = uri + base64(workbookXML);
      link.download = wbname || 'JUPITER_Inventory.xls';
      link.target = '_blank';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  })();
	
  </script>




</Head>
<body>
<%

Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV WHERE WAREHOUSE = 'JUPITER' ORDER BY AISLE ASC, RACK ASC, SHELF ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection

PreviousAisle = "XXX"
PreviousRack = "XXX"
CurrentAisle = "XXXX"
CurrentRack = "XXXX"
	If CurrentAisle = "" then
		CurrentAisle = "Blank"
	end if
	If CurrentRack = "" then
		CurrentRack = "Blank"
	end if		
	If PreviousAisle = "" then
		CurrentAisle = "Blank"
	end if
	If PreviousRack = "" then
		CurrentRack = "Blank"
	end if	

Dim TableName(36)
Dim ColumnName(36) 
TableNum = 1

'FIRST TAB IS MISSING 'AA_1   
'PRINT BUTTON SHOULD HHAVE AISLE NAME NOT NUM

Response.write "<table id='"& CurrentAisle & "_" & CurrentRack & "' class='table2excel' border='1'>"
Response.write "<tr>"
Response.write "<tr><td colspan = 6> JUPITER "& CurrentAisle & " " & CurrentRack & "</td><td>BY: </td><td></td><td>CHECKED: </td><td></td><td>ENTERED:</td><td></td></tr>"
Response.write "<th>Aisle</th><th>Rack</th><th>Shelf</th><th>Part</th><th>PO</th><th>Bundle</th><th>External Bundle</th><th>Colour</th><th>Input Date</th><th>Sheet Size</th><th>Quantity</th><th>Variance</th>"
Response.write "</tr>"

Do While not rs.eof
	PreviousAisle = CurrentAisle
	PreviousRack = CurrentRack
	CurrentAisle = trim(rs("Aisle"))
	CurrentRack = trim(rs("Rack"))
	If CurrentAisle = "" then
		CurrentAisle = "Blank"
	end if
	If CurrentRack = "" then
		CurrentRack = "Blank"
	end if	
	
	
		if (PreviousAisle = CurrentAisle) AND (PreviousRack = CurrentRack) then
		'Same Aisle and Rack need no Special changes, just get added to table
		else
			if PreviousAisle = CurrentAisle then
			'Same Aisle gets a new Rack, as a tab
			else
				if TableName(TableNum) = "" then
				else
					TableName(TableNum) = TableName(TableNum) & ",'" & PreviousAisle & "_" & PreviousRack & "'"& "]"
				end if
				ColumnName(TableNum) = PreviousAisle
				TableNum = TableNum + 1
				
				
			end if
			
			if TableName(TableNum) = "" then
				TableName(TableNum) = TableName(TableNum) & "[" 
		
			else
				if TableName(TableNum) = "[" then
					TableName(TableNum) = TableName(TableNum) & "'" & PreviousAisle & "_" & PreviousRack & "'"
				else
					TableName(TableNum) = TableName(TableNum) & ",'" & PreviousAisle & "_" & PreviousRack & "'"
				end if	
					
			end if

			
			
			Response.write "</Table>"
			Response.write "<table id='"& CurrentAisle & "_" & CurrentRack & "' class='table2excel' border = '1'>"
			Response.write "<tr><td>JUPITER</td><td>"& CurrentAisle & "</td><td>" & CurrentRack & "</td><td></td><td></td><td></td><td>BY: </td><td></td><td>CHECKED: </td><td></td><td>ENTERED:</td><td></td></tr>"
			Response.write "<tr>"
			Response.write "<th>Aisle</th><th>Rack</th><th>Shelf</th><th>Part</th><th>PO</th><th>Bundle</th><th>External Bundle</th><th>Colour</th><th>Input Date</th><th>Sheet Size</th><th>Quantity</TH><TH>Delta</TH>"
			Response.write "</tr>"
		
		end if
		
		Response.write "<TR>"
		Response.write "<TD>" & rs("Aisle") & "</TD>"
		Response.write "<TD>" & rs("Rack") & "</TD>"
		Response.write "<TD>" & rs("Shelf") & "</TD>"
		Response.write "<TD>" & rs("Part") & "</TD>"
		Response.write "<TD>" & rs("PO") & "</TD>"
		Response.write "<TD>" & rs("Bundle") & "</TD>"
		Response.write "<TD>" & rs("ExBundle") & "</TD>"
		Response.write "<TD>" & rs("Colour") & "</TD>"
		Response.write "<TD>" & rs("DateIn") & "</TD>"
		Response.write "<TD>"
			if int(rs("width")) >1 then
				response.write RS("width") & " by " & RS("height") 
			else 
				response.write " " 
			end if 
		Response.write "</TD>"
		
		Response.write "</TR>"
		
	RS.movenext
	Loop

	TableName(TableNum) = TableName(TableNum) & ",'" & CurrentAisle & "_" & CurrentRack & "'"& "]"
	ColumnName(TableNum) = CurrentAisle
	Response.write "</Table>"
	
	TestValue = 2
	Do until TestValue > TableNum
		if Left(TableName(TestValue),3) = "[,'" then
			TableName(TestValue) = replace(TableName(TestValue),"[,'","['")
		end if
		Response.write TableName(TestValue)
		TestValue = TestValue + 1
	
	
	
	%>
	<button  onclick="tablesToExcel(<%response.write TableName(TestValue-1)%>, <%response.write TableName(TestValue-1)%>, 'JUPITER_Inventory_<%response.write ColumnName(TestValue-1)%>.xls', 'Excel')">Export to Excel <%response.write ColumnName(TestValue-1)%></button>
	<br>
	<%
	Loop
	
%>

<%
rs.close
set rs = nothing

DBConnection.close
Set DBConnection = nothing

%>

</body>


</HTML>