<!--#include file="dbpath.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- May 2019 -->
<!-- DoorStyle pages collect information about Door types to get Machining Data for Emmegi Saws -->
<!-- Programmed by Michelle Dungo - At request of Ariel Aziza, using PanelStyle Pages as a template -->
<!-- DoorStyle.asp (General View) -- DoorStyleEditForm.asp (Manage Form) -- DoorStyleEditConf.asp (Manage Submit) -- DoorStyleEnter.asp (Enter Form)-- DoorStyleConf.asp (Enter Submit)--DoorStyleByJob.asp (view By Job filter) -->
<!-- SQL Table StylesDoor - NOT IN ACCESS -->




<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>Panel Styles</title>
  <meta name="viewport" content="width=devicewidth; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;"/>
  <link rel="apple-touch-icon" href="/iui/iui-logo-touch-icon.png" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <link rel="stylesheet" href="/iui/iui.css" type="text/css" />

  <link rel="stylesheet" title="Default" href="/iui/t/default/default-theme.css"  type="text/css"/>
  <script type="application/x-javascript" src="/iui/iui.js"></script>
  <script type="text/javascript">
    iui.animOn = true;
    </script>
		<SCRIPT language="javascript">
		function addElement(parentId, elementTag, elementId, html) {
			// Adds an element to the document
			var p = document.getElementById(parentId);
			var newElement = document.createElement(elementTag);
			newElement.setAttribute('id', elementId);
			newElement.innerHTML = html;
			p.appendChild(newElement);
		}
		function removeElement(elementId) {
		// Removes an element from the document
		var element = document.getElementById(elementId);
		element.parentNode.removeChild(element);
		}
		var fileId = 0; // used by the addFile() function to keep track of IDs

		function addFile() {
		fileId++; // increment fileId to get a unique ID for the new element
		var html = '<input type="file" name="uploaded_files[]" /> ' +
               '<a href="" onclick="javascript:removeElement('file-' + fileId + ''); return false;">Remove</a>';
		addElement('files', 'p', 'file-' + fileId, html);
}

		function addRow(tableID) {

			var table = document.getElementById(tableID);

			var rowCount = table.rows.length;
			var row = table.insertRow(rowCount);

			var colCount = table.rows[0].cells.length;

			for(var i=0; i<colCount; i++) {

				var newcell	= row.insertCell(i);

				newcell.innerHTML = table.rows[0].cells[i].innerHTML;
				//alert(newcell.childNodes);
				switch(newcell.childNodes[0].type) {
					case "text":
							newcell.childNodes[0].value = "";
							break;
					case "checkbox":
							newcell.childNodes[0].checked = false;
							break;
					case "select-one":
							newcell.childNodes[0].selectedIndex = 0;
							break;
				}
			}
		}

		function deleteRow(tableID) {
			try {
			var table = document.getElementById(tableID);
			var rowCount = table.rows.length;

			for(var i=0; i<rowCount; i++) {
				var row = table.rows[i];
				var chkbox = row.cells[0].childNodes[0];
				if(null != chkbox && true == chkbox.checked) {
					if(rowCount <= 1) {
						alert("Cannot delete all the rows.");
						break;
					}
					table.deleteRow(i);
					rowCount--;
					i--;
				}


			}
			}catch(e) {
				alert(e);
			}
		}

	</SCRIPT>
 
    </head>
<body>
    <div class="toolbar">
        <h1 id="pageTitle"></h1>
        <a class="button leftButton" type="cancel" href="DoorStyle.asp" target="_self">Styles</a>
    </div>
   
	<form id="enter" title="Enter New Door Style" class="panel" name="enter" action="DoorStyleConf.asp" target="_self" selected="true">
		<h2>Enter New Door Style:</h2>
		
		<fieldset>
			<div class="row">
				<label>Job</label>
				<input type="text" name='Job' id='Job' >
			</div>
			
			<div class="row">
				<label>Name</label>
				<input type="text" name='Name' id='Name' >
			</div>
			
			<div class="row">
				<label>Int</label>
				<Select name='IntDoorType'>
					<option value="Fapim">Fapim</option>
					<option value="Metra">Metra</option>
				</Select>				
			</div>			

			<div class="row">
				<label>Ext</label>
				<Select name='ExtDoorType'>
					<option value="Fapim">Fapim</option>
					<option value="Hopi">Hopi</option>					
					<option value="Metra">Metra</option>
					<option value="None">None</option>					
				</Select>				
			</div>
			


            <a class="whiteButton" href="javascript:enter.submit()">Submit</a>
		</fieldset>
  
    </form>
                
<%
DBConnection.close
Set DBConnection = nothing
%>		 
               
</body>
</html>
