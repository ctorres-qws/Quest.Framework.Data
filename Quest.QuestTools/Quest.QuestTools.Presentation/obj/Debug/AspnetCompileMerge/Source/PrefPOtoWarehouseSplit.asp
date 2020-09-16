<!--#include file="dbpath2.asp"-->
                       
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
<html xmlns="http://www.w3.org/1999/xhtml">		 
 <!-- Using Original code from - http://www.webdeveloper.com/forum/showthread.php?217602-How-to-create-form-with-button-to-generate-another-text-field-->
<head>
<title>Receive to Warehouse</title>


<% 
'Collect PO Number, PO Items, and PO Item Quantities from PREF

Dim PO
PO = request.QueryString("PO")
'PO = "7147"

'Access Database
Set rs = Server.CreateObject("adodb.recordset")
strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
rs.Cursortype = 2
rs.Locktype = 3
rs.Open strSQL, DBConnection


rs.filter = "PO ='" & PO & " '"

Dim PartsNumber 
PartsNumber = 0
Dim POQuantity()
Dim POPart()

if rs.eof or rs.bof then 
else

	Do while not rs.eof
		PartsNumber = PartsNumber + 1
	rs.movenext
	Loop

	rs.movefirst
	ReDim POPart(PartsNumber)
	ReDim POQuantity(PartsNumber)

	Do while not rs.eof
		POPart(PartsNumber) = rs.fields("Part")
		POQuantity(PartsNumber) = rs.fields("qty")
	rs.movenext
	Loop
end if


%>
<script>
var PartFieldWidth = "200"; // Field width for Part
var AddFieldWidth = "50"; // Field width for Qty, Aisle, Rack, Shelf

var partFieldNumber = new Array();
var numberOfParts = <% response.write PartsNumber %> // number of parts is automated from the Classic ASP code above

numberOfParts++ // (Arrays start at 0, so to have the partnumber match the array number, Add 1)


for (var i=0;i<numberOfParts;i++)
{ 
partFieldNumber[i] = 0;
}



function addRow(element, partnumber)
{
var elementClassName = element.className; // this is the class name of the button that was clicked
var fieldNumber = elementClassName.substr(3, elementClassName.length);

partFieldNumber[partnumber] = partFieldNumber[partnumber] +1 ;

var partTitle = document.getElementById("partName" + partnumber).value

// For each loop to get the part number and then collect the amount 


var amountTitle = partFieldNumber[partnumber];
// amountTitle should receive the Amount of partnumber, for now, newFieldNumber


// current partName width, take off the px on the end
var fieldWidth = document.getElementById("partName" + fieldNumber).style.width; 
fieldWidth = fieldWidth.substr(0, fieldWidth.indexOf("px"));

// get the new field number (incremented)


var rowContainer = element.parentNode; // get the surrounding div so we can add new elements

// create text field for Part

var partName = document.createElement("input");
partName.type = "text";
partName.setAttribute("value", partTitle);
partName.setAttribute("name", "Part" + partnumber + "." + partFieldNumber[partnumber]);
partName.id = "partName" + partnumber + "." + partFieldNumber[partnumber];
partName.setAttribute("style","width:" + PartFieldWidth + "px; margin-left: 41px "); 
partName.setAttribute("readonly", true);



// create text field for Amount
var Amount = document.createElement("input");
Amount.type = "text";
Amount.setAttribute("value", amountTitle);
Amount.setAttribute("name", "Amount" + partnumber + "." + partFieldNumber[partnumber]);
Amount.setAttribute("placeholder", "Quantity");
Amount.id = "Amount" + partnumber;
Amount.setAttribute("style","width:" + AddFieldWidth + "px ; margin-left: 58px ");


// create text field for Aisle
var Aisle = document.createElement("input");
Aisle.type = "text";
Aisle.setAttribute("value", "");
Aisle.setAttribute("name", "Aisle" + partnumber + "." + partFieldNumber[partnumber]);
Aisle.setAttribute("placeholder", "Aisle");
Aisle.id = "Aisle" + partnumber;
Aisle.setAttribute("style","width:" + AddFieldWidth + "px ; margin-left: 37px ");


// create text field for Rack
var Rack = document.createElement("input");
Rack.type = "text";
Rack.setAttribute("value", "");
Rack.setAttribute("name", "Rack" + partnumber + "." + partFieldNumber[partnumber]);
Rack.setAttribute("placeholder", "Rack");
Rack.id = "Rack" + partnumber;
Rack.setAttribute("style","width:" + AddFieldWidth + "px ; margin-left: 39px ");

// create text field for Shelf
var Shelf = document.createElement("input");
Shelf.type = "text";
Shelf.setAttribute("value", "");
Shelf.setAttribute("name", "Shelf" + partnumber + "." + partFieldNumber[partnumber]);
Shelf.setAttribute("placeholder", "Shelf");
Shelf.id = "Shelf" + partnumber;
Shelf.setAttribute("style","width:" + AddFieldWidth + "px ; margin-left: 36px ");


// add elements to page
rowContainer.appendChild(document.createElement("BR")); // add line break
rowContainer.appendChild(document.createTextNode(" ")); // add space
rowContainer.appendChild(partName);
rowContainer.appendChild(document.createTextNode(" ")); // add space
rowContainer.appendChild(Amount);
rowContainer.appendChild(document.createTextNode(" ")); // add space
rowContainer.appendChild(Aisle);
rowContainer.appendChild(document.createTextNode(" ")); // add space
rowContainer.appendChild(Rack);
rowContainer.appendChild(document.createTextNode(" ")); // add space
rowContainer.appendChild(Shelf);

}


function removeRow(element, partnumber) {
	//Declare the Form to remove existing form elements
	var elementClassName = element.className; // this is the class name of the button that was clicked
	var rowContainer = element.parentNode;
	
	// if loop not to remove if not split
	
	if (partFieldNumber[partnumber] > 0)
	{
	// Set up the Removal of the most recent item
		var RemovePart = document.getElementsByName("Part" + partnumber + "." + partFieldNumber[partnumber])[0];
		var RemoveQuantity = document.getElementsByName("Amount" + partnumber + "." + partFieldNumber[partnumber])[0]; 
		var RemoveAisle = document.getElementsByName("Aisle" + partnumber + "." + partFieldNumber[partnumber])[0];
		var RemoveRack = document.getElementsByName("Rack" + partnumber + "." + partFieldNumber[partnumber])[0];
		var RemoveShelf = document.getElementsByName("Shelf" + partnumber + "." + partFieldNumber[partnumber])[0]; 

	
	//Remove the Fields
		rowContainer.removeChild(RemovePart);
		rowContainer.removeChild(RemoveQuantity);
		rowContainer.removeChild(RemoveAisle);
		rowContainer.removeChild(RemoveRack);
		rowContainer.removeChild(RemoveShelf);
		

		
		
	//Reset the Most recent button by 1 to clear the addition
		partFieldNumber[partnumber] = partFieldNumber[partnumber] -1 ;

		var breaks = rowContainer.getElementsByTagName('BR');
		rowContainer.removeChild(breaks[partFieldNumber[partnumber]]);


	}

}


/*
function getTotal()
<!-- Do we need this function? -->
{
for (var i=0;i<numberOfParts;i++)
{ 

}

document.getElementById("Amount" + fieldNumber)
}
*/




</script>
</head>
<body>
	<div id="main-container">
		<form id="wreceive" title="Receive to Warehouse" class="panel" name="wreceive" action="PrefPOtoWarehouseReceive.asp" method="post" target="_self" selected="true" >
			<br>
			
			
<%

response.write "<h2> Parts from PO - " & PO & " </h2>"
response.write "<input type='hidden' id='pieces' name='pieces' value='"& PartsNumber & "' /> "
response.write "<input type='hidden' id='PO' name='PO' value='"& PO & "' /> "

'Creates each row for the parts in the PO - using the Part number as the Variable 		
For currentPart=1 To PartsNumber
response.write "<div class='row'>"
response.write "<label>Part " & currentPart &"</label> "
response.write "<input type='text' id='partName" & currentPart & "' name='partName" & currentPart & "' value='"& POPart(currentPart) & "' style='width:200px;' readonly /> "
response.write "<label> Quantity: </label> <input type='text' id='Amount' name='Amount" & currentPart & "' value='" & POQuantity(currentPart) & "' style='width:50px;'  />"
response.write "<label> Aisle: </label> <input type='text' id='Aisle' name='Aisle" & currentPart & "' value='' style='width:50px;'  />"
response.write "<label> Rack: </label> <input type='text' id='Rack' name='Rack" & currentPart & "' value='' style='width:50px;'  />"
response.write "<label> Shelf: </label> <input type='text' id='Shelf' name='Shelf" & currentPart & "' value='' style='width:50px;' />"
response.write "<input type='button' class='row" & currentPart & "' value='+' onclick='addRow(this, " & currentPart & ")'> "
response.write "<input type='button' class='row" & currentPart & "' value='-' onclick='removeRow(this, " & currentPart & ")'> "
response.write "<label> Is Filled: </label><input type='checkbox' id='Filled' name='Filled" & currentPart & "'  value='Checked' />"
response.write "<label id='original'> Original PO-" & PO & " Quantity of: " & POQuantity(currentPart) & "<label>"
response.write "</div><br>"

Next

%>		
		
	
				
				
<!-- Plain HTML code - translated above to ASP for each item in the PO (Variable number of items) - This is kept as a template
			<div class="row">
				<label>Part 2</label>
				<input type="text" id="partName 2" name='partName 2' value="Piece2" style="width:500px;" readonly />
				<label>Quantity: </label>
				<input type="text" id="Amount" name='Amount2' value="50" style="width:50px;"  />
				<label>Aisle: </label>
				<input type="text" id="Aisle" name='Aisle2' value="" style="width:50px;"  />
				<label>Rack: </label>
				<input type="text" id="Rack" name='Rack2' value="" style="width:50px;"  />
				<label>Shelf: </label>
				<input type="text" id="Shelf" name='Shelf2' value="" style="width:50px;" />
				<input type="button" class="row2" value="Break into Parts" onclick="addRow(this, '2')">
				<label> Original PO-XXXX Quantity of: 50<label>
			</div><br>
-->			
			
			
			<input type="button" class="whiteButton" value="Receive to Warehouse" onclick="javascript:wreceive.submit()">

		</form>
	</div>

</body>
</html>