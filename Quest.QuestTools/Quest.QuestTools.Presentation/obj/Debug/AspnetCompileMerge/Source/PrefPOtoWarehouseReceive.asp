<!--#include file="dbpath2.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
         "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
		 
		 
<html xmlns="http://www.w3.org/1999/xhtml">
 <!-- Receive information from PrefPOtoWarehouseSplit.asp and then send to Warehouse-->
 <head>
 <title>Received and Sent</title>  
  
  </script>
 </head>
<body>
<%
' First get the Number of pieces that are being collected in order to receive the right amount of information from the form
dim NumberOfPieces
NumberOfPieces = Request.Form("pieces") 

'Read the PO Number for entry into the database
dim PO
PO = Request.Form("PO")

' Read the Is Filled Button



'Declare Arrays for each item
dim part()
dim quantity1()
dim aisle1()
dim rack1()
dim shelf1()
dim isFilled()
dim Total()

redim part(NumberOfPieces)
redim quantity(NumberOfPieces)
redim aisle(NumberOfPieces)
redim rack(NumberOfPieces)
redim shelf(NumberOfPieces)
redim isFilled(NumberOfPieces)
redim total(NumberOfPieces)


'Declare "Split" Array to check if Whole value or Split into Parts
dim Split()
redim Split(NumberOfPieces)


'Fill in the array with the base values
For current=1 To NumberOfPieces

	part(current)= Request.Form("partname" & current)
	quantity(current)= Request.Form("amount" & current)
	aisle(current)= Request.Form("aisle" & current)
	rack(current)= Request.Form("rack" & current)
	shelf(current)= Request.Form("shelf" & current)
	
	
	if Request.Form("Filled" & current) <> "" then
		isFilled(current)= "Yes"
	else
		isFilled(current)= "No"
	End if
	
	

	if Request.form("Amount" & current & ".1") <> "" then
		Split(current) = "Split"

	else 
		Split(current) = "Whole"
	end if


Next





%>
<!--Displays the Base values--> 
<h2>Moved to Warehouse for PO: <%response.write PO %> </h2>
<table border="1">
<tr><th>Part</th><th>Quantity</th><th>Aisle</th><th>Rack</th><th>Shelf</th><th>PO</th></tr>

<%

Dim CurrentAmount
Dim CurrentAisle
Dim CurrentRack
Dim CurrentShelf

For current=1 To NumberOfPieces	

	
		Response.write "<tr><td> " & part(current)
		Response.write "</td><td>" & quantity(current)
		Response.write "</td><td>" & aisle(current)
		Response.write "</td><td>" & rack(current)
		Response.write "</td><td>" & shelf(current)
		Response.write "</td><td>" & PO
		Response.write "</td></tr>"
	
	if Split(current) = "Split" then
		i=1
		do until Request.form("Amount" & current & "." & i) = "" OR Request.form("Amount" & current & "." & i) = "0"

			CurrentAmount = Request.form("Amount" & current & "." & i)
			CurrentAisle = Request.form("Aisle" & current & "." & i)
			CurrentRack = Request.form("Rack" & current & "." & i)
			CurrentShelf = Request.form("Shelf" & current & "." & i)
			Total(current) = CLng(Total(current)) + CLng(CurrentAmount)
			
		Response.write "<tr><td>" & part(current)
		Response.write "</td><td>" & currentAmount
		Response.write "</td><td>" & currentAisle
		Response.write "</td><td>" & currentRack
		Response.write "</td><td>" & currentShelf
		Response.write "</td><td>"  & PO
		Response.write "</td></tr>"

		i=i+1
		loop
	end if
	
	
	
	
Next
%>
</table>
<h2>Sent to Pref: <%response.write PO %> </h2>
<table border="1">
<tr><th>PO</th><th>Part</th><th>Quantity</th><th>Is Filled</th></tr>

<%
For current=1 To NumberOfPieces	

		Response.write "<tr><td>" & PO
		Response.write "</td><td> " & part(current)
		if Split(current) = "Split" then
			Response.write "</td><td>" & Total(current) + quantity(current)
		else
			Response.write "</td><td>" & quantity(current)
		end if
		
		Response.write "</td><td>" & isFilled(current)
		Response.write "</td></tr>"


Next


%>


</body>
</html>