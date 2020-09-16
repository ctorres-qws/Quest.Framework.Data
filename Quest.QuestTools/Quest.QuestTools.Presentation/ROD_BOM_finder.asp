<%
' BIll or Material Listing for Connecting Rods - Rod_Array dictates the possible BOM lists
' Created by Daniel Zalcman, Coded by Michael Bernholtz, at Behest of Jody Cash
' Reads BOM number and allocates correct number of materials to each BOM List
' Used to Create Job/Floor Hardware requirements.
  
'Different Rules for Awning and Casement - August 2017 - Just Awning
' Awning Code

'BOM

'---------------------------------------------------------------BOM Values--------------------------------------------------

AKeeper = 0
APin = 0
AOpeningMech = 0
ATransmission = 0
AShavedTransmission = 0

Select Case BOM
'-----------------------------------------------------------Shaved Pin
CASE "101"
	AKeeper = 7
	APin = 6
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1
CASE "102"
	AKeeper = 2
	APin = 2
	AOpeningMech = 1
	ATransmission = 0
	AShavedTransmission = 1
CASE "103"
	AKeeper = 3
	APin = 3
	AOpeningMech = 1
	ATransmission = 0
	AShavedTransmission = 1
CASE "104"	
	AKeeper = 5
	APin = 4
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 1
CASE "105"
	AKeeper = 3
	APin = 3
	AOpeningMech = 2
	ATransmission = 0
	AShavedTransmission = 1
CASE "106"
	AKeeper = 4
	APin = 3
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 1
CASE "107"
	AKeeper = 6
	APin = 5
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 1
CASE "108"
	AKeeper = 6
	APin = 5
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1
CASE "109"
	AKeeper = 8
	APin = 7
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1
CASE "110"
	AKeeper = 4
	APin = 4
	AOpeningMech = 2
	ATransmission = 0
	AShavedTransmission = 1
CASE "111"
	AKeeper = 4
	APin = 3
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1
CASE "112"
	AKeeper = 5
	APin = 4
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1
CASE "113"
	AKeeper = 12
	APin = 11
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1
CASE "114"
	AKeeper = 4
	APin = 3
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1
CASE "115"
	AKeeper = 5
	APin = 5
	AOpeningMech = 2
	ATransmission = 0
	AShavedTransmission = 1
CASE "116"
	AKeeper = 9
	APin = 8
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 1

'111 and '114 appear to be the same - September 2018	
	
	
'--------------------------200-------------------
CASE "201"
	AKeeper = 2
	APin = 2
	AOpeningMech = 1
	ATransmission = 0
	AShavedTransmission = 0
CASE "202"
	AKeeper = 2
	APin = 2
	AOpeningMech = 1
	ATransmission = 0
	AShavedTransmission = 1
CASE "203"
	AKeeper = 3
	APin = 2
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 0
CASE "204"
	AKeeper = 2
	APin = 2
	AOpeningMech = 2
	ATransmission = 0
	AShavedTransmission = 0
CASE "205"
	AKeeper = 3
	APin = 2
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 0
CASE "206"
	AKeeper = 6
	APin = 4
	AOpeningMech = 1
	ATransmission = 2
	AShavedTransmission = 0
CASE "207"
	AKeeper = 6
	APin = 4
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0
CASE "208"
	AKeeper = 8
	APin = 6
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0
CASE "209"
	AKeeper = 3
	APin = 3
	AOpeningMech = 1
	ATransmission = 0
	AShavedTransmission = 1
CASE "210"
	AKeeper = 4
	APin = 2
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0
CASE "211"
	AKeeper = 5
	APin = 4
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 0
CASE "212"
	AKeeper = 4
	APin = 2
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 1
CASE "213"
	AKeeper = 6
	APin = 4
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 1
CASE "214"
	AKeeper = 12
	APin = 10
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0
CASE "215"
	AKeeper = 4
	APin = 2
	AOpeningMech = 1
	ATransmission = 2
	AShavedTransmission = 0
CASE "216"
	AKeeper = 4
	APin = 4
	AOpeningMech = 2
	ATransmission = 0
	AShavedTransmission = 0
CASE "217"
	AKeeper = 7
	APin = 5
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0
CASE "218"
	AKeeper = 9
	APin = 7
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0	
	
	
	
'--------------------------300-------------------
CASE "301"
	AKeeper = 2
	APin = 2
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 0 
CASE "302"
	AKeeper = 2
	APin = 2
	AOpeningMech = 1
	ATransmission = 0
	AShavedTransmission = 1 
CASE "303"
	AKeeper = 3
	APin = 2
	AOpeningMech = 1 
	ATransmission = 1
	AShavedTransmission = 0 
CASE "304"
	AKeeper = 2
	APin = 2
	AOpeningMech = 2 
	ATransmission = 1
	AShavedTransmission = 0 
CASE "305"
	AKeeper = 3
	APin = 2
	AOpeningMech = 2 
	ATransmission = 1
	AShavedTransmission = 0 
CASE "306"
	AKeeper = 6
	APin = 4
	AOpeningMech = 1 
	ATransmission = 2
	AShavedTransmission = 0 
CASE "307"
	AKeeper = 6
	APin = 4
	AOpeningMech = 2 
	ATransmission = 2
	AShavedTransmission = 0 
CASE "308"
	AKeeper = 8
	APin = 6
	AOpeningMech = 2 
	ATransmission = 2
	AShavedTransmission = 0 
CASE "309"
	AKeeper = 3
	APin = 3
	AOpeningMech = 1 
	ATransmission = 0
	AShavedTransmission = 1 
CASE "310"
	AKeeper = 4
	APin = 2
	AOpeningMech = 2 
	ATransmission = 2
	AShavedTransmission = 0 
CASE "311"
	AKeeper = 5
	APin = 4
	AOpeningMech = 2 
	ATransmission = 1
	AShavedTransmission = 0
CASE "312"
	AKeeper = 4
	APin = 2
	AOpeningMech = 1 
	ATransmission = 2
	AShavedTransmission = 0 
CASE "313"
	AKeeper = 6
	APin = 4
	AOpeningMech = 1
	ATransmission = 1
	AShavedTransmission = 1 
CASE "314"
	AKeeper = 12
	APin = 10
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0 
CASE "315"
	AKeeper = 4
	APin = 3
	AOpeningMech = 2
	ATransmission = 1
	AShavedTransmission = 0 
CASE "316"
	AKeeper = 6
	APin = 4
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0 
CASE "317"
	AKeeper = 7
	APin = 5
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0 
CASE "318"
	AKeeper = 9
	APin = 7
	AOpeningMech = 2
	ATransmission = 2
	AShavedTransmission = 0 
End Select
	
		
			
' Casement Code


%>