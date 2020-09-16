<!--#include file="@common.asp"-->
<%
' Create  DSN Less connection to Access Database FOR QC
' QCDatabase holds Glass Inventory in a seperate Access Database QualityControlDB.
' QC Inventory and Master Inventory Items for Glass, Spacer, Sealant, Misc
' One Master Table for each type, but Canada and USA have Seperate Databases
'Create DBConnection Object

Set DBConnection = Server.CreateObject("adodb.connection")
DSN = GetConnectionStrQC(b_SQL_Server) 'method in @common.asp

'Added Michael Bernholtz November 2018
DBConnection.ConnectionTimeout = 1000
DBConnection.CommandTimeout = 1000
'End add

DBConnection.Open DSN

' Location Marker
%>
<!--#include file="CountryLocation.INC"-->

	
