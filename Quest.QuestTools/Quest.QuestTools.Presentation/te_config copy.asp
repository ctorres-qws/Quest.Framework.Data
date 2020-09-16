<%
	'-------------------------------------------------------------
	'TableEditoR 0.6 Beta
	'http://www.2enetworx.com/dev/projects/tableeditor.asp
	
	'File: te_includes.asp
	'Description: Constants and Public Functions
	'Written By Hakan Eskici on Nov 01, 2000

	'You may use the code for any purpose
	'But re-publishing is discouraged.
	'See License.txt for additional information	

	'Change Log:
	'-------------------------------------------------------------
	'# Nov 16, 2000 by Kevin Yochum
	'Added switches for converting null values
	'-------------------------------------------------------------


	'Define your total number of connections here
	const iTotalConnections = 1

	'How many records in a page?
	const cPerPage = 10
	
	'Encode HTML tags?
	'Turn this on if you have problems with displaying 
	'records with html content.
	const bEncodeHTML = True
	
	'Maximum number of chars to display in table view (0 : no limit)
	'Warning: If you have HTML content in your fields;
	'you should set bEncodeHTML to True if you specify a limit
	const lMaxShowLen = 0
		 
   ' Should blank fields be converted to NULL when the field is nullable?
   ' Convert '' to null in non-numeric and non-date fields?
   Const bConvertNull = False
   ' Convert '' and 0 to null in numeric fields?
   Const bConvertNumericNull = False
   ' Convert '' and 0 to null in date fields?
   Const bConvertDateNull = False

	'Using Redim causes a performance degredation
	'But it's OK since array size is small
	redim arrDbs(iTotalConnections)
	redim arrDesc(iTotalConnections)
	redim arrConn(iTotalConnections)

	'Define database locations
	arrDBs(0) = "database/teadmin.mdb"
	arrDBs(1) = "database/rainingcash.mdb"
	
	'Define descriptions for corresponding connections
	arrDesc(0) = "Table Editor User Administration"
	arrDesc(1) = "RainingCash.com"

	'Construct connection strings
	for iConnection = 0 to iTotalConnections
		arrConn(iConnection) = "Provider=Microsoft.Jet.OLEDB.4.0;" &_
	     "Persist Security Info=False;" &_
	     "Data Source=" & Server.MapPath(arrDBs(iConnection))
	next
	
%><!--#include file="te_includes.asp"-->