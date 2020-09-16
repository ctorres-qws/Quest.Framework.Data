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
	'# Nov 15, 2000 by Hakan Eskici
	'Added permission assignment for Field functions
	'Added constants for fields
	'-------------------------------------------------------------

	if response.buffer = false then
	response.buffer = false
	end if
	
	const bProtected = true

	'If protection is on, make sure that user has logged in
	if instr(request.servervariables("script_name"), "index.asp") = 0 then
		if bProtected then
			if session("teUserName") = "" then
				response.redirect "index.asp?comebackto=" & request.servervariables("script_name") & "?" & server.urlencode(request.querystring)
			end if
		end if
	end if

	if bProtected then
		'If protection is on, get permissions
		'for the user from the session
		bAdmin = session("rAdmin")
		bTracking = session("rTracking")
		bSupplier = session("rSupplier")
		bMidLevel = session("rMidLevel")
		bQueryExec = session("rQueryExec")
		bSQLExec = session("rSQLExec")
		bTableAdd = session("rTableAdd")
		bTableEdit = session("rTableEdit")
		bTableDel = session("rTableDel")
		bFldAdd = session("rFldAdd")
		bFldEdit = session("rFldEdit")
		bFldDel = session("rFldDel")
		bFullName = session("rFullName")
	else
		'Not protected, give Full control
		bAdmin = True
		bTracking = False
		bSupplier = False
		bMidLevel = False
		bQueryExec = True
		bSQLExec = True
		bTableAdd = True
		bTableEdit = True
		bTableDel = True
		bFldAdd = True
		bFldEdit = True
		bFldDel = True
	end if

	'Pre-create connection and recordset objects
	set conn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	
	'Opens a given connection and initializes rs
	sub OpenRS(sConn)
		conn.open sConn
		set rs.ActiveConnection = conn
		rs.CursorType = adOpenStatic
	end sub
	
	'Closes open connections and releases objects
	sub CloseRS()
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end sub

	'---- CursorTypeEnum Values ----
	Const adOpenForwardOnly = 0
	Const adOpenKeyset = 1
	Const adOpenDynamic = 2
	Const adOpenStatic = 3

	'---- CursorLocationEnum Values ----
	Const adUseServer = 2
	Const adUseClient = 3

	'---- CommandTypeEnum Values ----
	Const adCmdUnknown = &H0008
	Const adCmdText = &H0001
	Const adCmdTable = &H0002
	Const adCmdStoredProc = &H0004
	Const adCmdFile = &H0100
	Const adCmdTableDirect = &H0200
	
	'---- SchemaEnum Values ----
	Const adSchemaTables = 20
	Const adSchemaPrimaryKeys = 28
	Const adSchemaIndexes = 12
	
	'---- DataTypeEnum Values ----
	Const adEmpty = 0
	Const adTinyInt = 16
	Const adSmallInt = 2
	Const adInteger = 3
	Const adBigInt = 20
	Const adUnsignedTinyInt = 17
	Const adUnsignedSmallInt = 18
	Const adUnsignedInt = 19
	Const adUnsignedBigInt = 21
	Const adSingle = 4
	Const adDouble = 5
	Const adCurrency = 6
	Const adDecimal = 14
	Const adNumeric = 131
	Const adBoolean = 11
	Const adError = 10
	Const adUserDefined = 132
	Const adVariant = 12
	Const adIDispatch = 9
	Const adIUnknown = 13
	Const adGUID = 72
	Const adDate = 7
	Const adDBDate = 133
	Const adDBTime = 134
	Const adDBTimeStamp = 135
	Const adBSTR = 8
	Const adChar = 129
	Const adVarChar = 200
	Const adLongVarChar = 201
	Const adWChar = 130
	Const adVarWChar = 202
	Const adLongVarWChar = 203
	Const adBinary = 128
	Const adVarBinary = 204
	Const adLongVarBinary = 205
	Const adChapter = 136
	Const adFileTime = 64
	Const adPropVariant = 138
	Const adVarNumeric = 139
	Const adArray = &H2000	
	
	'---- FieldAttributeEnum Values ----
	Const adFldMayDefer = &H00000002
	Const adFldUpdatable = &H00000004
	Const adFldUnknownUpdatable = &H00000008
	Const adFldFixed = &H00000010
	Const adFldIsNullable = &H00000020
	Const adFldMayBeNull = &H00000040
	Const adFldLong = &H00000080
	Const adFldRowID = &H00000100
	Const adFldRowVersion = &H00000200
	Const adFldCacheDeferred = &H00001000
	Const adFldIsChapter = &H00002000
	Const adFldNegativeScale = &H00004000
	Const adFldKeyColumn = &H00008000
	Const adFldIsRowURL = &H00010000
	Const adFldIsDefaultStream = &H00020000
	Const adFldIsCollection = &H00040000	
	
	
	Const adColFixed = 1
	Const adColNullable = 2	

%>