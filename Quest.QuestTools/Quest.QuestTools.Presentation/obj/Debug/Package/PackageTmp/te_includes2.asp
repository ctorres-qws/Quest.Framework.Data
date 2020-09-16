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

	response.buffer = false
	

	'If protection is on, make sure that user has logged in
	if instr(request.servervariables("script_name"), "index.asp") = 0 then
		if bProtected then
			if session("teUserName") = "" then
				response.redirect "index.asp?comebackto=" & request.servervariables("script_name") & "?" & server.urlencode(request.querystring)
			end if
		end if
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


%>