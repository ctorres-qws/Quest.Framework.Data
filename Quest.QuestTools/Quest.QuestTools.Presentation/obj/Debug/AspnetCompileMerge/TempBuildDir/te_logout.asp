<%
	'-------------------------------------------------------------
	'TableEditoR 0.5 Beta
	'http://www.2enetworx.com/dev/projects/tableeditor.asp
	
	'File: te_logout.asp
	'Description: Ends a session
	'Written By Hakan Eskici on Nov 01, 2000

	'You may use the code for any purpose
	'But re-publishing is discouraged.
	'See License.txt for additional information	
	'-------------------------------------------------------------

	session.abandon
	response.redirect "index.asp"
%>