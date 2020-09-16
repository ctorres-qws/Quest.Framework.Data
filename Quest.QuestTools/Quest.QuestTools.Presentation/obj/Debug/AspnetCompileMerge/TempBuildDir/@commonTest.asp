<!--#include file="@@siteTest.asp"-->
<%

	gbMyDebug = true
	ga_Months = Split(",JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC",",")
	gstr_WareHouses = ""


	If gi_Env = c_ENV_Dev Then
		'Response.Write("<style> body > .panel { background: #FFCC99 url(pinstripes.png) !important; }</style>")
		Response.Write("<script language='javascript'> window.addEventListener('load', function(event) { document.title = document.title + ' ***SANDBOX***'; });</script>")
		Response.Write("<div style='font-size: 16px; font-weight: bold; display: block !important; z-index:9999; position: absolute; left: 300px; top: 69px; color: tomato;'>*** DEV - ENVIRONMENT - DO NOT USE ***&nbsp;&nbsp;*** DEV - ENVIRONMENT - DO NOT USE ***&nbsp;&nbsp;*** DEV - ENVIRONMENT - DO NOT USE ***</div>")
	End If

	If gi_Env = c_ENV_Home Or gi_Env = c_ENV_Local Then
		'Response.Write("<meta http-equiv='X-UA-Compatible' content='IE=edge' />")
		'Response.Write("<meta http-equiv='X-UA-Compatible' content='IE=10' />")
	End If

	If gi_Env = c_ENV_Home Then
		Response.Write("<meta http-equiv='X-UA-Compatible' content='IE=11'>")
	End If

	g_SyncID = true 'Turn this off when running in SQL Server Mode

	If gi_Env = c_ENV_Home Or gi_Env = c_ENV_Local Or gi_Env = c_ENV_Dev Then
		Server.ScriptTimeout = 300
		'On Error Resume Next
		'Response.Clear()
		'Response.Buffer = false
		'On Error GoTo 0
	End If

	gstr_OmitActivity = "|iflash.asp|flashreport.asp|index.asp|"
	gstr_OmitDebugMsg = ""
	gstr_DateFormat = "M/D/YYYY"
	gstr_DateFormat2 = "YYYY-MM-DD"
	gstr_SQLDateType = "date"

	gstr_SQLDB_Primary = "[qws_Dev].[dbo]"
	gstr_SQLDB_Secondary = "[qws_Dev_Secondary].[dbo]"

	If gi_Env = c_ENV_Prod Then
		gstr_SQLDB_Primary = "[qws_Prod].[dbo]"
		gstr_SQLDB_Secondary = "[qws_Prod_Secondary].[dbo]"
	End If

	If gi_Env = c_ENV_Home Then
		gstr_DateFormat = "D/M/YYYY"
		gstr_SQLDateType = "datetime"
	End If

	If gi_Env = c_ENV_Home Or gi_Env = c_ENV_Local Then
		gstr_TestDate = "22/08/2017"  ' DD/MM/YYYY
	End If

	gstr_ID1 = ""
	gstr_ID2 = ""
	gstr_ID3 = ""
	gstr_ID4 = ""
	gstr_ID5 = ""
	gstr_ID6 = ""
	gstr_ErrMsg = ""        ' Only Used for Insert & Update Errors

	'UseMode(c_MODE_HYBRID)  ' **** Set Mode Here - c_MODE_ACCESS, c_MODE_HYBRID, c_MODE_SQL_SERVER

	UpdateActivity          ' **** Used for logging.

	Function GetConnectionStr(b_SQLServer)
		Dim str_Ret

		If (b_SQLServer) Then

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = gstr_DB_SQL_Home
				Case c_ENV_Prep
					str_Ret = gstr_DB_SQL_Prep
				Case c_ENV_Prod
					str_Ret = gstr_DB_SQL_Prod
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = gstr_DB_SQL_Dev
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: SQL Server</div>"
					str_DB = "Mode: SQL Server - " & gstr_TestDate
			End Select

		Else

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = gstr_DB_Access_Home
				Case c_ENV_Prep
					str_Ret = gstr_DB_Access_Prep
				Case c_ENV_Prod
					str_Ret = gstr_DB_Access_Prod
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = gstr_DB_Access_Dev
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: MsAccess</div>"
					str_DB = "Mode: MsAccess - " & gstr_TestDate
			End Select

		End If
		MyDebug(str_Ret)
		GetConnectionStr = str_Ret
	End Function

	Function GetConnectionStrDefault(b_SQLServer)
		Dim str_Ret
		Dim b_SQL
		b_SQL = b_SQLServer

		If gi_Mode = c_MODE_ACCESS Then
			b_SQLServer = True
		End If

		If (b_SQLServer) Then

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = "Provider=SQLOLEDB; Data Source=DEV-004;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS_dev"
				Case c_ENV_Prep
					str_Ret = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS-dev"
				Case c_ENV_Prod
					str_Ret = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS_Dev; Password=QWSDev;Initial Catalog=QWS_Dev"
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = "Provider=SQLOLEDB; Data Source=tackleberry\SQLEXPRESS;User Id=QWS-Test; Password=Test123;Initial Catalog=qws_dev"
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: SQL Server</div>"
					str_DB = "Mode: SQL Server - " & gstr_TestDate
			End Select

		Else

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\quest.mdb;"
				Case c_ENV_Prep
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\prod\quest.mdb;"
				Case c_ENV_Prod
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=F:\database\quest.mdb;"
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\quest.mdb;"
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: MsAccess</div>"
					str_DB = "Mode: MsAccess - " & gstr_TestDate
			End Select

		End If

		GetConnectionStr = str_Ret
	End Function

	Function GetConnectionStrSecondary(b_SQLServer)
		Dim str_Ret

		If (b_SQLServer) Then

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = "Provider=SQLOLEDB; Data Source=DEV-004;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS_dev_secondary"
				Case c_ENV_Prep
					str_Ret = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS_dev_secondary"
				Case c_ENV_Prod
					str_Ret = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS_Dev; Password=QWSDev;Initial Catalog=QWS_prod_secondary"
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = "Provider=SQLOLEDB; Data Source=tackleberry\SQLEXPRESS;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS_dev_secondary"
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: SQL Server</div>"
					str_DB = "Mode: SQL Server - " & gstr_TestDate
			End Select

		Else

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\InventoryReports.mdb;"
				Case c_ENV_Prep
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\prod\InventoryReports.mdb;"
				Case c_ENV_Prod
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=F:\database\InventoryReports.mdb;"
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\InventoryReports.mdb;"
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: MsAccess</div>"
					str_DB = "Mode: MsAccess - " & gstr_TestDate
			End Select

		End If

		GetConnectionStrSecondary = str_Ret
	End Function

	Function GetConnectionStrQC(b_SQLServer)
		Dim str_Ret

		If (b_SQLServer) Then

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = "Provider=SQLOLEDB; Data Source=DEV-004;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS_dev"
				Case c_ENV_Prep
					str_Ret = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS-dev"
				Case c_ENV_Prod
					str_Ret = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS_Dev; Password=QWSDev;Initial Catalog=QWS_Dev"
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = "Provider=SQLOLEDB; Data Source=tackleberry\SQLEXPRESS;User Id=QWS-Test; Password=Test123;Initial Catalog=qws_dev"
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: SQL Server</div>"
					str_DB = "Mode: SQL Server - " & gstr_TestDate
			End Select

		Else

			Select Case(gi_Env)
				Case c_ENV_Home
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\QualityControlDB.mdb;"
				Case c_ENV_Prep
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\prod\QualityControlDB.mdb;"
				Case c_ENV_Prod
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=F:\database\QualityControlDB.mdb;"
				Case Else
					'c_ENV_Dev, c_ENV_Local
					str_Ret = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\QualityControlDB.mdb;"
					str_DB = "<div style='position: absolute; top: 0px; left: 600px; font-family: arial; font-size: 8px; font-weight: bold; color: black;'>Mode: MsAccess</div>"
					str_DB = "Mode: MsAccess - " & gstr_TestDate
			End Select

		End If

		GetConnectionStrQC = str_Ret
	End Function

	Sub CheckJobTable(str_Job)
		If b_SQL_Server Then

		On Error Resume Next
		
		Dim cn_DB, rs_Data
		Dim str_SQL
		
		str_SQL = Replace("select columnproperty(object_id('{0}'),'ID','IsIdentity') as IsIdentity, Max(ID) as MaxID FROM {0}", "{0}", str_Job)

		Set cn_DB = Server.CreateObject("adodb.connection")
		cn_DB.Open GetConnectionStr(true)

		Set rs_Data = Server.CreateObject("adodb.recordset")
		rs_Data.Open str_SQL, cn_DB
						
		If Not rs_Data.EOF Then
			If rs_Data("IsIdentity") = 0 Then
				str_SQL = "Alter Table [" & str_Job & "] Add Id_new Int Identity(1, 1);" & vbCrLf
				
				str_SQL = str_SQL & "Alter Table [" & str_Job & "] Drop Column ID;" & vbCrLf
				
				str_SQL = str_SQL & Replace("Exec sp_rename '{0}.Id_new', 'ID', 'Column';", "{0}", str_Job) & vbCrLf
					
				cn_DB.Execute str_SQL
			End If
		End If
		
		rs_Data.Close
		cn_DB.Close
		
		On Error Goto 0
		
		End If
	
	End Sub

	Sub UseMode(i_Mode)
		gi_Mode = i_Mode
		Select Case(gi_Mode)
			Case c_MODE_ACCESS
				b_SQL_Server = False
			Case c_MODE_HYBRID
				b_SQL_Server = True
			Case c_MODE_SQL_SERVER
				b_SQL_Server = True 
		End Select
		isSQLServer = b_SQL_Server
	End Sub 

	Sub UseSQL(b_SQL)
		b_SQL_Server = b_SQL
		isSQLServer = b_SQL_Server
	End Sub

	Sub DebugMsg(str_Debug)
		If (gb_Debug) Then
			'On Error Resume Next
			Dim str_Page: str_Page = ""
			str_Page = Replace(Request.ServerVariables("SCRIPT_NAME"), "/", "")

			If Instr(1, gstr_OmitDebugMsg, "|" & str_Page & "|", 1) < 1 Then			
				Response.Write(str_Debug & "<br/>")
			End If
		End If
	End Sub 

	Sub DebugCode(str_Data)
		If (gb_Debug) Then
			Response.Write(Err.Description)
			Response.Write(str_Data & "<br/>" & "Date:" & Now)
			Response.End()
		End If
	End Sub

	Function IsSQLServerDefault
		Dim b_Ret
		b_Ret = False

		If b_SQL_Server Then
			If c_MODE_HYBRID Then
				b_Ret = True
			End If

			If c_MODE_SQL_SERVER Then
				b_Ret = True
			End If
		End If

		IsSQLServerDefault = b_Ret
	End Function

	Function DbClose(cn_DB)
		'On Error Resume Next
		If Not cn_DB Is Nothing Then
			If cn_DB.State And adStateOpen = adStateOpen Then
				cn_DB.Close 
				'Set cn_DB = Nothing
			End If
		End If
		'Set cn_DB = Server.CreateObject("adodb.connection")
		'On Error Goto 0
		'If Err.Number <> 0 Then Err.Clear
	End Function

	Function DbOpen(cn_DB, isSQLServer)
		'DBClose(cn_DB)
		Set cn_DB = Server.CreateObject("adodb.connection")
		cn_DB.Open GetConnectionStr(isSQLServer)
	End Function 

	Function DbOpenAccess(cn_DB, isSQLServer)
		'DBClose(cn_DB)
		Set cn_DB = Server.CreateObject("adodb.connection")
		If gi_Mode = c_MODE_ACCESS Or gi_Mode=c_MODE_HYBRID Then
			cn_DB.Open GetConnectionStr(False)
		Else
			cn_DB.Open GetConnectionStr(True)
		End If
	End Function 

	Function DbOpenSecondary(cn_DB, isSQLServer)
		'DBClose(cn_DB)
		Set cn_DB = Server.CreateObject("adodb.connection")
		cn_DB.Open GetConnectionStrSecondary(isSQLServer)
	End Function 

	Function DbOpenQC(cn_DB, isSQLServer)
		'DBClose(cn_DB)
		Set cn_DB = Server.CreateObject("adodb.connection")
		cn_DB.Open GetConnectionStrQC(isSQLServer)
	End Function 

	Function DBOpenRS(cn_DB, str_SQL, i_CursorType, i_LockType)
		Dim rs_Data
		Set rs_Data = Server.CreateObject("adodb.recordset")
		rs_Data.Cursortype = i_CursorType
		rs_Data.Locktype = i_LockType
		rs_Data.Open str_SQL, cn_DB

		Set DBOpenRS = rs_Data
	End Function

	Function DbOpenEnv(cn_DB, isSQLServer)
		'DBClose(cn_DB)
		Set cn_DB = Server.CreateObject("adodb.connection")
		cn_DB.Open GetConnectionStr(isSQLServer)
	End Function 

	Function RsClose(rs_Data)
		If Not rs_Data Is Nothing Then
			'DebugMsg("RsClose: Not Nothing" + "<br />")
			If rs_Data.State And adStateOpen = adStateOpen Then
				rs_Data.Close
			Else
				rs_Data.Close
			End If
		Else
			'DebugMsg("RsClose: Nothing" + "<br />")
		End If
	End Function

	Function DropTable(cn_DB, str_Table)	
		'cn_DB.Execute Replace("IF EXISTS(select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '[0]') DROP TABLE [[0]]", "[0]", str_Table)
	End Function

	Function DropTableV2(cn_DB, str_Table)	
		cn_DB.Execute Replace("IF EXISTS(select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '[0]') DROP TABLE [[0]]", "[0]", str_Table)
	End Function

	Function DbCloseAll
		On Error Resume Next

		rs.close
		set rs=nothing
		rs2.close
		set rs2=nothing
		rs3.close
		set rs3=nothing
		rs4.close
		set rs4=nothing
		rs5.close
		set rs5=nothing
		rs6.close
		set rs6=nothing
		rs7.close
		set rs7=nothing
		rs8.close
		set rs8=nothing
		rs9.close
		set rs9=nothing
		

		DBConnection.close
		set DBConnection=nothing

		Err.Clear 

		On Error Goto 0

	End Function

	Function DbCloseAllAndSecondary
		On Error Resume Next

		DbCloseAll

		DBConnection2.close
		set DBConnection2 =nothing

		Err.Clear 

		On Error Goto 0

	End Function

	Function DbCloseAllOnly
		On Error Resume Next

		rs.close
		set rs=nothing
		rs2.close
		set rs2=nothing
		rs3.close
		set rs3=nothing
		rs4.close
		set rs4=nothing

		DBConnection.close

		Err.Clear 

		On Error Goto 0

	End Function

	Sub StoreID1(isSQLServer, str_ID)
		If gi_Mode = c_MODE_HYBRID And isSQLServer = False Then
			gstr_ID1 = str_ID
		End If
	End Sub

	Sub StoreID2(isSQLServer, str_ID)
		If gi_Mode = c_MODE_HYBRID And isSQLServer = False Then
			gstr_ID2 = str_ID
		End If
	End Sub

	Sub StoreID(isSQLServer, str_ID, i_ID)
		If gi_Mode = c_MODE_HYBRID And isSQLServer = False Then
			Select Case(i_ID)
				Case 1
					gstr_ID1 = str_ID
				Case 2
					gstr_ID2 = str_ID
				Case 3
					gstr_ID3 = str_ID
				Case 4
					gstr_ID4 = str_ID
				Case 5
					gstr_ID5 = str_ID
				Case 6
					gstr_ID6 = str_ID
			End Select
		End If
	End Sub

	Function GetID(isSQLServer, i_ID)
		Dim str_Ret
		str_Ret = ""

		If gi_Mode = c_MODE_HYBRID And isSQLServer Then
			Select Case(i_ID)
				Case 1
					str_Ret = gstr_ID1
				Case 2
					str_Ret = gstr_ID2
				Case 3
					str_Ret = gstr_ID3
				Case 4
					str_Ret = gstr_ID4
				Case 5
					str_Ret = gstr_ID5
				Case 6
					str_Ret = gstr_ID6
			End Select
		End If

		GetID = str_Ret
	End Function

	Function GetIDByTable(isSQLServer, i_ID, str_Table)
		Dim str_Ret
		str_Ret = ""

		If gi_Mode = c_MODE_HYBRID And isSQLServer Then
			Select Case(i_ID)
				Case 1
					str_Ret = gstr_ID1
				Case 2
					str_Ret = gstr_ID2
				Case 3
					str_Ret = gstr_ID3
				Case 4
					str_Ret = gstr_ID4
				Case 5
					str_Ret = gstr_ID5
				Case 6
					str_Ret = gstr_ID6
			End Select
		End If

		GetIDByTable = str_Ret
	End Function

	Function TableExists(str_Table, str_MainLink, b_Exit)
		Dim b_Ret
		b_Ret = False
		Set rs_Data = Server.CreateObject("adodb.recordset")
		rs_Data = DBConnection.Execute("SELECT COUNT(*) as TableExists from INFORMATION_SCHEMA.TABLES WHERE Table_Name LIKE '" + str_Table + "'")

		If rs_Data(0) > 0 Then
			b_Ret = True
		End If

		If str_MainLink <> "" And b_Ret = False Then
			str_MainLink = Replace(str_MainLink, "{0}", str_Table)
			Response.Write(str_MainLink)
		End If

		If b_Exit And b_Ret = False Then
			Response.End()
		End If

		TableExists = b_Ret
	End Function

	Function GetAccessID(str_JobName, str_Tag, str_Floor)
		On Error Resume Next
		Dim str_ID: str_ID = ""
		Dim cn_DB, rs_Data

		Set cn_DB = Server.CreateObject("adodb.connection")
		cn_DB.Open GetConnectionStr(False)

		Set rs_Data = Server.CreateObject("adodb.recordset")
		rs_Data.Open "SELECT ID FROM [" & str_JobName & "] WHERE Tag='" & str_Tag & "' AND Floor='" & str_Floor & "'", cn_DB

		If Not rs_Data.EOF Then
			str_ID = rs_Data("ID")
		End If
		rs_Data.close
		cn_DB.Close
		If Err.Number <> 0 Then
			Response.Write(Err.Description)
		End If
		GetAccessID = str_ID
		On Error Goto 0
	End Function

	Sub UpdateActivity
		On Error Resume Next
		Dim cn_DB, rs_Data
		Dim str_Page: str_Page = ""
		Dim str_QueryString: str_QueryString = ""
		Dim b_Process: b_Process = True

		If Request("REQUEST_FROM_TESTAPP") <> "TRUE" Then

		str_Page = Replace(Request.ServerVariables("SCRIPT_NAME"), "/", "")
		str_QueryString = Request.QueryString

			If Instr(1, gstr_OmitActivity, "|" & str_Page & "|", 1) < 1 Then
	
					If str_QueryString  <> "" Then
						Set cn_DB = Server.CreateObject("adodb.connection")
						cn_DB.Open GetConnectionStr(true)
				
						Set rs_Data = Server.CreateObject("adodb.recordset")
						rs_Data.Cursortype = 2
						rs_Data.Locktype = 3
						rs_Data.Open "SELECT * FROM [_qws_ActivityManagerUser_Test] WHERE ID=-1", cn_DB
						rs_Data.Addnew
						rs_Data("User") = session("teUserName")
						rs_Data("Action") = ""
						rs_Data("Page") = str_Page
						rs_Data("QueryString") = str_QueryString
						rs_Data("IP") = Request.ServerVariables("REMOTE_ADDR")
						rs_Data("SessionID") = Session.SessionID
						rs_Data("Referer") = Left(Request.ServerVariables("HTTP_REFERER"), 50)
						rs_Data.Update
						cn_DB.Close
					End If
			End If
		End If
				'On Error Goto 0
	End Sub	

	Sub UpdateActivityAction(str_Action)
		On Error Resume Next
		
		Dim cn_DB, rs_Data
		Dim str_Page: str_Page = ""
		Dim str_QueryString: str_QueryString = ""
		Dim b_Process: b_Process = True

		str_Page = Replace(Request.ServerVariables("SCRIPT_NAME"), "/", "")
		str_QueryString = Request.QueryString

		If Instr(1, gstr_OmitActivity, "|" & str_Page & "|", 1) < 1 Then
			If str_QueryString  <> "" Then
				Set cn_DB = Server.CreateObject("adodb.connection")
				cn_DB.Open GetConnectionStr(true)

				Set rs_Data = Server.CreateObject("adodb.recordset")
				rs_Data.Cursortype = 2
				rs_Data.Locktype = 3
				rs_Data.Open "SELECT * FROM [_qws_ActivityUser] WHERE ID=-1", cn_DB
				rs_Data.Addnew
				rs_Data("User") = session("teUserName")
				rs_Data("Action") = str_Action
				rs_Data("Page") = str_Page
				rs_Data("QueryString") = str_QueryString
				rs_Data("IP") = Request.ServerVariables("REMOTE_ADDR")
				rs_Data("SessionID") = Session.SessionID
				rs_Data.Update
				cn_DB.Close
			End If
		End If

		On Error Goto 0
	End Sub

	Function FormatDateToSQLStr(str_Date, str_FormatIn)
		Dim a_Parts
		Dim a_Months
		Dim str_Ret

		a_Months = Split(",JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC",",")

		If Instr(str_Date, "-") > 0 Then
			a_Parts = Split(str_Date, "-")
		Else
			a_Parts = Split(str_Date, "/")
		End If

		Select Case(str_FormatIn)
			Case "M/D/YYYY"
				str_Ret = a_Parts(1) & "-" & a_Months(CLng(a_Parts(0))) & "-" & a_Parts(2)
			Case "D/M/YYYY"
				str_Ret = a_Parts(0) & "-" & a_Months(CLng(a_Parts(1))) & "-" & a_Parts(2)
			Case "YYYY/M/D"
				str_Ret = a_Parts(2) & "-" & a_Months(CLng(a_Parts(1))) & "-" & a_Parts(0)
			Case "YYYY-MM-DD"
				str_Ret = a_Parts(2) & "-" & a_Months(CLng(a_Parts(1))) & "-" & a_Parts(0)
		End Select

		FormatDateToSQLStr = str_Ret
	End Function

	Function FormatDateToSQL(str_Date, str_FormatIn)
		Dim str_Ret

		str_Ret = FormatDateToSQLStr(str_Date, str_FormatIn)

		FormatDateToSQL = " Convert(" & gstr_SQLDateType & ", '" + str_Ret + "') "
	End Function

	Function FormatDateToSQLCheck(str_Date, str_FormatIn, isSQLServer, str_Quotes)
		Dim str_Ret
		str_Ret = str_Date

		If isSQLServer Then
			str_Ret = FormatDateToSQL(str_Date, str_FormatIn)
		End If

		If Not isSQLServer Then
			str_Ret = str_Quotes & str_Ret & str_Quotes
		End If

		FormatDateToSQLCheck = str_Ret
	End Function

	Function FixSQLCheckOld(str_SQL, isSQLServer)
		Dim str_Ret
		str_Ret = str_SQL
		If isSQLServer Then
			str_Ret = FixSQL(str_SQL)
		End If
		FixSQLCheckOld = str_Ret
	End Function

	Function FixSQLCheck(str_SQL, isSQLServer)
		Dim str_Ret
		str_Ret = str_SQL
		If isSQLServer Then

			a_Parts = Split(str_SQL, "#")

			For i = 0 To UBound(a_Parts) - 1
				a_Date = Split(a_Parts(i),"/")
				If UBound(a_Date) = 2 Then 'And IsDate(a_Parts(i)) Then
					a_Parts(i) = FormatDateToSQL(Join(a_Date,"/"), gstr_DateFormat)
				Else
					a_Date = Split(a_Parts(i),"-")
					If UBound(a_Date) = 2 Then 'And IsDate(a_Parts(i)) Then
						a_Parts(i) = FormatDateToSQL(Join(a_Date,"/"), "YYYY/M/D")
					Else
						If (UCase(Right(a_Parts(i), 3)) = " PM" Or UCase(Right(a_Parts(i), 3)) = " AM") And Instr(1, a_Parts(i), ":") > 0 Then
							a_Parts(i) = "'" & a_Parts(i) & "'"
						End If
					End If
				End If
			Next

			str_Ret = Join(a_Parts, " ")

			If Instr(1, UCase(str_Ret), " FALSE ", 1) > 0 Then
				str_Ret = Replace(UCase(str_Ret), " FALSE ", " 0 ")
			End If

			If Instr(1, UCase(str_Ret), " TRUE ", 1) > 0 Then
				str_Ret = Replace(UCase(str_Ret), " TRUE ", " 1 ")
			End If

			If Instr(1, UCase(str_Ret), "DELETE * ", 1) > 0 Then
				str_Ret = Replace(UCase(str_Ret), "DELETE * ", "DELETE ")
			End If

			str_Ret = FixSQLIsNull(str_Ret)

		End If
		FixSQLCheck = str_Ret
	End Function

	Function FixSQL(str_SQL)
		Dim i
		Dim a_Parts
		Dim a_Date
		Dim str_Ret

		str_Ret = str_SQL

		If (b_SQL_Server) Then
			a_Parts = Split(str_SQL, "#")

			For i = 0 To UBound(a_Parts) - 1
				a_Date = Split(a_Parts(i),"/")
				If UBound(a_Date) = 2 Then 'And IsDate(a_Parts(i)) Then
					a_Parts(i) = FormatDateToSQL(Join(a_Date,"/"), gstr_DateFormat)
				Else
					a_Date = Split(a_Parts(i),"-")
					If UBound(a_Date) = 2 Then 'And IsDate(a_Parts(i)) Then
						a_Parts(i) = FormatDateToSQL(Join(a_Date,"/"), "YYYY/M/D")
					Else
						If (UCase(Right(a_Parts(i), 3)) = " PM" Or UCase(Right(a_Parts(i), 3)) = " AM") And Instr(1, a_Parts(i), ":") > 0 Then
							a_Parts(i) = "'" & a_Parts(i) & "'"
						End If
					End If
				End If
			Next

			str_Ret = Join(a_Parts, " ")

			If Instr(1, UCase(str_Ret), " FALSE ", 1) > 0 Then
				str_Ret = Replace(UCase(str_Ret), " FALSE ", " 0 ")
			End If

			If Instr(1, UCase(str_Ret), " TRUE ", 1) > 0 Then
				str_Ret = Replace(UCase(str_Ret), " TRUE ", " 1 ")
			End If

			If Instr(1, UCase(str_Ret), "DELETE * ", 1) > 0 Then
				str_Ret = Replace(UCase(str_Ret), "DELETE * ", "DELETE ")
			End If

			str_Ret = FixSQLIsNull(str_Ret)

		End If

		FixSQL = str_Ret
	End Function

	Function FixSQLIsNull(str_SQL)
		Dim str_Ret
		Dim str_Page
		Dim str_Search, str_Replace

		str_Ret = str_SQL

		str_Page = Replace(Request.ServerVariables("SCRIPT_NAME"), "/", "")

		' ***********  Used in GlassReceivedSelect.asp  ***********
		str_Search = " ISNULL(SHIPDATE) = TRUE "
		str_Replace = " ISNULL(SHIPDATE,'') = '' "
		If Instr(1, "|glassreceivedselect.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in GlassReceivedSelect.asp  ***********
		str_Search = " ISNULL(SHIPDATE) "
		str_Replace = " ISNULL(SHIPDATE,'') = '' "
		If Instr(1, "|glassreceivedexpected.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in GlassOptimaUndo.asp  ***********
		str_Search = " ISNULL(COMPLETEDDATE) "
		str_Replace = " ISNULL(COMPLETEDDATE,'') = '' "
		If Instr(1, "|glassoptimaundo.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in OptimizationLogGlass.asp  ***********
		str_Search = " ISNULL(SHIFT) "
		str_Replace = " ISNULL(SHIFT,'') = '' "
		If Instr(1, "|optimizationlogglass.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in QCInventoryReportActive.asp  ***********
		str_Search = " ISNULL(G.CONSUMEDATE) "
		str_Replace = " ISNULL(G.CONSUMEDATE,'') = '' "
		If Instr(1, "|qcinventoryreportactive.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in QCInventoryReportActive.asp  ***********
		str_Search = " ISNULL(SP.CONSUMEDATE) "
		str_Replace = " ISNULL(SP.CONSUMEDATE,'') = '' "
		If Instr(1, "|qcinventoryreportactive.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in QCInventoryReportActive.asp  ***********
		str_Search = " ISNULL(SE.CONSUMEDATE) "
		str_Replace = " ISNULL(SE.CONSUMEDATE,'') = '' "
		If Instr(1, "|qcinventoryreportactive.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in QCInventoryReportActive.asp  ***********
		str_Search = " ISNULL(M.CONSUMEDATE) "
		str_Replace = " ISNULL(M.CONSUMEDATE,'') = '' "
		If Instr(1, "|qcinventoryreportactive.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		' ***********  Used in GlassReportShipped.asp  ***********
		str_Search = "SHIPDATE <> NULL"
		str_Replace = " ISNULL(SHIPDATE,'') <> '' "
		If Instr(1, "|glassreportshipped.asp|", "|" & str_Page & "|", 1) > 0 AND Instr(1, str_SQL, str_Search, 1) > 0 Then
			str_Ret = Replace(UCase(str_Ret), UCase(str_Search), str_Replace)
		End If

		FixSQLIsNull = str_Ret
	End Function

' Cursor Types
' Const adOpenForwardOnly = 0
' Const adOpenKeyset = 1
' Const adOpenDynamic = 2
' Const adOpenStatic = 3

' Lock Types
' Const adLockReadOnly = 1
' Const adLockPessimistic = 2
' Const adLockOptimistic = 3
' Const adLockBatchOptimistic = 4

'Add - adOpenKeyset, adLockOptimistic
'Query - adOpenForwardOnly, adLockReadOnly


	Function GetDBCursorType
		GetDBCursorType = 0 				' adOpenForwardOnly		'New: 0, Previous: 2
	End Function

	Function GetDBLockType
		GetDBLockType = 1 					' adLockReadOnly		  'New: 1, Previous: 3
	End Function

	Function GetDBCursorTypeInsert
		GetDBCursorTypeInsert = 1			' adOpenKeyset
	End Function

	Function GetDBLockTypeInsert
		GetDBLockTypeInsert = 3  					' adLockOptimistic
	End Function

	Sub SetTestDate(str_Day, str_Month, str_Year)
		Dim a_Parts
		On Error Resume Next

		If gi_Env = c_ENV_Home Or gi_Env = c_ENV_Local Then
			a_Parts = Split(gstr_TestDate,"/")
			str_Day = a_Parts(0)
			str_Month = a_Parts(1)
			str_Year = a_Parts(2)
		End If

		On Error Goto 0
	End Sub

	Function FixSQLBool(str_Bool, isSQLServer)
		Dim str_Ret
		str_Ret = str_Bool

		If isSQLServer Then
			If UCase(str_Bool) = "TRUE" Then 
				str_Ret = 1
			Else
				str_Ret = 0
			End If
		End If
		FixSQLBool = str_Ret
	End Function

	Function FixSQLDate(str_Date, isSQLServer)
		Dim str_Ret
		str_Ret = str_Date

		If isSQLServer Then
			If Instr(str_Date, "-") > 0 Then
				str_Ret = FormatDateToSQLStr(str_Date, gstr_DateFormat2)
			Else
				str_Ret = FormatDateToSQLStr(str_Date, gstr_DateFormat)
			End If
		Else
			str_Ret = "" & str_Ret & ""
		End If

		FixSQLDate = str_Ret
	End Function

	Function GetScriptName
		GetScriptName = Replace(Request.ServerVariables("SCRIPT_NAME"),"/","")
	End Function

	Function GetScriptNameOnly(str_URL)
		Dim str_Ret
		On Error Resume Next
		Dim a_Parts: a_Parts = Split(str_URL, "/")

		str_Ret = a_Parts(UBound(a_Parts))

		GetScriptNameOnly = str_Ret
		On Error Goto 0
	End Function

	Function GetScriptNameOnlyV2(str_URL)
		Dim str_Ret
		On Error Resume Next
		Dim a_Parts: a_Parts = Split(str_URL, "/")

		str_Ret = a_Parts(UBound(a_Parts))

		a_Parts = Split(str_Ret & "?", "?")
		str_Ret = a_Parts(0)

		GetScriptNameOnlyV2 = str_Ret
		On Error Goto 0
	End Function

	Function MyDebug(str_Msg)

		If Request.ServerVariables("REMOTE_ADDR") = gstr_DevIP1 Or Request.ServerVariables("REMOTE_ADDR") = "192.168.1.162" Then
			Dim str_Page
			str_Page = LCase(GetScriptName)
			If Instr(1, "|index.asp|", "|" & str_Page & "|", 1) < 1 Then
				If gbMyDebug Then Response.Write(str_Msg)
			End If
		End If

		MyDebug = ""
	End Function

	Function IsDebug()
		Dim b_Ret: b_Ret = False
		If Request.ServerVariables("REMOTE_ADDR") = gstr_DevIP1 Then
			b_Ret = True
		End If
		IsDebug = b_Ret
	End Function

	Function GetDisconnectedRS(strSQL, cn_Conn)
	'this function returns a disconnected RS
	
	'---- CursorLocationEnum Values ----
		Const adUseServer = 2
		Const adUseClient = 3
	
	'Set some constants
	'---- CursorTypeEnum Values ----
		Const adOpenForwardOnly = 0
		Const adOpenKeyset = 1
		Const adOpenDynamic = 2
		Const adOpenStatic = 3
	
	'---- LockTypeEnum Values ----
		Const adLockReadOnly = 1
		Const adLockPessimistic = 2
		Const adLockOptimistic = 3
		Const adLockBatchOptimistic = 4
	
		'Declare our variables
		Dim oRS
	
		'Create the Recordset object
		Set oRS = Server.CreateObject("ADODB.Recordset")
	
		oRS.CursorLocation = adUseClient
		oRS.Open strSQL, cn_Conn, adOpenStatic, adLockBatchOptimistic
	
		'Disconnect the Recordset
		Set oRS.ActiveConnection = Nothing
	
		'Return the Recordset
		Set GetDisconnectedRS = oRS
	
		'Clean up...
		'oRS.Close
		Set oRS = Nothing
	End Function

	Function DebugLog(str_Ref, str_Msg, i_Tab)
		Dim o_FSO
		Set o_FSO = Server.CreateObject("Scripting.FileSystemObject") 
		Const fsoForReading = 1
		Const fsoForWriting = 2
		Const fsoForAppending = 8

		Dim o_TextStream

		Set o_TextStream = o_FSO.OpenTextFile(Server.MapPath("_Logs\Debug.log"), fsoForAppending, true)

		o_TextStream.WriteLine(String(i_Tab, vbTab) & str_Ref & " - " & Now & ": " & str_Msg)
		o_TextStream.Close
		Set o_TextStream = Nothing
		Set o_FSO = Nothing
	End Function

	Function ReadFile(str_File)
		Dim str_Ret
		Dim o_FSO
		Set o_FSO = Server.CreateObject("Scripting.FileSystemObject") 
		Const fsoForReading = 1 

		If o_FSO.FileExists(str_File) Then
			Dim o_TextStream
			Set o_TextStream = o_FSO.OpenTextFile(str_File, fsoForReading)
			str_Ret = o_TextStream.ReadAll
			o_TextStream.Close
			Set o_TextStream = Nothing
		End If

		Set o_FSO = Nothing
		ReadFile = str_Ret
	End Function

	Function TestClass()
	Set objAA = new cls_Window
	End Function

	Sub SendEmail(str_To, str_Subject, str_Body)

		Set o_Mail = CreateObject("CDO.Message")
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="smtp.gmail.com"
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 300
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="quest.windows@gmail.com" 'You can also use you email address that's setup through google apps.
		o_Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="gNpXe9fewe9N"
		o_Mail.Configuration.Fields.Update

		o_Mail.Subject = str_Subject
		o_Mail.From = "quest.windows@gmail.com" 
		'Mail.To="jodycash@gmail.com, shaunl@questwindows.com, lev@questwindows.com, ragnew@questwindows.com, shill@questwindows.com, mcash@questwindows.com, ken@questwindows.com, shipping@questwindows.com, aahmed@questwindows.com, vdavid@questwindows.com, valdi@questwindows.com"
		o_Mail.To = str_To
		'Mail.To="michael@questwindows.com" ' for testing
		'Mail.Bcc="someoneelse@somedomain.com" 'Carbon Copy
		'Mail.Cc="someoneelse2@somedomain.com" 'Blind Carbon Copy

		'**Below are different options for the Body of an email. *Only one of the below body types can be sent.
		'sMail.TextBody="Report"

		o_Mail.HTMLBody = str_Body
		'Mail.CreateMHTMLBody "http://172.18.13.31:8081/BARCODERTVEmail.asp" 'Sends an email which has a body of a specified webpage
		'Mail.CreateMHTMLBody "file://c:/mydocuments/email.htm" 'Sends an email which has a body of an html file that's stored on your computer. This MUST be on the server that this script is being served from.
		' How to add an attachment
		'Mail.AddAttachment "c:\mydocuments\test.txt" 'Again this must be on the server that is serving this script.
		o_Mail.Send
		Set o_Mail = Nothing

	End Sub

	Function GetColourQuest(cn_DB, str_Colour)
		Dim str_Ret

		Dim rs_Data
		Set rs_Data = cn_DB.Execute("SELECT Code FROM [qws_Prod].[dbo].y_Color WHERE Project='" & str_Colour & "'")
		If Not rs_Data.EOF Then
			str_Ret = rs_Data(0)
		End If

		rs_Data.Close: Set rs_Data = Nothing

		GetColourQuest = str_Ret
	End Function

	Function GetColourPrefOld(cn_DB, str_Part, str_Colour, str_ColourCode)
		Dim str_Ret

		Dim rs_Data
		Set rs_Data = cn_DB.Execute("SELECT ConfigurationCode FROM ColorConfigurations WHERE ColorName='" & str_Colour & "'")
		If Not rs_Data.EOF Then
			str_Ret = rs_Data(0)
		Else
			Set rs_Data = cn_DB.Execute("SELECT ConfigurationCode FROM ColorConfigurations WHERE ColorName='" & str_ColourCode & "'")
			If Not rs_Data.EOF Then
				str_Ret = rs_Data(0)
			End If
		End If

		rs_Data.Close: Set rs_Data = Nothing

		GetColourPrefOld = str_Ret
	End Function

	Function GetColourPref(cn_DB, str_Part, str_Colour, str_ColourCode)
		Dim b_Found: b_Found = False
		Dim str_Ret
		Dim rs_Data
		Set rs_Data = cn_DB.Execute("SELECT ConfigurationCode FROM [Quest].[dbo].ColorConfigurations WHERE ColorName='" & str_Colour & "'")
		If Not rs_Data.EOF Then
			Dim str_TmpColour
			str_TmpColour = rs_Data(0)
			Set rs_Data = cn_DB.Execute("SELECT COUNT(*) FROM [Quest].[dbo].Materiales WHERE Referencia='" & str_Part & " " & str_Colour & "'")
			If rs_Data(0) > 0 Then
				str_Ret = str_TmpColour
				b_Found = True
			End If
		End If

		If b_Found = False Then
			Set rs_Data = cn_DB.Execute("SELECT ConfigurationCode FROM [Quest].[dbo].ColorConfigurations WHERE ColorName='" & str_ColourCode & "'")
			If Not rs_Data.EOF Then
				str_Ret = rs_Data(0)
			End If
		End If

		rs_Data.Close: Set rs_Data = Nothing

		GetColourPref = str_Ret
	End Function

	Function GetPartPref(b_Debug, cn_DB, str_Part, str_Colour, str_ColourCode)
		Dim str_Ret
		Dim str_Step

		Dim rs_Data
		Set rs_Data = cn_DB.Execute("SELECT Referencia FROM Materiales WHERE ReferenciaBase='" & str_Part & "' AND Color='" & str_Colour & "'")
		If Not rs_Data.EOF Then
			str_Ret = str_Part & " " &  str_Colour
			str_Step = "1 - SELECT Referencia FROM Materiales WHERE ReferenciaBase='" & str_Part & "' AND Color='" & str_Colour & "'"
		Else
			Set rs_Data = cn_DB.Execute("SELECT Referencia FROM Materiales WHERE ReferenciaBase='" & str_Part & "' AND Color='" & str_ColourCode & "'")
			If Not rs_Data.EOF Then
				str_Ret = str_Part & " " &  str_ColourCode
				str_Step = "2 - SELECT Referencia FROM Materiales WHERE ReferenciaBase='" & str_Part & "' AND Color='" & str_ColourCode & "'"
			Else
				If str_Part <> "" And str_Colour <> "" Then
					Set rs_Data = cn_DB.Execute("SELECT Referencia FROM Materiales WHERE Referencia='" & str_Part & " " & str_Colour & "'")
					If Not rs_Data.EOF Then
						str_Ret = str_Part & " " &  str_Colour
						str_Step = "3 - SELECT Referencia FROM Materiales WHERE Referencia='" & str_Part & " " & str_Colour & "'"
					Else
						Set rs_Data = cn_DB.Execute("SELECT Referencia FROM Materiales WHERE Referencia='" & str_Part & " " & str_ColourCode & "'")
						If Not rs_Data.EOF Then
							str_Ret = str_Part & " " & str_ColourCode
							str_Step = "4 - SELECT Referencia FROM Materiales WHERE Referencia='" & str_Part & " " & str_ColourCode & "'"
						End If
					End If
				End If
			End If
		End If
		If b_Debug Then Response.Write("<br/>" & str_Step & "<br/>")
		rs_Data.Close: Set rs_Data = Nothing

		GetPartPref = str_Ret
	End Function

	Function GetNewID(cn_DB)
		Dim str_Ret

		Dim rs_Data
		Set rs_Data = cn_DB.Execute("SELECT NEWID()")
		str_Ret = rs_Data(0)

		rs_Data.Close: Set rs_Data = Nothing

		GetNewID = str_Ret
	End Function 

	Function SetControlDeStock(str_Part, cn_DB)
		cn_DB.Execute("UPDATE [Quest].dbo.Materiales SET ControlDeStock=1 WHERE ReferenciaBase='" & str_Part & "' AND ControlDeStock=0 ")
	End Function

	Function PrefAddTransfer(str_ID, str_Warehouse_Out, str_Warehouse_In, str_Part, str_Qty, str_Colour, str_LengthFt, b_Debug, str_Page, str_Note, str_PO, str_Referrer)
		' Called in StockMoveConf.asp & StockByRackEditConf.asp
		If b_Debug = False Then
			On Error Resume Next
		End If
		Dim str_Phase: str_Phase = 1
		Dim i_SuccessCode: i_SuccessCode = 0
		Dim i_DocCode: i_DocCode= 0
		Dim str_PrefPart
		Dim str_PrefColour
		Dim str_SQL
		Dim str_Step: str_Step = ""

		Dim b_Process: b_Process = True

		Dim cn_DBQuest: Set cn_DBQuest = Server.CreateObject("ADODB.Connection")
		cn_DBQuest.Open GetConnectionStr(True)

		Dim cn_DB: Set cn_DB = Server.CreateObject("ADODB.Connection")
		cn_DB.Open gstr_DB_Pref ' "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS_Dev; Password=QWSDev;Initial Catalog=QWS_Dev"
		cn_DB.BeginTrans

		Dim rs_WDoc: Set rs_WDoc = Server.CreateObject("ADODB.Recordset")
		rs_WDoc.Cursortype = GetDBCursorType
		rs_WDoc.Locktype = GetDBLockType
		rs_WDoc.Open "SELECT (ISNULL(Max(DocumentCode),0) + 1) as WarehouseDoc FROM [Quest].[dbo].WarehouseDocuments", cn_DB

		If Not rs_WDoc.EOF Then
			i_DocCode = rs_WDoc("WarehouseDoc")
		End If

		Dim i_WorkerCode: i_WorkerCode = 47
		Dim i_UsedState: i_UsedState = 1
		Dim i_WarehouseEntry, i_WarehouseExit

		If UCase(str_Colour) = "WHITE" Then str_Colour = "K1285"
		If b_Debug Then Response.Write("<br/>---------------------------------------------------<br/>")
		str_Part = GetFixedPart(str_Part)

		Select Case(Trim(UCase(str_Warehouse_Out)))
			Case "NASHUA"
				i_WarehouseExit = 1
			Case "GOREWAY"
				i_WarehouseExit = 7
			Case "HORNER"
				i_WarehouseExit = 8
			Case "WINDOW PRODUCTION"
				i_WarehouseExit = 9
			Case "NPREP"
				i_WarehouseExit = -1
			Case "NASHUA_ADJUSTMENT"
				i_WarehouseExit = -1
			Case "ADJUSTMENT"
				i_WarehouseExit = -1
			Case "DURAPAINT"
				i_WarehouseExit = 3
				If str_Warehouse_In <> "ADJUSTMENT" AND str_Warehouse_In <> "WINDOW PRODUCTION" Then b_Process = False
			Case "METRA", "CAN-ART"
				b_Process = False
			Case "DURAPAINT(WIP)","TILTON(WIP)"
				b_Process = False
				i_WarehouseExit = -1
				If UCase(str_Colour) = "MILL" Then
					b_Process = False
				End If
			Case "DEPENDABLE"
				b_Process = False
			Case Else
				i_WarehouseExit = -2
		End Select

		Select Case(UCase(str_Warehouse_In))
			Case "NASHUA"
				i_WarehouseEntry = 1
				If i_WarehouseExit = -2 Then i_WarehouseExit = -1
			Case "GOREWAY"
				i_WarehouseEntry = 7
				If i_WarehouseExit = -2 Then i_WarehouseExit = -1
			Case "HORNER"
				i_WarehouseEntry = 8
				If i_WarehouseExit = -2 Then i_WarehouseExit = -1
			Case "DURAPAINT"
				i_WarehouseEntry = 3
				If i_WarehouseExit = -2 Then i_WarehouseExit = -1
			Case "DURAPAINT(WIP)"
				i_WarehouseEntry = -1 'Use Inventory Exit
			Case "WINDOW PRODUCTION"
				i_WarehouseEntry = 9
				If i_WarehouseExit = -2 Then i_WarehouseExit = -1
			Case "SCRAP"
				If str_Warehouse_Out = "DURAPAINT" Then
					b_Process = True
				End If
				i_WarehouseEntry = -1
			Case "NPREP"
				i_WarehouseEntry = -1
			Case "TILTON(WIP)"
				i_WarehouseEntry = -1
			Case "DEPENDABLE"
				i_WarehouseEntry = -1
			Case "TILTON(WIP)"
				i_WarehouseEntry = -1
			Case "NASHUA_ADJUSTMENT"
				i_WarehouseEntry = -1
			Case "ADJUSTMENT"
				i_WarehouseEntry = -1
			Case Else
				i_WarehouseEntry = -2
		End Select

		str_GUID = GetNewID(cn_DB)

		str_ColourCode = Trim(GetColourQuest(cn_DB, str_Colour) & "")

		str_Colour = Replace(str_Colour, " ", "-")
		str_Colour = Replace(str_Colour, ".", "")
		str_Colour = Replace(str_Colour, "--", "-")
		str_Colour = Trim(str_Colour)

		str_PrefColour = GetColourPref(cn_DB, str_Part, str_Colour, str_ColourCode)
		If b_Debug Then Response.Write("GetColourPref: str_Part: " & str_Part & ", Colour: " & str_Colour & ", ColourCode: " & str_ColourCode & ", Return: " & str_PrefColour & "<br/>")

		str_Part = Trim(str_Part)

		str_PrefPart = GetPartPref(b_Debug, cn_DB, str_Part, str_Colour, str_ColourCode)
		If b_Debug Then Response.Write("GetPartPref: str_Part: " & str_Part & ", Colour: " & str_Colour & ", ColourCode: " & str_ColourCode & ", Return: " & str_PrefPart & "<br/>")

		If b_Debug Then Response.Write("PrefPart: " & str_PrefPart & "<br/>")

'CUSTOM CODE
		If b_Debug Then Response.Write("<br/>Part: |" & str_Part & "|, PrefPart: " & str_PrefPart & ", ColourCode: |" & str_ColourCode & "|, Pref Colour: " & str_PrefColour & "|, Colour: " & str_Colour & "<br/>")		

		If Trim(str_PrefPart & "") = "" AND str_ColourCode = "UC82989" Then
			str_ColourCode = str_ColourCode & "XL"
			str_PrefColour = GetColourPref(cn_DB, str_Part, str_Colour, str_ColourCode)
			
			str_PrefPart = GetPartPref(b_Debug, cn_DB, str_Part, str_Colour, str_ColourCode)
		End If

		If str_Colour = "CLR-ANDZD" Then
			str_PrefColour = GetColourPref(cn_DB, str_Part, str_Colour, "Clear Anodized")
		End If

		If (Trim(str_PrefPart & "") = "" AND UCase(str_Colour) = "CLEAR/ANOD") OR UCase(str_Colour) = "CLEAR-AND-EXT" Then
			str_PrefColour = GetColourPref(cn_DB, str_Part, str_Colour, "CLEAR AND-EXT")
			str_PrefPart = str_Part & " CLR-EXT"
		End If

		If str_PrefPart = "" Then
			b_Process = False
		End If

		If UCase(str_Part) = "QUE-50" Then
			str_PrefColour = GetColourPref(cn_DB, str_Part, str_Colour, "UCFX10053")
			str_PrefPart = "Que-50 UCFX10053"
		End If

		'If Trim(str_ColourCode) = "" Then
			'str_PrefColour = GetColourPref(cn_DB, str_Part, str_Colour, Replace(str_Colour, "-EXT", ""))
			'If b_Debug Then Response.Write("GetPrefColour II: str_Part: " & str_Part & ", Colour: " & str_Colour & ", ColourCode: " & str_ColourCode & ", Return: " & str_PrefPart & "<br/>")
		'End If

		If str_ColourCode = "" And str_Colour <> "" Then str_ColourCode = str_Colour

		If b_Debug Then Response.Write("Color:" & str_ColourCode)

		'If (UCase(str_Part) = "NC70043" Or UCase(str_Part) = "QUE-60" Or UCase(str_Part) = "QUE-64" Or UCase(str_Part) = "NC70001" Or UCase(str_Part) = "NC70007" Or UCase(str_Part) = "QUE-168" Or UCase(str_Part) = "QUE-108" Or UCase(str_Part) = "QUE-142" Or UCase(str_Part) = "QUE-147" Or UCase(str_Part) = "QUE-146" Or UCase(str_Part) = "QUE-189" Or UCase(str_Part) = "QUE-62") And str_ColourCode = "UC114626" And UCase(str_Colour) <> "SJU-INT" And UCase(str_Colour) <> "SJX-INT" Then
		'	str_PrefColour = GetColourPref(cn_DB, str_Part, str_Colour, "UC114626")
		'	str_PrefPart = str_Part & " UC11426"
		'End If

		If i_WarehouseExit = -2 Or i_WarehouseEntry = -2 Then
			b_Process = False
		End If

		If i_WarehouseEntry = i_WarehouseExit Then
			b_Process = False
		End If

		If CLng(str_Qty) = 0 Then
			b_Process = False
		End If

		str_SQL = ""
		str_SQL = str_SQL & "INSERT INTO [qws_prod].[dbo]._qws_PrefInventorySync (GUID,RecID, Warehouse_Out, Warehouse_In, Part, Qty, Colour, LengthFt, ColourCode, ColourPref, PartPref, Page, DocID, Note, PO, ReferrerPage, IP) VALUES "
		str_SQL = str_SQL & "('" & str_GUID & "'," & str_ID & ",'" & str_Warehouse_Out & "','" & str_Warehouse_In & "','" & str_Part & "'," & str_Qty & ",'" & str_Colour & "'," & str_LengthFt & ",'" & str_ColourCode & "','" & str_PrefColour & "','" & str_PrefPart & "','" & str_Page & "'," & i_DocCode & ",'" & str_Note & "','" & Replace(str_PO, "'", "") & "','" & Left(str_Referrer, 250) & "','" & Request.ServerVariables("REMOTE_ADDR") & "')"

		cn_DBQuest.Execute(str_SQL)

		If Not b_Process Then

			str_SQL = ""
			str_SQL = str_SQL & "UPDATE [qws_prod].[dbo]._qws_PrefInventorySync SET StatusPhase = 'Skipped', Status=-1 WHERE Guid='" & str_GUID & "'" 
			cn_DBQuest.Execute(str_SQL)

			cn_DB.RollbackTrans
		Else

			Dim cmd_Data
			Set cmd_Data = Server.CreateObject("ADODB.Command")
			Set cmd_Data.ActiveConnection = cn_DB
			cmd_Data.CommandText = "pa_WarehouseDocuments_AddDocument"
			cmd_Data.CommandType = &H0004

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("retVal", 3, 4)

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("nDocumentCode", 3, &H0001)    ' adParamInput=&H0001, adParamOutput=&H0004, adParamReturnValue = &H0004   adInteger
			cmd_Data("nDocumentCode") = i_DocCode

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("tDocumentDate", 133, &H0001)
			cmd_Data("tDocumentDate") = Now

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("nEntryWarehouse", 3, &H0001)
			cmd_Data("nEntryWarehouse") = i_WarehouseEntry

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("nExitWarehouse", 3, &H0001)
			cmd_Data("nExitWarehouse") = i_WarehouseExit

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("nWorkerCode", 3, &H0001)
			cmd_Data("nWorkerCode") = i_WorkerCode

			cmd_Data.Parameters.Append cmd_Data.CreateParameter("nUsedState", 3, &H0001)
			cmd_Data("nUsedState") = i_UsedState

			cmd_Data.Execute
			i_SuccessCode = cmd_Data("retVal")

			Dim d_Length

			If str_LengthFT = "21.33" Then
				d_Length = 6502.4 '6502.4, 6502
			Else
				d_Length = CDbl(str_LengthFt) * 304.8
			End If

			If b_Debug Then
				Response.Write("<br/>------------------------------------<br/>")
				Response.Write("ID: " & str_ID & "<br/>")
				Response.Write("DocCode: " & i_DocCode & "<br/>")
				Response.Write("Row: " & i_Row & "<br/>")
				Response.Write("Part: " & str_Part & "<br/>")
				Response.Write("PartPref: " & str_PrefPart & "<br/>")
				Response.Write("Colour: " & str_Colour & "<br/>")
				Response.Write("Colour Pref: " & str_PrefColour & "<br/>")
				Response.Write("Qty: " & str_Qty & "<br/>")
				Response.Write("Length(ft): " & str_LengthFt & "<br/>")
				Response.Write("Length(mm): " & d_Length & "<br/>")
				Response.Write("Warehouse Entry: " & i_WarehouseEntry & "<br/>")
				Response.Write("Warehouse Entry: " & str_Warehouse_In & "<br/>")
				Response.Write("Warehouse Exit: " & i_WarehouseExit & "<br/>")
				Response.Write("Warehouse Exit:" & str_Warehouse_Out & "<br/>")
				Response.Write("Date:" & Now & "<br/>")
				Response.Write("" & "" & "<br/>")
				Response.Write("" & "" & "<br/>")
			End If

			If b_Debug Then Response.Write("<br/>1:" & i_SuccessCode)

' Detail - S

			If  i_SuccessCode = 0 Then
				str_Phase = "2"
				Dim i_Row: i_Row = 0
				i_Row = i_Row + 1

				Call SetControlDeStock(str_Part, cn_DBQuest)

				Set cmd_Data = Server.CreateObject("ADODB.Command")
				Set cmd_Data.ActiveConnection = cn_DB
				cmd_Data.CommandText = "pa_WarehouseDocuments_AddDetailLine"
				cmd_Data.CommandType = &H0004

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("retVal", 3, 4)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("guidRowId", 72, &H0001)
				cmd_Data("guidRowId") = str_GUID

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("nDocumentCode", 3, &H0001, 0, i_DocCode)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("nID", 3, &H0001, 0, i_Row)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("strReference", 200, &H0001, 50, str_PrefPart)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("iColorConfiguration", 200, &H0001, 50, str_PrefColour)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("fQuantity", 3, &H0001, 0, str_Qty)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("fLength", 5, &H0001, 0, d_Length)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("fHeight", 3, &H0001, 0, 0)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("fPurchasePrice", 3, &H0001, 0, 0)

				cmd_Data.Parameters.Append cmd_Data.CreateParameter("strPurchasePriceCurrency", 200, &H0001, 50, "Canadian")

				cmd_Data.Execute
				i_SuccessCode = cmd_Data("retVal")
				If b_Debug Then Response.Write("<br/>2:" & i_SuccessCode)
			End If

' Detail - E

			If i_SuccessCode = 0 Then
				str_Phase = "3"
				Set cmd_Data = Server.CreateObject("ADODB.Command")
				Set cmd_Data.ActiveConnection = cn_DB
				cmd_Data.CommandText = "pa_WarehouseDocuments_UpdateStock"
				cmd_Data.CommandType = &H0004
	
				cmd_Data.Parameters.Append cmd_Data.CreateParameter("retVal", 3, 4)
	
				cmd_Data.Parameters.Append cmd_Data.CreateParameter("nDocumentCode", 3, &H0001)
				cmd_Data("nDocumentCode") = i_DocCode
		
				cmd_Data.Parameters.Append cmd_Data.CreateParameter("bAllowNegativeStock", 3, &H0001)
				If b_Debug Then
					cmd_Data("bAllowNegativeStock") = 0 '0 - Don't allow negative stock, 1 Allow negative stock
				Else
					cmd_Data("bAllowNegativeStock") = 0
				End If

				'Durapaint -> Nashua
				If i_WarehouseExit = 3 And i_WarehouseEntry = 1 Then
					cmd_Data("bAllowNegativeStock") = 1 ' Allow Negative Stock
				End If

				If str_Warehouse_Out = "WINDOW PRODUCTION" Then cmd_Data("bAllowNegativeStock") = 1

				cmd_Data.Execute
				i_SuccessCode = cmd_Data("retVal")
				If b_Debug Then Response.Write("<br/>3:" & i_SuccessCode & "<br/>")
			End If

			If i_SuccessCode = 0 Then
				cn_DB.CommitTrans
				'cn_DB.RollbackTrans
			Else
				cn_DB.RollbackTrans
			End If

			str_SQL = ""
			str_SQL = str_SQL & "UPDATE [qws_prod].[dbo]._qws_PrefInventorySync SET StatusPhase = '" & str_Phase & "', Status=" & i_SuccessCode & " WHERE Guid='" & str_GUID & "'" 
			cn_DBQuest.Execute(str_SQL)

			PrefAddTransfer = i_SuccessCode

		End If

		cn_DBQuest.Close

		cn_DB.Close
		
		On Error Goto 0
		
	End Function 

	Function GetFixedPart(str_Part)
		Dim str_Ret
		Select Case(Trim(UCase(str_Part)))
			Case "C-801"
				str_Ret = "QUE-C801"
			Case "C-802"
				str_Ret = "QUE-C802"
			Case "C-803"
				str_Ret = "QUE-C803"
			Case "C-804"
				str_Ret = "QUE-C804"
			Case "C-805"
				str_Ret = "QUE-C805"
			Case "QUE-115 F"
				str_Ret = "QUE-115F"
			Case "AH 27172"
				str_Ret = "AH-27172"
			Case "CS-75155"
				str_Ret = "CS75155"
			Case "AS-7769"
				str_Ret = "AS 7769"
			Case Else
				str_Ret = str_Part
		End Select
		GetFixedPart = str_Ret
	End Function

	Function FixFloat(o_Val)
		If o_Val = "" Then o_Val = 0
		FixFloat = o_Val
	End Function



class cls_Window
	Dim X
	Dim Y
end class

%>
