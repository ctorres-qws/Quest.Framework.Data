<%

	gstr_DevIP1 = "192.168.0.46"

	c_ENV_Home = 1
	c_ENV_Local = 2
	c_ENV_Dev = 3
	c_ENV_Prep = 4
	c_ENV_Prod = 5

	gi_Env = c_ENV_Prod

	b_SQL_Server = false
	isSQLServer = b_SQL_Server

	gstr_FolderUploadRecords = "C:\_Websites\Prod\QWS_Tools\UploadRecords\"

	' DB: SQL Server
	gstr_DB_SQL_Home = "Provider=SQLOLEDB; Data Source=192.168.1.167;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS_dev"
	gstr_DB_SQL_Prep = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Test; Password=Test123;Initial Catalog=QWS-dev"
	'gstr_DB_SQL_Prod = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Prod; Password=Test123;Initial Catalog=QWS_prod"
	gstr_DB_SQL_Prod = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS_Dev; Password=QWSDev;Initial Catalog=QWS_Dev"
	gstr_DB_SQL_Dev = "Provider=SQLOLEDB; Data Source=tackleberry\SQLEXPRESS;User Id=QWS-Test; Password=Test123;Initial Catalog=qws_dev"

	gstr_DB_Pref = "Provider=SQLOLEDB; Data Source=qwtordb1\quest;User Id=QWS-Prod; Password=Test123;Initial Catalog=Quest"
	' DB: ACCESS
	gstr_DB_Access_Home = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\quest-2017-09-16.mdb;"
	gstr_DB_Access_Prep = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\prod\quest.mdb;"
	gstr_DB_Access_Prod = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=F:\database\quest.mdb;"
	gstr_DB_Access_Dev = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=c:\_databases\dev\quest-2017-09-18.mdb;"

	gstr_DB_Access_Admin = "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=\\172.18.13.31\quest\database\teadmin.mdb;"

	c_MODE_ACCESS = 1
	c_MODE_HYBRID = 2
	c_MODE_SQL_SERVER = 3

	gb_Debug = false
	gb_DebugResumeNext = true
	str_DB = ""

	gstr_DebugTopRecs = " TOP 2000 "
	If gi_Env = c_ENV_Home Or gi_Env = c_ENV_Local Or gi_Env = c_ENV_Dev Then
		gb_Debug = true
	End If

	gi_Mode = c_MODE_HYBRID ' Set Default

	gstr_IIS_Log_Folder_Order = "C:\WINDOWS\system32\LogFiles\W3SVC1311942458\"
	gstr_IIS_Log_Folder_Order_Archive = "C:\_Websites\_Logs\Order_458\"
	gstr_IIS_Log_Folder_Tools = "C:\WINDOWS\system32\LogFiles\W3SVC2033936928\"
	gstr_IIS_Log_Folder_Tools_Archive = "C:\_Websites\_Logs\Tools_928\"

	gstr_IIS_Log_Prefix = ""

%>