<%
'Create  DSN Less connection to Access Database
'Create DBConnection Object
Set DBConnection = Server.CreateObject("adodb.connection")
DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; "
DSN = DSN & "DBQ=" & Server.Mappath("database2/quest.mdb")
DSN = DSN & ";PWD=stewart"
DBConnection.Open DSN




		'Dim Objcmd As SqlCommand = Nothing
        'Dim strSQL As String
		'strSQL = "Select M.ReferenciaBase FROM STOCK AS S INNER JOIN Materiales AS M ON S.reference = M.Referencia WHERE S.Warehouse LIKE '7' order by M.ReferenciaBase"
        'Objcmd = New SqlCommand(strSQL, dbconn)

        'Dim ds As DataSet = New DataSet()
        'Dim Adapter As SqlDataAdapter = New SqlDataAdapter(strSQL, dbconn)
         '   Adapter.Fill(ds, "STOCK")

		' Dim tableRow As DataRow
		' For Each tableRow In ds.Tables("Stock").Rows
		 
		 
		 
'Connect to LASSARD for PREF Database
Set DBConnection2 = Server.CreateObject("adodb.connection")
DSN2 = "DRIVER={SQL Server}; "
DSN2 = DSN2 & "server=qwtordb1;UID=qws-dev;PWD=welcome1;Database=Quest-dev"
DBConnection2.Open DSN2

'Access Database connection String looks like:
'Set rs = Server.CreateObject("adodb.recordset")
'strSQL = "SELECT * FROM Y_INV ORDER BY PART ASC"
'rs.Cursortype = 2
'rs.Locktype = 3
'rs.Open strSQL, DBConnection



'Expired Permission user WEB
'server=qwtordb1;DRIVER=SQL SERVER;DATABASE=Quest;UID=Anonymous;PWD=somepass
%>

