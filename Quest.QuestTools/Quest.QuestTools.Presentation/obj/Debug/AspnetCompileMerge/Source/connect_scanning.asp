<!--#include file="dbpath.asp"-->
<%
'Create a Query
    SQL = "SELECT * FROM [" & tablename & "] WHERE (((Floor) = '" & floor & "')) ORDER by Tag ASC"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
'Create a Query
    SQL = "Select * From X_Employees"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL)
%>