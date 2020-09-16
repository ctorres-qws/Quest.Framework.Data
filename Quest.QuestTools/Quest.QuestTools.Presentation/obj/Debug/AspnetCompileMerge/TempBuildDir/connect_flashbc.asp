<!--#include file="dbpath.asp"-->
<%
'Create a Query
    SQL = "SELECT Job, Floor, Tag, Style, X, Y, Y2, H1, H2, H3, H4, H5, H6, H7, H8, LM0, LM1, RM0, RM1, VH, VW FROM [" & tablename & "] WHERE (((Floor) = '" & floor & "')) ORDER by Tag ASC"
'Get a Record Set
    Set RS = DBConnection.Execute(SQL)
'Create a Query
    SQL = "Select * From X_Employees"
'Get a Record Set
    Set RS2 = DBConnection.Execute(SQL)
%>