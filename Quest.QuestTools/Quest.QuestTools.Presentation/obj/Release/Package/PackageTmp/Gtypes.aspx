<%@ Page Title="Home Page" Language="VB" AspCompat="true" EnableEventValidation="false" Debug="true" %>
<%@ Import Namespace="System.Data.SQLClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDB" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Web.Script.Serialization" %>
<%@ Import Namespace="Newtonsoft.Json.Linq" %>
<%@ Import Namespace="Newtonsoft.Json" %>
<%@ Import Namespace="qws_Tools" %>
<%@ Import Namespace="System.Web.Services"  %>
<%@ Import Namespace="System.Web.Script.Services" %>
<%@ Import Namespace="Newtonsoft.Json" %>

<script runat="server">

    Dim str_URL as String
    Sub Page_Load(Src As Object, E As EventArgs)
        If Not Me.IsPostBack Then
            ddlJobs.Text = GetJobs
            ddlExtGlassOptions.Text = GetGlassTypes
            ddlIntGlassOptions.Text = ddlExtGlassOptions.Text
            ddlOTOptions.Text = GetOverallThickness
            tableGTypes.Text = ShowGTypes
        End If
    End Sub

    Private Function ShowGTypes

        Dim cn_SQL As New OleDb.OleDbConnection With {
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;"
        }

        Dim cmd_Data As New OleDb.OleDbCommand
        Dim rdr_Data As OleDb.OleDbDataReader
        Dim str_SQL As String
        Dim sb_Results As StringBuilder = New StringBuilder()
        Dim str_Row As String = ""
        str_SQL = "SELECT A.Job, A.Gtype, B.Description AS ExtGlass, ExtLowE, IIF(IsNull(ExtSurface), '', ExtSurface) AS _ExtSurface, ExtTempered, ExtHeatStrengthened, ExtFritPattern, IIF(IsNull(ExtPatternID), '', ExtPatternID) AS _ExtPatternID, IIF(IsNull(ExtFritPatternSurface), '', ExtFritPatternSurface) AS _ExtFritPatternSurface," & vbCrLf
        str_SQL += "C.Description AS IntGlass, IntLowE, IIF(IsNull(IntSurface), '', IntSurface) As _IntSurface, IntTempered, IntHeatStrengthened, IntFritPattern, IIF(IsNull(IntPatternID), '', IntPatternID) AS _IntPatternID,IIF(IsNull(IntFritPatternSurface), '', IntFritPatternSurface) AS _IntFritPatternSurface," & vbCrLf
        str_SQL += "D.OT,IIF(IsNull(SpacerType), '', SpacerType) AS _SpacerType,IIF(IsNull(SpacerColor), '', SpacerColor) AS _SpacerColor,IIF(IsNull(SpacerSize), '', SpacerSize) AS _SpacerSize," & vbCrLf
        str_SQL += "IIF(IsNull(GasFill), '', GasFill) AS _GasFill, IIF(IsNull(SpacerPrimarySeal), '', SpacerPrimarySeal) AS _SpacerPrimarySeal,IIF(IsNull(SpacerPrimarySealColor), '', SpacerPrimarySealColor) AS _SpacerPrimarySealColor," & vbCrLf
        str_SQL += "IIF(IsNull(SpacerSecondarySeal), '', SpacerSecondarySeal) AS _SpacerSecondarySeal,IIF(IsNull(SpacerSecondarySealColor), '', SpacerSecondarySealColor) AS _SpacerSecondarySealColor,ExtIGUSpandrel, IIF(IsNull(ExtSpandrelColor), '', ExtSpandrelColor) AS _ExtSpandrelColor, IIF(IsNull(ExtSpandrelSurface), '', ExtSpandrelSurface) As _ExtSpandrelSurface," & vbCrLf
        str_SQL += "IntIGUSpandrel, IIF(IsNull(IntSpandrelColor), '', IntSpandrelColor) AS _IntSpandrelColor, IIF(IsNull(IntSpandrelSurface), '', IntSpandrelSurface) As _IntSpandrelSurface" & vbCrLf
        str_SQL += "FROM (((zGTypes A" & vbCrLf
        str_SQL += "INNER JOIN XQSU_GlassTypes B ON B.ID = A.ExtGlass)" & vbCrLf
        str_SQL += "INNER JOIN XQSU_GlassTypes C ON C.ID = A.IntGlass)" & vbCrLf
        str_SQL += "INNER JOIN XQSU_OTSpacer D ON D.ID = A.OT)"
        cn_SQL.Open()

        cmd_Data.Connection = cn_SQL
        cmd_Data.CommandText = str_SQL
        cmd_Data.CommandType = CommandType.Text
        rdr_Data = cmd_Data.ExecuteReader
        Dim listGtypes As New List(Of GtypeFormatted)
        'str_Row = "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td><button type=""button"" class=""btn btn-sm btn-primary"" onclick=""getGtype("{0}","{1}")"">Edit</button></td></tr>"
        Do While rdr_Data.Read
            Dim model As New GtypeFormatted
            model.Job = rdr_Data("Job")
            model.GType = rdr_Data("Gtype")
            model.ExtGlass = rdr_Data("ExtGlass")
            model.ExtLowE = rdr_Data("ExtLowE")
            model.ExtSurface = rdr_Data("_ExtSurface")
            model.ExtTempered = rdr_Data("ExtTempered")
            model.ExtHeatStrengthened = rdr_Data("ExtHeatStrengthened")
            model.ExtFritPattern = rdr_Data("ExtFritPattern")
            model.ExtPatternID = rdr_Data("_ExtPatternID")
            model.ExtFritPatternSurface = rdr_Data("_ExtFritPatternSurface")

            model.IntGlass = rdr_Data("IntGlass")
            model.IntLowE = rdr_Data("IntLowE")
            model.IntSurface = rdr_Data("_IntSurface")
            model.IntTempered = rdr_Data("IntTempered")
            model.IntHeatStrengthened = rdr_Data("IntHeatStrengthened")
            model.IntFritPattern = rdr_Data("IntFritPattern")
            model.IntPatternID = rdr_Data("_IntPatternID")
            model.IntFritPatternSurface = rdr_Data("_IntFritPatternSurface")

            model.OT = rdr_Data("OT")
            model.SpacerType = rdr_Data("_SpacerType")
            model.SpacerColor = rdr_Data("_SpacerColor")
            model.SpacerSize = rdr_Data("_SpacerSize")
            model.GasFill = rdr_Data("_GasFill")
            model.SpacerPrimarySeal = rdr_Data("_SpacerPrimarySeal")
            model.SpacerPrimarySealColor = rdr_Data("_SpacerPrimarySealColor")
            model.SpacerSecondarySeal = rdr_Data("_SpacerSecondarySeal")
            model.SpacerSecondarySealColor = rdr_Data("_SpacerSecondarySealColor")

            model.ExtIGUSpandrel = rdr_Data("ExtIGUSpandrel")
            model.ExtSpandrelColor = rdr_Data("_ExtSpandrelColor")
            model.ExtSpandrelSurface = rdr_Data("_ExtSpandrelSurface")

            model.IntIGUSpandrel = rdr_Data("IntIGUSpandrel")
            model.IntSpandrelColor = rdr_Data("_IntSpandrelColor")
            model.IntSpandrelSurface = rdr_Data("_IntSpandrelSurface")

            listGtypes.Add(model)
        Loop

        rdr_Data.Close()
        cn_SQL.Close()
        Return JsonConvert.SerializeObject(listGtypes)

    End Function

    Private Function GetGlassTypes()

        Dim cn_SQL As New OleDb.OleDbConnection With {
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;"
        }

        Try

            Dim cmd_Data As New OleDb.OleDbCommand
            Dim rdr_Data As OleDb.OleDbDataReader
            Dim str_SQL As String

            str_SQL = "SELECT ID, Description FROM XQSU_GlassTypes ORDER BY ID"


            cn_SQL.Open()


            cmd_Data.Connection = cn_SQL
            cmd_Data.CommandText = str_SQL
            cmd_Data.CommandType = CommandType.Text

            rdr_Data = cmd_Data.ExecuteReader

            Dim sb_Results As StringBuilder = New StringBuilder()

            Dim str_Row As String = ""
            Dim str_test As String = ""

            Do While rdr_Data.Read
                str_Row = "<option value=""{0}"">{1}</option>"
                sb_Results.Append(String.Format(str_Row, rdr_Data("ID"), rdr_Data("Description")))

            Loop

            rdr_Data.Close()
            Return sb_Results.ToString()

        Catch ex As Threading.ThreadAbortException
        Catch ex As Exception
            Response.Write(ex.StackTrace & "" & ex.Message)
            Response.End()
        Finally
            cn_SQL.Close()
        End Try
    End Function

    Private Function GetJobs()

        Dim cn_SQL As New OleDb.OleDbConnection With {
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;"
        }

        Try

            Dim cmd_Data As New OleDb.OleDbCommand
            Dim rdr_Data As OleDb.OleDbDataReader
            Dim str_SQL As String

            str_SQL = "SELECT DISTINCT Parent FROM Z_Jobs Where Parent <> '' and Completed = False"


            cn_SQL.Open()


            cmd_Data.Connection = cn_SQL
            cmd_Data.CommandText = str_SQL
            cmd_Data.CommandType = CommandType.Text

            rdr_Data = cmd_Data.ExecuteReader

            Dim sb_Results As StringBuilder = New StringBuilder()

            Dim str_Row As String = ""
            Dim str_test As String = ""

            Do While rdr_Data.Read
                str_Row = "<option value=""{0}"">{0}</option>"
                sb_Results.Append(String.Format(str_Row, rdr_Data("Parent"), rdr_Data("Parent")))

            Loop

            rdr_Data.Close()
            Return sb_Results.ToString()

        Catch ex As Threading.ThreadAbortException
        Catch ex As Exception
            Response.Write(ex.StackTrace & "" & ex.Message)
            Response.End()
        Finally
            cn_SQL.Close()
        End Try
    End Function

    Private Function GetOverallThickness()

        Dim cn_SQL As New OleDb.OleDbConnection With {
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;"
        }

        Try

            Dim cmd_Data As New OleDb.OleDbCommand
            Dim rdr_Data As OleDb.OleDbDataReader
            Dim str_SQL As String

            str_SQL = "SELECT ID, OT FROM XQSU_OTSpacer"


            cn_SQL.Open


            cmd_Data.Connection = cn_SQL
            cmd_Data.CommandText = str_SQL
            cmd_Data.CommandType = CommandType.Text

            rdr_Data = cmd_Data.ExecuteReader

            Dim sb_Results as StringBuilder = New StringBuilder()

            Dim str_Row As String = ""
            Dim str_test As String = ""

            Do While rdr_Data.Read
                str_Row = "<option value=""{0}"">{1}</option>"
                sb_Results.Append(String.Format(str_Row, rdr_Data("ID"), rdr_Data("OT")))
            Loop

            rdr_Data.Close
            Return sb_Results.ToString()

        Catch ex As Threading.ThreadAbortException
        Catch ex As Exception
            Response.Write(ex.StackTrace & "" & ex.Message)
            Response.End()
        Finally
            cn_SQL.Close
        End Try
    End Function

    Public Class GtypeFormatted
        Public Job As String
        Public GType As String

        Public ExtGlass As String
        Public ExtLowE As Boolean
        Public ExtSurface As String
        Public ExtTempered As Boolean
        Public ExtHeatStrengthened As Boolean
        Public ExtFritPattern As Boolean
        Public ExtPatternID As String
        Public ExtFritPatternSurface As String

        Public IntGlass As String
        Public IntLowE As Boolean
        Public IntSurface As String
        Public IntTempered As Boolean
        Public IntHeatStrengthened As Boolean
        Public IntFritPattern As Boolean
        Public IntPatternID As String
        Public IntFritPatternSurface As String

        Public SpacerType As String
        Public SpacerColor As String
        Public SpacerSize As String
        Public OT As String
        Public GasFill As String
        Public SpacerPrimarySeal As String
        Public SpacerPrimarySealColor As String
        Public SpacerSecondarySeal As String
        Public SpacerSecondarySealColor As String

        Public ExtIGUSpandrel As Boolean
        Public ExtSpandrelColor As String
        Public ExtSpandrelSurface As String

        Public IntIGUSpandrel As Boolean
        Public IntSpandrelColor As String
        Public IntSpandrelSurface As String
    End Class
    Public Class Gtype
        Public Job As String
        Public GType As String

        Public ExtGlass As Integer
        Public ExtLowE As Boolean
        Public ExtSurface As String
        Public ExtTempered As Boolean
        Public ExtHeatStrengthened As Boolean
        Public ExtFritPattern As Boolean
        Public ExtPatternID As String
        Public ExtFritPatternSurface As String

        Public IntGlass As Integer
        Public IntLowE As Boolean
        Public IntSurface As String
        Public IntTempered As Boolean
        Public IntHeatStrengthened As Boolean
        Public IntFritPattern As Boolean
        Public IntPatternID As String
        Public IntFritPatternSurface As String

        Public SpacerType As String
        Public SpacerColor As String
        Public SpacerSize As String
        Public OT As Integer
        Public GasFill As String
        Public SpacerPrimarySeal As String
        Public SpacerPrimarySealColor As String
        Public SpacerSecondarySeal As String
        Public SpacerSecondarySealColor As String

        Public ExtIGUSpandrel As Boolean
        Public ExtSpandrelColor As String
        Public ExtSpandrelSurface As String

        Public IntIGUSpandrel As Boolean
        Public IntSpandrelColor As String
        Public IntSpandrelSurface As String

        Public Function AddNewGType(Model as Gtype)
            Dim cn_SQL As New OleDb.OleDbConnection With {
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;"
            }
            Dim cmd_Data As New OleDb.OleDbCommand
            Dim rdr_Data As OleDb.OleDbDataReader
            Dim str_SQL As String
            Dim sb_Results as StringBuilder = New StringBuilder()

            str_SQL = "INSERT INTO zGTypes(Job,Gtype,ExtGlass,ExtLowE,ExtSurface,ExtTempered,ExtHeatStrengthened,ExtFritPattern,ExtPatternID,ExtFritPatternSurface,"&vbCrLf
            str_SQL += "IntGlass,IntLowE,IntSurface,IntTempered,IntHeatStrengthened,IntFritPattern,IntPatternID,IntFritPatternSurface,"&vbCrLf
            str_SQL += "OT,SpacerType,SpacerColor,SpacerSize,GasFill,SpacerPrimarySeal,SpacerPrimarySealColor,SpacerSecondarySeal,SpacerSecondarySealColor,ExtIGUSpandrel,ExtSpandrelColor,ExtSpandrelSurface,IntIGUSpandrel,IntSpandrelColor,IntSpandrelSurface)"&vbCrLf
            str_SQL += "VALUES('{0}','{1}',{2},{3},'{4}',{5},{6},{7},'{8}','{9}',"&vbCrLf
            str_SQL += "{10},{11},'{12}',{13},{14},{15},'{16}','{17}',"&vbCrLf
            str_SQL += "{18},'{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}',{27},'{28}','{29}',{30},'{31}','{32}')"

            sb_Results.Append(String.Format(str_SQL, Model.Job,Model.GType,Model.ExtGlass,Model.ExtLowE,Model.ExtSurface,Model.ExtTempered,Model.ExtHeatStrengthened,Model.ExtFritPattern,Model.ExtPatternID,Model.ExtFritPatternSurface,
            Model.IntGlass,Model.IntLowE,Model.IntSurface,Model.IntTempered,Model.IntHeatStrengthened,Model.IntFritPattern,Model.IntPatternID,Model.IntFritPatternSurface,
            Model.OT,Model.SpacerType,Model.SpacerColor,Model.SpacerSize,Model.GasFill,Model.SpacerPrimarySeal,Model.SpacerPrimarySealColor,Model.SpacerSecondarySeal,Model.SpacerSecondarySealColor,Model.ExtIGUSpandrel,Model.ExtSpandrelColor,Model.ExtSpandrelSurface,
            Model.IntIGUSpandrel,Model.IntSpandrelColor,Model.IntSpandrelSurface
            ))

            cn_SQL.Open

            cmd_Data.Connection = cn_SQL
            cmd_Data.CommandText = sb_Results.ToString()
            cmd_Data.CommandType = CommandType.Text

            cmd_Data.ExecuteNonQuery
            cmd_Data.Dispose

        End Function

        Public Function EditGType(Model As Gtype)
            Dim cn_SQL As New OleDb.OleDbConnection With {
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;"
            }

            Dim cmd_Data As New OleDb.OleDbCommand
            Dim rdr_Data As OleDb.OleDbDataReader
            Dim str_SQL As String
            Dim sb_Results As StringBuilder = New StringBuilder()

            str_SQL = "UPDATE zGTypes SET ExtGlass = {2},ExtLowE = {3},ExtSurface = '{4}',ExtTempered = {5},ExtHeatStrengthened = {6},ExtFritPattern = {7},ExtPatternID = '{8}',ExtFritPatternSurface = '{9}'," & vbCrLf
            str_SQL += "IntGlass = {10},IntLowE = {11},IntSurface = '{12}',IntTempered = {13},IntHeatStrengthened = {14},IntFritPattern = {15},IntPatternID = '{16}',IntFritPatternSurface = '{17}'," & vbCrLf
            str_SQL += "OT = {18},SpacerType = '{19}',SpacerColor = '{20}',SpacerSize = '{21}',GasFill = '{22}',SpacerPrimarySeal = '{23}',SpacerPrimarySealColor = '{24}',SpacerSecondarySeal = '{25}',SpacerSecondarySealColor = '{26}'," & vbCrLf
            str_SQL += "ExtIGUSpandrel = {27},ExtSpandrelColor = '{28}',ExtSpandrelSurface = '{29}',IntIGUSpandrel = {30},IntSpandrelColor = '{31}',IntSpandrelSurface = '{32}'" & vbCrLf
            str_SQL += "WHERE Job = '{0}' AND Gtype = '{1}'"

            sb_Results.Append(String.Format(str_SQL, Model.Job, Model.GType, Model.ExtGlass, Model.ExtLowE, Model.ExtSurface, Model.ExtTempered, Model.ExtHeatStrengthened, Model.ExtFritPattern, Model.ExtPatternID, ExtFritPatternSurface,
            Model.IntGlass, Model.IntLowE, Model.IntSurface, Model.IntTempered, Model.IntHeatStrengthened, Model.IntFritPattern, Model.IntPatternID, Model.IntFritPatternSurface,
            Model.OT, Model.SpacerType, Model.SpacerColor, Model.SpacerSize, Model.GasFill, Model.SpacerPrimarySeal, Model.SpacerPrimarySealColor, Model.SpacerSecondarySeal, Model.SpacerSecondarySealColor,
            Model.ExtIGUSpandrel, Model.ExtSpandrelColor, Model.ExtSpandrelSurface, Model.IntIGUSpandrel, Model.IntSpandrelColor, Model.IntSpandrelSurface
            ))

            cn_SQL.Open

            cmd_Data.Connection = cn_SQL
            cmd_Data.CommandText = sb_Results.ToString()
            cmd_Data.CommandType = CommandType.Text

            cmd_Data.ExecuteNonQuery
            cmd_Data.Dispose

        End Function

        Public Function GetGtype(Job as String, Type as String)

            Dim cn_SQL As New OleDb.OleDbConnection With {
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;"
            }

            Dim cmd_Data As New OleDb.OleDbCommand
            Dim rdr_Data As OleDb.OleDbDataReader
            Dim str_SQL As String
            Dim sb_Results as StringBuilder = New StringBuilder()
            str_SQL = "SELECT TOP 1 Job, Gtype, ExtGlass, ExtLowE, IIF(IsNull(ExtSurface), '', ExtSurface) AS ExtSurface, ExtTempered, ExtHeatStrengthened, ExtFritPattern, IIF(IsNull(ExtPatternID), '', ExtPatternID) AS ExtPatternID,IIF(IsNull(ExtFritPatternSurface), '', ExtFritPatternSurface) AS ExtFritPatternSurface," & vbCrLf
            str_SQL += "IntGlass, IntLowE, IIF(IsNull(IntSurface), '', IntSurface) As IntSurface, IntTempered, IntHeatStrengthened, IntFritPattern, IIF(IsNull(IntPatternID), '', IntPatternID) AS IntPatternID,IIF(IsNull(IntFritPatternSurface), '', IntFritPatternSurface) AS IntFritPatternSurface,"&vbCrLf
            str_SQL += "OT,IIF(IsNull(SpacerType), '', SpacerType) AS SpacerType,IIF(IsNull(SpacerColor), '', SpacerColor) AS SpacerColor,IIF(IsNull(SpacerSize), '', SpacerSize) AS SpacerSize,"&vbCrLf
            str_SQL += "IIF(IsNull(GasFill), '', GasFill) AS GasFill, IIF(IsNull(SpacerPrimarySeal), '', SpacerPrimarySeal) AS SpacerPrimarySeal,IIF(IsNull(SpacerPrimarySealColor), '', SpacerPrimarySealColor) AS SpacerPrimarySealColor,"&vbCrLf
            str_SQL += "IIF(IsNull(SpacerSecondarySeal), '', SpacerSecondarySeal) AS SpacerSecondarySeal,IIF(IsNull(SpacerSecondarySealColor), '', SpacerSecondarySealColor) AS SpacerSecondarySealColor,"&vbCrLf
            str_SQL += "ExtIGUSpandrel, IIF(IsNull(ExtSpandrelColor), '', ExtSpandrelColor) AS ExtSpandrelColor, IIF(IsNull(ExtSpandrelSurface), '', ExtSpandrelSurface) As ExtSpandrelSurface," & vbCrLf
            str_SQL += "IntIGUSpandrel, IIF(IsNull(IntSpandrelColor), '', IntSpandrelColor) AS IntSpandrelColor, IIF(IsNull(IntSpandrelSurface), '', IntSpandrelSurface) As IntSpandrelSurface" & vbCrLf
            str_SQL += "FROM zGTypes WHERE Job = '{0}' AND Gtype = '{1}'" & vbCrLf

            sb_Results.Append(String.Format(str_SQL, Job, Type))
            cn_SQL.Open

            cmd_Data.Connection = cn_SQL
            cmd_Data.CommandText = sb_Results.ToString()
            cmd_Data.CommandType = CommandType.Text
            rdr_Data = cmd_Data.ExecuteReader

            Dim model As New Gtype
            Do While rdr_Data.Read

                model.Job = rdr_Data("Job")
                model.GType = rdr_Data("Gtype")
                model.ExtGlass = rdr_Data("ExtGlass")
                model.ExtLowE = rdr_Data("ExtLowE")
                model.ExtSurface = rdr_Data("ExtSurface")
                model.ExtTempered = rdr_Data("ExtTempered")
                model.ExtHeatStrengthened = rdr_Data("ExtHeatStrengthened")
                model.ExtFritPattern = rdr_Data("ExtFritPattern")
                model.ExtPatternID = rdr_Data("ExtPatternID")
                model.ExtFritPatternSurface = rdr_Data("ExtFritPatternSurface")

                model.IntGlass = rdr_Data("IntGlass")
                model.IntLowE = rdr_Data("IntLowE")
                model.IntSurface = rdr_Data("IntSurface")
                model.IntTempered = rdr_Data("IntTempered")
                model.IntHeatStrengthened = rdr_Data("IntHeatStrengthened")
                model.IntFritPattern = rdr_Data("IntFritPattern")
                model.IntPatternID = rdr_Data("IntPatternID")
                model.IntFritPatternSurface = rdr_Data("IntFritPatternSurface")

                model.OT = rdr_Data("OT")
                model.SpacerType = rdr_Data("SpacerType")
                model.SpacerColor = rdr_Data("SpacerColor")
                model.SpacerSize = rdr_Data("SpacerSize")
                model.GasFill = rdr_Data("GasFill")
                Model.SpacerPrimarySeal = rdr_Data("SpacerPrimarySeal")
                Model.SpacerPrimarySealColor = rdr_Data("SpacerPrimarySealColor")
                Model.SpacerSecondarySeal = rdr_Data("SpacerSecondarySeal")
                Model.SpacerSecondarySealColor = rdr_Data("SpacerSecondarySealColor")

                Model.ExtIGUSpandrel = rdr_Data("ExtIGUSpandrel")
                Model.ExtSpandrelColor = rdr_Data("ExtSpandrelColor")
                Model.ExtSpandrelSurface = rdr_Data("ExtSpandrelSurface")

                Model.IntIGUSpandrel = rdr_Data("IntIGUSpandrel")
                Model.IntSpandrelColor = rdr_Data("IntSpandrelColor")
                Model.IntSpandrelSurface = rdr_Data("IntSpandrelSurface")

            Loop

            rdr_Data.Close
            Return JsonConvert.SerializeObject(model)

            cn_SQL.Close
        End Function

    End Class

    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Shared Function btnSaveResource(ByVal parDesc As String) As String

        Dim d As String

        Try

            Dim model As Gtype
            model = New JavaScriptSerializer().Deserialize(Of Gtype)(parDesc)
            model.AddNewGType(model)

            d = "Success"
        Catch ex As Exception
            d = "Failed" & ex.StackTrace & "" & ex.Message
        End Try

        Return d + "test"
    End Function
    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Shared Function btnSaveEdit(ByVal parDesc As String) As String

        Dim d As String

        Try

            Dim model As Gtype
            model = New JavaScriptSerializer().Deserialize(Of Gtype)(parDesc)

            model.EditGType(model)

            d = "Success"
        Catch ex As Exception
            d = "Failed" & ex.StackTrace & "" & ex.Message
        End Try

        Return d + "test"
    End Function
    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Json)> _
    Public Shared Function GetGtype(job As String, gtype As String) As String

        Dim d As String

        Try

            Dim model As New Gtype

            d = model.GetGtype(job, gtype)

        Catch ex As Exception
            d = "Failed " & ex.StackTrace & "" & ex.Message
        End Try

        Return d
    End Function
</script>

<html>
	<head>
		<link href="stylePopUp.css?v=1" rel="stylesheet">
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css"
              asp-fallback-test-class="sr-only" asp-fallback-test-property="position" asp-fallback-test-value="absolute"
              crossorigin="anonymous"
              integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" />
    	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css" />
    	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/malihu-custom-scrollbar-plugin/3.1.5/jquery.mCustomScrollbar.min.css">
		<link href="https://gitcdn.github.io/bootstrap-toggle/2.2.2/css/bootstrap-toggle.min.css" rel="stylesheet">
    	<link rel="stylesheet" href="https://cdn.datatables.net/1.10.20/css/dataTables.bootstrap4.min.css" />
		<link rel="stylesheet" href="https://raw.github.com/daneden/animate.css/master/animate.css" />
		
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.3.1/jquery.min.js"
                asp-fallback-test="window.jQuery"
                crossorigin="anonymous"
                integrity="sha256-FgpCb/KJQlLNfOu91ta32o/NMZxltwRo8QtmkMRdAu8=">
        </script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.bundle.min.js"
                asp-fallback-test="window.jQuery && window.jQuery.fn && window.jQuery.fn.modal"
                crossorigin="anonymous"
                integrity="sha384-xrRywqdh3PHs8keKZN+8zzc5TX0GRTLCcmivcbNJWm2rs5C8PRhcEn3czEjhAO9o">
        </script>
    	<script src="https://cdn.datatables.net/1.10.20/js/jquery.dataTables.min.js"></script>
    	<script src="https://cdn.datatables.net/1.10.20/js/dataTables.bootstrap4.min.js"></script>
    	<script defer src="https://use.fontawesome.com/releases/v5.0.13/js/solid.js" integrity="sha384-tzzSw1/Vo+0N5UhStP3bvwWPq+uvzCMfrN1fEFe+xBmv1C/AtVX5K0uZtmcHitFZ" crossorigin="anonymous"></script>
    	<script defer src="https://use.fontawesome.com/releases/v5.0.13/js/fontawesome.js" integrity="sha384-6OIrr52G08NpOFSZdxxz1xdNSndlD4vdcf/q2myIUVO0VsqaGHJsB0RaBE01VTOY" crossorigin="anonymous"></script>
    	<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.0/umd/popper.min.js" integrity="sha384-cs/chFZiN24E4KMATLdqdvsezGxaGsi4hLGOzlXwp5UZB1LY//20VyM2taTB4QvJ" crossorigin="anonymous"></script>
		<script src="https://gitcdn.github.io/bootstrap-toggle/2.2.2/js/bootstrap-toggle.min.js"></script>
    	<script src="https://cdnjs.cloudflare.com/ajax/libs/malihu-custom-scrollbar-plugin/3.1.5/jquery.mCustomScrollbar.concat.min.js"></script>
		<script src="js/bootstrap-notify.js"></script>
<style>
	body { font-family: Arial; font-size: 12px; }
	center{
		font-size: 1rem;
	}
	text { font-size: 12px !important; }
	.navbar-nav > li > a { color: #15a9df !important }
	.btn-primary { background-color: #15a9df !important;     
					/*margin-right: 15px !important;*/
    				padding-left: 30px !important;
    				padding-right: 30px !important;
    				float: right;}
					.btn-primary:hover {
    					color: rgb(199, 199, 199);
					}	
	.title-main { background-color: rgb(255,255,255) !important; padding-left: 0px; }  /* 237 */
	div.dt-buttons {
     	float: right; 
		 background-image: -webkit-linear- none !important;
		 background-image: none !important;
	}
</style>

<style>
	h2{
		font-family: Raleway,sans-serif;
    	font-size: 24px;
    	font-style: normal;
    	font-weight: 500;
    	letter-spacing: 4px;
	}
	/*.csTable tr:nth-child(odd){
		background-color: #eaeaea;
		color: #0;
	}*/
	.csTable tr{
		background-color: #fafafa;
		color: #0;
	}
	.trDetails{
		background-color: #fefefe;
    	border: solid 1px #dedede;
		display: inline-flex;
		width: 100%;
	}
	.trDetails th, .trDetails td{
		border: solid 1px #dedede;
	}
	.trDetails tr{
		background-color: #fefefe !important;
		color: #0;
	}
	.trDetails th{
		background-color: #DDDDDD;
	}
	.csTable { border: 1px solid #cccccc; }

	/*.csTable tr:nth-child(even){
		background-color: #fff;
		color: #0;
	}*/
	
	.csTable td, .csTable th { padding-right: 10px; height: 30px; font-size: 14px; padding-left: 10px;}

	input { width: 120px !important; border: 1px solid rgb(221,221,221) !important; border-radius: 5px; margin-bottom: 5px !important; height: 25px !important; background-color: white;}
	input:disabled {
    cursor: default;
    background-color: #eaeaea;
	}
	#dialog-form > fieldset > input { width: 90% !important; }

	.ui-dialog .ui-dialog-content { overflow: hidden !important; }

	fieldset { border: 1px solid rgb(221,221,221); border-radius: 5px; margin-bottom: 11px; }

	.csDialogRow { xborder: 1px solid black; clear: both;}
	.csDialogRow > label { width: 120px; float: left !important; xborder: 1px solid black; }
	.csDialogRow > input { width: 440px !important; float: left !important; xborder: 1px solid black; margin-left: 0px !important; margin-top: 0px !important; padding-left: 5px !important; }

	.row > input:not([type|=radio]):not([type|=checkbox]) {
		padding-left: 5px !important;
	}

	.row > table > tbody > tr > td > input:not([type|=radio]):not([type|=checkbox]) {
		padding-left: 5px !important;
		
	}
	
	select { border: 1px solid rgb(221,221,221) !important; border-radius: 5px; padding: 3px 3px 3px 3px; margin-bottom: 5px;
	height: 25px;
    font-size: 14px; }
	body.landscape > .toolbar > h1 { width: 350px !important; }

	.csSearch > tbody > tr > td  { padding-left: 20px; }

	.csTData { text-align: center; font-size: 13px; }
	.csTDataAmt { text-align: right; padding-right: 5px; font-size: 13px; }

	.csCheckBox { width: 20px !important; }
	.dataTables_filter {
text-align: left !important;
}
	.container-fluid
	{
		padding: 0px;
	}
	.navbar-brand
	{
		padding: 15px;
	}
	#searchArea {
	}
	#searchArea input {
    	padding: 5px;
		margin-left: 10px;
		margin-top: 3px;
	}
	#searchArea label {
    	display: inline-flex;
		align-items: center;
	}
	#GtypesTable button{
		padding: 0px;
		height: 75%;
	}
	.text-responsive {
  		font-size: calc(100% + .2vw + .2vh);
	}
	.text-responsive-small {
  		font-size: calc(100% + .05vw + .05vh);
	}
	.modal-body label{
		font-size: 14px;
		font-variant: normal;
    	font-variant-ligatures: normal;
    	font-variant-caps: normal;
    	font-variant-numeric: normal;
    	font-variant-east-asian: normal;
		white-space: normal;
    	line-height: normal;
    	font-weight: normal;
    	font-style: normal;
		text-transform: capitalize;
		font-family: Arial;
		letter-spacing: 1px;
	}
	.toggle-btn-selector span{
		    background-color: #fff;
    		border-color: #ccc;
	}
	.toggle-group label{
		font-size: 12px;
	}
	.toggle.btn {
    	min-height: 25px;
	}
	.btn-toggle-argon{
    	color: #212529;
    	background-color: rgb(226, 224, 0);
    	    border-color: rgb(199, 199, 199);
	}
	.btn-toggle-air{
    	color: #fff;
    	background-color: #15a9df !important;
    	border-color: rgb(199, 199, 199);
	}
	.custom-checkbox .custom-control-input:checked~.custom-control-label::before{
  		background-color: #15a9df;
	}
	#ddlOT{
		max-width: 63%;
    	margin-left: -15px;
	}
	.modal-dialog {
    	max-width: max-content;
    	margin: 1.75rem auto;
	}
	.modal-body select {
		padding-left: 0rem;
		padding-right: 1rem;
		font-size: 13px;
	}
	.row-modal {
		margin-bottom: 1rem;
	}
	.mr-3-neg{
		margin-right: -1.1rem;
	}
	td.details-control {
    background: url('https://datatables.net/examples/resources/details_open.png') no-repeat center center;
    cursor: pointer;
}
tr.shown td.details-control {
    background: url('https://datatables.net/examples/resources/details_close.png') no-repeat center center;
}
	.tableDetails td{
		text-align: center;
	}
	.tableDetails img{
		width: 1rem;
	}
	.tableDetailsTitle{
		text-align: center;
    	margin-top: 1rem;
    	height: 30px;
    	font-weight: bold;
		background-color: #CCCCCC;
		margin: 1rem 1rem 0rem 1rem;
    	padding-top: .35rem;
	}
	.imageGrayscale{
		filter: grayscale(100%) brightness(175%);
		width: .5rem !important;
	}
	legend{
		width: auto;
		border: none;
		margin-left: 1.5rem;
	} 
	.mw-45{
		max-width: 45%;
	}
	.col-1-and-a-half{
		flex: 0 0 8.333333%;
    	max-width: 8.333333%;
	}
	select:disabled{
		background-color: #eaeaea;
	}
</style>
<script>

	var Model;
	var PageModeFunction;
	$(document).ready(function () {
		
		PageModeFunction = add;
		initializeTable();
    	$("#txtExtPatternID").prop('disabled', true);
		$("#ddlExtFritPatternSurface").prop('disabled', true);
		$("#txtIntPatternID").prop('disabled', true);
		$("#ddlIntFritPatternSurface").prop('disabled', true);
		$("#ddlExtSurface").prop('disabled', true);
		$("#ddlIntSurface").prop('disabled', true);
		$("#ddlExtSpandrelColor").prop('disabled', true);
		$("#ddlExtSpandrelSurface").prop('disabled', true);
		$("#ddlIntSpandrelColor").prop('disabled', true);
		$("#ddlIntSpandrelSurface").prop('disabled', true);
		
		$("#cbExtFritPattern").change(function () {
            if ($("#cbExtFritPattern").is(":checked")) {
                $("#txtExtPatternID").prop('disabled', false);
				$("#ddlExtFritPatternSurface").prop('disabled', false);
            } else {
                $("#txtExtPatternID").prop('disabled', true);
				$("#txtExtPatternID").val("");
				$("#ddlExtFritPatternSurface").prop('disabled', true);
				$("#ddlExtFritPatternSurface").val("");
            }
		
        });

		$("#cbIntFritPattern").change(function () {
            if ($("#cbIntFritPattern").is(":checked")) {
                $("#txtIntPatternID").prop('disabled', false);
				$("#ddlIntFritPatternSurface").prop('disabled', false);
            } else {
                $("#txtIntPatternID").prop('disabled', true);
				$("#txtIntPatternID").val("");
				$("#ddlIntFritPatternSurface").prop('disabled', true);
				$("#ddlIntFritPatternSurface").val("");
            }
		
        });
		
		$("#cbExtLowE").change(function () {
			debugger;
            if ($("#cbExtLowE").is(":checked")) {
                $("#ddlExtSurface").prop('disabled', false);
				$("#ddlExtSurface").val("2");
            } else {
                $("#ddlExtSurface").prop('disabled', true);
				$("#ddlExtSurface").val("");
            }
		
        });
		$("#cbIntLowE").change(function () {
            if ($("#cbIntLowE").is(":checked")) {
                $("#ddlIntSurface").prop('disabled', false);
				$("#ddlIntSurface").val("3");
            } else {
                $("#ddlIntSurface").prop('disabled', true);
				$("#ddlIntSurface").val("");
            }
		
        });
		$("#cbExtIGUSpandrel").change(function () {
            if ($("#cbExtIGUSpandrel").is(":checked")) {
                $("#ddlExtSpandrelColor").prop('disabled', false);
				$("#ddlExtSpandrelSurface").prop('disabled', false);
            } else {
                $("#ddlExtSpandrelColor").prop('disabled', true);
				$("#ddlExtSpandrelSurface").prop('disabled', true);
				$("#ddlExtSpandrelColor").val("");
				$("#ddlExtSpandrelSurface").val("");
            }
		
        });
		$("#cbIntIGUSpandrel").change(function () {
            if ($("#cbIntIGUSpandrel").is(":checked")) {
                $("#ddlIntSpandrelColor").prop('disabled', false);
				$("#ddlIntSpandrelSurface").prop('disabled', false);
            } else {
                $("#ddlIntSpandrelColor").prop('disabled', true);
				$("#ddlIntSpandrelSurface").prop('disabled', true);
				$("#ddlIntSpandrelColor").val("");
				$("#ddlIntSpandrelSurface").val("");
            }
		
        });

		$("#ddlOT").change(function () {
			if($("#ddlOT option:selected").text().indexOf('/') > 0)
            	$("#txtSpacerSize").val($("#ddlOT option:selected").text().split('/')[2].trim());
			else
				$("#txtSpacerSize").val("")
        });
		$('#btnAddNew').on('click', function (e) {
			Model.Job = $("#ddlJobs").val();
			Model.GType = $("#ddlGTypes").val();

			if($("#ddlGTypes").val() != "")
			{
				isGTypeRegistered(Model.Job,Model.GType, function(job,gtype){
					showDangerAlert('GType ' + gtype + ' for Job ' + job + ' has already been registered.');
				}, function(job,gtype){
					clearControls();
					$("#txtJob").val(job);
					$("#txtGType").val(gtype);
					ShowModal(true);
				});
			}
			else
			{
				showDangerAlert('Please select a Job and a Type');
			}
		});
		initializeModel();
		$('#gTypesModal').on('show.bs.modal', function () {
			
			debugger;
			if(PageModeFunction == edit && json != undefined)
				fillControls(json);
		});
		$('#gTypesModal').on('shown.bs.modal', function () {
			
		});
		$('#gTypesModal').on('hidden.bs.modal', function () {
			clearControls();
		});
		
	});
	function isGTypeRegistered(job,gtype, registered, notRegistered){
		$.ajax({
            type: "POST",
            url: "Gtypes.aspx/GetGtype",
            data: JSON.stringify({job: job, gtype: gtype}),
            contentType: "application/json; charset=utf-8",
            dataType: "json",
			success: function(msg) {
				debugger;
				var x = JSON.parse(msg.d);
				if(x.Job != undefined && x.Job != "")
						registered(job,gtype);						
					else{
						notRegistered(job,gtype);
						
					}
            }
		});   
	}
	function initializeTable(){
		debugger;
		var data = JSON.parse($("#tableGTypes").text());
		var table = $('#GtypesTable').DataTable({
			"paging":   false,
        	"ordering": true,
        	"info":     false,
			"language": {
            	"searchPlaceholder": "Job filter"
        	},
			initComplete : function() {
        		$("#GtypesTable_filter").detach().appendTo('#searchArea');
    		},
        	"data": data,
        	"columns": [
            {
                "className":      'details-control',
                "orderable":      false,
                "data":           null,
                "defaultContent": ''
            },
            { "data": "Job" },
            { "data": "GType" },
            { "data": "ExtGlass" },
            { "data": "IntGlass" },
			{ "data": "OT" }
        ],
			"columnDefs": [ {
            "targets": 6,
			"data": null,
            "defaultContent": "<button type=\"button\" class=\"btn btn-sm btn-primary\">Edit</button>"
        	} ],
        "order": [[1, 'asc']]
    	} );
     
    // Add event listener for opening and closing details
    	$('#GtypesTable tbody').on('click', 'td.details-control', function () {
        var tr = $(this).closest('tr');
        var row = table.row( tr );
 
        if ( row.child.isShown() ) {
            // This row is already open - close it
            row.child.hide();
            tr.removeClass('shown');
        }
        else {
            // Open this row
            row.child( format(row.data()) ).show();
            tr.addClass('shown');
        }
    } );
		$('#GtypesTable tbody').on( 'click', 'button', function () {
        	var data = table.row( $(this).parents('tr') ).data();
			getGtype(data.Job,data.GType);
    	} );

	}
	function format ( item ) {
    // `d` is the original data object for the row
    return '<div class="trDetails" style="display: inline-flex;"><div><table class="tableDetails" cellpadding="5" cellspacing="0" border="0" style="margin: 1rem;">'+
			'<tr><th style="background-color: #CCCCCC;text-align: center;">Glass</th><th>Low E Surface</th><th>Tempered</th><th>Heat Strengthened</th><th>Frit Pattern</th><th>Frit Pattern Surface</th><th>Spandrel Color</th><th>Spandrel Surface</th></tr>'+
        	'<tr>'+
            	'<th>Exterior</th>'+
            	'<td>'+ (item.ExtLowE ? item.ExtSurface : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
        		'<td>'+ (item.ExtTempered ? '<img src="images/checkmark.png" />' : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.ExtHeatStrengthened ? '<img src="images/checkmark.png" />' : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.ExtFritPattern ? item.ExtPatternID : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.ExtFritPattern ? item.ExtFritPatternSurface : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.ExtIGUSpandrel ? item.ExtSpandrelColor : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.ExtIGUSpandrel ? item.ExtSpandrelSurface : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
        	'</tr>'+
			'<tr>'+
            	'<th>Interior</th>'+
            	'<td>' + (item.IntLowE ? item.IntSurface : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
        		'<td>'+ (item.IntTempered ? '<img src="images/checkmark.png" />' : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.IntHeatStrengthened ? '<img src="images/checkmark.png" />' : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.IntFritPattern ? item.IntPatternID : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.IntFritPattern ? item.IntFritPatternSurface : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.IntIGUSpandrel ? item.IntSpandrelColor : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
				'<td>'+ (item.IntIGUSpandrel ? item.IntSpandrelSurface : '<img class="imageGrayscale" src="images/minus.png" />') +'</td>'+
        	'</tr>'+
    	'</table></div>'+
		
		'<div><div class="tableDetailsTitle">Spacer</div><table class="tableDetails" cellpadding="5" cellspacing="0" border="0" style="margin: 0 1rem 1rem 1rem;">'+			
			'<tr><th>Type</th><th>Color</th><th>Size</th><th>Gas Fill</th><th>Primary Seal</th><th>Primary Seal Color</th><th>Secondary Seal</th><th>Secondary Seal Color</th></tr>'+
        	'<tr>'+
            	'<td>'+ item.SpacerType +'</td>'+
				'<td>'+ item.SpacerColor +'</td>'+
				'<td>'+ item.SpacerSize +'</td>'+
				'<td>'+ item.GasFill +'</td>'+
				'<td>'+ item.SpacerPrimarySeal +'</td>'+
				'<td>'+ item.SpacerPrimarySealColor +'</td>'+
				'<td>'+ item.SpacerSecondarySeal +'</td>'+
				'<td>'+ item.SpacerSecondarySealColor+'</td>'+          		
        	'</tr>'+
    	'</table></div></div>'
		;
	}
	function initializeModel(){
		Model = {
			Job: "",
			GType: "",

			ExtGlass: "",
			ExtLowE: "",
			ExtSurface: "",
			ExtTempered: "",
			ExtHeatStrengthened: "",
			ExtFritPattern: "",
			ExtPatternID: "",
			ExtFritPatternSurface: "",

			IntGlass: "",
			IntLowE: "",
			IntSurface: "",
			IntTempered: "",
			IntHeatStrengthened: "",
			IntFritPattern: "",
			IntPatternID: "",
			IntFritPatternSurface: "",

			SpacerType: "",
			SpacerColor: "",
			SpacerSize: "",
			OT: "",
			GasFill: "",
			SpacerPrimarySeal: "",
			SpacerPrimarySealColor: "",
			SpacerSecondarySeal: "",
			SpacerSecondarySealColor: "",
		
			ExtIGUSpandrel: "",
			ExtSpandrelColor: "",
			ExtSpandrelSurface: "",
			IntIGUSpandrel: "",
			IntSpandrelColor: "",
			IntSpandrelSurface: "",
			ButtonHTML: ""	
		}
	}

	function fillModel()
	{
		Model.Job = $("#txtJob").val();
		Model.GType = $("#txtGType").val();
		Model.ExtGlass = parseInt($("#ddlExtGlass").val());
		Model.ExtLowE = $("#cbExtLowE").is(":checked");
		Model.ExtSurface = $("#ddlExtSurface").val() ?? "";
		Model.ExtTempered = $("#cbExtTempered").is(":checked");
		Model.ExtHeatStrengthened = $("#cbExtHeatStrengthened").is(":checked");
		Model.ExtFritPattern = $("#cbExtFritPattern").is(":checked");
		Model.ExtPatternID = $("#txtExtPatternID").val();
		Model.ExtFritPatternSurface = $("#ddlExtFritPatternSurface").val() ?? "";

		Model.IntGlass = parseInt($("#ddlIntGlass").val());
		Model.IntLowE = $("#cbIntLowE").is(":checked");
		Model.IntSurface = $("#ddlIntSurface").val() ?? "";
		Model.IntTempered = $("#cbIntTempered").is(":checked");
		Model.IntHeatStrengthened = $("#cbIntHeatStrengthened").is(":checked");
		Model.IntFritPattern = $("#cbIntFritPattern").is(":checked");
		Model.IntPatternID = $("#txtIntPatternID").val();
		Model.IntFritPatternSurface = $("#ddlIntFritPatternSurface").val() ?? "";

		Model.SpacerType = $("#txtSpacerType").val();
		Model.SpacerColor = $("#ddlSpacerColor").val();
		Model.SpacerSize = $("#txtSpacerSize").val();
		Model.OT = parseInt($("#ddlOT").val());
		Model.GasFill = $("#cbGasFill").is(":checked") ? "Air" : "Argon";

		Model.SpacerPrimarySeal = $("#ddlSpacerPrimarySeal").val() ?? "";
		Model.SpacerPrimarySealColor = $("#ddlSpacerPrimarySealColor").val() ?? "";
		Model.SpacerSecondarySeal = $("#ddlSpacerSecondarySeal").val() ?? "";
		Model.SpacerSecondarySealColor = $("#ddlSpacerSecondarySealColor").val() ?? "";
		
		Model.ExtIGUSpandrel = $("#cbExtIGUSpandrel").is(":checked");
		Model.ExtSpandrelColor = $("#ddlExtSpandrelColor").val();
		Model.ExtSpandrelSurface = $("#ddlExtSpandrelSurface").val() ?? "";	

		Model.IntIGUSpandrel = $("#cbIntIGUSpandrel").is(":checked");
		Model.IntSpandrelColor = $("#ddlIntSpandrelColor").val();
		Model.IntSpandrelSurface = $("#ddlIntSpandrelSurface").val() ?? "";	
		
	}
	function fillControls(json)
	{
		$("#txtJob").val(json.Job);
		$("#txtGType").val(json.GType);
		$("#ddlExtGlass").val(json.ExtGlass);
		$("#cbExtLowE").prop('checked',json.ExtLowE);
		$("#ddlExtSurface").val(json.ExtSurface);
		$("#cbExtTempered").prop('checked',json.ExtTempered);
		$("#cbExtHeatStrengthened").prop('checked',json.ExtHeatStrengthened);
		$("#cbExtFritPattern").prop('checked',json.ExtFritPattern);
		$("#txtExtPatternID").val(json.ExtPatternID);
		$("#ddlExtFritPatternSurface").val(json.ExtFritPatternSurface);
		if(json.ExtFritPattern){
			$("#txtExtPatternID").prop('disabled', false);
			$("#ddlExtFritPatternSurface").prop('disabled', false);
		}
		else{
			$("#txtExtPatternID").prop('disabled', true);
			$("#ddlExtFritPatternSurface").prop('disabled', true);
		}

		$("#ddlIntGlass").val(json.IntGlass);
		$("#cbIntLowE").prop('checked',json.IntLowE);
		$("#ddlIntSurface").val(json.IntSurface);
		$("#cbIntTempered").prop('checked',json.IntTempered);
		$("#cbIntHeatStrengthened").prop('checked',json.IntHeatStrengthened);
		$("#cbIntFritPattern").prop('checked',json.IntFritPattern);		
		$("#txtIntPatternID").val(json.IntPatternID);
		$("#ddlIntFritPatternSurface").val(json.IntFritPatternSurface);
		if(json.IntFritPattern){
			$("#txtIntPatternID").prop('disabled', false);
			$("#ddlIntFritPatternSurface").prop('disabled', false);
		}
		else
		{
			$("#txtIntPatternID").prop('disabled', true);
			$("#ddlIntFritPatternSurface").prop('disabled', true);
		}

		$("#txtSpacerType").val(json.SpacerType);
		$("#ddlSpacerColor").val(json.SpacerColor);
		$("#txtSpacerSize").val(json.SpacerSize);
		$("#ddlOT").val(json.OT);
		
		if(json.GasFill == "Air")
			$("#cbGasFill").bootstrapToggle('on');
		else
			$("#cbGasFill").bootstrapToggle('off');

		
		
		$("#ddlSpacerPrimarySeal").val(json.SpacerPrimarySeal);
		$("#ddlSpacerPrimarySealColor").val(json.SpacerPrimarySealColor);
		$("#ddlSpacerSecondarySeal").val(json.SpacerSecondarySeal); 
		$("#ddlSpacerSecondarySealColor").val(json.SpacerSecondarySealColor); 
		
		$("#cbExtIGUSpandrel").prop('checked',json.ExtIGUSpandrel);
		$("#ddlExtSpandrelColor").val(json.ExtSpandrelColor);
		$("#ddlExtSpandrelSurface").val(json.ExtSpandrelSurface); 

		if(json.ExtIGUSpandrel)
		{
			$("#ddlExtSpandrelColor").prop('disabled', false);
			$("#ddlExtSpandrelSurface").prop('disabled', false);
		}	

		$("#cbIntIGUSpandrel").prop('checked',json.IntIGUSpandrel);
		$("#ddlIntSpandrelColor").val(json.IntSpandrelColor);
		$("#ddlIntSpandrelSurface").val(json.IntSpandrelSurface); 	

		if(json.IntIGUSpandrel)
		{
			$("#ddlIntSpandrelColor").prop('disabled', false);
			$("#ddlIntSpandrelSurface").prop('disabled', false);
		}	
	}
	function clearControls()
	{
		$("#txtJob").val("");
		$("#txtGType").val("");
		$("#ddlExtGlass").val(0);
		$("#cbExtLowE").prop('checked',false);
		$("#ddlExtSurface").val("");
		$("#ddlExtSurface").prop('disabled', true); 
		$("#cbExtTempered").prop('checked',false);
		$("#cbExtHeatStrengthened").prop('checked',false);
		$("#cbExtFritPattern").prop('checked',false);
		$("#txtExtPatternID").val("");
		$("#txtExtPatternID").prop('disabled', true);
		$("#ddlExtFritPatternSurface").val(0);
		$("#ddlExtFritPatternSurface").prop('disabled', true);

		$("#ddlIntGlass").val(0);
		$("#cbIntLowE").prop('checked',false);
		$("#ddlIntSurface").val("");
		$("#ddlIntSurface").prop('disabled', true); 
		$("#cbIntTempered").prop('checked',false);
		$("#cbIntHeatStrengthened").prop('checked',false);
		$("#cbIntFritPattern").prop('checked',false);		
		$("#txtIntPatternID").val("");
		$("#txtIntPatternID").prop('disabled', true);
		$("#ddlIntFritPatternSurface").val(0);
		$("#ddlIntFritPatternSurface").prop('disabled', true);

		$("#txtSpacerType").val("");
		$("#ddlSpacerColor").val(0);
		$("#txtSpacerSize").val("");
		$("#ddlOT").val(0);
		$("#cbGasFill").bootstrapToggle('on');

		$("#ddlSpacerPrimarySeal").val("");
		$("#ddlSpacerPrimarySealColor").val(0);
		$("#ddlSpacerSecondarySeal").val(""); 
		$("#ddlSpacerSecondarySealColor").val(0); 
		
		$("#cbExtIGUSpandrel").prop('checked', false);
		$("#ddlExtSpandrelColor").val("");
		$("#ddlExtSpandrelColor").prop('disabled', true); 	
		$("#ddlExtSpandrelSurface").val("");
		$("#ddlExtSpandrelSurface").prop('disabled', true); 

		$("#cbIntIGUSpandrel").prop('checked', false);
		$("#ddlIntSpandrelColor").val("");
		$("#ddlIntSpandrelColor").prop('disabled', true); 	
		$("#ddlIntSpandrelSurface").val("");
		$("#ddlIntSpandrelSurface").prop('disabled', true); 	
	}
	function ShowModal(IsAdding)
	{
		if(IsAdding)
		{			
			$("#idTitle").text("New GType")
			PageModeFunction = add;
		}
		else
		{			
			$("#idTitle").text("Edit")
			PageModeFunction = edit;
		}
		//$("#btnSaveChanges").on('click',PageModeFunction);
		$('#gTypesModal').modal('show')
	}
	var pageLog = "";
	function add()
	{
		if(validateModel())
		{
			fillModel();
			$.ajax({
            	type: "POST",
            	url: "Gtypes.aspx/btnSaveResource",
            	data: JSON.stringify({parDesc: JSON.stringify(Model)}),
            	contentType: "application/json; charset=utf-8",
            	dataType: "json",
            	success: function(msg) {
					debugger;
					if(msg.d.toLowerCase().indexOf('failed') < 0)
					{
						showSuccessAlert("The changes have been saved");     
						setTimeout(updatePage, 2000);						
					}
					else
					{
						pageLog = msg.d;
						showDangerAlert("An error ocurred, please contact the support team")
					}
            	}
        	});
		}
	}
	
	function edit()
	{
		if(validateModel())
		{
			fillModel();
			$.ajax({
            	type: "POST",
            	url: "Gtypes.aspx/btnSaveEdit",
            	data: JSON.stringify({parDesc: JSON.stringify(Model)}),
            	contentType: "application/json; charset=utf-8",
            	dataType: "json",
            	success: function(msg) {
					debugger;
                	if(msg.d.toLowerCase().indexOf('failed') < 0)
					{
						showSuccessAlert("The changes have been saved");     
						setTimeout(updatePage, 2000);						
					}
					else
					{
						pageLog = msg.d;
						showDangerAlert("An error ocurred, please contact the support team")
					}
            	}
        	});
		}
	}
	function updatePage()
	{
		location.reload();
	}
	function validateModel(){
		if($("#ddlExtGlass").val() == null || $("#ddlExtGlass").prop('selectedIndex') <= 0){
			showDangerAlert('Please select the Exterior Glass');
			return false;
		}	

		if($("#cbExtLowE").is(":checked") && $("#ddlExtSurface").val() == null){
			showDangerAlert('Please select the Exterior Surface');
			return false;
		}
						
		if($("#cbExtFritPattern").is(":checked")){
			if($("#txtExtPatternID").val() == null || $("#txtExtPatternID").val() == "")
			{
				showDangerAlert('Please enter the Exterior Pattern ID');
				return false;
			}
			if($("#ddlExtFritPatternSurface").val() == null || $("#ddlExtFritPatternSurface").val() == "")
			{
				showDangerAlert('Please select the Exterior Pattern Surface');
				return false;
			}
		}

		if($("#ddlIntGlass").val() == null || $("#ddlIntGlass").prop('selectedIndex') <= 0){
			showDangerAlert('Please select the Interior Glass');
			return false;
		}

		if($("#cbIntLowE").is(":checked") && $("#ddlIntSurface").val() == null){
			showDangerAlert('Please select the Interior Surface');
			return false;
		}
			
		if($("#cbIntFritPattern").is(":checked")){
			if($("#txtIntPatternID").val() == null || $("#txtIntPatternID").val() == "")
			{
				showDangerAlert('Please enter the Interior Pattern ID');
				return false;
			}
			if($("#ddlIntFritPatternSurface").val() == null || $("#ddlIntFritPatternSurface").val() == "")
			{
				showDangerAlert('Please select the Interior Pattern Surface');
				return false;
			}
		}

		if($("#txtSpacerType").val() == ""){
			showDangerAlert('Please enter the Spacer Type');
			return false;
		}

		if($("#ddlSpacerColor").val() == null || $("#ddlSpacerColor").val() == "Not selected"){
			showDangerAlert('Please select the Spacer Color');
			return false;
		}

		/*if($("#txtSpacerSize").val() == ""){
			showDangerAlert('Please enter the Spacer Size');
			return false;
		}*/

		if($("#ddlOT").val() == null || $("#ddlOT").val() == "Not selected"){
			showDangerAlert('Please select the Overall Thickness');
			return false;
		}

		if($("#ddlSpacerPrimarySeal").val() == null || $("#ddlSpacerPrimarySeal").val() == ""){
			showDangerAlert('Please select the Spacer Primary Seal');
			return false;
		}
		if($("#ddlSpacerPrimarySealColor").val() == null || $("#ddlSpacerPrimarySealColor").val() == ""){
			showDangerAlert('Please select the Spacer Primary Seal Color');
			return false;
		}

		if($("#ddlSpacerSecondarySeal").val() == null || $("#ddlSpacerSecondarySeal").val() == ""){
			showDangerAlert('Please select the Spacer Secondary Seal');
			return false;
		}
		if($("#ddlSpacerSecondarySealColor").val() == null || $("#ddlSpacerSecondarySealColor").val() == ""){
			showDangerAlert('Please select the Spacer Secondary Seal Color');
			return false;
		}
		
		if($("#cbExtIGUSpandrel").is(":checked"))
		{
			if($("#ddlExtSpandrelColor").val() == null || $("#ddlExtSpandrelColor").val() == ""){
				showDangerAlert('Please enter an Exterior Spandrel Color');
				return false;
			}
			if($("#ddlExtSpandrelSurface").val() == null || $("#ddlExtSpandrelSurface").val() == "Not selected"){
				showDangerAlert('Please select an Exterior Spandrel Surface');
				return false;
			}
		}

		if($("#cbIntIGUSpandrel").is(":checked"))
		{
			if($("#ddlIntSpandrelColor").val() == null || $("#ddlIntSpandrelColor").val() == ""){
				showDangerAlert('Please select an Interior Spandrel Color');
				return false;
			}
			if($("#ddlIntSpandrelSurface").val() == null || $("#ddlIntSpandrelSurface").val() == "Not selected"){
				showDangerAlert('Please select an Interior Spandrel Surface');
				return false;
			}
		}
		return true;
	}
	var json="";
	function getGtype(job,gtype)
	{
		PageModeFunction = edit;
		$.ajax({
            type: "POST",
            url: "Gtypes.aspx/GetGtype",
            data: JSON.stringify({job: job, gtype: gtype}),
            contentType: "application/json; charset=utf-8",
            dataType: "json"
			})
			.done(function(msg) {
				debugger;
				json = JSON.parse(msg.d);
				ShowModal(false);
            });        
	}
	
	function showDangerAlert(msg) {
    $.notify({
        title: '<center><strong>Important</strong></center>',
        message: '<center>' + msg + '<center>'
    	}, {
            type: 'danger',
            z_index: 5000,
            placement: {
                align: "center"
            },
            offset: {
                y: 50,
                x: 200
            },
            animate: {
				enter: 'animated fadeInRight',
				exit: 'animated fadeOutRight'
			},
            delay: 1000
			
        });
	}
	function showSuccessAlert(msg) {
    	$.notify({
        	title: '<center><strong>Success!</strong></center>',
        	message: '<center>' + msg + '<center>'
    	}, {
            type: 'success',
            z_index: 3000,
            placement: {
                align: "center"
            },
            offset: {
                y: 300,
                x: 200
            },
            animate: {
                enter: 'animated fadeInUp',
                exit: 'animated fadeOutDown'
            },
			delay: 5000,
			timer: 1000
        });
	}
</script>

	</head>
<body>
<form id="fMain" runat="server">
<div id="serverResponse"></div>
<div id="root">
	<div id="tableGTypes" style="display:none"><asp:Literal id="tableGTypes" runat="server"></asp:Literal></div>
	<div class="container-fluid">
		<div id="pageHeader">
					<div class="header-banner-small" ></div>
					<nav class="navbar-white navbar-nobottom navbar-default">
						<div class="container-fluid">
							<div class="navbar-header">
								<a class="navbar-brand" href="index.html#_Job"><img src="images/full-logo.png" class="App-logo" alt="logo"></a>
								<button class="navbar-toggle" data-toggle="collapse" data-target=".bsNavCollapse">
									<span class="sr-only">Toggle navigation</span>
									<span class="icon-bar"></span>
									<span class="icon-bar"></span>
									<span class="icon-bar"></span>
								</button>
							</div>
						</div>
					</nav>
		</div>
		<div style="opacity: 1; background-color: white; padding: 5px; margin: 5px;">
			<h2 runat="server">Gtypes</h2>	
			<fieldset >
				<div class="row m-0" style="padding: 5px;display: inline-flex;width:100%;">
					<div class="col-1 pl-0" style="margin: 3px;width: 10%;margin-top: 5px;margin-bottom: 0px;">
						<div class="row">
							<label class="col-4 pt-1" for="ddlJobs">Job</label>
							<select class="col-8 custom-control pl-1" id="ddlJobs">
								<asp:Literal id="ddlJobs" runat="server"></asp:Literal> 
							</select>
						</div>
					</div>
					<div class="col-2" style="margin: 3px;width: 11%;margin-top: 5px;margin-bottom: 0px;">
						<div class="row">
							<label class="col-5 pt-1" for="ddlGTypes">G Type</label>
							<input class="col-7 type="text" class="custom-control" id="ddlGTypes" placeholder="Enter a type"/>
						</div>
					</div>
					<div class="col-8" style="margin-right: -15px"></div>
					<div class="col-1 pl-0 pr-0" style="margin-top: 2px;"><button id="btnAddNew" type="button" class="btn btn-sm btn-primary">Add New Type</button></div>
				</div>
			</fieldset>
		
			<div class="toolbar">
				<div id="searchArea" style="display: inline-flex"></div>
				<%-- <asp:Literal id="windowProductionData" runat="server"></asp:Literal> --%>
				<table id="GtypesTable" class='sortable csTable tablePaginated' width='100%'>
					<thead>
						<tr style="background-color: #CCCCCC;border: solid 1px #555555;">
							<th></th>
							<th>Job</th>
							<th>Type</th>
							
							<th>Exterior</th>
							<th>Interior</th>
							<th>OT</th>
							<th></th>						
						</tr>
					</thead>
					<tbody>	
						
					</tbody>
				</table>
			</div>
		</div>
		

	</div>

</div>
<div class="modal fade" id="gTypesModal" tabindex="-1" role="dialog" aria-labelledby="gTypesModalTitle" >
	<div class="modal-dialog modal-xl modal-dialog-centered" role="document">
    	<div class="modal-content">
      		<div class="modal-header">
        		<h4 class="modal-title" id="idTitle"></h4>
        		<button type="button" class="close" data-dismiss="modal" aria-label="Close">
          			<span aria-hidden="true">&times;</span>
        		</button>
      		</div>
      		<div class="modal-body">
			  	
  				<div class="container">
				  <div class="row m-2">
				  	<div class="col-3">
						<div class="row">
			  				<label class="modal-title col-4">Job:</label>
							<input class="modal-title col-8" id="txtJob" disabled/>
						</div>
					</div>
					<div class="col-1"></div>
					<div class="col-3">
						<div class="row">
							<label class="modal-title col-4">GType:</label>
							<input class="modal-title col-8" id="txtGType" disabled/>
						</div>
					</div>
				</div>
    				<div class="row mt-2 pt-2 pb-3">
      					<div class="col-4 d-flex justify-content-center text-responsive">EXT</div>
						<div class="col-4 d-flex justify-content-center text-responsive">Spacer</div>
      					<div class="col-4 ml-auto d-flex justify-content-center text-responsive">INT</div>
    				</div>
					<div class="row m-2 pt-1 pb-1">
      					<div class="col-3">
							<div class="row row-modal">
								<label class="col-4" for="ddlExtGlass">Glass</label>
								<select class="col-8" id="ddlExtGlass" placeholder>
									<option>Select an option</option>
									<asp:Literal id="ddlExtGlassOptions" runat="server"></asp:Literal> 
								</select>
							</div>
							<div class="row row-modal">
								<div class="col-4 custom-control custom-checkbox ml-3 mr-3">
    								<input type="checkbox" class="custom-control-input" id="cbExtLowE">
    								<label class="custom-control-label pt-1" for="cbExtLowE">Low E</label>
								</div>
								<div class="col pull-right">
						  			<div class="row">						  	
										<label class="col pt-1" for="ddlExtSurface">Surface</label>
										<select class="col" id="ddlExtSurface" disabled>
											<option>2</option>							
										</select>
									</div>
								</div>							
							</div>
							<div class="row row-modal">
								<div class="col-6 custom-control custom-checkbox ml-3">
						  			<input type="checkbox" class="custom-control-input" id="cbExtTempered">
    								<label class="custom-control-label pt-1" for="cbExtTempered">Tempered</label>
								</div>
								<div class="col-5 custom-control custom-checkbox">
    								<input type="checkbox" class="custom-control-input" id="cbExtHeatStrengthened">
    								<label class="custom-control-label pt-1" for="cbExtHeatStrengthened">Heat Strengthened</label>
								</div>
							</div>
							<div class="row row-modal">
								<div class="col-6 custom-control custom-checkbox ml-3 mw-45">
    								<input type="checkbox" class="custom-control-input" id="cbExtFritPattern">
    								<label class="custom-control-label pt-1" for="cbExtFritPattern">Frit Pattern</label>
								</div>
								<div class="col-6 mw-45">
									<input type="text" id="txtExtPatternID" class="custom-control pl-2" disabled/>
								</div>
							</div>
							<div class="row mb-2">
								<label class="col-9" for="ddlExtFritPatternSurface">Frit Pattern Surface</label>
								<select class="col-3" id="ddlExtFritPatternSurface">
									<option>1</option>
									<option>2</option>
								</select>
							</div>
							<div class="row row-modal">
      							<div class="col-12 custom-control custom-checkbox ml-3">
    								<input type="checkbox" class="custom-control-input" id="cbExtIGUSpandrel">
    								<label class="custom-control-label pt-1" for="cbExtIGUSpandrel">IGU Spandrel</label>
								</div>
							</div>
							<div class="row">
								<div class="col-6">
									<div class="row">
										<label class="col-4" for="ddlExtSpandrelColor" style="margin-right: 1.5rem;">Color</label>
										<select class="col-5" id="ddlExtSpandrelColor" disabled>
											<option>Black</option>
											<option>Grey</option>							
										</select>
									</div>
								</div>
      							<div class="col-6">
    								<div class="row">						  	
										<label class="col-7" for="ddlExtSpandrelSurface">Surface</label>
										<select class="col-5" id="ddlExtSpandrelSurface" disabled>
											<option>1</option>
											<option>2</option>
											<option>3</option>	
											<option>4</option>							
										</select>
									</div>
								</div>
							</div>							
						</div>
						<div class="col-1-and-a-half"></div>
						<div class="col-3">
							<div class="row row-modal">
								<label class="col-4" for="txtSpacerType">Type</label>
								<input type="text" class="col-8" id="txtSpacerType" />
							</div>
							<div class="row row-modal">
								<label class="col-4" for="ddlSpacerColor">Color</label>
								<select class="col" id="ddlSpacerColor">
									<option>Not selected</option>
									<option>Black</option>
									<option>Grey</option>
								</select>
							</div>
							<div class="row row-modal">
								<label class="col-7" for="ddlOT">Overall thickness</label>
								<select class="col-5 ml-0" id="ddlOT">
									<option>Not selected</option>
									<asp:Literal id="ddlOTOptions" runat="server"></asp:Literal>
								</select>
							</div>
							<div class="row row-modal">
								<div class="col-5">
									<div class="row">
										<label class="col-6" for="txtSpacerSize">Size</label>
										<input type="text" class="col-5" id="txtSpacerSize" disabled/>
									</div>
								</div>
								<div class="col pl-1 toggle-btn-selector">
									<div class="row">
										<div class="col-3"></div>
										<label class="col-3 pl-1" for="cbGasFill">Gas</label>
										<div class="col-6 pl-0">
    										<input type="checkbox" id="cbGasFill" checked data-toggle="toggle" data-on="Air" data-off="Argon" data-onstyle="toggle-air" data-offstyle="toggle-argon" data-width="87" data-height="20">
    									</div>
									</div>
								</div>
							</div>
							<div class="row row-modal">
								<label class="col-5" for="ddlSpacerPrimarySeal">Primary Seal</label>
								<select class="col" id="ddlSpacerPrimarySeal">
									<option>Butyl</option>
								</select>
								<select class="col ml-1 pr-0" id="ddlSpacerPrimarySealColor" placeholder="Color">
									<option disabled selected hidden>Color</option>
									<option>Black</option>
									<option>Grey</option>
								</select>
							</div>
							<div class="row row-modal">
								<label class="col-5" for="ddlSpacerSecondarySeal">Secondary Seal</label>
								<select class="col" id="ddlSpacerSecondarySeal" >
									<option>Polysulfide</option>
									<option>Silicone</option>
								</select>
								<select class="col ml-1 pr-0" id="ddlSpacerSecondarySealColor" placeholder="Color">
									<option disabled selected hidden>Color</option>
									<option>Black</option>
									<option>Grey</option>
								</select>
							</div>		
						</div>
						<div class="col-1-and-a-half"></div>
      					<div class="col-3">
						  	<div class="row row-modal">
								<label class="col-4" for="ddlIntGlass">Glass</label>
								<select class="col-8" id="ddlIntGlass">
									<option>Select an option</option>
									<asp:Literal id="ddlIntGlassOptions" runat="server"></asp:Literal>
								</select>
							</div>
							<div class="row row-modal">
								<div class="col-4 custom-control custom-checkbox ml-3 mr-3">
    								<input type="checkbox" class="custom-control-input" id="cbIntLowE">
    								<label class="custom-control-label pt-1" for="cbIntLowE">Low E</label>
								</div>
								<div class="col pull-right">
						  			<div class="row">						  	
										<label class="col pt-1" for="ddlIntSurface">Surface</label>
										<select class="col" id="ddlIntSurface" disabled>
											<option>3</option>								
										</select>
									</div>
								</div>
							</div>
							<div class="row row-modal">
								<div class="col-6 custom-control custom-checkbox ml-3">
    								<input type="checkbox" class="custom-control-input" id="cbIntTempered">
    								<label class="custom-control-label pt-1" for="cbIntTempered">Tempered</label>
								</div>
								<div class="col-5 custom-control custom-checkbox">
    								<input type="checkbox" class="custom-control-input" id="cbIntHeatStrengthened">
    								<label class="custom-control-label pt-1" for="cbIntHeatStrengthened">Heat Strengthened</label>
								</div>
							</div>
							<div class="row row-modal">
							  	<div class="col-6 custom-control custom-checkbox ml-3 mw-45">
    								<input type="checkbox" class="custom-control-input" id="cbIntFritPattern">
    								<label class="custom-control-label pt-1" for="cbIntFritPattern">Frit Pattern</label>
								</div>
								<div class="col-6 mw-45">
									<input type="text" id="txtIntPatternID" class="custom-control pl-2"  disabled/>
								</div>
								
							</div>
							<div class="row mb-2">
								<label class="col-9" for="ddlIntFritPatternSurface">Frit Pattern Surface</label>
								<select class="col-3" id="ddlIntFritPatternSurface">
									<option>3</option>
									<option>4</option>
								</select>
							</div>
							<div class="row row-modal">
      							<div class="col-12 custom-control custom-checkbox ml-3">
    								<input type="checkbox" class="custom-control-input" id="cbIntIGUSpandrel">
    								<label class="custom-control-label pt-1" for="cbIntIGUSpandrel">IGU Spandrel</label>
								</div>
							</div>
							<div class="row">
								<div class="col-6">
									<div class="row">
										<label class="col-4" for="ddlIntSpandrelColor" style="margin-right: 1.5rem;">Color</label>
										<select class="col-5" id="ddlIntSpandrelColor" disabled>
											<option>Black</option>
											<option>Grey</option>						
										</select>
									</div>
								</div>
      							<div class="col-6">
    								<div class="row">						  	
										<label class="col-7" for="ddlIntSpandrelSurface">Surface</label>
										<select class="col-5" id="ddlIntSpandrelSurface" disabled>
											<option>1</option>
											<option>2</option>
											<option>3</option>	
											<option>4</option>							
										</select>
									</div>
								</div>
							</div>		
						</div>
    				</div>				
			</div>
      		<div class="modal-footer">
        		<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        		<button id="btnSaveChanges" type="button" onclick="PageModeFunction()" class="btn btn-primary">Save changes</button>
      		</div>
    	</div>
  	</div>
</div>
</form>

</body>
</html>
