Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class OpeMotivos

    Public Shared Function GetDescripcion(ByVal Codigo As String) As String

        Dim sSql As StringBuilder = New StringBuilder
        sSql.Append("SELECT [DESCRIPCION]")
        sSql.Append(" FROM [MOTIVOS] WITH ( NOLOCK )")
        sSql.Append(" WHERE [IDMOTIVO] = @IDMOTIVO")
        sSql.Append(" AND [ESTATUS] = 'A'")

        sSql.Replace("@IDMOTIVO", db.AdjustSql(Codigo))

        Dim oConn As New SqlConnection(db.GetConnectionString())
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.Text
            .CommandText = sSql.ToString
        End With

        Try
            oConn.Open()
            Return oCmd.ExecuteScalar
        Catch ex As Exception
            Return Nothing
        Finally
            oConn.Close()
        End Try

    End Function

End Class
