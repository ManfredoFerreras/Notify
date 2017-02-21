Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class OpeConfiguracion

    ''' <summary>
    ''' Get Header
    ''' </summary>
    ''' <returns>Encabezado</returns>
    ''' <remarks></remarks>
    Public Shared Function GetHeader() As String

        Dim sSql As StringBuilder = New StringBuilder
        sSql.Append("SELECT TOP 1 [ENCABEZAMIENTO]")
        sSql.Append(" FROM [CONFIGURACION] WITH ( NOLOCK )")
        sSql.Append(" WHERE [COM_CODIGO] = 1")
        sSql.Append(" AND [SUC_CODIGO] = '001'")
        Dim wrkstr As String = sSql.ToString

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
