Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class OpeCambiosBultos

    Public Shared Function GetValorActual(ByVal Codigo As String, _
                                          ByVal TipoTransaccion As String) As String

        Dim sSql As StringBuilder = New StringBuilder
        sSql.Append("SELECT TOP 1 a.CAB_VALOR_ACTUAL ")
        sSql.Append(" FROM CAMBIOS_BULTOS AS a WITH ( NOLOCK )")
        sSql.Append(" INNER JOIN TIPOS_TRANSACCIONES AS b WITH ( NOLOCK )")
        sSql.Append(" ON a.TTR_CODIGO = b.TTR_CODIGO")
        sSql.Append(" WHERE a.CAB_CLAVE = @CAB_CLAVE")
        sSql.Append(" AND b.TTR_DESCRIPCION = @TTR_DESCRIPCION")
        sSql.Append(" ORDER BY CAB_FECHA DESC")

        sSql.Replace("@CAB_CLAVE", String.Format("'{0}'", db.AdjustSql(Codigo)))
        sSql.Replace("@TTR_DESCRIPCION", String.Format("'{0}'", db.AdjustSql(TipoTransaccion)))

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
