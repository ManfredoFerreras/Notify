Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class EventosWeb

    ''' <summary>
    ''' Add To Evento Web
    ''' </summary>
    ''' <param name="Mensaje">Mensage</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function Add(ByVal Mensaje As String) As Boolean

        Dim sTipoDeEvento As String = "I"
        Dim sNumeroEPS As String = String.Empty
        Dim sTipoDeTransaccion As String = "200"
        Dim sCategoriaDeEvento As String = "SIS"
        Dim nWebAppID As Integer = My.Settings.ApplicationWebAppID
        Dim sLocalhost As String = DNSUtility.GetIPAddress(My.Computer.Name)

        Return Add(sTipoDeEvento, sNumeroEPS, sTipoDeTransaccion, Mensaje, sCategoriaDeEvento, nWebAppID, sLocalhost)

    End Function

    ''' <summary>
    ''' Add To Evento Web
    ''' </summary>
    ''' <param name="TipoDeEvento">Tipo de Evento</param>
    ''' <param name="NumeroEPS">Numero de EPS</param>
    ''' <param name="TipoDeTransaccion">Tipo de Transaccion</param>
    ''' <param name="Mensaje">Mensage</param>
    ''' <param name="CategoriaDeEvento">Categoria de Evento</param>
    ''' <param name="WebAppID">Aplicacion Web ID</param>
    ''' <param name="Localhost">Direccion de IP</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function Add(ByVal TipoDeEvento As String, _
                                         ByVal NumeroEPS As String, _
                                         ByVal TipoDeTransaccion As String, _
                                         ByVal Mensaje As String, _
                                         ByVal CategoriaDeEvento As String, _
                                         ByVal WebAppID As Integer, _
                                         ByVal Localhost As String) As Boolean

        Dim oConn As New SqlConnection(db.GetConnectionString)
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "dbo.proc_LOG_EVENTOS_WEB"

            '@EWT_CODIGO char(1)
            db.AddParameter(oCmd, "@EWT_CODIGO", TipoDeEvento, 1)

            ' @CTE_NUMERO_EPS char(12)
            db.AddParameter(oCmd, "@CTE_NUMERO_EPS", NumeroEPS, 12)

            '@TTR_CODIGO char(3)
            db.AddParameter(oCmd, "@TTR_CODIGO", TipoDeTransaccion, 3)

            ' @EVW_MENSAJE varchar(255)
            db.AddParameter(oCmd, "@EVW_MENSAJE", Mensaje, 255)

            '@EWC_CODIGO char(3)
            db.AddParameter(oCmd, "@EWC_CODIGO", CategoriaDeEvento, 3)

            '@PAG_ID int
            db.AddParameter(oCmd, "@PAG_ID", db.DataType.eInteger, db.Direction.eInput, WebAppID)

            '@EVW_DIRECCION_IP varchar(16)
            db.AddParameter(oCmd, "@EVW_DIRECCION_IP", Localhost, 16)
        End With

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

End Class
