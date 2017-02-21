Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class ewErrorHandler

    Private Const nMesajeMaxLenght As Integer = 195
    Private Const nDetallesMaxLenght As Integer = 3872

    ''' <summary>
    ''' Notifica Error
    ''' </summary>
    ''' <param name="Mensaje">Mensaje de error o exitoso</param>
    ''' <param name="Detalles">Detalle del error o si (n/a) que no aplica</param>
    ''' <param name="Severidad">3 = Alto, 2 = Medio, 1 = Bajo, 0 = Exitoso</param>
    ''' <remarks></remarks>
    Public Shared Sub NotificaError(ByVal Mensaje As String, ByVal Detalles As String, ByVal Severidad As Integer)
        Dim nMsgCount As Integer = 0
        Dim sSendMensaje As String = String.Empty
        Dim sSendDetalles As String = String.Empty
        Dim sCurDetalles As String = String.Empty
        Dim nCurDetallesLenght As Integer = 0
        Dim nProcesoId As Integer = My.Settings.ApplicationNotifyErrorProcessID

        If Not String.IsNullOrEmpty(Mensaje) Then
            Mensaje = Mensaje.Replace(Chr(34), Chr(34) & Chr(34))
            If Mensaje.Length > nMesajeMaxLenght Then Mensaje = Mensaje.Substring(0, nMesajeMaxLenght - 1)
        End If

        If Not String.IsNullOrEmpty(Detalles) Then
            Detalles = Detalles.Replace(Chr(34), Chr(34) & Chr(34))
        End If

        ' Enviar mensaje de la notificacíon dependiendo el tamaño
        If Not String.IsNullOrEmpty(Detalles) AndAlso Detalles.Length > nDetallesMaxLenght Then
            ' Enviar notificaciones multiples
            Try
                sCurDetalles = String.Copy(Detalles)

                While sCurDetalles.Length > 0
                    nMsgCount += 1

                    ' Mensaje de la notificacion a enviar
                    sSendMensaje = Mensaje & " [" & nMsgCount & "]"

                    ' Detalles de la notificacion actual
                    ' Si el tamaño del detalle actual es mayor que el tamaño maximo
                    If sCurDetalles.Length > nDetallesMaxLenght Then
                        ' Actualizar el detalle de la notificacion a enviar
                        sSendDetalles = sCurDetalles.Substring(0, nDetallesMaxLenght)

                        ' Buscar el tamaño actual del detalles
                        nCurDetallesLenght = sCurDetalles.Length - sSendDetalles.Length

                        ' Actualizar el detalle de la notificacion actual detalle restante
                        sCurDetalles = sCurDetalles.Substring(sSendDetalles.Length, nCurDetallesLenght)
                    Else
                        ' Detalles de la notificacion a enviar
                        sSendDetalles = sCurDetalles

                        ' Actualizar el detalle de la notificaion actual en blanco
                        sCurDetalles = String.Empty
                    End If

                    ' Enviar notificacion
                    Notificar(sSendMensaje.Trim, sSendDetalles.Trim, Severidad, nProcesoId)
                End While

            Catch ex As Exception
                ' Do nothing
            End Try
        Else
            ' Enviar notificacion simples
            Notificar(Mensaje, Detalles, Severidad, nProcesoId)
        End If
    End Sub

    ''' <summary>
    ''' Notifica Error por Consola
    ''' </summary>
    ''' <param name="Mensaje">Mensaje de error o exitoso</param>
    ''' <param name="Detalles">Detalle del error o si (n/a) que no aplica</param>
    ''' <param name="Severidad">3 = Alto, 2 = Medio, 1 = Bajo, 0 = Exitoso</param>
    ''' <param name="ProcesoId">ID de Aplicacion y/o Proceso</param>
    ''' <remarks></remarks>
    Private Shared Sub NotificaPorConsole(ByVal Mensaje As String, ByVal Detalles As String, ByVal Severidad As Integer, ByVal ProcesoId As Integer)
        Dim sPath As String = String.Empty
        Dim sAppPath As String = My.Settings.ApplicationNotifyErrorPath

        ' Replace quotes
        Mensaje = Mensaje.Replace(Chr(34), Chr(34) & Chr(34))
        Detalles = Detalles.Replace(Chr(34), Chr(34) & Chr(34))

        ' Set max length
        If nMesajeMaxLenght > 0 Then
            If Mensaje.Length > nMesajeMaxLenght Then Mensaje = Mensaje.Substring(0, nMesajeMaxLenght)
        End If
        If nDetallesMaxLenght > 0 Then
            If Detalles.Length > nDetallesMaxLenght Then Detalles = Detalles.Substring(0, nDetallesMaxLenght)
        End If

        ' Set path and parameters
        sPath = sAppPath & " " & ProcesoId.ToString & " " & Severidad.ToString & " " & _
                Chr(34) & Mensaje & Chr(34) & " " & _
                Chr(34) & Detalles & Chr(34)

        ' Send path and parameters
        Try
            Shell(sPath, AppWinStyle.Hide, False)
        Catch ex As Exception
            ' Do nothing
        End Try
    End Sub

    ''' <summary>
    ''' Notifica Error en la Base de Datos
    ''' </summary>
    ''' <param name="Mensaje">Mensaje de error o exitoso</param>
    ''' <param name="Detalles">Detalle del error o si (n/a) que no aplica</param>
    ''' <param name="Severidad">3 = Alto, 2 = Medio, 1 = Bajo, 0 = Exitoso</param>
    ''' <param name="ProcesoId">ID de Aplicacion y/o Proceso</param>
    ''' <remarks></remarks>
    Private Shared Function Notificar(ByVal Mensaje As String, ByVal Detalles As String, ByVal Severidad As Integer, ByVal ProcesoId As Integer) As Boolean

        'Dim nReturn As Integer = 0
        'Dim oConn As New SqlConnection(db.GetConnectionString("NotifyErrorConnectionString"))
        'Dim oCmd As New SqlCommand

        'With oCmd
        '    .Connection = oConn
        '    .CommandType = CommandType.StoredProcedure
        '    .CommandText = "dbo.proc_AlertasNotificar"

        '    ' @AlertaId INT
        '    db.AddParameter(oCmd, "@AlertaId", db.DataType.eInteger, db.Direction.eOutput, 0)

        '    ' @ProcesoId INT
        '    db.AddParameter(oCmd, "@ProcesoId", db.DataType.eInteger, db.Direction.eInput, ProcesoId)

        '    ' @FechaHora DATETIME
        '    db.AddParameter(oCmd, "@FechaHora", db.DataType.eDateTime, db.Direction.eInput, db.ewDateTimeSql())

        '    ' @Severidad INT
        '    db.AddParameter(oCmd, "@Severidad", db.DataType.eInteger, db.Direction.eInput, Severidad)

        '    ' @Mensaje VARCHAR(255)
        '    db.AddParameter(oCmd, "@Mensaje", Mensaje, 255)

        '    ' @Detalle VARCHAR(4000)
        '    db.AddParameter(oCmd, "@Detalle", Detalles, 4000)
        'End With

        'Try
        '    oConn.Open()
        '    oCmd.ExecuteNonQuery()

        '    ' To get a procedure's return value
        '    nReturn = CType(oCmd.Parameters("@AlertaId").Value, Integer)
        '    If nReturn > 0 Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'Catch ex As Exception
        '    Return False
        'End Try

    End Function


End Class
