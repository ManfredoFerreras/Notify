Imports System
Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Module MainModule

    Sub Main()

        ' Verificar y procesar la aplicacion
        CheckApplication()

    End Sub

#Region "Default Variables"

    Public Const Divider As String = "--------------------------------------------------------------------"

    Public bDebug As Boolean = False

    Public CurDateTime As Date = DateTime.Now()
    Public spacer As String = ""

    Public nRecCount As Integer = 0
    Public nRecCountSgda As Integer = 0
    Public nRecCountEra As Integer = 0

    Public sErrorMessage As String = String.Empty
    Public bInternetConnection As Boolean = False

    ' Application Info
    Public nApplicationID As Integer = My.Settings.ApplicationNotifyErrorProcessID
    Public sApplicationCompanyName As String = My.Application.Info.CompanyName
    Public sApplicationDescription As String = My.Application.Info.Description
    Public sApplicationVersion As String = My.Application.Info.Version.ToString


    ' Cantidad de días a procesar
    Public nAppProcesarDias As Integer = My.Settings.ApplicationProcesarDias

    ' Cantidad de registro a procesar
    Public nAppProcesarRegistros As Integer = My.Settings.ApplicationProcesarRegistros

    Public LogFile As StreamWriter

    Public sCondicionSeg As String = Nothing
    Public sCondicionTerc As String = Nothing

#End Region

    ''' <summary>
    ''' Check Application
    ''' </summary>
    ''' <remarks></remarks>
    Sub CheckApplication()

        Dim LogFilePath As String = My.Settings.LogFilePath

        ' Create log files
        CreateFile(LogFilePath)

        Const sPleaseWait As String = "Por favor espere un momento..."

        Dim CheckDatabase As Boolean = False
        Dim sMessage As String = String.Empty

        ' Desplegar mensaje
        PrintLine(String.Format("{0}: {1} - [Version {2}]", _
                                sApplicationCompanyName, _
                                sApplicationDescription, _
                                sApplicationVersion))

        ' Desplegar mensaje
        PrintDobleLine("(C)2011 Express Parcel Services, Int. Todos los derechos reservados.")

        ' Desplegar mensaje
        PrintDobleLine(String.Format("Comenzó: {0}", CurDateTime.ToString))

        ' Desplegar mensaje
        PrintDobleLine(sPleaseWait)

        ' Desplegar mensaje
        PrintLine("Conectando a la Base de Datos...")

        ' Verificar que exista una conección con la Base de Datos
        CheckDatabase = db.VerifyConnection("SELECT FECHA_SISTEMA = GETDATE()")

        If CheckDatabase Then

            ' Desplegar mensaje
            PrintDobleLine("Conexión a la Base de Datos exitosa!")

            If My.Settings.ApplicationCheckConnection = 1 Then

                ' Verificar conexión a internet
                PrintLine("Conectando a Internet...")

                If Internet.CheckConnection(My.Settings.ApplicationCheckConnectionURL) Then
                    sMessage = "Conexión a Internet exitosa!"
                    bInternetConnection = True
                Else
                    sMessage = "SE HA PRODUCIDO UN ERROR Y NO SE PUEDE CONECTAR A INTERNET."
                End If
                PrintDobleLine(sMessage)

            End If

            ' Procesar aplicación
            ProcessApplication()

            ' Desplegar mensaje
            PrintDobleLine(String.Format("Terminó: {0}", DateTime.Now.ToString))

            If Not bDebug Then

                ' Log de Eventos Web
                EventosWeb.Add(sApplicationDescription)

                ' Enviar reporte si exite direcciones de correos electronicas invalidas
                If nRecCount > 0 Then
                    sMessage = "Proceso ejecutado con existo"
                Else
                    sMessage = "No existen elementos para procesar"
                End If
                PrintDobleLine(sMessage)

                PrintLine("Enviando alerta del sistema...")

                ' Close log files before reading file
                LogFile.Close()

                ewErrorHandler.NotificaError(sMessage, ReadFile(LogFilePath), 0)
            Else
                ' Close log files
                LogFile.Close()
            End If

        Else

            sMessage = "SE HA PRODUCIDO UN ERROR Y NO SE PUEDE CONECTAR A LA BASE DE DATOS."
            PrintDobleLine(sMessage)
            ewErrorHandler.NotificaError("Se ha producido un error (DB)", sMessage, 3)

            ' Close log files
            LogFile.Close()

        End If

        If bDebug Then
            Console.ReadLine()
        End If

    End Sub

    ''' <summary>
    ''' Process Main Application
    ''' </summary>
    ''' <remarks></remarks>
    Sub ProcessApplication()

        ' PROCESAR MENSAJES POR TEMPLATES
        PrintLine(Divider)
        PrintDobleLine("*** PROCESANDO MENSAJES DE NOTIFICACIONES POR TEMPLATES")
        ProcessApplicationByTemplates()
        ' ProcessApplicationByTemplatesDa()
        'ProcessApplicationByTemplatesEra()

        ' ASIGNACION DE MENSAJES
        PrintLine(Divider)
        PrintDobleLine("*** ASIGNACION DE MENSAJES A REPRESENTANTES DE SERVICIOS")
        AsignacionDeMensajes(nAppProcesarDias)

        'Console.ReadKey()

    End Sub

    ''' <summary>
    ''' Procesar Aplicacion por Templates
    ''' </summary>
    ''' <remarks></remarks>
    Sub ProcessApplicationByTemplates()

        ' Buscar Encabezado de la Empresa en la Base de Datos
        Dim sCompanyHeaderInfo As String = String.Empty
        Dim sCurCompanyHeaderInfor As String = String.Empty
        sCompanyHeaderInfo = OpeConfiguracion.GetHeader()

        Dim nTemplateID As Integer = 0 ' ID del E-mail Template
        Dim sTemplateDescripcion As String = String.Empty ' Nombre/Descripcion del E-mail Template

        Dim bIsHtml As Boolean = False ' Formato del E-mail Template es HTML
        Dim sEmailFormat As String = String.Empty ' Formato del E-mail Template
        Dim sEmailTo As String = String.Empty ' Direccion de E-mail (Opcional)
        Dim sEmailCcList As String = String.Empty ' Lista de Direcciones de E-mail de Copia (Opcional, Separador por Coma)
        Dim sEmailBccList As String = String.Empty ' Lista de Direcciones de E-mail BCC (Opcional, Separador por Coma)
        Dim sEmailFrom As String = String.Empty ' Direccion de E-mail de Envio (De)
        Dim sEmailFromName As String = String.Empty ' Nombre de Envio
        Dim sEmailReplyTo As String = String.Empty ' Direccion de E-mail de Respuesta
        Dim sEmailSubject As String = String.Empty ' Asunto del Mensaje

        Dim sEmailTemplateHeader As String = String.Empty ' Encabezado del Mensaje
        Dim sEmailTemplateBody As String = String.Empty ' Cuerpo del Mensaje
        Dim sEmailTemplateFooter As String = String.Empty ' Pie de Pagina del Mensaje

        Dim sEmailBody As String = String.Empty ' Mensaje completo

        Dim sImportarMensajes As String = String.Empty ' Importar mensajes de notificaciones
        Dim sActualizaMensajes As String = String.Empty ' Actualizar mensajes de notificaciones sin responder
        Dim sNotificaRepresentantes As String = String.Empty ' Notificar Representantes de Servicios
        Dim dFechaCorrida As Object = Nothing ' Fecha de la ultima corrida del proceso

        Dim sEjecutarInicio As String = String.Empty ' Nombre de Stored Procedure a ejecutar antes de generar la notificacion
        Dim sEjecutarFinalizar As String = String.Empty ' Nombre de Stored Procedure a ejecutar despues de generar la notificacion
        Dim sEjecutarTabla As String = String.Empty ' Nombre de la tabla principal para procesar los datos


        Try
            Dim ds As New DataSet
            ds = GetWebMailTemplateDataSet()

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                nTemplateID = db.ewToInteger(dr("TPL_EMAIL_ID"))
                sTemplateDescripcion = db.ewToString(dr("TPL_DESCRIPCION"))

                sEmailFormat = db.ewToString(dr("TPL_EMAIL_FORMAT"))
                sEmailTo = db.ewToString(dr("TPL_EMAIL_TO"))
                sEmailCcList = db.ewToString(dr("TPL_EMAIL_CC_LIST"))
                sEmailBccList = db.ewToString(dr("TPL_EMAIL_BCC_LIST"))
                sEmailFrom = db.ewToString(dr("TPL_EMAIL_FROM"))
                sEmailFromName = db.ewToString(dr("TPL_FROMNAME"))
                sEmailReplyTo = db.ewToString(dr("TPL_EMAIL_REPLYTO"))
                sEmailSubject = db.ewToString(dr("TPL_SUBJECT"))

                sEmailTemplateHeader = db.ewToString(dr("TPL_HEADER"))
                sEmailTemplateBody = db.ewToString(dr("TPL_BODY"))
                sEmailTemplateFooter = db.ewToString(dr("TPL_FOOTER"))

                sImportarMensajes = db.ewToStringUpper(dr("TPL_IMPORTAR_MENSAJES"))
                sActualizaMensajes = db.ewToStringUpper(dr("TPL_ACTUALIZA_MENSAJES"))
                sNotificaRepresentantes = db.ewToStringUpper(dr("TPL_NOTIFICA_REPRESENTANTES"))

                sEjecutarInicio = db.ewToString(dr("TPL_EJECUTAR_INICIO"))
                sEjecutarFinalizar = db.ewToString(dr("TPL_EJECUTAR_FINALIZAR"))
                sEjecutarTabla = db.ewToString(dr("TPL_EJECUTAR_TABLA"))

                ' Valores por defecto
                If String.IsNullOrEmpty(sImportarMensajes) Then sImportarMensajes = "N"
                If String.IsNullOrEmpty(sActualizaMensajes) Then sActualizaMensajes = "N"
                If String.IsNullOrEmpty(sNotificaRepresentantes) Then sNotificaRepresentantes = "N"

                ' Crear Email Template
                If String.IsNullOrEmpty(sEmailFormat) Then sEmailFormat = "TEXT"

                ' Validar si el mensaje es HTML
                bIsHtml = IIf(sEmailFormat.ToUpper.Trim = "HTML", True, False)

                If bIsHtml = True Then
                    sCurCompanyHeaderInfor = sCompanyHeaderInfo.Replace(Chr(10), "<br />")
                Else
                    sCurCompanyHeaderInfor = sCompanyHeaderInfo
                End If

                ' Crear Email Template
                sEmailBody = String.Empty
                If Not String.IsNullOrEmpty(sEmailTemplateHeader) Then sEmailBody += sEmailTemplateHeader
                If Not String.IsNullOrEmpty(sEmailTemplateBody) Then sEmailBody += sEmailTemplateBody
                If Not String.IsNullOrEmpty(sEmailTemplateFooter) Then sEmailBody += sEmailTemplateFooter

                ' Formatear template a enviar
                '------------------------------------------------------------------------
                FormatTemplateText(sEmailBody, "CompanyHeaderInfo", sCurCompanyHeaderInfor)
                FormatTemplateText(sEmailBody, "COMPANY_HEADER_INFO", sCurCompanyHeaderInfor)
                FormatTemplateText(sEmailBody, "PAGE_TITLE", "<span style=""color:#FFFFFF;font:bold 16px Arial,Helvetica,sans-serif;padding:5px;"">&nbsp;</span>")
                FormatTemplateText(sEmailBody, "TPL_DESCRIPCION", sTemplateDescripcion)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_TO", sEmailTo)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_CC", sEmailCcList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_CC_LIST", sEmailCcList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_BCC", sEmailBccList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_BCC_LIST", sEmailBccList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_FROM", sEmailFrom)
                FormatTemplateText(sEmailBody, "TPL_FROMNAME", sEmailFromName)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_REPLYTO", sEmailReplyTo)
                FormatTemplateText(sEmailBody, "TPL_SUBJECT", sEmailSubject)
                FormatTemplateText(sEmailBody, "TPL_YEAR", CurDateTime.Year)

                '------------------------------------------------------------------------
                PrintLine(Divider)
                PrintDobleLine(String.Format("{0} - {1}", nTemplateID, sTemplateDescripcion))

                ' IMPORTAR MENSAJES DE NOTIFICACIONES POR TEMPLATES
                '------------------------------------------------------------------------
                If sImportarMensajes = "S" Then ImportMessages(nTemplateID, "", "")

                ' PROCESAR MENSAJES DE NOTIFICACIONES POR TEMPLATES
                '------------------------------------------------------------------------
                Select Case sEjecutarTabla.ToUpper.Trim
                    Case "BULTOS"
                        ProcessMensajesNotificacionesBultos(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)
                        'ProcessMensajesNotificacionesBultosSgda(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)
                        'ProcessMensajesNotificacionesBultosEra(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)
                        ' Procesar mensajes de notificaciones sin responder por templates
                        If sActualizaMensajes = "S" Then UpdateWebMailNotificacionesSinResponder(nTemplateID, nAppProcesarDias)

                    Case "CLIENTES_CREDITO"
                        ProcessMensajesNotificacionesClientesCredito(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)

                    Case "CLIENTE_NUEVO"
                        ProcessMensajesNotiClientesNuevo(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)

                    Case Else ' CLIENTES
                        ProcessMensajesNotificacionesClientes(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)
                End Select

                ' ACTUALIZAR CORRIDA POR TEMPLATES
                '------------------------------------------------------------------------
                PrintDobleLine("- Actualizando corrida...")
                UpdateWebMailTemplatesCorrida(nTemplateID)
            Next
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessApplicationByTemplates(): " & sEx, 2)
        End Try

    End Sub
    Sub ProcessApplicationByTemplatesDa()

        ' Buscar Encabezado de la Empresa en la Base de Datos
        Dim sCompanyHeaderInfo As String = String.Empty
        Dim sCurCompanyHeaderInfor As String = String.Empty
        sCompanyHeaderInfo = OpeConfiguracion.GetHeader()

        Dim nTemplateID As Integer = 0 ' ID del E-mail Template
        Dim sTemplateDescripcion As String = String.Empty ' Nombre/Descripcion del E-mail Template

        Dim bIsHtml As Boolean = False ' Formato del E-mail Template es HTML
        Dim sEmailFormat As String = String.Empty ' Formato del E-mail Template
        Dim sEmailTo As String = String.Empty ' Direccion de E-mail (Opcional)
        Dim sEmailCcList As String = String.Empty ' Lista de Direcciones de E-mail de Copia (Opcional, Separador por Coma)
        Dim sEmailBccList As String = String.Empty ' Lista de Direcciones de E-mail BCC (Opcional, Separador por Coma)
        Dim sEmailFrom As String = String.Empty ' Direccion de E-mail de Envio (De)
        Dim sEmailFromName As String = String.Empty ' Nombre de Envio
        Dim sEmailReplyTo As String = String.Empty ' Direccion de E-mail de Respuesta
        Dim sEmailSubject As String = String.Empty ' Asunto del Mensaje

        Dim sEmailTemplateHeader As String = String.Empty ' Encabezado del Mensaje
        Dim sEmailTemplateBody As String = String.Empty ' Cuerpo del Mensaje
        Dim sEmailTemplateFooter As String = String.Empty ' Pie de Pagina del Mensaje

        Dim sEmailBody As String = String.Empty ' Mensaje completo

        Dim sImportarMensajes As String = String.Empty ' Importar mensajes de notificaciones
        Dim sActualizaMensajes As String = String.Empty ' Actualizar mensajes de notificaciones sin responder
        Dim sNotificaRepresentantes As String = String.Empty ' Notificar Representantes de Servicios
        Dim dFechaCorrida As Object = Nothing ' Fecha de la ultima corrida del proceso

        Dim sEjecutarInicio As String = String.Empty ' Nombre de Stored Procedure a ejecutar antes de generar la notificacion
        Dim sEjecutarFinalizar As String = String.Empty ' Nombre de Stored Procedure a ejecutar despues de generar la notificacion
        Dim sEjecutarTabla As String = String.Empty ' Nombre de la tabla principal para procesar los datos


        Try
            Dim ds As New DataSet
            ds = GetWebMailTemplateDataSet()

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                nTemplateID = db.ewToInteger(dr("TPL_EMAIL_ID"))
                sTemplateDescripcion = db.ewToString(dr("TPL_DESCRIPCION"))

                sEmailFormat = db.ewToString(dr("TPL_EMAIL_FORMAT"))
                sEmailTo = db.ewToString(dr("TPL_EMAIL_TO"))
                sEmailCcList = db.ewToString(dr("TPL_EMAIL_CC_LIST"))
                sEmailBccList = db.ewToString(dr("TPL_EMAIL_BCC_LIST"))
                sEmailFrom = db.ewToString(dr("TPL_EMAIL_FROM"))
                sEmailFromName = db.ewToString(dr("TPL_FROMNAME"))
                sEmailReplyTo = db.ewToString(dr("TPL_EMAIL_REPLYTO"))
                sEmailSubject = db.ewToString(dr("TPL_SUBJECT"))

                sEmailTemplateHeader = db.ewToString(dr("TPL_HEADER"))
                sEmailTemplateBody = db.ewToString(dr("TPL_BODY"))
                sEmailTemplateFooter = db.ewToString(dr("TPL_FOOTER"))

                sImportarMensajes = db.ewToStringUpper(dr("TPL_IMPORTAR_MENSAJES"))
                sActualizaMensajes = db.ewToStringUpper(dr("TPL_ACTUALIZA_MENSAJES"))
                sNotificaRepresentantes = db.ewToStringUpper(dr("TPL_NOTIFICA_REPRESENTANTES"))

                sEjecutarInicio = db.ewToString(dr("TPL_EJECUTAR_INICIO"))
                sEjecutarFinalizar = db.ewToString(dr("TPL_EJECUTAR_FINALIZAR"))
                sEjecutarTabla = db.ewToString(dr("TPL_EJECUTAR_TABLA"))

                ' Valores por defecto
                If String.IsNullOrEmpty(sImportarMensajes) Then sImportarMensajes = "N"
                If String.IsNullOrEmpty(sActualizaMensajes) Then sActualizaMensajes = "N"
                If String.IsNullOrEmpty(sNotificaRepresentantes) Then sNotificaRepresentantes = "N"

                ' Crear Email Template
                If String.IsNullOrEmpty(sEmailFormat) Then sEmailFormat = "TEXT"

                ' Validar si el mensaje es HTML
                bIsHtml = IIf(sEmailFormat.ToUpper.Trim = "HTML", True, False)

                If bIsHtml = True Then
                    sCurCompanyHeaderInfor = sCompanyHeaderInfo.Replace(Chr(10), "<br />")
                Else
                    sCurCompanyHeaderInfor = sCompanyHeaderInfo
                End If

                ' Crear Email Template
                sEmailBody = String.Empty
                If Not String.IsNullOrEmpty(sEmailTemplateHeader) Then sEmailBody += sEmailTemplateHeader
                If Not String.IsNullOrEmpty(sEmailTemplateBody) Then sEmailBody += sEmailTemplateBody
                If Not String.IsNullOrEmpty(sEmailTemplateFooter) Then sEmailBody += sEmailTemplateFooter

                ' Formatear template a enviar
                '------------------------------------------------------------------------
                FormatTemplateText(sEmailBody, "CompanyHeaderInfo", sCurCompanyHeaderInfor)
                FormatTemplateText(sEmailBody, "COMPANY_HEADER_INFO", sCurCompanyHeaderInfor)
                FormatTemplateText(sEmailBody, "PAGE_TITLE", "<span style=""color:#FFFFFF;font:bold 16px Arial,Helvetica,sans-serif;padding:5px;"">&nbsp;</span>")
                FormatTemplateText(sEmailBody, "TPL_DESCRIPCION", sTemplateDescripcion)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_TO", sEmailTo)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_CC", sEmailCcList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_CC_LIST", sEmailCcList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_BCC", sEmailBccList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_BCC_LIST", sEmailBccList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_FROM", sEmailFrom)
                FormatTemplateText(sEmailBody, "TPL_FROMNAME", sEmailFromName)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_REPLYTO", sEmailReplyTo)
                FormatTemplateText(sEmailBody, "TPL_SUBJECT", sEmailSubject)
                FormatTemplateText(sEmailBody, "TPL_YEAR", CurDateTime.Year)

                '------------------------------------------------------------------------
                PrintLine(Divider)
                PrintDobleLine(String.Format("{0} - {1}", nTemplateID, sTemplateDescripcion))

                ' IMPORTAR MENSAJES DE NOTIFICACIONES POR TEMPLATES
                '------------------------------------------------------------------------
                If sImportarMensajes = "S" Then ImportMessages(nTemplateID, "-2", "")

                ' PROCESAR MENSAJES DE NOTIFICACIONES POR TEMPLATES
                '------------------------------------------------------------------------
                Select Case sEjecutarTabla.ToUpper.Trim
                    Case "BULTOS"
                        ProcessMensajesNotificacionesBultosSgda(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)

                        ' Procesar mensajes de notificaciones sin responder por templates
                        If sActualizaMensajes = "S" Then UpdateWebMailNotificacionesSinResponderSgda(nTemplateID, nAppProcesarDias)

                    
                End Select

                ' ACTUALIZAR CORRIDA POR TEMPLATES
                '------------------------------------------------------------------------
                PrintDobleLine("- Actualizando corrida...")
                UpdateWebMailTemplatesCorrida(nTemplateID)
            Next
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessApplicationByTemplates(): " & sEx, 2)
        End Try

    End Sub
    Sub ProcessApplicationByTemplatesEra()

        ' Buscar Encabezado de la Empresa en la Base de Datos
        Dim sCompanyHeaderInfo As String = String.Empty
        Dim sCurCompanyHeaderInfor As String = String.Empty
        sCompanyHeaderInfo = OpeConfiguracion.GetHeader()

        Dim nTemplateID As Integer = 0 ' ID del E-mail Template
        Dim sTemplateDescripcion As String = String.Empty ' Nombre/Descripcion del E-mail Template

        Dim bIsHtml As Boolean = False ' Formato del E-mail Template es HTML
        Dim sEmailFormat As String = String.Empty ' Formato del E-mail Template
        Dim sEmailTo As String = String.Empty ' Direccion de E-mail (Opcional)
        Dim sEmailCcList As String = String.Empty ' Lista de Direcciones de E-mail de Copia (Opcional, Separador por Coma)
        Dim sEmailBccList As String = String.Empty ' Lista de Direcciones de E-mail BCC (Opcional, Separador por Coma)
        Dim sEmailFrom As String = String.Empty ' Direccion de E-mail de Envio (De)
        Dim sEmailFromName As String = String.Empty ' Nombre de Envio
        Dim sEmailReplyTo As String = String.Empty ' Direccion de E-mail de Respuesta
        Dim sEmailSubject As String = String.Empty ' Asunto del Mensaje

        Dim sEmailTemplateHeader As String = String.Empty ' Encabezado del Mensaje
        Dim sEmailTemplateBody As String = String.Empty ' Cuerpo del Mensaje
        Dim sEmailTemplateFooter As String = String.Empty ' Pie de Pagina del Mensaje

        Dim sEmailBody As String = String.Empty ' Mensaje completo

        Dim sImportarMensajes As String = String.Empty ' Importar mensajes de notificaciones
        Dim sActualizaMensajes As String = String.Empty ' Actualizar mensajes de notificaciones sin responder
        Dim sNotificaRepresentantes As String = String.Empty ' Notificar Representantes de Servicios
        Dim dFechaCorrida As Object = Nothing ' Fecha de la ultima corrida del proceso

        Dim sEjecutarInicio As String = String.Empty ' Nombre de Stored Procedure a ejecutar antes de generar la notificacion
        Dim sEjecutarFinalizar As String = String.Empty ' Nombre de Stored Procedure a ejecutar despues de generar la notificacion
        Dim sEjecutarTabla As String = String.Empty ' Nombre de la tabla principal para procesar los datos


        Try
            Dim ds As New DataSet
            ds = GetWebMailTemplateDataSet()

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                nTemplateID = db.ewToInteger(dr("TPL_EMAIL_ID"))
                sTemplateDescripcion = db.ewToString(dr("TPL_DESCRIPCION"))

                sEmailFormat = db.ewToString(dr("TPL_EMAIL_FORMAT"))
                sEmailTo = db.ewToString(dr("TPL_EMAIL_TO"))
                sEmailCcList = db.ewToString(dr("TPL_EMAIL_CC_LIST"))
                sEmailBccList = db.ewToString(dr("TPL_EMAIL_BCC_LIST"))
                sEmailFrom = db.ewToString(dr("TPL_EMAIL_FROM"))
                sEmailFromName = db.ewToString(dr("TPL_FROMNAME"))
                sEmailReplyTo = db.ewToString(dr("TPL_EMAIL_REPLYTO"))
                sEmailSubject = db.ewToString(dr("TPL_SUBJECT"))

                sEmailTemplateHeader = db.ewToString(dr("TPL_HEADER"))
                sEmailTemplateBody = db.ewToString(dr("TPL_BODY"))
                sEmailTemplateFooter = db.ewToString(dr("TPL_FOOTER"))

                sImportarMensajes = db.ewToStringUpper(dr("TPL_IMPORTAR_MENSAJES"))
                sActualizaMensajes = db.ewToStringUpper(dr("TPL_ACTUALIZA_MENSAJES"))
                sNotificaRepresentantes = db.ewToStringUpper(dr("TPL_NOTIFICA_REPRESENTANTES"))

                sEjecutarInicio = db.ewToString(dr("TPL_EJECUTAR_INICIO"))
                sEjecutarFinalizar = db.ewToString(dr("TPL_EJECUTAR_FINALIZAR"))
                sEjecutarTabla = db.ewToString(dr("TPL_EJECUTAR_TABLA"))

                ' Valores por defecto
                If String.IsNullOrEmpty(sImportarMensajes) Then sImportarMensajes = "N"
                If String.IsNullOrEmpty(sActualizaMensajes) Then sActualizaMensajes = "N"
                If String.IsNullOrEmpty(sNotificaRepresentantes) Then sNotificaRepresentantes = "N"

                ' Crear Email Template
                If String.IsNullOrEmpty(sEmailFormat) Then sEmailFormat = "TEXT"

                ' Validar si el mensaje es HTML
                bIsHtml = IIf(sEmailFormat.ToUpper.Trim = "HTML", True, False)

                If bIsHtml = True Then
                    sCurCompanyHeaderInfor = sCompanyHeaderInfo.Replace(Chr(10), "<br />")
                Else
                    sCurCompanyHeaderInfor = sCompanyHeaderInfo
                End If

                ' Crear Email Template
                sEmailBody = String.Empty
                If Not String.IsNullOrEmpty(sEmailTemplateHeader) Then sEmailBody += sEmailTemplateHeader
                If Not String.IsNullOrEmpty(sEmailTemplateBody) Then sEmailBody += sEmailTemplateBody
                If Not String.IsNullOrEmpty(sEmailTemplateFooter) Then sEmailBody += sEmailTemplateFooter

                ' Formatear template a enviar
                '------------------------------------------------------------------------
                FormatTemplateText(sEmailBody, "CompanyHeaderInfo", sCurCompanyHeaderInfor)
                FormatTemplateText(sEmailBody, "COMPANY_HEADER_INFO", sCurCompanyHeaderInfor)
                FormatTemplateText(sEmailBody, "PAGE_TITLE", "<span style=""color:#FFFFFF;font:bold 16px Arial,Helvetica,sans-serif;padding:5px;"">&nbsp;</span>")
                FormatTemplateText(sEmailBody, "TPL_DESCRIPCION", sTemplateDescripcion)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_TO", sEmailTo)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_CC", sEmailCcList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_CC_LIST", sEmailCcList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_BCC", sEmailBccList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_BCC_LIST", sEmailBccList)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_FROM", sEmailFrom)
                FormatTemplateText(sEmailBody, "TPL_FROMNAME", sEmailFromName)
                FormatTemplateText(sEmailBody, "TPL_EMAIL_REPLYTO", sEmailReplyTo)
                FormatTemplateText(sEmailBody, "TPL_SUBJECT", sEmailSubject)
                FormatTemplateText(sEmailBody, "TPL_YEAR", CurDateTime.Year)

                '------------------------------------------------------------------------
                PrintLine(Divider)
                PrintDobleLine(String.Format("{0} - {1}", nTemplateID, sTemplateDescripcion))

                ' IMPORTAR MENSAJES DE NOTIFICACIONES POR TEMPLATES
                '------------------------------------------------------------------------
                If sImportarMensajes = "S" Then ImportMessages(nTemplateID, "", "-3")

                ' PROCESAR MENSAJES DE NOTIFICACIONES POR TEMPLATES
                '------------------------------------------------------------------------
                Select Case sEjecutarTabla.ToUpper.Trim
                    Case "BULTOS"
                        ProcessMensajesNotificacionesBultosEra(nTemplateID, sEmailFormat, sEmailTo, sEmailCcList, sEmailBccList, sEmailFrom, sEmailFromName, sEmailReplyTo, sEmailSubject, sEmailBody, sNotificaRepresentantes)

                        ' Procesar mensajes de notificaciones sin responder por templates
                        If sActualizaMensajes = "S" Then UpdateWebMailNotificacionesSinResponder(nTemplateID, nAppProcesarDias)

                   
                End Select

                ' ACTUALIZAR CORRIDA POR TEMPLATES
                '------------------------------------------------------------------------
                PrintDobleLine("- Actualizando corrida...")
                UpdateWebMailTemplatesCorrida(nTemplateID)
            Next
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessApplicationByTemplates(): " & sEx, 2)
        End Try

    End Sub

#Region "Procesar Notificaciones"

    ''' <summary>
    ''' Procesar Mensajes de Notificaciones para Paquetes o Bultos
    ''' </summary>
    ''' <param name="TemplateID">ID del E-mail Template</param>
    ''' <param name="EmailFormat">Formato del E-mail Template</param>
    ''' <param name="EmailTo">Direccion de E-mail (Opcional)</param>
    ''' <param name="EmailCcList">Lista de Direcciones de E-mail de Copia (Opcional, Separador por Coma)</param>
    ''' <param name="EmailBccList">Lista de Direcciones de E-mail BCC (Opcional, Separador por Coma)</param>
    ''' <param name="EmailFrom">Direccion de E-mail de Envio (De)</param>
    ''' <param name="EmailFromName">Nombre de Envio</param>
    ''' <param name="EmailReplyTo">Direccion de E-mail de Respuesta</param>
    ''' <param name="EmailSubject">Asunto del Mensaje</param>
    ''' <param name="EmailBody">Mensaje completo</param>
    ''' <param name="NotificaRepresentantes">Notificar Representantes de Servicios</param>
    ''' <remarks></remarks>
    Sub ProcessMensajesNotificacionesBultos(ByVal TemplateID As Integer, _
                                            ByVal EmailFormat As String, _
                                            ByVal EmailTo As String, _
                                            ByVal EmailCcList As String, _
                                            ByVal EmailBccList As String, _
                                            ByVal EmailFrom As String, _
                                            ByVal EmailFromName As String, _
                                            ByVal EmailReplyTo As String, _
                                            ByVal EmailSubject As String, _
                                            ByVal EmailBody As String, _
                                            ByVal NotificaRepresentantes As String)

        Dim sCurName As String = "Paquete o Bulto"
        Dim bSendEmail As Boolean = False

        Dim sTmp As String = String.Empty
        Dim sMessage As String = String.Empty

        Dim sCurEmailTo As String = String.Empty
        Dim sCurEmailSubject As String = String.Empty
        Dim sCurEmailBoby As String = String.Empty
        Dim sCurEmailAsesor As String = String.Empty

        Dim nMensajeID As String = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim sMensajeFechaCreado As String = Nothing

        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sTipoClienteCodigo As String = Nothing
        Dim sSubTipoClienteCodigo As String = Nothing
        Dim sClienteEstado As String = Nothing
        Dim sClienteEstatus As String = Nothing
        Dim sCedula As String = Nothing
        Dim sRNC As String = Nothing
        Dim sPasaporte As String = Nothing
        Dim sClienteEmail As String = Nothing
        Dim sEnviarEmail As String = Nothing
        Dim nCompaniaID As String = Nothing
        Dim sSucursalCodigo As String = Nothing
        Dim sAgenciaCodigo As String = Nothing
        Dim sAgenciaEmail As String = Nothing
        Dim sOficialCodigo As String = Nothing
        Dim sOficialNombre As String = Nothing
        Dim sAsesorCodigo As String = Nothing
        Dim sAsesorNombre As String = Nothing
        Dim sAsesorEmail As String = Nothing
        Dim sClienteTieneCredito As String = Nothing
        Dim sClienteTieneCorrespondecia As String = Nothing

        Dim nBultoNumero As String = Nothing
        Dim sCodigoBarra As String = Nothing
        Dim sTrackingNumber As String = Nothing
        Dim sGuiaHija As String = Nothing
        Dim sGuiaMadre As String = Nothing
        Dim sManifiesto As String = Nothing
        Dim sServicioCodigo As String = Nothing
        Dim sServicio As String = Nothing
        Dim sOrigen As String = Nothing
        Dim sOrdenNo As String = Nothing
        Dim sFacturaNo As String = Nothing
        Dim sFechaRecepcion As String = Nothing
        Dim nPiezas As String = Nothing
        Dim nPeso As String = Nothing
        Dim sContenido As String = Nothing
        Dim sRemitente As String = Nothing
        Dim sDestinatario As String = Nothing
        Dim sLocalizacion As String = Nothing
        Dim sCondicionPrimera As String = Nothing
        Dim bDocumentoDisponible As Boolean = Nothing
        Dim sDocumentoReferencia As String = Nothing

        ' Validar si el mensaje es HTML
        Dim bIsHtml As Boolean = IIf(EmailFormat.ToUpper.Trim = "HTML", True, False)

        ' Desplegar mensaje
        PrintDobleLine(String.Format("- Buscando notificaciones {0} para enviar, por favor espere...", sCurName))

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetMensajesNotificacionesBultosDataSet(TemplateID, nAppProcesarDias)

            ' Desplegar mensaje
            PrintDobleLine(String.Format("- Procesando notificaciones de {0} para enviar...", sCurName))

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                ' Enviar correo
                bSendEmail = True

                ' Buscar los datos de la notificacion
                '------------------------------------
                nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
                sMensajeGUID = db.ewToString(dr("WMM_MENSAJE_GUID"))
                sMensajeFechaCreado = db.ewToString(dr("WMM_FECHA_CREADO"))

                sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
                sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))
                sTipoClienteCodigo = db.ewToStringUpper(dr("CTE_TIPO"))
                sSubTipoClienteCodigo = db.ewToStringUpper(dr("STC_CODIGO"))
                sClienteEstado = db.ewToStringUpper(dr("CTE_ESTADO"))
                sClienteEstatus = db.ewToStringUpper(dr("ESTATUS"))
                sCedula = db.ewToString(dr("CTE_CEDULA"))
                sRNC = db.ewToString(dr("CTE_RNC"))
                sPasaporte = db.ewToString(dr("CTE_PASAPORTE"))
                sClienteEmail = db.ewToStringLower(dr("CTE_EMAIL"))
                sEnviarEmail = db.ewToStringUpper(dr("CTE_ENVIAR_EMAIL"))
                nCompaniaID = db.ewToString(dr("COM_CODIGO"))
                sSucursalCodigo = db.ewToStringUpper(dr("SUC_CODIGO"))
                sAgenciaCodigo = db.ewToStringUpper(dr("AGE_CODIGO"))
                sAgenciaEmail = db.ewToStringLower(dr("AGENCIA_EMAIL"))
                sOficialCodigo = db.ewToStringUpper(dr("CTE_VENDEDOR"))
                sOficialNombre = db.ewToStringUpper(dr("OFICIAL"))
                sAsesorCodigo = db.ewToStringUpper(dr("RES_CODIGO"))
                sAsesorNombre = db.ewToStringUpper(dr("ASESOR"))
                sAsesorEmail = db.ewToStringLower(dr("ASESOR_EMAIL"))
                sClienteTieneCredito = db.ewToStringUpper(dr("CTE_CREDITO"))
                sClienteTieneCorrespondecia = db.ewToStringUpper(dr("CTE_CORRESPONDENCIA"))

                nBultoNumero = db.ewToString(dr("BLT_NUMERO"))
                sCodigoBarra = db.ewToStringUpper(dr("BLT_CODIGO_BARRA"))
                sTrackingNumber = db.ewToStringUpper(dr("BLT_TRACKING_NUMBER"))
                sGuiaHija = db.ewToStringUpper(dr("BLT_GUIA_HIJA"))
                sGuiaMadre = db.ewToStringUpper(dr("MAN_GUIA"))
                sManifiesto = db.ewToStringUpper(dr("MAN_MANIFIESTO"))
                sServicioCodigo = db.ewToStringUpper(dr("PRO_CODIGO"))
                sServicio = db.ewToStringUpper(dr("PRO_DESCRIPCION"))
                sOrigen = db.ewToStringUpper(dr("ORIGEN"))
                sOrdenNo = db.ewToStringUpper(dr("BLT_PONUMBER"))
                sFacturaNo = db.ewToStringUpper(dr("BLT_FACTURA_SUPLIDOR"))
                sFechaRecepcion = db.ewToString(dr("BLT_FECHA_RECEPCION"))
                nPiezas = db.ewToString(dr("BLT_PIEZAS"))
                nPeso = db.ewToString(dr("BLT_PESO"))
                sContenido = db.ewToStringUpper(dr("CONTENIDO"))
                sRemitente = db.ewToStringUpper(dr("REMITENTE"))
                sDestinatario = db.ewToStringUpper(dr("DESTINATARIO"))
                sLocalizacion = db.ewToStringUpper(dr("LOCALIZACION"))
                ' sCondicion = db.ewToStringUpper(dr("CONDICION"))

                Dim condiciones = GetCondiciones(sCodigoBarra, sTrackingNumber)

                sCondicionPrimera = db.ewToStringUpper(condiciones(0))


                'If (condiciones(1) <> Nothing) Then
                '    ' sCondicionSeg = db.ewToStringUpper(condiciones(1))
                '    ProcessApplicationByTemplatesDa()
                'End If

                'If (condiciones(2) <> Nothing) Then
                '    ' sCondicionTerc = db.ewToStringUpper(condiciones(2))
                '    ProcessApplicationByTemplatesEra()
                'End If

                bDocumentoDisponible = db.ewToBool(dr("DOC_DISPONIBLE"))
                sDocumentoReferencia = db.ewToStringNullable(dr("DOC_REFERENCIA"))

                ' SetUp Cliente
                ' -------------
                SetUpClienteOficial(sOficialCodigo, sOficialNombre, sSucursalCodigo, sAgenciaCodigo)
                SetUpClienteAsesor(sAsesorCodigo, sAsesorNombre, sAsesorEmail, sSucursalCodigo, sAgenciaCodigo)
                If NotificaRepresentantes = "S" Then sCurEmailAsesor = sAsesorEmail ' Enviar notificacion al asesor de la cuenta

                ' HTML Encoding
                '--------------
                If bIsHtml = True Then
                    If Not String.IsNullOrEmpty(sContenido) Then sContenido = ew_EncodeText(sContenido)
                    If Not String.IsNullOrEmpty(sRemitente) Then sRemitente = ew_EncodeText(sRemitente)
                    If Not String.IsNullOrEmpty(sDestinatario) Then sDestinatario = ew_EncodeText(sDestinatario)
                End If

                ' Set Default Values
                '-------------------
                If String.IsNullOrEmpty(sContenido) = True Then sContenido = "N/A"
                If String.IsNullOrEmpty(sRemitente) = True Then sRemitente = "N/A"
                If String.IsNullOrEmpty(sDestinatario) = True Then sDestinatario = "N/A"

                ' Buscar dirección de envio del mensaje
                '------------------------------------------------------------------------
                sCurEmailTo = ProcessSendToEmailAddressList(EmailTo, EmailCcList, EmailBccList, sClienteEmail, sCurEmailAsesor)

                ' Crear el asunto del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailSubject = EmailSubject

                FormatTemplateText(sCurEmailSubject, "EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "NUMERO_EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "CTE_NUMERO_EPS", sNumeroEPS)

                FormatTemplateText(sCurEmailSubject, "CODIGO_BARRA", sCodigoBarra)
                FormatTemplateText(sCurEmailSubject, "BLT_CODIGO_BARRA", sCodigoBarra)

                FormatTemplateText(sCurEmailSubject, "TRACKING_NUMBER", sTrackingNumber)
                FormatTemplateText(sCurEmailSubject, "BLT_TRACKING_NUMBER", sTrackingNumber)

                ' Formatear cuerpo del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailBoby = EmailBody

                ' Formatear Template Datos Generales
                FormatTemplateTextDatosGenerales(sCurEmailBoby, TemplateID, bIsHtml, dr)

                FormatTemplateText(sCurEmailBoby, "EMAIL_FROM", EmailFrom)

                ' Paquete o bulto codigo de barra
                FormatTemplateText(sCurEmailBoby, "CODIGO_BARRA", sCodigoBarra)
                FormatTemplateText(sCurEmailBoby, "BLT_CODIGO_BARRA", sCodigoBarra)

                ' Paquete o bulto tracking number
                FormatTemplateText(sCurEmailBoby, "TRACKING_NUMBER", sTrackingNumber)
                FormatTemplateText(sCurEmailBoby, "BLT_TRACKING_NUMBER", sTrackingNumber)

                ' Paquete o bulto guia hija
                FormatTemplateText(sCurEmailBoby, "GUIA_HIJA", sGuiaHija)
                FormatTemplateText(sCurEmailBoby, "BLT_GUIA_HIJA", sGuiaHija)

                ' Paquete o bulto guia madre
                FormatTemplateText(sCurEmailBoby, "MAN_GUIA", sGuiaMadre)

                ' Paquete o bulto manifiesto
                FormatTemplateText(sCurEmailBoby, "MAN_MANIFIESTO", sManifiesto)

                ' Paquete o bulto servicio codigo y descripcion
                FormatTemplateText(sCurEmailBoby, "PRO_CODIGO", sServicioCodigo)
                FormatTemplateText(sCurEmailBoby, "PRO_DESCRIPCION", sServicio)
                FormatTemplateText(sCurEmailBoby, "PRODUCTO", sServicio)
                FormatTemplateText(sCurEmailBoby, "SERVICIO", sServicio)

                ' Paquete o bulto origen codigo y descripcion
                FormatTemplateText(sCurEmailBoby, "ORI_CODIGO", sOrigen)
                FormatTemplateText(sCurEmailBoby, "ORIGEN", sOrigen)

                ' Paquete o bulto orden numero
                FormatTemplateText(sCurEmailBoby, "BLT_PONUMBER", sOrdenNo)
                FormatTemplateText(sCurEmailBoby, "PONUMBER", sOrdenNo)
                FormatTemplateText(sCurEmailBoby, "ORDEN_NO", sOrdenNo)

                ' Paquete o bulto factura numero del suplidor
                FormatTemplateText(sCurEmailBoby, "BLT_FACTURA_SUPLIDOR", sFacturaNo)
                FormatTemplateText(sCurEmailBoby, "FACTURA_SUPLIDOR", sFacturaNo)
                FormatTemplateText(sCurEmailBoby, "FACTURA_NO", sFacturaNo)

                ' Paquete o bulto fecha de recepcion
                FormatTemplateText(sCurEmailBoby, "BLT_FECHA_RECEPCION", sFechaRecepcion)
                FormatTemplateText(sCurEmailBoby, "FECHA_RECEPCION", sFechaRecepcion)

                ' Paquete o bulto piezas
                FormatTemplateText(sCurEmailBoby, "BLT_PIEZAS", nPiezas)
                FormatTemplateText(sCurEmailBoby, "PIEZAS", nPiezas)

                ' Paquete o bulto peso
                FormatTemplateText(sCurEmailBoby, "BLT_PESO", nPeso)
                FormatTemplateText(sCurEmailBoby, "PESO", nPeso)

                ' Paquete o bulto contenido
                FormatTemplateText(sCurEmailBoby, "CONTENIDO", sContenido)

                ' Paquete o bulto remitente
                FormatTemplateText(sCurEmailBoby, "SUPLIDOR", sRemitente)
                FormatTemplateText(sCurEmailBoby, "REMITENTE", sRemitente)

                ' Paquete o bulto destinatario
                FormatTemplateText(sCurEmailBoby, "DESTINATARIO", sDestinatario)

                ' Paquete o bulto localizacion
                FormatTemplateText(sCurEmailBoby, "LOCALIZACION", sLocalizacion)

                ' Paquete o bulto condicion
                FormatTemplateText(sCurEmailBoby, "CONDICION", sCondicionPrimera)

                ' Procesar mensaje dependiendo el template
                '------------------------------------------------------------------------
                Select Case TemplateID
                    Case 5 ' Pendiente autorización pago Impuestos
                        If ProcessImpuestosAduanales(sNumeroEPS, sCodigoBarra, sCurEmailBoby) = False Then
                            bSendEmail = False
                        End If
                    Case 3, 4   'paquete sin factura aduana
                        If ProcesaLinkSubirFactura(sNumeroEPS, sCodigoBarra, sCurEmailBoby) = False Then
                            bSendEmail = False
                        End If
                    Case Else
                        ' Do nothing
                End Select

                ' Crear enlace para ver documento
                '------------------------------------------------------------------------
                If sCurEmailBoby.Contains("%URL_DOCUMENTO%") = True Then
                    If bDocumentoDisponible = True AndAlso Not String.IsNullOrEmpty(sDocumentoReferencia) Then
                        Dim sCurDocumentoUrl As String = String.Empty
                        sCurDocumentoUrl = String.Format(My.Settings.ApplicationDocumentoURL, "{" & sDocumentoReferencia.ToUpper & "}")

                        Select Case TemplateID
                            Case 4, 18
                                sTmp = String.Format("<p><a href=""{0}"">Hacer clic aqu&iacute; para ver factura</a><br><a href=""{0}"">{0}</a></p>", sCurDocumentoUrl)
                            Case 5 ' Pendiente autorización pago Impuestos
                                sTmp = String.Format("<p><a href=""{0}"">Hacer clic aqu&iacute; para ver planilla anexo</a><br><a href=""{0}"">{0}</a></p>", sCurDocumentoUrl)
                            Case Else
                                sTmp = String.Empty
                        End Select
                        FormatTemplateText(sCurEmailBoby, "URL_DOCUMENTO", sTmp)
                    Else
                        FormatTemplateText(sCurEmailBoby, "URL_DOCUMENTO", String.Empty)
                    End If
                End If

                ' Enviar el correo electrónico
                '------------------------------------------------------------------------

                SendEmail(TemplateID, EmailFormat, sCurEmailTo, EmailFrom, EmailFromName, EmailReplyTo, sCurEmailSubject & "--", sCurEmailBoby, bSendEmail, dr)
                ' Incrementar contador de registros procesados
                nRecCount += 1
            Next

            ' Desplegar mensaje que no existe email para procesar
            If nRecCount = 0 Then PrintDobleLine(String.Format("> No existen notificaciones de {0} para enviar", sCurName))

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessMensajesNotificacionesBultos(): " & sEx, 2)
        End Try

    End Sub

    Private Function ProcesaLinkSubirFactura(ByVal NumeroEPS As String, ByVal CodigoBarra As String, ByRef EmailBody As String) As Boolean
        Dim bReturn As Boolean = False

        Dim sKey As String = ""

        Dim ds As New DataSet
        sKey = GetDatosSubirFacturaDataSet(NumeroEPS, CodigoBarra)
        If sKey = Nothing Then
            Return False
        End If
        EmailBody = EmailBody.Replace("%RowId%", sKey)

        Return True

    End Function

    Private Function GetDatosSubirFacturaDataSet(ByVal NumeroEPS As String, ByVal CodigoBarra As String) As String
        'ID       ROWID      BLT_NUMERO  BLT_CODIGO_BARRA      CARGADO TRACKING  CARRIER   DESCRIPCION   VALOR
        Dim dt1 As New DataTable()
        Dim dt2 As New DataTable()


        Dim sSql = " SELECT B.BLT_NUMERO, ISNULL(BLT_CARRIER,' ') BLT_CARRIER ,ISNULL(BLT_VALOR_FOB,0) BLT_VALOR_FOB, ISNULL(C.COB_CONTENIDO,' ') CONTENIDO , BLT_TRACKING_NUMBER" & _
                 " FROM BULTOS B INNER JOIN CONTENIDO_BULTOS C ON B.BLT_NUMERO = C.BLT_NUMERO " & _
                 " WHERE BLT_CODIGO_BARRA  = '" & CodigoBarra & "'"

        Try
            dt1 = db.ewGetDataSet(sSql).Tables(0)

            If dt1.Rows.Count = 0 Then
                Return Nothing
            End If

            Dim sSql2 As String = "IF (SELECT COUNT(1) FROM  AVISO_FACT_CORREO " & _
                                  " WHERE BLT_CODIGO_BARRA =  '" & CodigoBarra & "')=0" & _
                 " BEGIN " & _
                "   INSERT INTO AVISO_FACT_CORREO ( BLT_NUMERO , BLT_CODIGO_BARRA,      CARGADO, TRACKING,  CARRIER,   DESCRIPCION,   VALOR ) " & _
                                  " Values (@BLT_NUMERO , '@BLT_CODIGO_BARRA',@CARGADO, '@TRACKING',  '@CARRIER',   '@DESCRIPCION',   @VALOR); " &
                 "END " & _
                 "SELECT TOP 1 ROWID  FROM AVISO_FACT_CORREO WHERE BLT_CODIGO_BARRA =  '" & CodigoBarra & "'"


            sSql2 = sSql2.Replace("@BLT_NUMERO", dt1.Rows(0).Item(0).ToString())
            sSql2 = sSql2.Replace("@BLT_CODIGO_BARRA", CodigoBarra)
            sSql2 = sSql2.Replace("@CARGADO", "0")
            sSql2 = sSql2.Replace("@TRACKING", dt1.Rows(0).Item("BLT_TRACKING_NUMBER").ToString())
            sSql2 = sSql2.Replace("@CARRIER", dt1.Rows(0).Item("BLT_CARRIER").ToString())
            sSql2 = sSql2.Replace("@DESCRIPCION", dt1.Rows(0).Item("CONTENIDO").ToString().Trim)
            sSql2 = sSql2.Replace("@VALOR", dt1.Rows(0).Item("BLT_VALOR_FOB").ToString().Replace(",", "."))

            dt2 = db.ewGetDataSet(sSql2).Tables(0)

            Return dt2.Rows(0).Item(0).ToString()


        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetImpuestosAduanalesDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

    Sub ProcessMensajesNotificacionesBultosSgda(ByVal TemplateID As Integer, _
                                            ByVal EmailFormat As String, _
                                            ByVal EmailTo As String, _
                                            ByVal EmailCcList As String, _
                                            ByVal EmailBccList As String, _
                                            ByVal EmailFrom As String, _
                                            ByVal EmailFromName As String, _
                                            ByVal EmailReplyTo As String, _
                                            ByVal EmailSubject As String, _
                                            ByVal EmailBody As String, _
                                            ByVal NotificaRepresentantes As String)

        Dim sCurName As String = "Paquete o Bulto"
        Dim bSendEmail As Boolean = False

        Dim sTmp As String = String.Empty
        Dim sMessage As String = String.Empty

        Dim sCurEmailTo As String = String.Empty
        Dim sCurEmailSubject As String = String.Empty
        Dim sCurEmailBoby As String = String.Empty
        Dim sCurEmailAsesor As String = String.Empty

        Dim nMensajeID As String = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim sMensajeFechaCreado As String = Nothing

        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sTipoClienteCodigo As String = Nothing
        Dim sSubTipoClienteCodigo As String = Nothing
        Dim sClienteEstado As String = Nothing
        Dim sClienteEstatus As String = Nothing
        Dim sCedula As String = Nothing
        Dim sRNC As String = Nothing
        Dim sPasaporte As String = Nothing
        Dim sClienteEmail As String = Nothing
        Dim sEnviarEmail As String = Nothing
        Dim nCompaniaID As String = Nothing
        Dim sSucursalCodigo As String = Nothing
        Dim sAgenciaCodigo As String = Nothing
        Dim sAgenciaEmail As String = Nothing
        Dim sOficialCodigo As String = Nothing
        Dim sOficialNombre As String = Nothing
        Dim sAsesorCodigo As String = Nothing
        Dim sAsesorNombre As String = Nothing
        Dim sAsesorEmail As String = Nothing
        Dim sClienteTieneCredito As String = Nothing
        Dim sClienteTieneCorrespondecia As String = Nothing

        Dim nBultoNumero As String = Nothing
        Dim sCodigoBarra As String = Nothing
        Dim sTrackingNumber As String = Nothing
        Dim sGuiaHija As String = Nothing
        Dim sGuiaMadre As String = Nothing
        Dim sManifiesto As String = Nothing
        Dim sServicioCodigo As String = Nothing
        Dim sServicio As String = Nothing
        Dim sOrigen As String = Nothing
        Dim sOrdenNo As String = Nothing
        Dim sFacturaNo As String = Nothing
        Dim sFechaRecepcion As String = Nothing
        Dim nPiezas As String = Nothing
        Dim nPeso As String = Nothing
        Dim sContenido As String = Nothing
        Dim sRemitente As String = Nothing
        Dim sDestinatario As String = Nothing
        Dim sLocalizacion As String = Nothing
        Dim sCondicionSegunda As String = Nothing
        Dim bDocumentoDisponible As Boolean = Nothing
        Dim sDocumentoReferencia As String = Nothing

        ' Validar si el mensaje es HTML
        Dim bIsHtml As Boolean = IIf(EmailFormat.ToUpper.Trim = "HTML", True, False)

        ' Desplegar mensaje
        PrintDobleLine(String.Format("- Buscando notificaciones {0} para enviar, por favor espere...", sCurName))

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetMensajesNotificacionesBultosDataSetSgda(TemplateID, nAppProcesarDias)

            ' Desplegar mensaje
            PrintDobleLine(String.Format("- Procesando notificaciones de {0} para enviar...", sCurName))

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                ' Enviar correo
                bSendEmail = True

                ' Buscar los datos de la notificacion
                '------------------------------------
                nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
                sMensajeGUID = db.ewToString(dr("WMM_MENSAJE_GUID"))
                sMensajeFechaCreado = db.ewToString(dr("WMM_FECHA_CREADO"))

                sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
                sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))
                sTipoClienteCodigo = db.ewToStringUpper(dr("CTE_TIPO"))
                sSubTipoClienteCodigo = db.ewToStringUpper(dr("STC_CODIGO"))
                sClienteEstado = db.ewToStringUpper(dr("CTE_ESTADO"))
                sClienteEstatus = db.ewToStringUpper(dr("ESTATUS"))
                sCedula = db.ewToString(dr("CTE_CEDULA"))
                sRNC = db.ewToString(dr("CTE_RNC"))
                sPasaporte = db.ewToString(dr("CTE_PASAPORTE"))
                sClienteEmail = db.ewToStringLower(dr("CTE_EMAIL"))
                sEnviarEmail = db.ewToStringUpper(dr("CTE_ENVIAR_EMAIL"))
                nCompaniaID = db.ewToString(dr("COM_CODIGO"))
                sSucursalCodigo = db.ewToStringUpper(dr("SUC_CODIGO"))
                sAgenciaCodigo = db.ewToStringUpper(dr("AGE_CODIGO"))
                sAgenciaEmail = db.ewToStringLower(dr("AGENCIA_EMAIL"))
                sOficialCodigo = db.ewToStringUpper(dr("CTE_VENDEDOR"))
                sOficialNombre = db.ewToStringUpper(dr("OFICIAL"))
                sAsesorCodigo = db.ewToStringUpper(dr("RES_CODIGO"))
                sAsesorNombre = db.ewToStringUpper(dr("ASESOR"))
                sAsesorEmail = db.ewToStringLower(dr("ASESOR_EMAIL"))
                sClienteTieneCredito = db.ewToStringUpper(dr("CTE_CREDITO"))
                sClienteTieneCorrespondecia = db.ewToStringUpper(dr("CTE_CORRESPONDENCIA"))

                nBultoNumero = db.ewToString(dr("BLT_NUMERO"))
                sCodigoBarra = db.ewToStringUpper(dr("BLT_CODIGO_BARRA"))
                sTrackingNumber = db.ewToStringUpper(dr("BLT_TRACKING_NUMBER"))
                sGuiaHija = db.ewToStringUpper(dr("BLT_GUIA_HIJA"))
                sGuiaMadre = db.ewToStringUpper(dr("MAN_GUIA"))
                sManifiesto = db.ewToStringUpper(dr("MAN_MANIFIESTO"))
                sServicioCodigo = db.ewToStringUpper(dr("PRO_CODIGO"))
                sServicio = db.ewToStringUpper(dr("PRO_DESCRIPCION"))
                sOrigen = db.ewToStringUpper(dr("ORIGEN"))
                sOrdenNo = db.ewToStringUpper(dr("BLT_PONUMBER"))
                sFacturaNo = db.ewToStringUpper(dr("BLT_FACTURA_SUPLIDOR"))
                sFechaRecepcion = db.ewToString(dr("BLT_FECHA_RECEPCION"))
                nPiezas = db.ewToString(dr("BLT_PIEZAS"))
                nPeso = db.ewToString(dr("BLT_PESO"))
                sContenido = db.ewToStringUpper(dr("CONTENIDO"))
                sRemitente = db.ewToStringUpper(dr("REMITENTE"))
                sDestinatario = db.ewToStringUpper(dr("DESTINATARIO"))
                sLocalizacion = db.ewToStringUpper(dr("LOCALIZACION"))
                sCondicionSegunda = db.ewToStringUpper(dr("CONDICION"))

                'Dim condiciones = GetCondiciones(sCodigoBarra, sTrackingNumber)


                bDocumentoDisponible = db.ewToBool(dr("DOC_DISPONIBLE"))
                sDocumentoReferencia = db.ewToStringNullable(dr("DOC_REFERENCIA"))

                ' SetUp Cliente
                ' -------------
                SetUpClienteOficial(sOficialCodigo, sOficialNombre, sSucursalCodigo, sAgenciaCodigo)
                SetUpClienteAsesor(sAsesorCodigo, sAsesorNombre, sAsesorEmail, sSucursalCodigo, sAgenciaCodigo)
                If NotificaRepresentantes = "S" Then sCurEmailAsesor = sAsesorEmail ' Enviar notificacion al asesor de la cuenta

                ' HTML Encoding
                '--------------
                If bIsHtml = True Then
                    If Not String.IsNullOrEmpty(sContenido) Then sContenido = ew_EncodeText(sContenido)
                    If Not String.IsNullOrEmpty(sRemitente) Then sRemitente = ew_EncodeText(sRemitente)
                    If Not String.IsNullOrEmpty(sDestinatario) Then sDestinatario = ew_EncodeText(sDestinatario)
                End If

                ' Set Default Values
                '-------------------
                If String.IsNullOrEmpty(sContenido) = True Then sContenido = "N/A"
                If String.IsNullOrEmpty(sRemitente) = True Then sRemitente = "N/A"
                If String.IsNullOrEmpty(sDestinatario) = True Then sDestinatario = "N/A"

                ' Buscar dirección de envio del mensaje
                '------------------------------------------------------------------------
                sCurEmailTo = ProcessSendToEmailAddressList(EmailTo, EmailCcList, EmailBccList, sClienteEmail, sCurEmailAsesor)

                ' Crear el asunto del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailSubject = EmailSubject

                FormatTemplateText(sCurEmailSubject, "EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "NUMERO_EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "CTE_NUMERO_EPS", sNumeroEPS)

                FormatTemplateText(sCurEmailSubject, "CODIGO_BARRA", sCodigoBarra)
                FormatTemplateText(sCurEmailSubject, "BLT_CODIGO_BARRA", sCodigoBarra)

                FormatTemplateText(sCurEmailSubject, "TRACKING_NUMBER", sTrackingNumber)
                FormatTemplateText(sCurEmailSubject, "BLT_TRACKING_NUMBER", sTrackingNumber)

                ' Formatear cuerpo del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailBoby = EmailBody

                ' Formatear Template Datos Generales
                FormatTemplateTextDatosGenerales(sCurEmailBoby, TemplateID, bIsHtml, dr)

                FormatTemplateText(sCurEmailBoby, "EMAIL_FROM", EmailFrom)

                ' Paquete o bulto codigo de barra
                FormatTemplateText(sCurEmailBoby, "CODIGO_BARRA", sCodigoBarra)
                FormatTemplateText(sCurEmailBoby, "BLT_CODIGO_BARRA", sCodigoBarra)

                ' Paquete o bulto tracking number
                FormatTemplateText(sCurEmailBoby, "TRACKING_NUMBER", sTrackingNumber)
                FormatTemplateText(sCurEmailBoby, "BLT_TRACKING_NUMBER", sTrackingNumber)

                ' Paquete o bulto guia hija
                FormatTemplateText(sCurEmailBoby, "GUIA_HIJA", sGuiaHija)
                FormatTemplateText(sCurEmailBoby, "BLT_GUIA_HIJA", sGuiaHija)

                ' Paquete o bulto guia madre
                FormatTemplateText(sCurEmailBoby, "MAN_GUIA", sGuiaMadre)

                ' Paquete o bulto manifiesto
                FormatTemplateText(sCurEmailBoby, "MAN_MANIFIESTO", sManifiesto)

                ' Paquete o bulto servicio codigo y descripcion
                FormatTemplateText(sCurEmailBoby, "PRO_CODIGO", sServicioCodigo)
                FormatTemplateText(sCurEmailBoby, "PRO_DESCRIPCION", sServicio)
                FormatTemplateText(sCurEmailBoby, "PRODUCTO", sServicio)
                FormatTemplateText(sCurEmailBoby, "SERVICIO", sServicio)

                ' Paquete o bulto origen codigo y descripcion
                FormatTemplateText(sCurEmailBoby, "ORI_CODIGO", sOrigen)
                FormatTemplateText(sCurEmailBoby, "ORIGEN", sOrigen)

                ' Paquete o bulto orden numero
                FormatTemplateText(sCurEmailBoby, "BLT_PONUMBER", sOrdenNo)
                FormatTemplateText(sCurEmailBoby, "PONUMBER", sOrdenNo)
                FormatTemplateText(sCurEmailBoby, "ORDEN_NO", sOrdenNo)

                ' Paquete o bulto factura numero del suplidor
                FormatTemplateText(sCurEmailBoby, "BLT_FACTURA_SUPLIDOR", sFacturaNo)
                FormatTemplateText(sCurEmailBoby, "FACTURA_SUPLIDOR", sFacturaNo)
                FormatTemplateText(sCurEmailBoby, "FACTURA_NO", sFacturaNo)

                ' Paquete o bulto fecha de recepcion
                FormatTemplateText(sCurEmailBoby, "BLT_FECHA_RECEPCION", sFechaRecepcion)
                FormatTemplateText(sCurEmailBoby, "FECHA_RECEPCION", sFechaRecepcion)

                ' Paquete o bulto piezas
                FormatTemplateText(sCurEmailBoby, "BLT_PIEZAS", nPiezas)
                FormatTemplateText(sCurEmailBoby, "PIEZAS", nPiezas)

                ' Paquete o bulto peso
                FormatTemplateText(sCurEmailBoby, "BLT_PESO", nPeso)
                FormatTemplateText(sCurEmailBoby, "PESO", nPeso)

                ' Paquete o bulto contenido
                FormatTemplateText(sCurEmailBoby, "CONTENIDO", sContenido)

                ' Paquete o bulto remitente
                FormatTemplateText(sCurEmailBoby, "SUPLIDOR", sRemitente)
                FormatTemplateText(sCurEmailBoby, "REMITENTE", sRemitente)

                ' Paquete o bulto destinatario
                FormatTemplateText(sCurEmailBoby, "DESTINATARIO", sDestinatario)

                ' Paquete o bulto localizacion
                FormatTemplateText(sCurEmailBoby, "LOCALIZACION", sLocalizacion)

                ' Paquete o bulto condicion

                FormatTemplateText(sCurEmailBoby, "CONDICION", sCondicionSegunda)

                ' Procesar mensaje dependiendo el template
                '------------------------------------------------------------------------
                Select Case TemplateID
                    Case 5 ' Pendiente autorización pago Impuestos
                        If ProcessImpuestosAduanales(sNumeroEPS, sCodigoBarra, sCurEmailBoby) = False Then
                            bSendEmail = False
                        End If
                    Case Else
                        ' Do nothing
                End Select

                ' Crear enlace para ver documento
                '------------------------------------------------------------------------
                If sCurEmailBoby.Contains("%URL_DOCUMENTO%") = True Then
                    If bDocumentoDisponible = True AndAlso Not String.IsNullOrEmpty(sDocumentoReferencia) Then
                        Dim sCurDocumentoUrl As String = String.Empty
                        sCurDocumentoUrl = String.Format(My.Settings.ApplicationDocumentoURL, "{" & sDocumentoReferencia.ToUpper & "}")

                        Select Case TemplateID
                            Case 4, 18
                                sTmp = String.Format("<p><a href=""{0}"">Hacer clic aqu&iacute; para ver factura</a><br><a href=""{0}"">{0}</a></p>", sCurDocumentoUrl)
                            Case 5 ' Pendiente autorización pago Impuestos
                                sTmp = String.Format("<p><a href=""{0}"">Hacer clic aqu&iacute; para ver planilla anexo</a><br><a href=""{0}"">{0}</a></p>", sCurDocumentoUrl)
                            Case Else
                                sTmp = String.Empty
                        End Select
                        FormatTemplateText(sCurEmailBoby, "URL_DOCUMENTO", sTmp)
                    Else
                        FormatTemplateText(sCurEmailBoby, "URL_DOCUMENTO", String.Empty)
                    End If
                End If

                ' Enviar el correo electrónico
                '------------------------------------------------------------------------
                If (String.IsNullOrEmpty(sCondicionSegunda) = False And sCondicionSegunda <> "N/A") Then
                    SendEmail(TemplateID, EmailFormat, sCurEmailTo, EmailFrom, EmailFromName, EmailReplyTo, sCurEmailSubject & "--", sCurEmailBoby, bSendEmail, dr)
                End If

                ' Incrementar contador de registros procesados
                nRecCountSgda += 1
            Next

            ' Desplegar mensaje que no existe email para procesar
            If nRecCountSgda = 0 Then PrintDobleLine(String.Format("> No existen notificaciones de {0} para enviar", sCurName))

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessMensajesNotificacionesBultos(): " & sEx, 2)
        End Try

    End Sub

    Sub ProcessMensajesNotificacionesBultosEra(ByVal TemplateID As Integer, _
                                            ByVal EmailFormat As String, _
                                            ByVal EmailTo As String, _
                                            ByVal EmailCcList As String, _
                                            ByVal EmailBccList As String, _
                                            ByVal EmailFrom As String, _
                                            ByVal EmailFromName As String, _
                                            ByVal EmailReplyTo As String, _
                                            ByVal EmailSubject As String, _
                                            ByVal EmailBody As String, _
                                            ByVal NotificaRepresentantes As String)

        Dim sCurName As String = "Paquete o Bulto"
        Dim bSendEmail As Boolean = False

        Dim sTmp As String = String.Empty
        Dim sMessage As String = String.Empty

        Dim sCurEmailTo As String = String.Empty
        Dim sCurEmailSubject As String = String.Empty
        Dim sCurEmailBoby As String = String.Empty
        Dim sCurEmailAsesor As String = String.Empty

        Dim nMensajeID As String = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim sMensajeFechaCreado As String = Nothing

        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sTipoClienteCodigo As String = Nothing
        Dim sSubTipoClienteCodigo As String = Nothing
        Dim sClienteEstado As String = Nothing
        Dim sClienteEstatus As String = Nothing
        Dim sCedula As String = Nothing
        Dim sRNC As String = Nothing
        Dim sPasaporte As String = Nothing
        Dim sClienteEmail As String = Nothing
        Dim sEnviarEmail As String = Nothing
        Dim nCompaniaID As String = Nothing
        Dim sSucursalCodigo As String = Nothing
        Dim sAgenciaCodigo As String = Nothing
        Dim sAgenciaEmail As String = Nothing
        Dim sOficialCodigo As String = Nothing
        Dim sOficialNombre As String = Nothing
        Dim sAsesorCodigo As String = Nothing
        Dim sAsesorNombre As String = Nothing
        Dim sAsesorEmail As String = Nothing
        Dim sClienteTieneCredito As String = Nothing
        Dim sClienteTieneCorrespondecia As String = Nothing

        Dim nBultoNumero As String = Nothing
        Dim sCodigoBarra As String = Nothing
        Dim sTrackingNumber As String = Nothing
        Dim sGuiaHija As String = Nothing
        Dim sGuiaMadre As String = Nothing
        Dim sManifiesto As String = Nothing
        Dim sServicioCodigo As String = Nothing
        Dim sServicio As String = Nothing
        Dim sOrigen As String = Nothing
        Dim sOrdenNo As String = Nothing
        Dim sFacturaNo As String = Nothing
        Dim sFechaRecepcion As String = Nothing
        Dim nPiezas As String = Nothing
        Dim nPeso As String = Nothing
        Dim sContenido As String = Nothing
        Dim sRemitente As String = Nothing
        Dim sDestinatario As String = Nothing
        Dim sLocalizacion As String = Nothing
        Dim sCondicionTercera As String = Nothing
        Dim bDocumentoDisponible As Boolean = Nothing
        Dim sDocumentoReferencia As String = Nothing

        ' Validar si el mensaje es HTML
        Dim bIsHtml As Boolean = IIf(EmailFormat.ToUpper.Trim = "HTML", True, False)

        ' Desplegar mensaje
        PrintDobleLine(String.Format("- Buscando notificaciones {0} para enviar, por favor espere...", sCurName))

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetMensajesNotificacionesBultosDataSetEra(TemplateID, nAppProcesarDias)

            ' Desplegar mensaje
            PrintDobleLine(String.Format("- Procesando notificaciones de {0} para enviar...", sCurName))

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                ' Enviar correo
                bSendEmail = True

                ' Buscar los datos de la notificacion
                '------------------------------------
                nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
                sMensajeGUID = db.ewToString(dr("WMM_MENSAJE_GUID"))
                sMensajeFechaCreado = db.ewToString(dr("WMM_FECHA_CREADO"))

                sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
                sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))
                sTipoClienteCodigo = db.ewToStringUpper(dr("CTE_TIPO"))
                sSubTipoClienteCodigo = db.ewToStringUpper(dr("STC_CODIGO"))
                sClienteEstado = db.ewToStringUpper(dr("CTE_ESTADO"))
                sClienteEstatus = db.ewToStringUpper(dr("ESTATUS"))
                sCedula = db.ewToString(dr("CTE_CEDULA"))
                sRNC = db.ewToString(dr("CTE_RNC"))
                sPasaporte = db.ewToString(dr("CTE_PASAPORTE"))
                sClienteEmail = db.ewToStringLower(dr("CTE_EMAIL"))
                sEnviarEmail = db.ewToStringUpper(dr("CTE_ENVIAR_EMAIL"))
                nCompaniaID = db.ewToString(dr("COM_CODIGO"))
                sSucursalCodigo = db.ewToStringUpper(dr("SUC_CODIGO"))
                sAgenciaCodigo = db.ewToStringUpper(dr("AGE_CODIGO"))
                sAgenciaEmail = db.ewToStringLower(dr("AGENCIA_EMAIL"))
                sOficialCodigo = db.ewToStringUpper(dr("CTE_VENDEDOR"))
                sOficialNombre = db.ewToStringUpper(dr("OFICIAL"))
                sAsesorCodigo = db.ewToStringUpper(dr("RES_CODIGO"))
                sAsesorNombre = db.ewToStringUpper(dr("ASESOR"))
                sAsesorEmail = db.ewToStringLower(dr("ASESOR_EMAIL"))
                sClienteTieneCredito = db.ewToStringUpper(dr("CTE_CREDITO"))
                sClienteTieneCorrespondecia = db.ewToStringUpper(dr("CTE_CORRESPONDENCIA"))

                nBultoNumero = db.ewToString(dr("BLT_NUMERO"))
                sCodigoBarra = db.ewToStringUpper(dr("BLT_CODIGO_BARRA"))
                sTrackingNumber = db.ewToStringUpper(dr("BLT_TRACKING_NUMBER"))
                sGuiaHija = db.ewToStringUpper(dr("BLT_GUIA_HIJA"))
                sGuiaMadre = db.ewToStringUpper(dr("MAN_GUIA"))
                sManifiesto = db.ewToStringUpper(dr("MAN_MANIFIESTO"))
                sServicioCodigo = db.ewToStringUpper(dr("PRO_CODIGO"))
                sServicio = db.ewToStringUpper(dr("PRO_DESCRIPCION"))
                sOrigen = db.ewToStringUpper(dr("ORIGEN"))
                sOrdenNo = db.ewToStringUpper(dr("BLT_PONUMBER"))
                sFacturaNo = db.ewToStringUpper(dr("BLT_FACTURA_SUPLIDOR"))
                sFechaRecepcion = db.ewToString(dr("BLT_FECHA_RECEPCION"))
                nPiezas = db.ewToString(dr("BLT_PIEZAS"))
                nPeso = db.ewToString(dr("BLT_PESO"))
                sContenido = db.ewToStringUpper(dr("CONTENIDO"))
                sRemitente = db.ewToStringUpper(dr("REMITENTE"))
                sDestinatario = db.ewToStringUpper(dr("DESTINATARIO"))
                sLocalizacion = db.ewToStringUpper(dr("LOCALIZACION"))
                sCondicionTercera = db.ewToStringUpper(dr("CONDICION"))

                Dim condiciones = GetCondiciones(sCodigoBarra, sTrackingNumber)


                bDocumentoDisponible = db.ewToBool(dr("DOC_DISPONIBLE"))
                sDocumentoReferencia = db.ewToStringNullable(dr("DOC_REFERENCIA"))

                ' SetUp Cliente
                ' -------------
                SetUpClienteOficial(sOficialCodigo, sOficialNombre, sSucursalCodigo, sAgenciaCodigo)
                SetUpClienteAsesor(sAsesorCodigo, sAsesorNombre, sAsesorEmail, sSucursalCodigo, sAgenciaCodigo)
                If NotificaRepresentantes = "S" Then sCurEmailAsesor = sAsesorEmail ' Enviar notificacion al asesor de la cuenta

                ' HTML Encoding
                '--------------
                If bIsHtml = True Then
                    If Not String.IsNullOrEmpty(sContenido) Then sContenido = ew_EncodeText(sContenido)
                    If Not String.IsNullOrEmpty(sRemitente) Then sRemitente = ew_EncodeText(sRemitente)
                    If Not String.IsNullOrEmpty(sDestinatario) Then sDestinatario = ew_EncodeText(sDestinatario)
                End If

                ' Set Default Values
                '-------------------
                If String.IsNullOrEmpty(sContenido) = True Then sContenido = "N/A"
                If String.IsNullOrEmpty(sRemitente) = True Then sRemitente = "N/A"
                If String.IsNullOrEmpty(sDestinatario) = True Then sDestinatario = "N/A"

                ' Buscar dirección de envio del mensaje
                '------------------------------------------------------------------------
                sCurEmailTo = ProcessSendToEmailAddressList(EmailTo, EmailCcList, EmailBccList, sClienteEmail, sCurEmailAsesor)

                ' Crear el asunto del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailSubject = EmailSubject

                FormatTemplateText(sCurEmailSubject, "EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "NUMERO_EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "CTE_NUMERO_EPS", sNumeroEPS)

                FormatTemplateText(sCurEmailSubject, "CODIGO_BARRA", sCodigoBarra)
                FormatTemplateText(sCurEmailSubject, "BLT_CODIGO_BARRA", sCodigoBarra)

                FormatTemplateText(sCurEmailSubject, "TRACKING_NUMBER", sTrackingNumber)
                FormatTemplateText(sCurEmailSubject, "BLT_TRACKING_NUMBER", sTrackingNumber)

                ' Formatear cuerpo del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailBoby = EmailBody

                ' Formatear Template Datos Generales
                FormatTemplateTextDatosGenerales(sCurEmailBoby, TemplateID, bIsHtml, dr)

                FormatTemplateText(sCurEmailBoby, "EMAIL_FROM", EmailFrom)

                ' Paquete o bulto codigo de barra
                FormatTemplateText(sCurEmailBoby, "CODIGO_BARRA", sCodigoBarra)
                FormatTemplateText(sCurEmailBoby, "BLT_CODIGO_BARRA", sCodigoBarra)

                ' Paquete o bulto tracking number
                FormatTemplateText(sCurEmailBoby, "TRACKING_NUMBER", sTrackingNumber)
                FormatTemplateText(sCurEmailBoby, "BLT_TRACKING_NUMBER", sTrackingNumber)

                ' Paquete o bulto guia hija
                FormatTemplateText(sCurEmailBoby, "GUIA_HIJA", sGuiaHija)
                FormatTemplateText(sCurEmailBoby, "BLT_GUIA_HIJA", sGuiaHija)

                ' Paquete o bulto guia madre
                FormatTemplateText(sCurEmailBoby, "MAN_GUIA", sGuiaMadre)

                ' Paquete o bulto manifiesto
                FormatTemplateText(sCurEmailBoby, "MAN_MANIFIESTO", sManifiesto)

                ' Paquete o bulto servicio codigo y descripcion
                FormatTemplateText(sCurEmailBoby, "PRO_CODIGO", sServicioCodigo)
                FormatTemplateText(sCurEmailBoby, "PRO_DESCRIPCION", sServicio)
                FormatTemplateText(sCurEmailBoby, "PRODUCTO", sServicio)
                FormatTemplateText(sCurEmailBoby, "SERVICIO", sServicio)

                ' Paquete o bulto origen codigo y descripcion
                FormatTemplateText(sCurEmailBoby, "ORI_CODIGO", sOrigen)
                FormatTemplateText(sCurEmailBoby, "ORIGEN", sOrigen)

                ' Paquete o bulto orden numero
                FormatTemplateText(sCurEmailBoby, "BLT_PONUMBER", sOrdenNo)
                FormatTemplateText(sCurEmailBoby, "PONUMBER", sOrdenNo)
                FormatTemplateText(sCurEmailBoby, "ORDEN_NO", sOrdenNo)

                ' Paquete o bulto factura numero del suplidor
                FormatTemplateText(sCurEmailBoby, "BLT_FACTURA_SUPLIDOR", sFacturaNo)
                FormatTemplateText(sCurEmailBoby, "FACTURA_SUPLIDOR", sFacturaNo)
                FormatTemplateText(sCurEmailBoby, "FACTURA_NO", sFacturaNo)

                ' Paquete o bulto fecha de recepcion
                FormatTemplateText(sCurEmailBoby, "BLT_FECHA_RECEPCION", sFechaRecepcion)
                FormatTemplateText(sCurEmailBoby, "FECHA_RECEPCION", sFechaRecepcion)

                ' Paquete o bulto piezas
                FormatTemplateText(sCurEmailBoby, "BLT_PIEZAS", nPiezas)
                FormatTemplateText(sCurEmailBoby, "PIEZAS", nPiezas)

                ' Paquete o bulto peso
                FormatTemplateText(sCurEmailBoby, "BLT_PESO", nPeso)
                FormatTemplateText(sCurEmailBoby, "PESO", nPeso)

                ' Paquete o bulto contenido
                FormatTemplateText(sCurEmailBoby, "CONTENIDO", sContenido)

                ' Paquete o bulto remitente
                FormatTemplateText(sCurEmailBoby, "SUPLIDOR", sRemitente)
                FormatTemplateText(sCurEmailBoby, "REMITENTE", sRemitente)

                ' Paquete o bulto destinatario
                FormatTemplateText(sCurEmailBoby, "DESTINATARIO", sDestinatario)

                ' Paquete o bulto localizacion
                FormatTemplateText(sCurEmailBoby, "LOCALIZACION", sLocalizacion)

                ' Paquete o bulto condicion
                FormatTemplateText(sCurEmailBoby, "CONDICION", sCondicionTercera)

                ' Procesar mensaje dependiendo el template
                '------------------------------------------------------------------------
                Select Case TemplateID
                    Case 5 ' Pendiente autorización pago Impuestos
                        If ProcessImpuestosAduanales(sNumeroEPS, sCodigoBarra, sCurEmailBoby) = False Then
                            bSendEmail = False
                        End If
                    Case Else
                        ' Do nothing
                End Select

                ' Crear enlace para ver documento
                '------------------------------------------------------------------------
                If sCurEmailBoby.Contains("%URL_DOCUMENTO%") = True Then
                    If bDocumentoDisponible = True AndAlso Not String.IsNullOrEmpty(sDocumentoReferencia) Then
                        Dim sCurDocumentoUrl As String = String.Empty
                        sCurDocumentoUrl = String.Format(My.Settings.ApplicationDocumentoURL, "{" & sDocumentoReferencia.ToUpper & "}")

                        Select Case TemplateID
                            Case 4, 18
                                sTmp = String.Format("<p><a href=""{0}"">Hacer clic aqu&iacute; para ver factura</a><br><a href=""{0}"">{0}</a></p>", sCurDocumentoUrl)
                            Case 5 ' Pendiente autorización pago Impuestos
                                sTmp = String.Format("<p><a href=""{0}"">Hacer clic aqu&iacute; para ver planilla anexo</a><br><a href=""{0}"">{0}</a></p>", sCurDocumentoUrl)
                            Case Else
                                sTmp = String.Empty
                        End Select
                        FormatTemplateText(sCurEmailBoby, "URL_DOCUMENTO", sTmp)
                    Else
                        FormatTemplateText(sCurEmailBoby, "URL_DOCUMENTO", String.Empty)
                    End If
                End If

                ' Enviar el correo electrónico
                '------------------------------------------------------------------------

                If (String.IsNullOrEmpty(sCondicionTercera) = False And sCondicionTercera <> "N/A") Then
                    SendEmail(TemplateID, EmailFormat, EmailTo, EmailFrom, EmailFromName, EmailFromName, sCurEmailSubject & "---", sCurEmailBoby, bSendEmail, dr)
                End If

                ' Incrementar contador de registros procesados
                nRecCountEra += 1
            Next

            ' Desplegar mensaje que no existe email para procesar
            If nRecCountEra = 0 Then PrintDobleLine(String.Format("> No existen notificaciones de {0} para enviar", sCurName))

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessMensajesNotificacionesBultos(): " & sEx, 2)
        End Try

    End Sub


    ''' <summary>
    ''' Procesar Mensajes de Notificaciones para Clientes
    ''' </summary>
    ''' <param name="TemplateID">ID del E-mail Template</param>
    ''' <param name="EmailFormat">Formato del E-mail Template</param>
    ''' <param name="EmailTo">Direccion de E-mail (Opcional)</param>
    ''' <param name="EmailCcList">Lista de Direcciones de E-mail de Copia (Opcional, Separador por Coma)</param>
    ''' <param name="EmailBccList">Lista de Direcciones de E-mail BCC (Opcional, Separador por Coma)</param>
    ''' <param name="EmailFrom">Direccion de E-mail de Envio (De)</param>
    ''' <param name="EmailFromName">Nombre de Envio</param>
    ''' <param name="EmailReplyTo">Direccion de E-mail de Respuesta</param>
    ''' <param name="EmailSubject">Asunto del Mensaje</param>
    ''' <param name="EmailBody">Mensaje completo</param>
    ''' <param name="NotificaRepresentantes">Notificar Representantes de Servicios</param>
    ''' <remarks></remarks>
    Sub ProcessMensajesNotificacionesClientes(ByVal TemplateID As Integer, _
                                              ByVal EmailFormat As String, _
                                              ByVal EmailTo As String, _
                                              ByVal EmailCcList As String, _
                                              ByVal EmailBccList As String, _
                                              ByVal EmailFrom As String, _
                                              ByVal EmailFromName As String, _
                                              ByVal EmailReplyTo As String, _
                                              ByVal EmailSubject As String, _
                                              ByVal EmailBody As String, _
                                              ByVal NotificaRepresentantes As String)

        Dim sCurName As String = "Clientes"
        Dim bSendEmail As Boolean = False

        Dim sTmp As String = String.Empty
        Dim sMessage As String = String.Empty

        Dim sCurEmailTo As String = String.Empty
        Dim sCurEmailSubject As String = String.Empty
        Dim sCurEmailBoby As String = String.Empty
        Dim sCurEmailAsesor As String = String.Empty

        Dim nMensajeID As String = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim sMensajeFechaCreado As String = Nothing

        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sTipoClienteCodigo As String = Nothing
        Dim sSubTipoClienteCodigo As String = Nothing
        Dim sClienteEstado As String = Nothing
        Dim sClienteEstatus As String = Nothing
        Dim sCedula As String = Nothing
        Dim sRNC As String = Nothing
        Dim sPasaporte As String = Nothing
        Dim sClienteEmail As String = Nothing
        Dim sEnviarEmail As String = Nothing
        Dim nCompaniaID As String = Nothing
        Dim sSucursalCodigo As String = Nothing
        Dim sAgenciaCodigo As String = Nothing
        Dim sAgenciaEmail As String = Nothing
        Dim sOficialCodigo As String = Nothing
        Dim sOficialNombre As String = Nothing
        Dim sAsesorCodigo As String = Nothing
        Dim sAsesorNombre As String = Nothing
        Dim sAsesorEmail As String = Nothing
        Dim sClienteTieneCredito As String = Nothing

        ' Validar si el mensaje es HTML
        Dim bIsHtml As Boolean = IIf(EmailFormat.ToUpper.Trim = "HTML", True, False)

        ' Desplegar mensaje
        PrintDobleLine(String.Format("- Buscando notificaciones de {0} para enviar, por favor espere...", sCurName))

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetMensajesNotificacionesClientesDataSet(TemplateID, nAppProcesarDias)

            ' Desplegar mensaje
            PrintDobleLine(String.Format("- Procesando notificaciones de {0} para enviar...", sCurName))

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                ' Enviar correo
                bSendEmail = True

                ' Buscar los datos de la notificacion
                '------------------------------------
                nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
                sMensajeGUID = db.ewToString(dr("WMM_MENSAJE_GUID"))
                sMensajeFechaCreado = db.ewToString(dr("WMM_FECHA_CREADO"))

                sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
                sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))
                sTipoClienteCodigo = db.ewToStringUpper(dr("CTE_TIPO"))
                sSubTipoClienteCodigo = db.ewToStringUpper(dr("STC_CODIGO"))
                sClienteEstado = db.ewToStringUpper(dr("CTE_ESTADO"))
                sClienteEstatus = db.ewToStringUpper(dr("ESTATUS"))
                sCedula = db.ewToString(dr("CTE_CEDULA"))
                sRNC = db.ewToString(dr("CTE_RNC"))
                sPasaporte = db.ewToString(dr("CTE_PASAPORTE"))
                sClienteEmail = db.ewToStringLower(dr("CTE_EMAIL"))
                sEnviarEmail = db.ewToStringUpper(dr("CTE_ENVIAR_EMAIL"))
                nCompaniaID = db.ewToString(dr("COM_CODIGO"))
                sSucursalCodigo = db.ewToStringUpper(dr("SUC_CODIGO"))
                sAgenciaCodigo = db.ewToStringUpper(dr("AGE_CODIGO"))
                sAgenciaEmail = db.ewToStringLower(dr("AGENCIA_EMAIL"))
                sOficialCodigo = db.ewToStringUpper(dr("CTE_VENDEDOR"))
                sOficialNombre = db.ewToStringUpper(dr("OFICIAL"))
                sAsesorCodigo = db.ewToStringUpper(dr("RES_CODIGO"))
                sAsesorNombre = db.ewToStringUpper(dr("ASESOR"))
                sAsesorEmail = db.ewToStringLower(dr("ASESOR_EMAIL"))
                sClienteTieneCredito = db.ewToStringUpper(dr("CTE_CREDITO"))

                ' SetUp Cliente
                ' -------------
                SetUpClienteOficial(sOficialCodigo, sOficialNombre, sSucursalCodigo, sAgenciaCodigo)
                SetUpClienteAsesor(sAsesorCodigo, sAsesorNombre, sAsesorEmail, sSucursalCodigo, sAgenciaCodigo)
                If NotificaRepresentantes = "S" Then sCurEmailAsesor = sAsesorEmail ' Enviar notificacion al asesor de la cuenta

                ' Buscar dirección de envio del mensaje
                '------------------------------------------------------------------------
                sCurEmailTo = ProcessSendToEmailAddressList(EmailTo, EmailCcList, EmailBccList, sClienteEmail, sCurEmailAsesor)

                ' Crear el asunto del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailSubject = EmailSubject

                FormatTemplateText(sCurEmailSubject, "EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "NUMERO_EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "CTE_NUMERO_EPS", sNumeroEPS)

                ' Formatear cuerpo del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailBoby = EmailBody

                ' Formatear Template Datos Generales
                FormatTemplateTextDatosGenerales(sCurEmailBoby, TemplateID, bIsHtml, dr)

                FormatTemplateText(sCurEmailBoby, "EMAIL_FROM", EmailFrom)

                ' Enviar el correo electrónico
                '------------------------------------------------------------------------
                SendEmail(TemplateID, EmailFormat, EmailTo, EmailFrom, EmailFromName, EmailFrom, sCurEmailSubject, sCurEmailBoby, bSendEmail, dr)

                ' Incrementar contador de registros procesados
                nRecCount += 1
            Next

            ' Desplegar mensaje que no existe email para procesar
            If nRecCount = 0 Then PrintDobleLine(String.Format("> No existen notificaciones de {0} para enviar", sCurName))

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessMensajesNotificacionesClientes(): " & sEx, 2)
        End Try

    End Sub

    Sub ProcessMensajesNotiClientesNuevo(ByVal TemplateID As Integer, _
                                             ByVal EmailFormat As String, _
                                             ByVal EmailTo As String, _
                                             ByVal EmailCcList As String, _
                                             ByVal EmailBccList As String, _
                                             ByVal EmailFrom As String, _
                                             ByVal EmailFromName As String, _
                                             ByVal EmailReplyTo As String, _
                                             ByVal EmailSubject As String, _
                                             ByVal EmailBody As String, _
                                             ByVal NotificaRepresentantes As String)

        Dim sCurName As String = "Clientes"
        Dim bSendEmail As Boolean = False

        Dim sTmp As String = String.Empty
        Dim sMessage As String = String.Empty

        Dim sCurEmailTo As String = String.Empty
        Dim sCurEmailSubject As String = String.Empty
        Dim sCurEmailBoby As String = String.Empty
        Dim sCurEmailAsesor As String = String.Empty

        Dim nMensajeID As String = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim sMensajeFechaCreado As String = Nothing

        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sNombre As String = Nothing
        Dim sApellido As String = Nothing
        Dim sTipoClienteCodigo As String = Nothing
        Dim sSubTipoClienteCodigo As String = Nothing
        Dim sClienteEstado As String = Nothing
        Dim sClienteEstatus As String = Nothing
        Dim sCedula As String = Nothing
        Dim sRNC As String = Nothing
        Dim sPasaporte As String = Nothing
        Dim sClienteEmail As String = Nothing
        Dim sEnviarEmail As String = Nothing
        Dim nCompaniaID As String = Nothing
        Dim sSucursalCodigo As String = Nothing
        Dim sAgenciaCodigo As String = Nothing
        Dim sAgenciaEmail As String = Nothing
        Dim sOficialCodigo As String = Nothing
        Dim sOficialNombre As String = Nothing
        Dim sAsesorCodigo As String = Nothing
        Dim sAsesorNombre As String = Nothing
        Dim sAsesorEmail As String = Nothing
        Dim sClienteTieneCredito As String = Nothing
        '
        Dim sDireccion1 As String
        Dim sDireccion2 As String
        Dim sCodigoVoice As String


        ' Validar si el mensaje es HTML
        Dim bIsHtml As Boolean = IIf(EmailFormat.ToUpper.Trim = "HTML", True, False)

        ' Desplegar mensaje
        PrintDobleLine(String.Format("- Buscando notificaciones de {0} para enviar, por favor espere...", sCurName))

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetMSGNotiCteNuevos(TemplateID, nAppProcesarDias)

            ' Desplegar mensaje
            PrintDobleLine(String.Format("- Procesando notificaciones de {0} para enviar...", sCurName))

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                ' Enviar correo
                bSendEmail = True

                ' Buscar los datos de la notificacion
                '------------------------------------
                nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
                sMensajeGUID = db.ewToString(dr("WMM_MENSAJE_GUID"))
                sMensajeFechaCreado = db.ewToString(dr("WMM_FECHA_CREADO"))

                sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
                sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))
                sNombre = db.ewToStringUpper(dr("CTE_NOMBRE"))
                sApellido = db.ewToStringUpper(dr("CTE_APELLIDO"))

                sTipoClienteCodigo = db.ewToStringUpper(dr("CTE_TIPO"))

             
                sClienteEmail = db.ewToStringLower(dr("CTE_EMAIL"))

                nCompaniaID = db.ewToString(dr("COM_CODIGO"))
                sSucursalCodigo = db.ewToStringUpper(dr("SUC_CODIGO"))
                sAgenciaCodigo = db.ewToStringUpper(dr("AGE_CODIGO"))
                sAgenciaEmail = db.ewToStringLower(dr("AGENCIA_EMAIL"))

                sCodigoVoice = db.ewToStringLower(dr("CTE_CODIGO_VOICE"))
                Dim sCorreoAgencia As String = db.ewToStringLower(dr("AGENCIA_EMAIL"))
                Dim sTelefonoAgencia As String = db.ewToStringLower(dr("AGE_TELEFONO"))


                ' ag.AGE_DIREC_POBOX_IND,
                ' ag.AGE_DIREC_PAQ_IND,
                ' ag.AGE_DIREC_POBOX_CORP,
                ' ag.AGE_DIREC_PAQ_CORP,
                ' vc.CTE_CODIGO_VOICE,
                         

                ' Buscar dirección de envio del mensaje
                '------------------------------------------------------------------------
                sCurEmailTo = ProcessSendToEmailAddressList(EmailTo, EmailCcList, EmailBccList, sClienteEmail, sCurEmailAsesor)

                ' Crear el asunto del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailSubject = EmailSubject

                ' Formatear cuerpo del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailBoby = EmailBody

                ' Formatear Template Datos Generales
                FormatTemplateTextDatosGenerales(sCurEmailBoby, TemplateID, bIsHtml, dr)

                FormatTemplateText(sCurEmailBoby, "EMAIL_FROM", EmailFrom)

                If sTipoClienteCodigo = "I" Then
                    Dim stmpDireccion As String = db.ewToStringLower(dr("AGE_DIREC_POBOX_IND"))
                    Dim aDir() As String = stmpDireccion.Split(Chr(10))

                    sDireccion1 = aDir(0)
                    sDireccion2 = aDir(2)


                    FormatTemplateText(sCurEmailBoby, "DIRECCION_C1", sDireccion1)
                    FormatTemplateText(sCurEmailBoby, "DIRECCION_C2", sDireccion2)

                    stmpDireccion = db.ewToStringLower(dr("AGE_DIREC_PAQ_IND"))
                    aDir = stmpDireccion.Split(Chr(10))

                    sDireccion1 = aDir(0)
                    sDireccion2 = aDir(2)


                    FormatTemplateText(sCurEmailBoby, "DIRECCION1", sDireccion1)
                    FormatTemplateText(sCurEmailBoby, "DIRECCION2", sDireccion2)

                Else

                    Dim stmpDireccion As String = db.ewToStringLower(dr("AGE_DIREC_POBOX_CORP"))
                    Dim aDir() As String = stmpDireccion.Split(Chr(10))

                    sDireccion1 = aDir(0)
                    sDireccion2 = aDir(2)


                    FormatTemplateText(sCurEmailBoby, "DIRECCION_C1", sDireccion1)
                    FormatTemplateText(sCurEmailBoby, "DIRECCION_C2", sDireccion2)

                    stmpDireccion = db.ewToStringLower(dr("AGE_DIREC_PAQ_CORP"))
                    aDir = stmpDireccion.Split(Chr(10))

                    sDireccion1 = aDir(0)
                    sDireccion2 = aDir(2)


                    FormatTemplateText(sCurEmailBoby, "DIRECCION1", sDireccion1)
                    FormatTemplateText(sCurEmailBoby, "DIRECCION2", sDireccion2)


                End If


                FormatTemplateText(sCurEmailBoby, "TELEFONO_AGENCIA", sTelefonoAgencia)
                FormatTemplateText(sCurEmailBoby, "CODIGO_VOICE", sCodigoVoice)
                FormatTemplateText(sCurEmailBoby, "CORREO_AGENCIA", sCorreoAgencia)
                FormatTemplateText(sCurEmailBoby, "NOMBRE_COMPLETO", sNombreCompleto)
                FormatTemplateText(sCurEmailBoby, "NOMBRE_COMPLETO", sNombreCompleto)
                FormatTemplateText(sCurEmailBoby, "CLIENTE_NOMBRE", sNombre)
                FormatTemplateText(sCurEmailBoby, "CLIENTE_APELLIDO", sApellido)



                ' Enviar el correo electrónico
                '------------------------------------------------------------------------
                SendEmail(TemplateID, EmailFormat, EmailTo, EmailFrom, EmailFromName, EmailFrom, sCurEmailSubject, sCurEmailBoby, bSendEmail, dr)

                ' Incrementar contador de registros procesados
                nRecCount += 1
            Next

            ' Desplegar mensaje que no existe email para procesar
            If nRecCount = 0 Then PrintDobleLine(String.Format("> No existen notificaciones de {0} para enviar", sCurName))

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessMensajesNotificacionesClientes(): " & sEx, 2)
        End Try

    End Sub

    ''' <summary>
    ''' Procesar Mensajes de Notificaciones para Clientes Credito
    ''' </summary>
    ''' <param name="TemplateID">ID del E-mail Template</param>
    ''' <param name="EmailFormat">Formato del E-mail Template</param>
    ''' <param name="EmailTo">Direccion de E-mail (Opcional)</param>
    ''' <param name="EmailCcList">Lista de Direcciones de E-mail de Copia (Opcional, Separador por Coma)</param>
    ''' <param name="EmailBccList">Lista de Direcciones de E-mail BCC (Opcional, Separador por Coma)</param>
    ''' <param name="EmailFrom">Direccion de E-mail de Envio (De)</param>
    ''' <param name="EmailFromName">Nombre de Envio</param>
    ''' <param name="EmailReplyTo">Direccion de E-mail de Respuesta</param>
    ''' <param name="EmailSubject">Asunto del Mensaje</param>
    ''' <param name="EmailBody">Mensaje completo</param>
    ''' <param name="NotificaRepresentantes">Notificar Representantes de Servicios</param>
    ''' <remarks></remarks>
    Sub ProcessMensajesNotificacionesClientesCredito(ByVal TemplateID As Integer, _
                                                     ByVal EmailFormat As String, _
                                                     ByVal EmailTo As String, _
                                                     ByVal EmailCcList As String, _
                                                     ByVal EmailBccList As String, _
                                                     ByVal EmailFrom As String, _
                                                     ByVal EmailFromName As String, _
                                                     ByVal EmailReplyTo As String, _
                                                     ByVal EmailSubject As String, _
                                                     ByVal EmailBody As String, _
                                                     ByVal NotificaRepresentantes As String)

        Dim sCurName As String = "Clientes Credito"
        Dim bSendEmail As Boolean = False

        Dim sTmp As String = String.Empty
        Dim sMessage As String = String.Empty

        Dim sCurEmailTo As String = String.Empty
        Dim sCurEmailSubject As String = String.Empty
        Dim sCurEmailBoby As String = String.Empty
        Dim sCurEmailAsesor As String = String.Empty

        Dim nMensajeID As String = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim sMensajeFechaCreado As String = Nothing

        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sTipoClienteCodigo As String = Nothing
        Dim sSubTipoClienteCodigo As String = Nothing
        Dim sClienteEstado As String = Nothing
        Dim sClienteEstatus As String = Nothing
        Dim sCedula As String = Nothing
        Dim sRNC As String = Nothing
        Dim sPasaporte As String = Nothing
        Dim sClienteEmail As String = Nothing
        Dim sEnviarEmail As String = Nothing
        Dim nCompaniaID As String = Nothing
        Dim sSucursalCodigo As String = Nothing
        Dim sAgenciaCodigo As String = Nothing
        Dim sAgenciaEmail As String = Nothing
        Dim sOficialCodigo As String = Nothing
        Dim sOficialNombre As String = Nothing
        Dim sAsesorCodigo As String = Nothing
        Dim sAsesorNombre As String = Nothing
        Dim sAsesorEmail As String = Nothing

        Dim sClienteTieneCredito As String = Nothing
        Dim sClienteTieneCreditoInlimitado As String = Nothing
        Dim nLimiteCredito As String = Nothing
        Dim nDiasCreditos As String = Nothing
        Dim nDiaCorte As String = Nothing
        Dim nCreditoDisponible As String = Nothing
        Dim nBalanceDisponible As String = Nothing
        Dim sCobradorCodigo As String = Nothing
        Dim sCobradorNombre As String = Nothing
        Dim sClasificacionCreditoCodigo As String = Nothing
        Dim sClasificacionCredito As String = Nothing

        ' Validar si el mensaje es HTML
        Dim bIsHtml As Boolean = IIf(EmailFormat.ToUpper.Trim = "HTML", True, False)

        ' Desplegar mensaje
        PrintDobleLine(String.Format("- Buscando notificaciones de {0} para enviar, por favor espere...", sCurName))

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetMensajesNotificacionesClientesCreditoDataSet(TemplateID, nAppProcesarDias)

            ' Desplegar mensaje
            PrintDobleLine(String.Format("- Procesando notificaciones de {0} para enviar...", sCurName))

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                ' Enviar correo
                bSendEmail = True

                ' Buscar los datos de la notificacion
                '------------------------------------
                nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
                sMensajeGUID = db.ewToString(dr("WMM_MENSAJE_GUID"))
                sMensajeFechaCreado = db.ewToString(dr("WMM_FECHA_CREADO"))

                sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
                sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))
                sTipoClienteCodigo = db.ewToStringUpper(dr("CTE_TIPO"))
                sSubTipoClienteCodigo = db.ewToStringUpper(dr("STC_CODIGO"))
                sClienteEstado = db.ewToStringUpper(dr("CTE_ESTADO"))
                sClienteEstatus = db.ewToStringUpper(dr("ESTATUS"))
                sCedula = db.ewToString(dr("CTE_CEDULA"))
                sRNC = db.ewToString(dr("CTE_RNC"))
                sPasaporte = db.ewToString(dr("CTE_PASAPORTE"))
                sClienteEmail = db.ewToStringLower(dr("CTE_EMAIL"))
                sEnviarEmail = db.ewToStringUpper(dr("CTE_ENVIAR_EMAIL"))
                nCompaniaID = db.ewToString(dr("COM_CODIGO"))
                sSucursalCodigo = db.ewToStringUpper(dr("SUC_CODIGO"))
                sAgenciaCodigo = db.ewToStringUpper(dr("AGE_CODIGO"))
                sAgenciaEmail = db.ewToStringLower(dr("AGENCIA_EMAIL"))
                sOficialCodigo = db.ewToStringUpper(dr("CTE_VENDEDOR"))
                sOficialNombre = db.ewToStringUpper(dr("OFICIAL"))
                sAsesorCodigo = db.ewToStringUpper(dr("RES_CODIGO"))
                sAsesorNombre = db.ewToStringUpper(dr("ASESOR"))
                sAsesorEmail = db.ewToStringLower(dr("ASESOR_EMAIL"))

                sClienteTieneCredito = db.ewToStringUpper(dr("CTE_CREDITO"))
                sClienteTieneCreditoInlimitado = db.ewToStringUpper(dr("CREDITO_INLIMITADO"))
                nLimiteCredito = db.ewToString(dr("CTE_LIMITE_CREDITO"))
                nDiasCreditos = db.ewToString(dr("CTE_DIAS_CREDITOS"))
                nDiaCorte = db.ewToString(dr("CTE_DIA_CORTE"))
                nCreditoDisponible = db.ewToString(dr("CTE_CREDITO_DISPONIBLE"))
                nBalanceDisponible = db.ewToString(dr("CTE_BALANCE_DISPONIBLE"))
                sCobradorCodigo = db.ewToStringUpper(dr("CTE_COBRADOR"))
                sCobradorNombre = db.ewToStringUpper(dr("COBRADOR"))
                sClasificacionCreditoCodigo = db.ewToStringUpper(dr("CRE_CODIGO"))
                sClasificacionCredito = db.ewToStringUpper(dr("CLASIFICACION_CREDITO"))

                ' SetUp Cliente
                ' -------------
                SetUpClienteOficial(sOficialCodigo, sOficialNombre, sSucursalCodigo, sAgenciaCodigo)
                SetUpClienteAsesor(sAsesorCodigo, sAsesorNombre, sAsesorEmail, sSucursalCodigo, sAgenciaCodigo)
                If NotificaRepresentantes = "S" Then sCurEmailAsesor = sAsesorEmail ' Enviar notificacion al asesor de la cuenta

                ' Buscar dirección de envio del mensaje
                '------------------------------------------------------------------------
                sCurEmailTo = ProcessSendToEmailAddressList(EmailTo, EmailCcList, EmailBccList, sClienteEmail, sCurEmailAsesor)

                ' Crear el asunto del mensaje a enviar
                '------------------------------------------------------------------------

                sCurEmailSubject = EmailSubject

                FormatTemplateText(sCurEmailSubject, "EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "NUMERO_EPS", sNumeroEPS)
                FormatTemplateText(sCurEmailSubject, "CTE_NUMERO_EPS", sNumeroEPS)

                ' Formatear cuerpo del mensaje a enviar
                '------------------------------------------------------------------------
                sCurEmailBoby = EmailBody

                ' Formatear Template Datos Generales
                FormatTemplateTextDatosGenerales(sCurEmailBoby, TemplateID, bIsHtml, dr)

                FormatTemplateText(sCurEmailBoby, "EMAIL_FROM", EmailFrom)

                FormatTemplateText(sCurEmailBoby, "LIMITE_CREDITO", nLimiteCredito)
                FormatTemplateText(sCurEmailBoby, "DIAS_CREDITOS", nDiasCreditos)
                FormatTemplateText(sCurEmailBoby, "DIA_CORTE", nDiaCorte)
                FormatTemplateText(sCurEmailBoby, "CREDITO_DISPONIBLE", nCreditoDisponible)
                FormatTemplateText(sCurEmailBoby, "BALANCE_DISPONIBLE", nBalanceDisponible)
                FormatTemplateText(sCurEmailBoby, "COBRADOR", sCobradorNombre)
                FormatTemplateText(sCurEmailBoby, "CLASIFICACION_CREDITO", sClasificacionCredito)

                ' Enviar el correo electrónico
                '------------------------------------------------------------------------
                SendEmail(TemplateID, EmailFormat, EmailTo, EmailFrom, EmailFromName, EmailFromName, sCurEmailSubject, sCurEmailBoby, bSendEmail, dr)

                ' Incrementar contador de registros procesados
                nRecCount += 1
            Next

            ' Desplegar mensaje que no existe email para procesar
            If nRecCount = 0 Then PrintDobleLine(String.Format("> No existen notificaciones de {0} para enviar", sCurName))

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessMensajesNotificacionesClientesCredito(): " & sEx, 2)
        End Try

    End Sub

#End Region

#Region "Traer Notificiaciones DataSets"

    ''' <summary>
    ''' Traer Mensajes de Notificaciones de Paquetes o Bultos DataSet
    ''' </summary>
    ''' <param name="TemplateID">Template ID</param>
    ''' <param name="Dias">Cantidad de dias para procesar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetMensajesNotificacionesBultosDataSet(ByVal TemplateID As Integer, ByVal Dias As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EPSWEBMAIL_MENSAJES_NOTIFICACIONES_BULTOS] @TPL_EMAIL_ID = {0}, @DIAS = {1}", _
                                           TemplateID, Dias)

        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetMensajesNotificacionesBultosDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

    Private Function GetMensajesNotificacionesBultosDataSetSgda(ByVal TemplateID As Integer, ByVal Dias As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EPSWEBMAIL_MENSAJES_NOTIFICACIONES_BULTOS2] @TPL_EMAIL_ID = {0}, @DIAS = {1}", _
                                           TemplateID, Dias)

        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetMensajesNotificacionesBultosDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

    Private Function GetMensajesNotificacionesBultosDataSetEra(ByVal TemplateID As Integer, ByVal Dias As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EPSWEBMAIL_MENSAJES_NOTIFICACIONES_BULTOS3] @TPL_EMAIL_ID = {0}, @DIAS = {1}", _
                                           TemplateID, Dias)

        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetMensajesNotificacionesBultosDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

    Public Function GetCondiciones(ByVal bltCodBarra As String, ByVal bltTrackingNumber As String) As String()
        'proc_InfoBultos_mfr2
        'proc_InfoBultos4_P
        Dim sSql As String = String.Format("EXEC [dbo].[proc_InfoBultos_mfr2] '{0}', '{1}'", _
                                         bltCodBarra, bltTrackingNumber)

        Dim condiciones(3) As String

        condiciones(0) = ""
        condiciones(1) = ""
        condiciones(2) = ""


        Try

            For Each fila As DataRow In db.ewGetDataSet(sSql).Tables(0).Rows
                condiciones(0) = fila("Condicion")
                condiciones(1) = fila("Condicion_2")
                System.Diagnostics.Debug.WriteLine(fila("Condicion_2"))
                condiciones(2) = fila("Condicion_3")
            Next
            'Return condiciones
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetMensajesNotificacionesBultosDataSet(): " & sEx, 2)
            ' Return Nothing
        End Try
        Return condiciones

    End Function

    ''' <summary>
    ''' Traer Mensajes de Notificaciones de Clientes Credtio DataSet
    ''' </summary>
    ''' <param name="TemplateID">Template ID</param>
    ''' <param name="Dias">Cantidad de dias para procesar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetMensajesNotificacionesClientesCreditoDataSet(ByVal TemplateID As Integer, ByVal Dias As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EPSWEBMAIL_MENSAJES_NOTIFICACIONES_CLIENTES_CREDITO] @TPL_EMAIL_ID = {0}, @DIAS = {1}", _
                                           TemplateID, Dias)

        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetMensajesNotificacionesClientesCreditoDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Traer Mensajes de Notificaciones de Clientes DataSet
    ''' </summary>
    ''' <param name="TemplateID">Template ID</param>
    ''' <param name="Dias">Cantidad de dias para procesar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetMensajesNotificacionesClientesDataSet(ByVal TemplateID As Integer, ByVal Dias As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EPSWEBMAIL_MENSAJES_NOTIFICACIONES_CLIENTES] @TPL_EMAIL_ID = {0}, @DIAS = {1}", _
                                           TemplateID, Dias)

        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetMensajesNotificacionesClientesDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function


    Private Function GetMSGNotiCteNuevos(ByVal TemplateID As Integer, ByVal Dias As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].proc_EPSWEBMAIL_MSG_NOTI_CTE_NUEVOS @TPL_EMAIL_ID = {0}, @DIAS = {1}", _
                                           TemplateID, Dias)

        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetMensajesNotificacionesClientesDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function


#End Region

#Region "WebMail Mensajes"

    ''' <summary>
    ''' Insertar WebMail Mensaje
    ''' </summary>
    ''' <param name="MensajeGUID"></param>
    ''' <param name="FechaCreado"></param>
    ''' <param name="TemplateID"></param>
    ''' <param name="NumeroEPS"></param>
    ''' <param name="Clave"></param>
    ''' <param name="Alterna"></param>
    ''' <param name="Tabla"></param>
    ''' <param name="NotaAdicional1"></param>
    ''' <param name="NotaAdicional2"></param>
    ''' <param name="NotaAdicional3"></param>
    ''' <param name="Estatus"></param>
    ''' <param name="FechaEnviado"></param>
    ''' <param name="Respondio"></param>
    ''' <param name="FechaRespondio"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function InsertWebMailMensaje(ByVal MensajeGUID As Object, _
                                          ByVal FechaCreado As Object, _
                                          ByVal TemplateID As Integer, _
                                          ByVal NumeroEPS As String, _
                                          ByVal Clave As String, _
                                          ByVal Alterna As String, _
                                          ByVal Tabla As String, _
                                          ByVal NotaAdicional1 As String, _
                                          ByVal NotaAdicional2 As String, _
                                          ByVal NotaAdicional3 As String, _
                                          ByVal Estatus As String, _
                                          ByVal FechaEnviado As Object, _
                                          ByVal Respondio As String, _
                                          ByVal FechaRespondio As Object) As Integer

        Dim nReturn As Integer = 0

        Dim oConn As New SqlConnection(db.GetConnectionString())
         Dim oCmd2 As New SqlCommand

        With oCmd2
            .Connection = oConn
            
            .CommandType = CommandType.Text
            .CommandText="set ARITHABORT on"
              End With
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[proc_EPSWEBMAIL_MENSAJESInsert]"

            ' '@WMM_MENSAJE_ID' int
            db.AddParameter(oCmd, "@WMM_MENSAJE_ID", db.DataType.eInteger, db.Direction.eOutput, 0)

            ' @WMM_MENSAJE_GUID UNIQUEIDENTIFIER
            db.AddParameter(oCmd, "@WMM_MENSAJE_GUID", db.DataType.eString, db.Direction.eInput, MensajeGUID)

            ' @WMM_FECHA_CREADO DATETIME
            db.AddParameter(oCmd, "@WMM_FECHA_CREADO", db.DataType.eDateTime, db.Direction.eInput, FechaCreado)

            ' @TPL_EMAIL_ID INT
            db.AddParameter(oCmd, "@TPL_EMAIL_ID", db.DataType.eInteger, db.Direction.eInput, TemplateID)

            ' @CTE_NUMERO_EPS CHAR(12)
            db.AddParameter(oCmd, "@CTE_NUMERO_EPS", NumeroEPS, 12)

            ' @WMM_CLAVE VARCHAR(16)
            db.AddParameter(oCmd, "@WMM_CLAVE", Clave, 16)

            ' @WMM_ALTERNA VARCHAR(22)
            db.AddParameter(oCmd, "@WMM_ALTERNA", Alterna, 22)

            ' @WMM_TABLA VARCHAR(32)
            db.AddParameter(oCmd, "@WMM_TABLA", Tabla, 32)

            ' @WMM_NOTA_ADICIONAL1 VARCHAR(255)
            db.AddParameter(oCmd, "@WMM_NOTA_ADICIONAL1", NotaAdicional1, 255)

            ' @WMM_NOTA_ADICIONAL2 VARCHAR(255)
            db.AddParameter(oCmd, "@WMM_NOTA_ADICIONAL2", NotaAdicional2, 255)

            ' @WMM_NOTA_ADICIONAL3 VARCHAR(255)
            db.AddParameter(oCmd, "@WMM_NOTA_ADICIONAL3", NotaAdicional3, 255)

            ' @WMM_ESTATUS CHAR(1)
            db.AddParameter(oCmd, "@WMM_ESTATUS", Estatus, 1)

            ' @WMM_FECHA_ENVIADO DATETIME
            db.AddParameter(oCmd, "@WMM_FECHA_ENVIADO", db.DataType.eDateTime, db.Direction.eInput, FechaEnviado)

            ' @WMM_RESPONDIO CHAR(1)
            db.AddParameter(oCmd, "@WMM_RESPONDIO", Respondio, 1)

            ' @WMM_FECHA_RESPONDIO DATETIME
            db.AddParameter(oCmd, "@WMM_FECHA_RESPONDIO", db.DataType.eDateTime, db.Direction.eInput, FechaRespondio)

        End With

        Try
            oConn.Open()
            oCmd2.ExecuteNonQuery()
            oCmd.ExecuteNonQuery()

            ' To get a procedure's return value
            nReturn = CType(oCmd.Parameters("@WMM_MENSAJE_ID").Value, Integer)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "InsertWebMailMensaje(): " & sEx, 2)
        Finally
            oConn.Close()
        End Try

        Return nReturn

    End Function

    ''' <summary>
    ''' Actualizar WebMail Mensajes Enviados
    ''' </summary>
    ''' <param name="MensajeID"></param>
    ''' <param name="Estatus"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UpdateWebMailMensajesEnviado(ByVal MensajeID As Integer, ByVal Estatus As String) As Boolean

        Dim bReturn As Boolean = True
        Dim oConn As New SqlConnection(db.GetConnectionString)
           Dim oCmd2 As New SqlCommand

        With oCmd2
            .Connection = oConn
            
            .CommandType = CommandType.Text
            .CommandText="set ARITHABORT on"
              End With
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[proc_EPSWEBMAIL_MENSAJES_UpdateEnviado]"

            ' @WMM_MENSAJE_ID int
            db.AddParameter(oCmd, "@WMM_MENSAJE_ID", db.DataType.eInteger, db.Direction.eInput, MensajeID)

            ' @WMM_ESTATUS char(1) = NULL
            db.AddParameter(oCmd, "@WMM_ESTATUS", Estatus, 1)

            ' @WMM_FECHA_ENVIADO datetime = NULL
            db.AddParameter(oCmd, "@WMM_FECHA_ENVIADO", db.DataType.eDateTime, db.Direction.eInput, Nothing)
        End With

        Try
            oConn.Open()
            oCmd2.ExecuteNonQuery()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "UpdateWebMailMensajesEnviado(): " & sEx, 2)
            bReturn = False
        Finally
            oConn.Close()
        End Try

        Return bReturn

    End Function

    ''' <summary>
    ''' Actualizar WebMail Mensajes Cliente Respondio
    ''' </summary>
    ''' <param name="MensajeID"></param>
    ''' <param name="Respondio"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UpdateWebMailMensajesRespondio(ByVal MensajeID As Integer, ByVal Respondio As String) As Boolean

        Dim bReturn As Boolean = True
        Dim oConn As New SqlConnection(db.GetConnectionString)
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[proc_EPSWEBMAIL_MENSAJES_UpdateRespondio]"

            ' @WMM_MENSAJE_ID int
            db.AddParameter(oCmd, "@WMM_MENSAJE_ID", db.DataType.eInteger, db.Direction.eInput, MensajeID)

            ' @WMM_RESPONDIO char(1) = NULL
            db.AddParameter(oCmd, "@WMM_RESPONDIO", Respondio, 1)

            ' @WMM_FECHA_RESPONDIO datetime = NULL
            db.AddParameter(oCmd, "@WMM_FECHA_RESPONDIO", db.DataType.eDateTime, db.Direction.eInput, Nothing)
        End With

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "UpdateWebMailMensajesRespondio(): " & sEx, 2)
            bReturn = False
        Finally
            oConn.Close()
        End Try

        Return bReturn

    End Function

    ''' <summary>
    ''' Actualizar WebMail Mensajes Sin Responder 
    ''' </summary>
    ''' <param name="TemplateID"></param>
    ''' <param name="Dias"></param>
    ''' <remarks></remarks>
    Private Sub UpdateWebMailNotificacionesSinResponder(ByVal TemplateID As Integer, ByVal Dias As Integer)

        ' Desplegar mensaje
        PrintDobleLine("- Procesando Mensajes de Notificaciones sin responder, por favor espere...")

        Dim oConn As New SqlConnection(db.GetConnectionString)
           Dim oCmd2 As New SqlCommand

        With oCmd2
            .Connection = oConn
            
            .CommandType = CommandType.Text
            .CommandText="set ARITHABORT on"
              End With
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[proc_EPSWEBMAIL_MENSAJES_NOTIFICACIONES_SINRESPONDER_P]"

            ' @TPL_EMAIL_ID int
            db.AddParameter(oCmd, "@TPL_EMAIL_ID", db.DataType.eInteger, db.Direction.eInput, TemplateID)

            ' @DIAS int
            db.AddParameter(oCmd, "@DIAS", db.DataType.eInteger, db.Direction.eInput, Dias)
        End With

        Try
            oConn.Open()
             oCmd2.ExecuteNonQuery()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "UpdateWebMailNotificacionesSinResponder(): " & sEx, 2)
        Finally
            oConn.Close()
        End Try

    End Sub
    Private Sub UpdateWebMailNotificacionesSinResponderSgda(ByVal TemplateID As Integer, ByVal Dias As Integer)

        ' Desplegar mensaje
        PrintDobleLine("- Procesando Mensajes de Notificaciones sin responder, por favor espere...")

        Dim oConn As New SqlConnection(db.GetConnectionString)
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[proc_EPSWEBMAIL_MENSAJES_NOTIFICACIONES_SINRESPONDER_P2]"

            ' @TPL_EMAIL_ID int
            db.AddParameter(oCmd, "@TPL_EMAIL_ID", db.DataType.eInteger, db.Direction.eInput, TemplateID)

            ' @DIAS int
            db.AddParameter(oCmd, "@DIAS", db.DataType.eInteger, db.Direction.eInput, Dias)
        End With

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "UpdateWebMailNotificacionesSinResponder(): " & sEx, 2)
        Finally
            oConn.Close()
        End Try

    End Sub
#End Region

#Region "WebMail Mensajes Importar"

    ''' <summary>
    ''' Import Messages
    ''' </summary>
    ''' <param name="TemplateID"></param>
    ''' <remarks></remarks>
    Sub ImportMessages(ByVal TemplateID As Integer, Optional ByVal distincion As String = "", Optional ByVal distincionEra As String = "")

        Dim sCurName As String = "Mensajes de Notificaciones"

        Dim sMessage As String = String.Empty


        Dim nReturn As Integer = 0
        Dim nImportCount As Integer = 0

        Dim nMensajeID As Integer = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim dFechaCreado As Object = Nothing
        Dim sNumeroEPS As String = Nothing
        Dim sClave As String = Nothing
        Dim sAlterna As String = Nothing
        Dim sTabla As String = Nothing
        Dim sNotaAdicional1 As String = Nothing
        Dim sNotaAdicional2 As String = Nothing
        Dim sNotaAdicional3 As String = Nothing
        Dim sEstatus As String = Nothing
        Dim dFechaEnviado As Object = Nothing
        Dim sRespondio As String = Nothing
        Dim dFechaRespondio As Object = Nothing
        Dim sOperacionReplica As String = Nothing
        Dim sReplicaID As String = Nothing

        ' Desplegar mensaje
        PrintDobleLine(String.Format("- Importando {0}, por favor espere", sCurName))

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetImportDataSet(TemplateID, nAppProcesarDias, nAppProcesarRegistros)

            ' Desplegar mensaje
            PrintDobleLine(String.Format("- Procesando importación {0}...", sCurName))

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                Dim sufijo As String = db.ewToString(dr("WMM_MENSAJE_GUID"))

                If String.IsNullOrEmpty(distincion) <> True Then
                    sufijo += distincion
                End If

                If (String.IsNullOrEmpty(distincionEra)) <> True Then
                    sufijo += distincionEra
                End If

                ' Buscar los datos de la importacion
                nMensajeID = db.ewToInteger(dr("WMM_MENSAJE_ID"))
                sMensajeGUID = sufijo
                dFechaCreado = db.ewToDateTime(dr("WMM_FECHA_CREADO"))
                sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
                sClave = db.ewToStringUpper(dr("WMM_CLAVE"))
                sAlterna = db.ewToStringUpper(dr("WMM_ALTERNA"))
                sTabla = db.ewToString(dr("WMM_TABLA"))
                sNotaAdicional1 = db.ewToString(dr("WMM_NOTA_ADICIONAL1"))
                sNotaAdicional2 = db.ewToString(dr("WMM_NOTA_ADICIONAL2"))
                sNotaAdicional3 = db.ewToString(dr("WMM_NOTA_ADICIONAL3"))
                sEstatus = db.ewToString(dr("WMM_ESTATUS"))
                dFechaEnviado = db.ewToDateTime(dr("WMM_FECHA_ENVIADO"))
                sRespondio = db.ewToString(dr("WMM_RESPONDIO"))
                dFechaRespondio = db.ewToDateTime(dr("WMM_FECHA_RESPONDIO"))
                sOperacionReplica = db.ewToString(dr("OPERACION_REPLICA"))
                sReplicaID = db.ewToString(dr("REP_ID"))

                ' Insertar WebMail Mensaje
                nReturn = InsertWebMailMensaje(sMensajeGUID, _
                                               dFechaCreado, _
                                               TemplateID, _
                                               sNumeroEPS, _
                                               sClave, _
                                               sAlterna, _
                                               sTabla, _
                                               sNotaAdicional1, _
                                               sNotaAdicional2, _
                                               sNotaAdicional3, _
                                               sEstatus, _
                                               dFechaEnviado, _
                                               sRespondio, _
                                               dFechaRespondio)
                If nReturn > 0 Then
                    ' Desplegar mensaje
                    PrintLine(String.Format("EPS: {0} -> CODIGO: {1}", sNumeroEPS.PadRight(12), sClave))

                    ' Actualizar replica
                    UpdateImportMessages(TemplateID, sNumeroEPS, sClave, sAlterna, nAppProcesarDias)

                    ' Incrementar contador de registros procesados
                    nImportCount += 1
                End If
            Next

            ' Desplegar mensaje que no existe email para procesar
            If nImportCount = 0 Then
                PrintDobleLine(String.Format("> No existen {0} para importar", sCurName))
            End If

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ImportMessages(): " & sEx, 2)
        End Try

    End Sub

    ''' <summary>
    ''' Traer Mensajes de Notificaciones a Importar DataSet
    ''' </summary>
    ''' <param name="TemplateID"></param>
    ''' <param name="Dias"></param>
    ''' <param name="Cantidad"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetImportDataSet(ByVal TemplateID As Integer, ByVal Dias As Integer, ByVal Cantidad As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EmailNotify_EPSWEBMAIL_MENSAJESLoadAll] @TPL_EMAIL_ID = {0}, @DIAS = {1}, @CANTIDAD = {2}", _
                                           TemplateID, Dias, Cantidad)

        Try
            Return db.ewGetDataSet(sSql, GetImportConnectionString())
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetImportDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Actualizar Mensajes de Notificaciones Importados
    ''' </summary>
    ''' <param name="TemplateID"></param>
    ''' <param name="NumeroEPS"></param>
    ''' <param name="Clave"></param>
    ''' <param name="Alterna"></param>
    ''' <param name="Dias"></param>
    ''' <remarks></remarks>
    Private Sub UpdateImportMessages(ByVal TemplateID As Integer, ByVal NumeroEPS As String, ByVal Clave As String, ByVal Alterna As String, ByVal Dias As Integer)

        Dim oConn As New SqlConnection(GetImportConnectionString())
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[proc_EmailNotify_EPSWEBMAIL_MENSAJESUpdate]"

            ' @TPL_EMAIL_ID int
            db.AddParameter(oCmd, "@TPL_EMAIL_ID", db.DataType.eInteger, db.Direction.eInput, TemplateID)

            ' @CTE_NUMERO_EPS CHAR(12)
            db.AddParameter(oCmd, "@CTE_NUMERO_EPS", NumeroEPS, 12)

            ' @WMM_CLAVE VARCHAR(16)
            db.AddParameter(oCmd, "@WMM_CLAVE", Clave, 16)

            ' @WMM_ALTERNA VARCHAR(22)
            db.AddParameter(oCmd, "@WMM_ALTERNA", Alterna, 22)

            ' @DIAS int
            db.AddParameter(oCmd, "@DIAS", db.DataType.eInteger, db.Direction.eInput, Dias)
        End With

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "UpdateImportMessages(): " & sEx, 2)
        Finally
            oConn.Close()
        End Try

    End Sub

#End Region

#Region "WebMail Templates"

    ''' <summary>
    ''' Actualizar WebMail Templates Corrida
    ''' </summary>
    ''' <param name="TemplateID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UpdateWebMailTemplatesCorrida(ByVal TemplateID As Integer) As Boolean

        Dim bReturn As Boolean = True
        Dim oConn As New SqlConnection(db.GetConnectionString)
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[proc_EPSWEBMAIL_TEMPLATESUpdateFechaCorrida]"

            ' @TPL_EMAIL_ID int
            db.AddParameter(oCmd, "@TPL_EMAIL_ID", db.DataType.eInteger, db.Direction.eInput, TemplateID)

            ' @TPL_FECHA_CORRIDA datetime = NULL
            db.AddParameter(oCmd, "@TPL_FECHA_CORRIDA", db.DataType.eDateTime, db.Direction.eInput, Nothing)
        End With

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "UpdateWebMailTemplatesCorrida(): " & sEx, 2)
            bReturn = False
        End Try

        Return bReturn
    End Function

    ''' <summary>
    ''' Traer DataSet de WebMail Template por Llave Primaria
    ''' </summary>
    ''' <param name="TemplateID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetWebMailTemplateByIdDataSet(ByVal TemplateID As Integer) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EPSWEBMAIL_TEMPLATESLoadByPrimaryKey] @TPL_EMAIL_ID = {0}", _
                                           TemplateID)

        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetWebMailTemplateByIdDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Traer DataSet de Todos los WebMail Template
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetWebMailTemplateDataSet() As DataSet

        Dim sSql As String = "EXEC [dbo].[proc_EPSWEBMAIL_TEMPLATESLoadAll]"

        '  sSql = "SELECT * FROM EPSWEBMAIL_TEMPLATES WHERE TPL_EMAIL_ID = 36 "


        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetWebMailTemplateDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

#End Region

#Region "SetUp and Format Template Text"

    ''' <summary>
    ''' Get Cliente Credito Contacto de la Empresa
    ''' </summary>
    ''' <param name="EncargadoPagos"></param>
    ''' <param name="Representante"></param>
    ''' <param name="SucursalCodigo"></param>
    ''' <param name="AgenciaCodigo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetClienteEmpresaContacto(ByVal EncargadoPagos As String, ByVal Representante As String, ByVal SucursalCodigo As String, ByVal AgenciaCodigo As String) As String

        If Not String.IsNullOrEmpty(EncargadoPagos) Then
            Return EncargadoPagos
        Else
            Return Representante
        End If

    End Function

    ''' <summary>
    ''' Get Cliente Telefono Oficina con Extension
    ''' </summary>
    ''' <param name="Telefono"></param>
    ''' <param name="Ext"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetClienteTelefonoOficina(ByVal Telefono As String, ByVal Ext As String) As String

        If Not String.IsNullOrEmpty(Telefono) And Not String.IsNullOrEmpty(Ext) Then
            Return String.Format("{0}, {1}", Telefono, Ext)
        Else
            Return Telefono
        End If

    End Function


    ''' <summary>
    ''' SetUp Cliente Oficial
    ''' </summary>
    ''' <param name="OficialCodigo">Asesor Codigo</param>
    ''' <param name="OficialNombre">Asesor Nombre</param>
    ''' <param name="SucursalCodigo">Sucursal Codigo</param>
    ''' <param name="AgenciaCodigo">Agencia Codigo</param>
    ''' <remarks></remarks>
    Private Sub SetUpClienteOficial(ByRef OficialCodigo As String, ByRef OficialNombre As String, ByVal SucursalCodigo As String, ByVal AgenciaCodigo As String)

        If Not String.IsNullOrEmpty(OficialNombre) Then OficialNombre = OficialNombre.Replace("¥", "Ñ")

    End Sub

    ''' <summary>
    ''' SetUp Cliente Asesor
    ''' </summary>
    ''' <param name="AsesorCodigo">Asesor Codigo</param>
    ''' <param name="AsesorNombre">Asesor Nombre</param>
    ''' <param name="AsesorEmail">Asesor Email</param>
    ''' <param name="SucursalCodigo">Sucursal Codigo</param>
    ''' <param name="AgenciaCodigo">Agencia Codigo</param>
    ''' <remarks></remarks>
    Private Sub SetUpClienteAsesor(ByRef AsesorCodigo As String, ByRef AsesorNombre As String, ByRef AsesorEmail As String, ByVal SucursalCodigo As String, ByVal AgenciaCodigo As String)

        If (SucursalCodigo.ToUpper.Trim = "001") AndAlso (AgenciaCodigo.ToUpper.Trim <> "EPS") Then
            AsesorCodigo = My.Settings.AgenciasAsesorCodigo
            AsesorNombre = My.Settings.AgenciasAsesorNombre
            AsesorEmail = My.Settings.AgenciasAsesorEmail
        End If

        If Not String.IsNullOrEmpty(AsesorNombre) Then AsesorNombre = AsesorNombre.Replace("¥", "Ñ")

    End Sub

    ''' <summary>
    ''' Formatear texto y reemplazar textos datos generales de la notificiacion
    ''' </summary>
    ''' <param name="sInput">Email Body</param>
    ''' <param name="TemplateID">Template ID</param>
    ''' <param name="bIsHtml">Is HTML?</param>
    ''' <param name="dr">Table Data Row</param>
    ''' <remarks></remarks>
    Private Sub FormatTemplateTextDatosGenerales(ByRef sInput As String, ByVal TemplateID As Integer, ByVal bIsHtml As Boolean, ByVal dr As DataRow)

        Dim sTmp As String = String.Empty

        Dim nMensajeID As String = Nothing
        Dim sMensajeGUID As String = Nothing
        Dim sMensajeFechaCreado As String = Nothing

        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sTipoClienteCodigo As String = Nothing
        Dim sTipoCliente As String = Nothing
        Dim sSubTipoClienteCodigo As String = Nothing
        Dim sSubTipoCliente As String = Nothing
        Dim sClienteEstado As String = Nothing
        Dim sClienteEstatus As String = Nothing
        Dim sCedula As String = Nothing
        Dim sRNC As String = Nothing
        Dim sPasaporte As String = Nothing
        Dim sClienteEmail As String = Nothing
        Dim sDireccion As String = Nothing
        Dim sDireccionCasa As String = Nothing
        Dim sDireccionOficina As String = Nothing
        Dim sSectorCodigo As String = Nothing
        Dim sSector As String = Nothing
        Dim sCiudadCodigo As String = Nothing
        Dim sCiudad As String = Nothing
        Dim sPaisCodigo As String = Nothing
        Dim sPais As String = Nothing
        Dim sTelefono As String = Nothing
        Dim sTelefonoCasa As String = Nothing
        Dim sTelefonoOficina As String = Nothing
        Dim sTelefonoOficinaExt As String = Nothing
        Dim sTelefonoOficinaCompleto As String = Nothing
        Dim sFax As String = Nothing
        Dim sCelular As String = Nothing
        Dim sFechaNacimiento As String = Nothing
        Dim nCompaniaID As String = Nothing
        Dim sCompania As String = Nothing
        Dim sSucursalCodigo As String = Nothing
        Dim sSucursal As String = Nothing
        Dim sAgenciaCodigo As String = Nothing
        Dim sAgencia As String = Nothing
        Dim sAgenciaEmail As String = Nothing
        Dim sOficialCodigo As String = Nothing
        Dim sOficialNombre As String = Nothing
        Dim sAsesorCodigo As String = Nothing
        Dim sAsesorNombre As String = Nothing
        Dim sAsesorEmail As String = Nothing
        Dim sEmpresaNombre As String = Nothing
        Dim sEmpresaRepresentante As String = Nothing
        Dim sEmpresaSecretaria As String = Nothing
        Dim sEmpresaEncargadoPagos As String = Nothing
        Dim sEmpresaContacto As String = Nothing

        Dim sNotaAdicional1 As String = Nothing
        Dim sNotaAdicional2 As String = Nothing
        Dim sNotaAdicional3 As String = Nothing

        ' Buscar los datos de la notificacion
        nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
        sMensajeGUID = db.ewToString(dr("WMM_MENSAJE_GUID"))
        sMensajeFechaCreado = db.ewToString(dr("WMM_FECHA_CREADO"))

        sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
        sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))
        sTipoClienteCodigo = db.ewToStringUpper(dr("CTE_TIPO"))
        sTipoCliente = db.ewToStringUpper(dr("TIPO_CLIENTE"))
        sSubTipoClienteCodigo = db.ewToStringUpper(dr("STC_CODIGO"))
        sSubTipoCliente = db.ewToStringUpper(dr("SUBTIPO_CLIENTE"))
        sClienteEstado = db.ewToStringUpper(dr("CTE_ESTADO"))
        sClienteEstatus = db.ewToStringUpper(dr("ESTATUS"))
        sCedula = db.ewToString(dr("CTE_CEDULA"))
        sRNC = db.ewToString(dr("CTE_RNC"))
        sPasaporte = db.ewToString(dr("CTE_PASAPORTE"))
        sClienteEmail = db.ewToStringLower(dr("CTE_EMAIL"))
        sDireccion = db.ewToStringUpper(dr("DIRECCION"))
        sDireccionCasa = db.ewToStringUpper(dr("CTE_DIRECCION_CASA"))
        sDireccionOficina = db.ewToStringUpper(dr("CTE_DIRECCION_OFICINA"))
        sSectorCodigo = db.ewToStringUpper(dr("CTE_SECTOR"))
        sSector = db.ewToStringUpper(dr("SECTOR"))
        sCiudadCodigo = db.ewToStringUpper(dr("CTE_CIUDAD"))
        sCiudad = db.ewToStringUpper(dr("CIUDAD"))
        sPaisCodigo = db.ewToStringUpper(dr("COD_PAIS"))
        sPais = db.ewToStringUpper(dr("PAIS"))
        sTelefono = db.ewToString(dr("TELEFONO"))
        sTelefonoCasa = db.ewToString(dr("CTE_TELEFONO_CASA"))
        sTelefonoOficina = db.ewToString(dr("CTE_TELEFONO_OFICINA"))
        sTelefonoOficinaExt = db.ewToString(dr("CTE_EXT_TELOFIC"))
        sFax = db.ewToString(dr("CTE_FAX"))
        sCelular = db.ewToString(dr("CTE_CELULAR"))
        sFechaNacimiento = db.ewToString(dr("CTE_FECHA_NACIMIENTO"))
        nCompaniaID = db.ewToString(dr("COM_CODIGO"))
        sCompania = db.ewToStringUpper(dr("COMPANIA"))
        sSucursalCodigo = db.ewToStringUpper(dr("SUC_CODIGO"))
        sSucursal = db.ewToStringUpper(dr("SUCURSAL"))
        sAgenciaCodigo = db.ewToStringUpper(dr("AGE_CODIGO"))
        sAgencia = db.ewToStringUpper(dr("AGENCIA"))
        sAgenciaEmail = db.ewToStringLower(dr("AGENCIA_EMAIL"))
        sOficialCodigo = db.ewToStringUpper(dr("CTE_VENDEDOR"))
        sOficialNombre = db.ewToStringUpper(dr("OFICIAL"))
        sAsesorCodigo = db.ewToStringUpper(dr("RES_CODIGO"))
        sAsesorNombre = db.ewToStringUpper(dr("ASESOR"))
        sAsesorEmail = db.ewToStringLower(dr("ASESOR_EMAIL"))
        sEmpresaNombre = db.ewToStringUpper(dr("CTE_NOMBRECOMPANIA"))
        sEmpresaRepresentante = db.ewToStringUpper(dr("CTE_REPRESENTANTE"))
        sEmpresaSecretaria = db.ewToStringUpper(dr("CTE_SECRETARIA"))
        sEmpresaEncargadoPagos = db.ewToStringUpper(dr("CTE_ENCARGADO_PAGOS"))

        sNotaAdicional1 = db.ewToStringNullable(dr("WMM_NOTA_ADICIONAL1"))
        sNotaAdicional2 = db.ewToStringNullable(dr("WMM_NOTA_ADICIONAL2"))
        sNotaAdicional3 = db.ewToStringNullable(dr("WMM_NOTA_ADICIONAL3"))

        ' SetUp Cliente
        ' -------------
        sTelefonoOficinaCompleto = GetClienteTelefonoOficina(sTelefonoOficina, sTelefonoOficinaExt)
        sEmpresaContacto = GetClienteEmpresaContacto(sEmpresaEncargadoPagos, sEmpresaRepresentante, sSucursalCodigo, sAgenciaCodigo)

        SetUpClienteOficial(sOficialCodigo, sOficialNombre, sSucursalCodigo, sAgenciaCodigo)
        SetUpClienteAsesor(sAsesorCodigo, sAsesorNombre, sAsesorEmail, sSucursalCodigo, sAgenciaCodigo)

        ' HTML Encoding
        ' -------------
        If bIsHtml = True Then
            If Not String.IsNullOrEmpty(sNombreCompleto) Then sNombreCompleto = ew_EncodeText(sNombreCompleto)
            If Not String.IsNullOrEmpty(sDireccion) Then sDireccion = ew_EncodeText(sDireccion)
            If Not String.IsNullOrEmpty(sDireccionCasa) Then sDireccion = ew_EncodeText(sDireccionCasa)
            If Not String.IsNullOrEmpty(sDireccionOficina) Then sDireccion = ew_EncodeText(sDireccionOficina)
            If Not String.IsNullOrEmpty(sOficialNombre) Then sOficialNombre = ew_EncodeText(sOficialNombre)
            If Not String.IsNullOrEmpty(sAsesorNombre) Then sAsesorNombre = ew_EncodeText(sAsesorNombre)
            If Not String.IsNullOrEmpty(sEmpresaNombre) Then sEmpresaNombre = ew_EncodeText(sEmpresaNombre)
            If Not String.IsNullOrEmpty(sEmpresaRepresentante) Then sEmpresaRepresentante = ew_EncodeText(sEmpresaRepresentante)
            If Not String.IsNullOrEmpty(sEmpresaSecretaria) Then sEmpresaSecretaria = ew_EncodeText(sEmpresaSecretaria)
            If Not String.IsNullOrEmpty(sEmpresaEncargadoPagos) Then sEmpresaEncargadoPagos = ew_EncodeText(sEmpresaEncargadoPagos)
            If Not String.IsNullOrEmpty(sNotaAdicional1) Then sNotaAdicional1 = ew_EncodeText(sNotaAdicional1)
            If Not String.IsNullOrEmpty(sNotaAdicional2) Then sNotaAdicional2 = ew_EncodeText(sNotaAdicional2)
            If Not String.IsNullOrEmpty(sNotaAdicional3) Then sNotaAdicional3 = ew_EncodeText(sNotaAdicional3)
        End If

        FormatTemplateText(sInput, "MENSAJE_ID", nMensajeID)
        FormatTemplateText(sInput, "MENSAJE_GUID", sMensajeGUID)
        FormatTemplateText(sInput, "FECHA_CREADO", Utilities.DataFormat.ewDateTimeFormat(17, "/", sMensajeFechaCreado))

        ' Numero de EPS
        FormatTemplateText(sInput, "EPS", sNumeroEPS)
        FormatTemplateText(sInput, "NUMERO_EPS", sNumeroEPS)
        FormatTemplateText(sInput, "CTE_NUMERO_EPS", sNumeroEPS)

        ' Cliente Nombre Completo
        FormatTemplateText(sInput, "NOMBRE_COMPLETO", sNombreCompleto)
        FormatTemplateText(sInput, "CLIENTE", String.Format("{0} / {1}", sNumeroEPS, sNombreCompleto))

        ' Cliente Tipo Descripcion
        FormatTemplateText(sInput, "CTE_TIPO", sTipoCliente)
        FormatTemplateText(sInput, "TIPO_CLIENTE", sTipoCliente)

        ' Cliente SubTipo Descripcion
        FormatTemplateText(sInput, "SUBTIPO_CLIENTE", sSubTipoCliente)

        ' Cliente Estado
        FormatTemplateText(sInput, "CTE_ESTADO", sClienteEstado)

        ' Cliente Estatus
        FormatTemplateText(sInput, "CTE_ESTATUS", sClienteEstatus)

        ' Cliente numero de Cedula
        FormatTemplateText(sInput, "CEDULA", sCedula)
        FormatTemplateText(sInput, "CTE_CEDULA", sCedula)

        ' Cliente numero de RNC
        FormatTemplateText(sInput, "RNC", sRNC)
        FormatTemplateText(sInput, "CTE_RNC", sRNC)

        ' Cliente numero de Pasaporte
        FormatTemplateText(sInput, "PASAPORTE", sPasaporte)
        FormatTemplateText(sInput, "CTE_PASAPORTE", sPasaporte)

        ' Clientes Emails
        FormatTemplateText(sInput, "EMAIL", sClienteEmail)
        FormatTemplateText(sInput, "CTE_EMAIL", sClienteEmail)

        ' Cliente Direccion Principal (Dependiendo del tipo de cliente)
        FormatTemplateText(sInput, "DIRECCION", sDireccion)
        FormatTemplateText(sInput, "CTE_DIRECCION", sDireccion)

        ' Cliente Direccion de la Casa
        FormatTemplateText(sInput, "DIRECCION_CASA", sDireccionCasa)
        FormatTemplateText(sInput, "CTE_DIRECCION_CASA", sDireccionCasa)

        ' Cliente Direccion de la Oficina
        FormatTemplateText(sInput, "DIRECCION_OFICINA", sDireccionOficina)
        FormatTemplateText(sInput, "CTE_DIRECCION_OFICINA", sDireccionOficina)

        ' Cliente Sector
        FormatTemplateText(sInput, "SECTOR", sSector)
        FormatTemplateText(sInput, "CTE_SECTOR", sSector)

        ' Cliente Ciudad
        FormatTemplateText(sInput, "CIUDAD", sCiudad)
        FormatTemplateText(sInput, "CTE_CIUDAD", sCiudad)

        ' Cliente Pais
        FormatTemplateText(sInput, "PAIS", sPais)
        FormatTemplateText(sInput, "CTE_PAIS", sPais)

        ' Cliente Telefono Principal (Dependiendo del tipo de cliente)
        FormatTemplateText(sInput, "TELEFONO", sTelefono)
        FormatTemplateText(sInput, "CTE_TELEFONO", sTelefono)
        FormatTemplateText(sInput, "CTE_TELEFONOS", sTelefono)

        ' Cliente Telefono de la Casa
        FormatTemplateText(sInput, "TELEFONO_CASA", sTelefonoCasa)
        FormatTemplateText(sInput, "CTE_TELEFONO_CASA", sTelefonoCasa)
        FormatTemplateText(sInput, "TEL_CASA", sTelefonoCasa)
        FormatTemplateText(sInput, "CTE_TEL_CASA", sTelefonoCasa)

        ' Cliente Telefono de la Oficina
        FormatTemplateText(sInput, "TELEFONO_OFICINA", sTelefonoOficina)
        FormatTemplateText(sInput, "CTE_TELEFONO_OFICINA", sTelefonoOficina)
        FormatTemplateText(sInput, "TEL_OFICINA", sTelefonoOficina)
        FormatTemplateText(sInput, "CTE_TEL_OFICINA", sTelefonoOficina)

        ' Cliente Telefono de la Oficina Extension
        FormatTemplateText(sInput, "EXT_TELOFIC", sTelefonoOficinaExt)
        FormatTemplateText(sInput, "CTE_EXT_TELOFIC", sTelefonoOficinaExt)
        FormatTemplateText(sInput, "EXT_TEL_OFIC", sTelefonoOficinaExt)
        FormatTemplateText(sInput, "CTE_EXT_TEL_OFIC", sTelefonoOficinaExt)

        ' Cliente Telefono de la Oficina Completo
        FormatTemplateText(sInput, "TELEFONO_OFICINA_COMPLETO", sTelefonoOficinaCompleto)
        FormatTemplateText(sInput, "CTE_TELEFONO_OFICINA_COMPLETO", sTelefonoOficinaCompleto)
        FormatTemplateText(sInput, "TEL_OFICINA_COMPLETO", sTelefonoOficinaCompleto)
        FormatTemplateText(sInput, "CTE_TEL_OFICINA_COMPLETO", sTelefonoOficinaCompleto)

        ' Cliente Fax
        FormatTemplateText(sInput, "FAX", sFax)
        FormatTemplateText(sInput, "CTE_FAX", sFax)

        ' Cliente Celular
        FormatTemplateText(sInput, "CELULAR", sCelular)
        FormatTemplateText(sInput, "CTE_CELULAR", sCelular)

        ' Compania codigo y descripcion
        FormatTemplateText(sInput, "COM_CODIGO", nCompaniaID)
        FormatTemplateText(sInput, "COM_DESCRIPCION", sCompania)
        FormatTemplateText(sInput, "COMPANIA", sCompania)

        ' Sucursal codigo y descripcion
        FormatTemplateText(sInput, "SUC_CODIGO", sSucursalCodigo)
        FormatTemplateText(sInput, "SUC_DESCRIPCION", sSucursal)
        FormatTemplateText(sInput, "SUCURSAL", sSucursal)

        ' Agencia codigo y descripcion
        FormatTemplateText(sInput, "AGE_CODIGO", sAgenciaCodigo)
        FormatTemplateText(sInput, "AGE_DESCRIPCION", sAgencia)
        FormatTemplateText(sInput, "AGENCIA", sAgencia)

        ' Agencia email
        FormatTemplateText(sInput, "AGENCIA_EMAIL", sAgenciaEmail)

        ' Oficial codigo y nombre
        FormatTemplateText(sInput, "CTE_VENDEDOR", sOficialCodigo)
        FormatTemplateText(sInput, "OFI_CODIGO", sOficialCodigo)
        FormatTemplateText(sInput, "OFICIAL", sOficialNombre)

        ' Asesor codigo y nombre
        FormatTemplateText(sInput, "RES_CODIGO", sAsesorCodigo)
        FormatTemplateText(sInput, "ASE_CODIGO", sAsesorCodigo)
        FormatTemplateText(sInput, "ASESOR", sAsesorNombre)

        ' Asesor email
        FormatTemplateText(sInput, "RES_EMAIL", sAsesorEmail)
        FormatTemplateText(sInput, "ASESOR_EMAIL", sAsesorEmail)

        ' Empresa nombre
        FormatTemplateText(sInput, "EMPRESA_NOMBRE", sEmpresaNombre)
        FormatTemplateText(sInput, "NOMBRECOMPANIA", sEmpresaNombre)

        ' Empresa representante
        FormatTemplateText(sInput, "EMPRESA_REPRESENTANTE", sEmpresaRepresentante)
        FormatTemplateText(sInput, "CTE_REPRESENTANTE", sEmpresaRepresentante)

        ' Empresa secretaria
        FormatTemplateText(sInput, "EMPRESA_SECRETARIA", sEmpresaSecretaria)
        FormatTemplateText(sInput, "CTE_SECRETARIA", sEmpresaSecretaria)

        ' Empresa encargado de pagos
        FormatTemplateText(sInput, "EMPRESA_ENCARGADO_PAGOS", sEmpresaEncargadoPagos)
        FormatTemplateText(sInput, "CTE_ENCARGADO_PAGOS", sEmpresaEncargadoPagos)

        ' Empresa contacto
        FormatTemplateText(sInput, "EMPRESA_CONTACTO", sEmpresaContacto)
        FormatTemplateText(sInput, "CTE_EMPRESA_CONTACTO", sEmpresaContacto)

        ' Nota adicionales
        FormatTemplateText(sInput, "MOTIVO", sNotaAdicional1) ' Formatear texto actualizar los motivos
        FormatTemplateText(sInput, "NOTA_ADICIONAL1", sNotaAdicional1)
        FormatTemplateText(sInput, "NOTA_ADICIONAL2", sNotaAdicional2)
        FormatTemplateText(sInput, "NOTA_ADICIONAL3", sNotaAdicional3)

        ' Crear mensaje de contactar dependiendo del tipo de cliente
        '------------------------------------------------------------------------
        If sInput.Contains("%CONTACTAR%") = True Then
            Select Case TemplateID
                Case 5
                    sTmp = "o SU ASESORA ASIGNADA"
                    If sTipoClienteCodigo = "C" AndAlso Not String.IsNullOrEmpty(sAsesorNombre) Then
                        sTmp = " o " & sAsesorNombre
                    End If
                Case Else
                    sTmp = "el departamento de servicio al cliente"
                    If sTipoClienteCodigo = "C" AndAlso Not String.IsNullOrEmpty(sAsesorNombre) Then
                        sTmp = sAsesorNombre & " o " & sTmp
                    End If
            End Select
            FormatTemplateText(sInput, "CONTACTAR", sTmp)
        End If

        ' Formatear texto actualizar las direcciones de Paquetes y PO Box
        FormatTemplateTextClienteDirecciones(sInput, bIsHtml, sNumeroEPS)

    End Sub

    ''' <summary>
    ''' Formatear texto y actualizar direcciones de Paquetes y POBox de la notificiacion
    ''' </summary>
    ''' <param name="sInput">Template Text</param>
    ''' <param name="bIsHtml">Texto en formato de HTML</param>
    ''' <param name="NumeroEPS">Numero de EPS</param>
    ''' <remarks></remarks>
    Private Sub FormatTemplateTextClienteDirecciones(ByRef sInput As String, ByVal bIsHtml As Boolean, ByVal NumeroEPS As String)
        Dim sFind As String = String.Empty
        Dim sTmp As String = String.Empty

        Try
            sFind = "DIRRECION_EPS_POBOX"
            If sInput.Contains(String.Format("<!--${0}-->", sFind)) = True Then
                sInput = sInput.Replace(String.Format("<!--${0}-->", sFind), String.Format("%{0}%", sFind))
            End If
            If sInput.Contains(String.Format("%{0}%", sFind)) = True Then
                sTmp = GetPackagesAddressAndPOBox(NumeroEPS, "C")
                If bIsHtml = True Then
                    sTmp = sTmp.Replace(Chr(10), "<br />")
                End If
                sInput = sInput.Replace(String.Format("%{0}%", sFind), sTmp)
            End If

            sFind = "DIRRECION_EPS_PAQUETES"
            If sInput.Contains(String.Format("<!--${0}-->", sFind)) = True Then
                sInput = sInput.Replace(String.Format("<!--${0}-->", sFind), String.Format("%{0}%", sFind))
            End If
            If sInput.Contains(String.Format("%{0}%", sFind)) = True Then
                sTmp = GetPackagesAddressAndPOBox(NumeroEPS, "P")
                If bIsHtml = True Then
                    sTmp = sTmp.Replace(Chr(10), "<br />")
                End If
                sInput = sInput.Replace(String.Format("%{0}%", sFind), sTmp)
            End If
        Catch ex As Exception
            ' Do nothing
        End Try
    End Sub

    ''' <summary>
    ''' Formatear y reemplazar texto
    ''' </summary>
    ''' <param name="sInput">Template Text</param>
    ''' <param name="sFind">Find Text</param>
    ''' <param name="sReplace">Replace Text</param>
    ''' <remarks></remarks>
    Private Sub FormatTemplateText(ByRef sInput As String, ByVal sFind As String, ByVal sReplace As Object)
        Try
            If sInput.Contains(String.Format("%{0}%", sFind)) = True Then
                sInput = sInput.Replace(String.Format("%{0}%", sFind), CType(sReplace, String))
            ElseIf sInput.Contains(String.Format("<!--${0}-->", sFind)) = True Then
                sInput = sInput.Replace(String.Format("<!--${0}-->", sFind), CType(sReplace, String))
            End If
        Catch ex As Exception
            ' Do nothing
        End Try
    End Sub

#End Region

#Region "Impuestos Aduanales"

    ''' <summary>
    ''' Procesar Impuestos Aduanales
    ''' </summary>
    ''' <param name="NumeroEPS"></param>
    ''' <param name="CodigoBarra"></param>
    ''' <param name="EmailBody"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ProcessImpuestosAduanales(ByVal NumeroEPS As String, ByVal CodigoBarra As String, ByRef EmailBody As String) As Boolean

        Dim bReturn As Boolean = False

        Dim sDeclaracion As String = Nothing
        Dim nGravamen As String = Nothing
        Dim nSelectivo As String = Nothing
        Dim nItbis As String = Nothing
        Dim nArt52 As String = Nothing
        Dim nTotalRD As String = Nothing
        Dim nServicioAduanero As String = Nothing
        Dim nValorDeclarado As String = Nothing
        Dim nValorAsignado As String = Nothing
        Dim sNota As String = Nothing
        Dim sPinPago As String = Nothing
        Dim sNota2 As String = Nothing
        Dim sNota3 As String = Nothing
        Dim sNota4 As String = Nothing
        Dim sNota5 As String = Nothing
        Dim sTieneDocumento As String = Nothing

        Try
            ' Create a new DataSet Object to fill with Data
            Dim ds As New DataSet
            ds = GetImpuestosAduanalesDataSet(NumeroEPS, CodigoBarra)

            Dim dr As DataRow
            For Each dr In ds.Tables(0).Rows

                ' Buscar los datos de los impuestos aduanales
                sDeclaracion = db.ewToStringUpper(dr("TPL_DECLARACION"))

                nGravamen = db.ewToDecimal(dr("TPL_GRAVAMEN")).ToString
                nSelectivo = db.ewToDecimal(dr("TPL_SELECTIVO")).ToString
                nItbis = db.ewToDecimal(dr("TPL_ITBIS")).ToString
                nArt52 = db.ewToDecimal(dr("TPL_ART52")).ToString
                nTotalRD = db.ewToDecimal(dr("TPL_TOTAL_RD")).ToString
                nServicioAduanero = db.ewToDecimal(dr("TPL_SERVICIO_ADUANERO")).ToString
                nValorDeclarado = db.ewToDecimal(dr("TPL_VALOR_DECLARADO")).ToString
                nValorAsignado = db.ewToDecimal(dr("TPL_VALOR_ASIGNADO")).ToString

                sPinPago = db.ewToStringUpper(dr("TPL_PINPAGO"))
                sTieneDocumento = db.ewToStringUpper(dr("TPL_TIENEDOCUMENTO"))

                sNota = db.ewToStringUpper(dr("TPL_NOTA"))
                sNota2 = db.ewToStringUpper(dr("TPL_NOTA2"))
                sNota3 = db.ewToStringUpper(dr("TPL_NOTA3"))
                sNota4 = db.ewToStringUpper(dr("TPL_NOTA4"))
                sNota5 = db.ewToStringUpper(dr("TPL_NOTA5"))

                ' Formatear numeros
                '------------------------------------------------------------------------
                If Not String.IsNullOrEmpty(nGravamen) And Utilities.DataFormat.ewCheckDecimal(nGravamen) = True Then
                    nGravamen = Utilities.DataFormat.ewNumberFormat(nGravamen, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(nSelectivo) And Utilities.DataFormat.ewCheckDecimal(nSelectivo) = True Then
                    nSelectivo = Utilities.DataFormat.ewNumberFormat(nSelectivo, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(nItbis) And Utilities.DataFormat.ewCheckDecimal(nItbis) = True Then
                    nItbis = Utilities.DataFormat.ewNumberFormat(nItbis, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(nArt52) And Utilities.DataFormat.ewCheckDecimal(nArt52) = True Then
                    nArt52 = Utilities.DataFormat.ewNumberFormat(nArt52, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(nTotalRD) And Utilities.DataFormat.ewCheckDecimal(nTotalRD) = True Then
                    nTotalRD = Utilities.DataFormat.ewNumberFormat(nTotalRD, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(nServicioAduanero) And Utilities.DataFormat.ewCheckDecimal(nServicioAduanero) = True Then
                    nServicioAduanero = Utilities.DataFormat.ewNumberFormat(nServicioAduanero, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(nValorDeclarado) And Utilities.DataFormat.ewCheckDecimal(nValorDeclarado) = True Then
                    nValorDeclarado = Utilities.DataFormat.ewNumberFormat(nValorDeclarado, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(nValorAsignado) And Utilities.DataFormat.ewCheckDecimal(nValorAsignado) = True Then
                    nValorAsignado = Utilities.DataFormat.ewNumberFormat(nValorAsignado, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(sNota2) And Utilities.DataFormat.ewCheckDecimal(sNota2) = True Then
                    sNota2 = Utilities.DataFormat.ewNumberFormat(sNota2, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(sNota3) And Utilities.DataFormat.ewCheckDecimal(sNota3) = True Then
                    sNota3 = Utilities.DataFormat.ewNumberFormat(sNota3, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(sNota4) And Utilities.DataFormat.ewCheckDecimal(sNota4) = True Then
                    sNota4 = Utilities.DataFormat.ewNumberFormat(sNota4, 2, 1, 0, 1)
                End If
                If Not String.IsNullOrEmpty(sNota5) And Utilities.DataFormat.ewCheckDecimal(sNota5) = True Then
                    sNota5 = Utilities.DataFormat.ewNumberFormat(sNota5, 2, 1, 0, 1)
                End If

                ' Crear el cuerpo del mensaje a enviar
                '------------------------------------------------------------------------
                If EmailBody.Contains("%DECLARACION%") = True Then
                    EmailBody = EmailBody.Replace("%DECLARACION%", sDeclaracion)
                End If
                If EmailBody.Contains("%GRAVAMEN%") = True Then
                    EmailBody = EmailBody.Replace("%GRAVAMEN%", nGravamen)
                End If
                If EmailBody.Contains("%SELECTIVO%") = True Then
                    EmailBody = EmailBody.Replace("%SELECTIVO%", nSelectivo)
                End If
                If EmailBody.Contains("%ITBIS%") = True Then
                    EmailBody = EmailBody.Replace("%ITBIS%", nItbis)
                End If
                If EmailBody.Contains("%ART52%") = True Then
                    EmailBody = EmailBody.Replace("%ART52%", nArt52)
                End If
                If EmailBody.Contains("%TOTAL_RD%") = True Then
                    EmailBody = EmailBody.Replace("%TOTAL_RD%", nTotalRD)
                End If
                If EmailBody.Contains("%SERVICIO_ADUANERO%") = True Then
                    EmailBody = EmailBody.Replace("%SERVICIO_ADUANERO%", nServicioAduanero)
                End If
                If EmailBody.Contains("%VALOR_DECLARADO%") = True Then
                    EmailBody = EmailBody.Replace("%VALOR_DECLARADO%", nValorDeclarado)
                End If
                If EmailBody.Contains("%VALOR_ASIGNADO%") = True Then
                    EmailBody = EmailBody.Replace("%VALOR_ASIGNADO%", nValorAsignado)
                End If
                If EmailBody.Contains("%PIN_PAGO%") = True Then
                    EmailBody = EmailBody.Replace("%PIN_PAGO%", "(" & sPinPago & ")")
                End If
                If EmailBody.Contains("%TIENE_DOCUMENTO%") = True Then
                    If sTieneDocumento = "S" Then
                        EmailBody = EmailBody.Replace("%TIENE_DOCUMENTO%", "%URL_DOCUMENTO%")
                    End If
                End If

                If EmailBody.Contains("%NOTA%") = True Then
                    If Not String.IsNullOrEmpty(sNota) Then
                        EmailBody = EmailBody.Replace("%NOTA%", "<p>NOTA (3):<br>" & sNota & "<p>")
                    Else
                        EmailBody = EmailBody.Replace("%NOTA%", String.Empty)
                    End If
                End If
                If EmailBody.Contains("%NOTA2%") = True Then
                    If Not String.IsNullOrEmpty(sNota2) Then
                        'If sNota2.StartsWith(",") = False Then
                        '    sNota2 = ", " & sNota2
                        'End If
                        EmailBody = EmailBody.Replace("%NOTA2%", sNota2)
                    Else
                        EmailBody = EmailBody.Replace("%NOTA2%", String.Empty)
                    End If
                End If
                If EmailBody.Contains("%NOTA3%") = True Then
                    If Not String.IsNullOrEmpty(sNota3) Then
                        EmailBody = EmailBody.Replace("%NOTA3%", sNota3)
                    Else
                        EmailBody = EmailBody.Replace("%NOTA3%", String.Empty)
                    End If
                End If
                If EmailBody.Contains("%NOTA4%") = True Then
                    If Not String.IsNullOrEmpty(sNota4) Then
                        EmailBody = EmailBody.Replace("%NOTA4%", sNota4)
                    Else
                        EmailBody = EmailBody.Replace("%NOTA4%", String.Empty)
                    End If
                End If
                If EmailBody.Contains("%NOTA5%") = True Then
                    If Not String.IsNullOrEmpty(sNota5) Then
                        EmailBody = EmailBody.Replace("%NOTA5%", sNota5)
                    Else
                        EmailBody = EmailBody.Replace("%NOTA5%", String.Empty)
                    End If
                End If

                bReturn = True
            Next

        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ProcessImpuestosAduanales(): " & sEx, 2)
        End Try

        Return bReturn

    End Function

    ''' <summary>
    ''' Traer Impuestos Aduanales DataSet
    ''' </summary>
    ''' <param name="NumeroEPS"></param>
    ''' <param name="CodigoBarra"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetImpuestosAduanalesDataSet(ByVal NumeroEPS As String, ByVal CodigoBarra As String) As DataSet

        Dim sSql As String = String.Format("EXEC [dbo].[proc_EPSWEBMAIL_IMPUESTOS_ADUANALESLoadByPrimaryKey] @CTE_NUMERO_EPS = '{0}', @BLT_CODIGO_BARRA = '{1}'", _
                                           NumeroEPS, CodigoBarra)
        Try
            Return db.ewGetDataSet(sSql)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "GetImpuestosAduanalesDataSet(): " & sEx, 2)
            Return Nothing
        End Try

    End Function

#End Region

#Region "Sistema de Cola de Llamadas"

    ''' <summary>
    ''' Asignacion de Mensajes 
    ''' </summary>
    ''' <param name="Dias"></param>
    ''' <remarks></remarks>
    Private Sub AsignacionDeMensajes(ByVal Dias As Integer)

        Dim oConn As New SqlConnection(db.GetConnectionString)
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "[dbo].[ASIGNACION_DE_MENSAJES]"

            ' @DIAS int
            db.AddParameter(oCmd, "@DIAS", db.DataType.eInteger, db.Direction.eInput, Dias)
        End With

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "AsignacionDeMensajes(): " & sEx, 2)
        Finally
            oConn.Close()
        End Try

    End Sub

#End Region

    ''' <summary>
    ''' Buscar Direcciones de Correspondencia y Paquetes 
    ''' </summary>
    ''' <param name="NumeroEPS">Numero de EPS</param>
    ''' <param name="TipoDireccion">Tipo de Direccion: C=Correo; P=Paquetes;</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPackagesAddressAndPOBox(ByVal NumeroEPS As String, ByVal TipoDireccion As String) As String

        Dim sSql As String = String.Empty
        sSql = String.Format("SELECT [dbo].[f_Direccion_POBOX] ('{0}', '{1}')", _
                             NumeroEPS, TipoDireccion)

        Dim oConn As New SqlConnection(db.GetConnectionString())
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.Text
            .CommandText = sSql
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

    ''' <summary>
    ''' Buscar la Conexión a la Base de Datos para Importar
    ''' </summary>
    ''' <returns>Database connection string</returns>
    ''' <remarks></remarks>
    Private Function GetImportConnectionString() As String
        Return db.GetConnectionString("ImportConnectionString")
    End Function

#Region "Send Email"

    ''' <summary>
    ''' Enviar Correo
    ''' </summary>
    ''' <param name="TemplateID">Template ID</param>
    ''' <param name="Format">Email Format</param>
    ''' <param name="EmailTo">Email Address To</param>
    ''' <param name="EmailFrom">Email Address From</param>
    ''' <param name="FromName">From Name</param>
    ''' <param name="EmailReplyTo">Email Address Reply To</param>
    ''' <param name="EmailSubject">Email Subject</param>
    ''' <param name="EmailBody">Email Body</param>
    ''' <param name="SendEmail">Send Email?</param>
    ''' <param name="dr">Table Data Row</param>
    ''' <remarks></remarks>
    Private Sub SendEmail(ByVal TemplateID As Integer, _
                          ByVal Format As String, _
                          ByVal EmailTo As String, _
                          ByVal EmailFrom As String, _
                          ByVal FromName As String, _
                          ByVal EmailReplyTo As String, _
                          ByVal EmailSubject As String, _
                          ByVal EmailBody As String, _
                          ByVal SendEmail As Boolean, _
                          ByVal dr As DataRow)

        Dim nMensajeID As String = Nothing
        Dim sNumeroEPS As String = Nothing
        Dim sNombreCompleto As String = Nothing
        Dim sClienteEmail As String = Nothing

        nMensajeID = db.ewToString(dr("WMM_MENSAJE_ID"))
        sNumeroEPS = db.ewToStringUpper(dr("CTE_NUMERO_EPS"))
        sNombreCompleto = db.ewToStringUpper(dr("NOMBRE_COMPLETO"))

        ' Desplegar mensaje
        PrintLine(String.Format("Cliente: {0} / {1}", sNumeroEPS, sNombreCompleto))
        PrintLine(String.Format("Mensaje: {0}", EmailSubject))

        ' Enviar el correo electrónico
        If SendEmail = True Then

            ' Validar si el mensaje es HTML
            Dim bIsHtml As Boolean = IIf(Format.ToUpper.Trim = "HTML", True, False)

            ' Shrink HTML
            If bIsHtml = True Then
                If Not String.IsNullOrEmpty(sNombreCompleto) Then sNombreCompleto = ew_EncodeText(sNombreCompleto)
                If Not String.IsNullOrEmpty(EmailBody) Then EmailBody = ew_ShrinkHtml(EmailBody)
            End If

            ' Enviar notificacion a EmailQueue
            If SendToEmailQueue(Format, EmailTo, EmailFrom, FromName, EmailReplyTo, EmailSubject, EmailBody, sNumeroEPS) = True Then
                ' Actualizar notificaciones como enviada
                UpdateWebMailMensajesEnviado(nMensajeID, "S")
            Else
                ' Actualizar notificaciones como no enviada
                UpdateWebMailMensajesEnviado(nMensajeID, "N")
            End If

        Else

            ' Actualizar notificaciones como no enviada
            UpdateWebMailMensajesEnviado(nMensajeID, "N")

            PrintDobleLine("-> ERROR: Mensajes de email sin enviar")
        End If

    End Sub

    ''' <summary>
    ''' Enviar Correo a EmailQueue
    ''' </summary>
    ''' <param name="Format">Email Format</param>
    ''' <param name="EmailTo">Email Address To</param>
    ''' <param name="EmailFrom">Email Address From</param>
    ''' <param name="FromName">From Name</param>
    ''' <param name="EmailReplyTo">Email Address Reply To</param>
    ''' <param name="EmailSubject">Email Subject</param>
    ''' <param name="EmailBody">Email Body</param>
    ''' <param name="NumeroEPS">Numero de EPS</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SendToEmailQueue(ByVal Format As String, _
                                      ByVal EmailTo As String, _
                                      ByVal EmailFrom As String, _
                                      ByVal FromName As String, _
                                      ByVal EmailReplyTo As String, _
                                      ByVal EmailSubject As String, _
                                      ByVal EmailBody As String, _
                                      ByVal NumeroEPS As String) As Boolean

        Dim sMessage As String = String.Empty
        Dim bReturn As Boolean = False
        Dim nReturn As Integer = 0
        Dim nTotalEmailSent As Integer = 0

        ' Validar dirección de correo (De)
        If String.IsNullOrEmpty(EmailFrom) Then
            PrintLine("==> Dirección Email (De) en blanco")
            Return bReturn
            Exit Function
        ElseIf Not Validators.IsEmail(EmailFrom) Then
            PrintLine("==> Dirección Email (De) invalida")
            Return bReturn
            Exit Function
        End If

        ' Procesar direcciones de E-mail
        If String.IsNullOrEmpty(EmailTo) Then
            PrintLine("==> Dirección Email (Para) en blanco")
            Return bReturn
            Exit Function
        Else
            EmailTo = ProcessEmailAddresses(EmailTo)

            ' Validar que por lo menos exista una dirección de correo (Para)
            If String.IsNullOrEmpty(EmailTo) Then
                PrintLine("==> Dirección Email (Para) invalida o no existen Email validos")
                Return bReturn
                Exit Function
            End If
        End If

        '// For Testing
        '// EmailTo = "efernandez@eps-int.com; edwinet@gmail.com; edwinet@yahoo.com; edwin fernandez"
        '// EmailTo = ProcessEmailAddresses(EmailTo)

        ' Separar los correos de los clientes para enviar uno por cada correo
        Dim sArrValidEmail As String = String.Empty
        Dim arrEmail() As String = EmailTo.Split(","c)
        For Each aEmail As String In arrEmail

            ' Desplegar mensaje
            PrintLine(String.Format("Email: {0}", aEmail))

            ' Agregar al EmailQueue
            nReturn = ewEmailQueue.AddToQueue(Format, EmailFrom, aEmail, EmailSubject, EmailBody, NumeroEPS, nApplicationID)

            ' Incrementar counter de Email enviados
            If nReturn > 0 Then
                nTotalEmailSent += 1
            End If

        Next

        ' Desplegar mensaje
        PrintLine(spacer)

        ' Si se envio más de un correo entonces exitoso
        If nTotalEmailSent > 0 Then
            bReturn = True
        End If

        Return bReturn

    End Function

    ''' <summary>
    ''' Procesar Dirección de E-mail
    ''' </summary>
    ''' <param name="Email"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ProcessEmailAddresses(ByVal Email As String) As String

        If String.IsNullOrEmpty(Email) Then
            Return String.Empty
            Exit Function
        End If

        Dim sArrEmail As String = Utilities.Email.EmailParser(Email)

        If (String.IsNullOrEmpty(sArrEmail)) Or (sArrEmail = "n/a") Then
            Return String.Empty
            Exit Function
        End If

        sArrEmail = sArrEmail.Replace("&", ",").Replace("\", ",").Replace(" ", ",")
        sArrEmail = sArrEmail.Replace("<", ",").Replace(">", ",")

        While sArrEmail.Contains(",,")
            sArrEmail = sArrEmail.Replace(",,", ",")
        End While

        Dim sArrValidEmail As String = String.Empty
        Dim arrEmail() As String = sArrEmail.Split(","c)
        For Each aEmail As String In arrEmail
            If (Not String.IsNullOrEmpty(aEmail.Trim)) Then
                If Validators.IsEmail(aEmail.Trim) Then
                    sArrValidEmail += aEmail.Trim & ","
                End If
            End If
        Next

        ' Remover el ultimo (,)
        If Not String.IsNullOrEmpty(sArrValidEmail) Then
            sArrValidEmail = Utilities.Email.ParseLastSeparator(sArrValidEmail)
        End If

        Return sArrValidEmail

    End Function

    ''' <summary>
    ''' Procesar Dirección de Envio
    ''' </summary>
    ''' <param name="EmailTo"></param>
    ''' <param name="EmailCcList"></param>
    ''' <param name="EmailBccList"></param>
    ''' <param name="EmailCliente"></param>
    ''' <param name="EmailAsesor"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ProcessSendToEmailAddressList(ByVal EmailTo As String, ByVal EmailCcList As String, ByVal EmailBccList As String, ByVal EmailCliente As String, ByVal EmailAsesor As String)

        Dim sEmailReturn As String = String.Empty

        ' Buscar dirección de envio del mensaje
        If Not String.IsNullOrEmpty(EmailTo) Then
            sEmailReturn = EmailTo ' Asignar la(s) dirección(es) de email del template del mensaje
        Else
            sEmailReturn = EmailCliente ' Asignar la(s) dirección(es) de email del cliente
        End If

        ' Notificar a la lista de direcciones de email de Cc
        sEmailReturn = AddEmailAddressToList(sEmailReturn, EmailCcList)

        ' Notificar a la lista de direcciones de email de Bcc
        sEmailReturn = AddEmailAddressToList(sEmailReturn, EmailBccList)

        ' Notificar Asesor
        If Not String.IsNullOrEmpty(EmailAsesor) AndAlso EmailAsesor <> "servicios@eps-int.com" Then
            sEmailReturn = AddEmailAddressToList(sEmailReturn, EmailAsesor)
        End If

        Return sEmailReturn

    End Function

    ''' <summary>
    ''' Agregar Dirección de E-mail a la Lista
    ''' </summary>
    ''' <param name="EmailList"></param>
    ''' <param name="AddEmailToList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddEmailAddressToList(ByVal EmailList As String, ByVal AddEmailToList As String) As String
        If String.IsNullOrEmpty(EmailList) AndAlso String.IsNullOrEmpty(AddEmailToList) Then
            Return String.Empty
        ElseIf Not String.IsNullOrEmpty(EmailList) AndAlso String.IsNullOrEmpty(AddEmailToList) Then
            Return EmailList
        ElseIf Not String.IsNullOrEmpty(EmailList) Then
            If EmailList.Contains(AddEmailToList) = False Then
                Return EmailList & ", " & AddEmailToList
            Else
                Return EmailList
            End If
        Else
            Return AddEmailToList
        End If
    End Function

#End Region

#Region "Create / Read files"

    ''' <summary>
    ''' Create File
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <remarks></remarks>
    Sub CreateFile(ByVal FileName As String)

        Try
            LogFile = New StreamWriter(FileName)
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "CreateFile(): " & sEx, 1)
        End Try

    End Sub

    ''' <summary>
    ''' Read File
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ReadFile(ByVal FileName As String) As String

        Dim sReturn As String = String.Empty
        Dim sFilePath As String = String.Empty

        If String.IsNullOrEmpty(FileName) Then
            Return sReturn
        End If

        If File.Exists(FileName) = True Then
            sFilePath = FileName
        Else
            sFilePath = GetFileFullPath(FileName)
        End If

        Try
            Dim file As New StreamReader(sFilePath)
            sReturn = file.ReadToEnd()
            file.Close()
        Catch ex As Exception
            Dim sEx As String = ex.Message.ToString
            PrintDobleLine("ERROR: " & sEx)
            ewErrorHandler.NotificaError("Se ha producido un error", "ReadFile(): " & sEx, 1)
        End Try

        Return sReturn

    End Function

    ''' <summary>
    ''' Get File Full Path
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetFileFullPath(ByVal FileName As String) As String

        Dim sReturn As String = String.Empty

        If String.IsNullOrEmpty(FileName) Then
            Return sReturn
        End If

        ' Get file path
        If Not FileName.Contains(":\") Then
            sReturn += Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location())

            If Not FileName.StartsWith("\") Then
                sReturn += "\"
            End If
        End If

        sReturn += FileName

        Return sReturn

    End Function

#End Region

#Region "HTML Functions"

    ''' <summary>
    ''' Formats a text for HTML viewing
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ew_FormatTextToHtml(ByVal str As String) As String
        If Not String.IsNullOrEmpty(str) Then
            Dim sTmp As String = str
            ' Change carriage returns and line feeds to HTML
            sTmp = sTmp.Replace(vbCrLf, "<br />")
            sTmp = sTmp.Replace(vbCr, "<br />")
            sTmp = sTmp.Replace(vbLf, "<br />")
            ' Change tabs to HTML
            sTmp = sTmp.Replace(vbTab, "&nbsp;&nbsp;&nbsp;")
            ' Change double-space to HTML
            While Not InStr(1, sTmp, "  ") = 0
                sTmp = sTmp.Replace("  ", "&nbsp; ")
            End While
            Return sTmp
        Else
            Return String.Empty
        End If
    End Function

    ''' <summary>
    ''' Remove HTML Tags
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ew_RemoveHtmlTags(ByVal str As String) As String
        If Not String.IsNullOrEmpty(str) Then
            Dim sTmp As String = str
            sTmp = sTmp.Replace("<html>" & vbCrLf, String.Empty)
            sTmp = sTmp.Replace("<body>" & vbCrLf, String.Empty)
            sTmp = sTmp.Replace("</body>" & vbCrLf, String.Empty)
            sTmp = sTmp.Replace("</html>", String.Empty)
            Return sTmp
        Else
            Return String.Empty
        End If
    End Function

    ''' <summary>
    ''' Shrink HTML
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ew_ShrinkHtml(ByVal str As String) As String
        If Not String.IsNullOrEmpty(str) Then
            Dim sTmp As String = str
            ' Change tabs to HTML
            sTmp = sTmp.Replace(vbTab, String.Empty)
            ' Change carriage returns and line feeds to HTML
            sTmp = sTmp.Replace(vbCrLf, String.Empty)
            sTmp = sTmp.Replace(vbCr, String.Empty)
            sTmp = sTmp.Replace(vbLf, String.Empty)
            Return sTmp
        Else
            Return String.Empty
        End If
    End Function

    ''' <summary>
    ''' Encode non-US-ASCII characters to name entities
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ew_EncodeText(ByVal str As String) As String
        Dim sTmp As String = String.Empty
        sTmp = str
        sTmp = sTmp.Replace("¡", "&iexcl;")
        sTmp = sTmp.Replace("¢", "&cent;")
        sTmp = sTmp.Replace("£", "&pound;")
        sTmp = sTmp.Replace("¤", "&curren;")
        sTmp = sTmp.Replace("¥", "&Ntilde;")
        sTmp = sTmp.Replace("¦", "&brvbar;")
        sTmp = sTmp.Replace("§", "&sect;")
        sTmp = sTmp.Replace("¨", "&uml;")
        sTmp = sTmp.Replace("©", "&copy;")
        sTmp = sTmp.Replace("ª", "&ordf;")
        sTmp = sTmp.Replace("«", "&laquo;")
        sTmp = sTmp.Replace("¬", "&not;")
        sTmp = sTmp.Replace("­", "&shy;")
        sTmp = sTmp.Replace("®", "&reg;")
        sTmp = sTmp.Replace("¯", "&macr;")
        sTmp = sTmp.Replace("°", "&deg;")
        sTmp = sTmp.Replace("±", "&plusmn;")
        sTmp = sTmp.Replace("²", "&sup2;")
        sTmp = sTmp.Replace("³", "&sup3;")
        sTmp = sTmp.Replace("´", "&acute;")
        sTmp = sTmp.Replace("µ", "&micro;")
        sTmp = sTmp.Replace("¶", "&para;")
        sTmp = sTmp.Replace("·", "&middot;")
        sTmp = sTmp.Replace("¸", "&cedil;")
        sTmp = sTmp.Replace("¹", "&sup1;")
        sTmp = sTmp.Replace("º", "&ordm;")
        sTmp = sTmp.Replace("»", "&raquo;")
        sTmp = sTmp.Replace("¼", "&frac14;")
        sTmp = sTmp.Replace("½", "&frac12;")
        sTmp = sTmp.Replace("¾", "&frac34;")
        sTmp = sTmp.Replace("¿", "&iquest;")
        sTmp = sTmp.Replace("À", "&Agrave;")
        sTmp = sTmp.Replace("Á", "&Aacute;")
        sTmp = sTmp.Replace("Â", "&Acirc;")
        sTmp = sTmp.Replace("Ã", "&Atilde;")
        sTmp = sTmp.Replace("Ä", "&Auml;")
        sTmp = sTmp.Replace("Å", "&Aring;")
        sTmp = sTmp.Replace("Æ", "&AElig;")
        sTmp = sTmp.Replace("Ç", "&Ccedil;")
        sTmp = sTmp.Replace("È", "&Egrave;")
        sTmp = sTmp.Replace("É", "&Eacute;")
        sTmp = sTmp.Replace("Ê", "&Ecirc;")
        sTmp = sTmp.Replace("Ë", "&Euml;")
        sTmp = sTmp.Replace("Ì", "&Igrave;")
        sTmp = sTmp.Replace("Í", "&Iacute;")
        sTmp = sTmp.Replace("Î", "&Icirc;")
        sTmp = sTmp.Replace("Ï", "&Iuml;")
        sTmp = sTmp.Replace("Ð", "&ETH;")
        sTmp = sTmp.Replace("Ñ", "&Ntilde;")
        sTmp = sTmp.Replace("Ò", "&Ograve;")
        sTmp = sTmp.Replace("Ó", "&Oacute;")
        sTmp = sTmp.Replace("Ô", "&Ocirc;")
        sTmp = sTmp.Replace("Õ", "&Otilde;")
        sTmp = sTmp.Replace("Ö", "&Ouml;")
        sTmp = sTmp.Replace("×", "&times;")
        sTmp = sTmp.Replace("Ø", "&Oslash;")
        sTmp = sTmp.Replace("Ù", "&Ugrave;")
        sTmp = sTmp.Replace("Ú", "&Uacute;")
        sTmp = sTmp.Replace("Û", "&Ucirc;")
        sTmp = sTmp.Replace("Ü", "&Uuml;")
        sTmp = sTmp.Replace("Ý", "&yacute;")
        sTmp = sTmp.Replace("Þ", "&THORN;")
        sTmp = sTmp.Replace("ß", "&szlig;")
        sTmp = sTmp.Replace("à", "&agrave;")
        sTmp = sTmp.Replace("á", "&aacute;")
        sTmp = sTmp.Replace("â", "&acirc;")
        sTmp = sTmp.Replace("ã", "&atilde;")
        sTmp = sTmp.Replace("ä", "&auml;")
        sTmp = sTmp.Replace("å", "&aring;")
        sTmp = sTmp.Replace("æ", "&aelig;")
        sTmp = sTmp.Replace("ç", "&ccedil;")
        sTmp = sTmp.Replace("è", "&egrave;")
        sTmp = sTmp.Replace("é", "&eacute;")
        sTmp = sTmp.Replace("ê", "&ecirc;")
        sTmp = sTmp.Replace("ë", "&euml;")
        sTmp = sTmp.Replace("ì", "&igrave;")
        sTmp = sTmp.Replace("í", "&iacute;")
        sTmp = sTmp.Replace("î", "&icirc;")
        sTmp = sTmp.Replace("ï", "&iuml;")
        sTmp = sTmp.Replace("ð", "&eth;")
        sTmp = sTmp.Replace("ñ", "&ntilde;")
        sTmp = sTmp.Replace("ò", "&ograve;")
        sTmp = sTmp.Replace("ó", "&oacute;")
        sTmp = sTmp.Replace("ô", "&ocirc;")
        sTmp = sTmp.Replace("õ", "&otilde;")
        sTmp = sTmp.Replace("ö", "&ouml;")
        sTmp = sTmp.Replace("÷", "&divide;")
        sTmp = sTmp.Replace("ø", "&oslash;")
        sTmp = sTmp.Replace("ù", "&ugrave;")
        sTmp = sTmp.Replace("ú", "&uacute;")
        sTmp = sTmp.Replace("û", "&ucirc;")
        sTmp = sTmp.Replace("ü", "&uuml;")
        sTmp = sTmp.Replace("ý", "&yacute;")
        sTmp = sTmp.Replace("þ", "&thorn;")
        sTmp = sTmp.Replace("ÿ", "&yuml;")
        sTmp = sTmp.Replace("ˆ", "&circ;")
        sTmp = sTmp.Replace("˜", "&tilde;")
        sTmp = sTmp.Replace("–", "&ndash;")
        sTmp = sTmp.Replace("—", "&mdash;")
        sTmp = sTmp.Replace("‘", "&lsquo;")
        sTmp = sTmp.Replace("’", "&rsquo;")
        sTmp = sTmp.Replace("†", "&dagger;")
        sTmp = sTmp.Replace("‡", "&Dagger;")
        sTmp = sTmp.Replace("•", "&bull;")
        sTmp = sTmp.Replace("‹", "&lsaquo;")
        sTmp = sTmp.Replace("›", "&rsaquo;")
        sTmp = sTmp.Replace("€", "&euro;")
        sTmp = sTmp.Replace("™", "&trade;")

        Return sTmp
    End Function

#End Region

#Region "Print Lines"

    ''' <summary>
    ''' Print a single line message
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <remarks></remarks>
    Private Sub PrintLine(ByVal Message As String)
        ' Write message to Log File
        LogFile.WriteLine(Message)

        ' Write a message to the Immediate Window and to the Console
        ewPrint.WriteLine(Message)
    End Sub

    ''' <summary>
    ''' Print a doble line message
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <remarks></remarks>
    Private Sub PrintDobleLine(ByVal Message As String)
        ' Write message to Log File
        LogFile.WriteLine(Message)
        LogFile.WriteLine("")

        ' Write a message to the Immediate Window and to the Console
        ewPrint.WriteDobleLine(Message)
    End Sub

#End Region

End Module
