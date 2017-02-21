Imports System
Imports System.Configuration
Imports System.Net.Mail
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic

Module Utilities

    Public Class DataFormat
        ''' <summary>
        ''' Get tri state
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function GetTriState(ByVal value As Integer) As TriState
            Select Case value
                Case 0
                    Return TriState.False
                Case -1
                    Return TriState.True
                Case -2
                    Return TriState.UseDefault
                Case Else
                    Return TriState.UseDefault
            End Select
        End Function

        ''' <summary>
        ''' Format Currency
        ''' </summary>
        ''' <param name="exp"></param>
        ''' <param name="NumDigitsAfterDecimal"></param>
        ''' <param name="IncludeLeadingDigit"></param>
        ''' <param name="UseParensForNegativeNumbers"></param>
        ''' <param name="GroupDigits"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCurrencyFormat(ByVal exp As Object, ByVal NumDigitsAfterDecimal As Integer, ByVal IncludeLeadingDigit As Integer, ByVal UseParensForNegativeNumbers As Integer, ByVal GroupDigits As Integer) As String
            Return Strings.FormatCurrency(exp, NumDigitsAfterDecimal, GetTriState(IncludeLeadingDigit), GetTriState(UseParensForNegativeNumbers), GetTriState(GroupDigits))
        End Function

        ''' <summary>
        ''' Format Number
        ''' </summary>
        ''' <param name="exp"></param>
        ''' <param name="NumDigitsAfterDecimal"></param>
        ''' <param name="IncludeLeadingDigit"></param>
        ''' <param name="UseParensForNegativeNumbers"></param>
        ''' <param name="GroupDigits"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewNumberFormat(ByVal exp As Object, ByVal NumDigitsAfterDecimal As Integer, ByVal IncludeLeadingDigit As Integer, ByVal UseParensForNegativeNumbers As Integer, ByVal GroupDigits As Integer) As String
            Return Strings.FormatNumber(exp, NumDigitsAfterDecimal, GetTriState(IncludeLeadingDigit), GetTriState(UseParensForNegativeNumbers), GetTriState(GroupDigits))
        End Function

        ''' <summary>
        ''' Format Percent
        ''' </summary>
        ''' <param name="exp"></param>
        ''' <param name="NumDigitsAfterDecimal"></param>
        ''' <param name="IncludeLeadingDigit"></param>
        ''' <param name="UseParensForNegativeNumbers"></param>
        ''' <param name="GroupDigits"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewPercentFormat(ByVal exp As Object, ByVal NumDigitsAfterDecimal As Integer, ByVal IncludeLeadingDigit As Integer, ByVal UseParensForNegativeNumbers As Integer, ByVal GroupDigits As Integer) As String
            Return Strings.FormatPercent(exp, NumDigitsAfterDecimal, GetTriState(IncludeLeadingDigit), GetTriState(UseParensForNegativeNumbers), GetTriState(GroupDigits))
        End Function

        ''' <summary>
        ''' Format DateTime
        ''' </summary>
        ''' <param name="NamedFormat">
        ''' 0 = 1/6/2008 3:30:15 PM  *** 
        ''' 1 = Sunday, January 06, 2008  *** 
        ''' 2 = 1/6/2008  *** 
        ''' 3 = 15:30:15  *** 
        ''' 4 = 15:30  *** 
        ''' 5 = 2008/1/6  *** 
        ''' 6 = 1/6/2008  *** 
        ''' 7 = 6/1/2008  *** 
        ''' 8 = 2008/1/6 15:30:15  *** 
        ''' 9 = 2008/01/06  *** 
        ''' 10 = 01/06/2008  *** 
        ''' 11 = 06/01/2008  *** 
        ''' 12 = 01/06/08  *** 
        ''' 13 = 06/01/08  *** 
        ''' 14 = 2008/01/06 15:30:15  *** 
        ''' 15 = 6-Enero-2008  *** 
        ''' 16 = 6-Ene-2008  *** 
        ''' 17 = 6 de Enero del 2008  *** 
        ''' 18 = Domingo, 6 de Enero del 2008 
        ''' </param>
        ''' <param name="DateSeparator">Date Separator: '/' or '-'</param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewDateTimeFormat(ByVal NamedFormat As Integer, ByVal DateSeparator As String, ByVal value As Object) As String
            If Information.IsDate(value) OrElse Regex.IsMatch(value.ToString, "\d\d:\d\d:\d\d") Then
                Select Case NamedFormat
                    Case 0 '// 1/6/2008 3:30:15 PM
                        Return Strings.FormatDateTime(Convert.ToDateTime(value), Microsoft.VisualBasic.DateFormat.GeneralDate)
                    Case 1 '// Sunday, January 06, 2008
                        Return Strings.FormatDateTime(Convert.ToDateTime(value), Microsoft.VisualBasic.DateFormat.LongDate)
                    Case 2 '// 1/6/2008
                        Return Strings.FormatDateTime(Convert.ToDateTime(value), Microsoft.VisualBasic.DateFormat.ShortDate)
                    Case 3 '// 15:30:15
                        Try
                            Dim dDate As Date = Convert.ToDateTime(value)
                            Return String.Format("{0:00}:{1:00}:{2:00}", dDate.Hour, dDate.Minute, dDate.Second)
                        Catch ex As Exception
                            Try
                                Dim dTime As TimeSpan = TimeSpan.Parse(value.ToString)
                                Return String.Format("{0:00}:{1:00}:{2:00}", dTime.Hours, dTime.Minutes, dTime.Seconds)
                            Catch generatedExceptionName As Exception
                                Return value.ToString
                            End Try
                        End Try
                    Case 4 '// 15:30
                        Try
                            Dim dDate As Date = Convert.ToDateTime(value)
                            Return String.Format("{0:00}:{1:00}", dDate.Hour, dDate.Minute)
                        Catch ex As Exception
                            Try
                                Dim dTime As TimeSpan = TimeSpan.Parse(value.ToString)
                                Return String.Format("{0:00}:{1:00}", dTime.Hours, dTime.Minutes)
                            Catch generatedExceptionName As Exception
                                Return value.ToString
                            End Try
                        End Try
                    Case 5 '// 2008/1/6
                        Try
                            Return String.Format("{0:yyyy" + DateSeparator + "M" + DateSeparator + "d}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 6 '// 1/6/2008
                        Try
                            Return String.Format("{0:M" + DateSeparator + "d" + DateSeparator + "yyyy}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 7 '// 6/1/2008
                        Try
                            Return String.Format("{0:d" + DateSeparator + "M" + DateSeparator + "yyyy}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 8 '// 2008/1/6 15:30:15
                        Try
                            Dim dDateTime As Date = Convert.ToDateTime(value)
                            Dim retValue As String = ewDateTimeFormat(5, "/", dDateTime)
                            If DateAndTime.Hour(dDateTime) <> 0 OrElse DateAndTime.Minute(dDateTime) <> 0 OrElse DateAndTime.Second(dDateTime) <> 0 Then
                                retValue += " " + ewDateTimeFormat(4, DateSeparator, dDateTime) + ":" + ewZeroPad(DateAndTime.Second(dDateTime), 2)
                            End If
                            Return retValue
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 9 '// 2008/01/06
                        Try
                            Return String.Format("{0:yyyy" + DateSeparator + "MM" + DateSeparator + "dd}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 10 '// 01/06/2008
                        Try
                            Return String.Format("{0:MM" + DateSeparator + "dd" + DateSeparator + "yyyy}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 11 '// 06/01/2008
                        Try
                            Return String.Format("{0:dd" + DateSeparator + "MM" + DateSeparator + "yyyy}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 12 '// 01/06/08
                        Try
                            Return String.Format("{0:MM" + DateSeparator + "dd" + DateSeparator + "yy}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 13 '// 06/01/08
                        Try
                            Return String.Format("{0:dd" + DateSeparator + "MM" + DateSeparator + "yy}", Convert.ToDateTime(value))
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 14 '// 2008/01/06 15:30:15
                        Dim dDateTime As Date = Convert.ToDateTime(value)
                        Dim retValue As String = ewDateTimeFormat(9, DateSeparator, dDateTime)
                        If DateAndTime.Hour(dDateTime) <> 0 OrElse DateAndTime.Minute(dDateTime) <> 0 OrElse DateAndTime.Second(dDateTime) <> 0 Then
                            retValue += " " + ewDateTimeFormat(4, DateSeparator, dDateTime) + ":" + ewZeroPad(DateAndTime.Second(dDateTime), 2)
                        End If
                        Return retValue
                    Case 15 '// 06-Enero-2008
                        Try
                            Dim dDateTime As Date = Convert.ToDateTime(value)
                            Dim dDay As String = dDateTime.Day.ToString
                            Dim dMonth As String = ewDateTimeMonthName(value)
                            Dim dYear As String = dDateTime.Year.ToString
                            Return String.Format("{0}" + DateSeparator + "{1}" + DateSeparator + "{2}", dDay, dMonth, dYear)
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 16 '// 06-Ene-2008
                        Try
                            Dim dDateTime As Date = Convert.ToDateTime(value)
                            Dim dDay As String = dDateTime.Day.ToString
                            Dim dMonth As String = ewDateTimeMonthName(value).Substring(0, 3)
                            Dim dYear As String = dDateTime.Year.ToString
                            Return String.Format("{0}" + DateSeparator + "{1}" + DateSeparator + "{2}", dDay, dMonth, dYear)
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 17 '// 6 de Enero del 2008
                        Try
                            Dim dDateTime As Date = Convert.ToDateTime(value)
                            Dim dDay As String = dDateTime.Day.ToString
                            Dim dMonth As String = ewDateTimeMonthName(value)
                            Dim dYear As String = dDateTime.Year.ToString
                            Return String.Format("{0} de {1} del {2}", dDay, dMonth, dYear)
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                    Case 18 '// Martes, 27 de Mayo del 2008
                        Try
                            Dim dDateTime As Date = Convert.ToDateTime(value)
                            Dim dDayOfWeek As String = ewDateTimeDayOfWeek(value)
                            Dim dDay As String = dDateTime.Day.ToString
                            Dim dMonth As String = ewDateTimeMonthName(value)
                            Dim dYear As String = dDateTime.Year.ToString
                            Return String.Format("{0}, {1} de {2} del {3}", dDayOfWeek, dDay, dMonth, dYear)
                        Catch ex As Exception
                            Return value.ToString
                        End Try
                End Select
            End If
            Return value.ToString
        End Function

        ''' <summary>
        ''' Get spanish month name from date
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function ewDateTimeMonthName(ByVal value As Object) As String
            If Information.IsDate(value) OrElse Regex.IsMatch(value.ToString, "\d\d:\d\d:\d\d") Then
                Dim dDateTime As Date = Convert.ToDateTime(value)
                Select Case dDateTime.Month
                    Case 1
                        Return "Enero"
                    Case 2
                        Return "Febrero"
                    Case 3
                        Return "Marzo"
                    Case 4
                        Return "Abril"
                    Case 5
                        Return "Mayo"
                    Case 6
                        Return "Junio"
                    Case 7
                        Return "Julio"
                    Case 8
                        Return "Agosto"
                    Case 9
                        Return "Septiembre"
                    Case 10
                        Return "Octubre"
                    Case 11
                        Return "Noviembre"
                    Case 12
                        Return "Diciembre"
                End Select
            End If
            Return value.ToString
        End Function

        ''' <summary>
        ''' Get spanish day name from date of the week
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function ewDateTimeDayOfWeek(ByVal value As Object) As String
            If Information.IsDate(value) OrElse Regex.IsMatch(value.ToString, "\d\d:\d\d:\d\d") Then
                Dim dDateTime As Date = Convert.ToDateTime(value)
                Select Case dDateTime.DayOfWeek
                    Case DayOfWeek.Sunday
                        Return "Domingo"
                    Case DayOfWeek.Monday
                        Return "Lunes"
                    Case DayOfWeek.Tuesday
                        Return "Martes"
                    Case DayOfWeek.Wednesday
                        Return "Miércoles"
                    Case DayOfWeek.Thursday
                        Return "Jueves"
                    Case DayOfWeek.Friday
                        Return "Viernes"
                    Case DayOfWeek.Saturday
                        Return "Sábado"
                End Select
            End If
            Return value.ToString
        End Function

        ''' <summary>
        ''' Unformat DateTime
        ''' </summary>
        ''' <param name="dDate"></param>
        ''' <param name="NamedFormat"></param>
        ''' <param name="DateSeparator"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewUnFormatDateTime(ByVal dDate As String, ByVal NamedFormat As Integer, ByVal DateSeparator As Char) As String
            Dim arDateTime As String()
            Dim arDate As String()
            Dim strDt As String = String.Empty
            dDate = dDate.Trim
            While dDate.IndexOf("  ") > 0
                dDate = dDate.Remove("  ", " ")
            End While
            arDateTime = dDate.Split(New Char() {" "c})
            If arDateTime.Length < 1 Then
                Return dDate
            End If
            arDate = arDateTime(0).Split(New Char() {DateSeparator})
            If arDate.Length = 3 Then
                If NamedFormat = 6 Then
                    strDt = arDate(2) + "/" + arDate(0) + "/" + arDate(1)
                ElseIf NamedFormat = 7 Then
                    strDt = arDate(2) + "/" + arDate(1) + "/" + arDate(0)
                Else ' NamedFormat = 5 or other
                    strDt = arDate(0) + "/" + arDate(1) + "/" + arDate(2)
                End If
                If arDateTime.Length > 1 Then
                    If ewCheckDateTime(arDateTime(1)) Then
                        Return strDt + " " + arDateTime(1) ' is time
                    End If
                End If
                Return strDt
            Else
                Return dDate
            End If
        End Function

        ''' <summary>
        ''' Check if not null
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewNotNull(ByVal sValue As Object) As Boolean
            If sValue IsNot Nothing Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Check Int16
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCheckInt16(ByVal sValue As Object) As Boolean
            If sValue Is Nothing Then Return True
            If sValue.ToString = String.Empty Then Return True
            Try
                Dim i As Integer = Convert.ToInt16(sValue)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check Int32
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCheckInt32(ByVal sValue As Object) As Boolean
            If sValue Is Nothing Then Return True
            If sValue.ToString = String.Empty Then Return True
            Try
                Dim i As Integer = Convert.ToInt32(sValue)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check Int64
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCheckInt64(ByVal sValue As Object) As Boolean
            If sValue Is Nothing Then Return True
            If sValue.ToString = String.Empty Then Return True
            Try
                Dim i As Long = Convert.ToInt64(sValue)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check Decimal
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCheckDecimal(ByVal sValue As Object) As Boolean
            If sValue Is Nothing Then Return True
            If sValue.ToString = String.Empty Then Return True
            Try
                Dim d As Decimal = Convert.ToDecimal(sValue)
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check Double
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCheckDouble(ByVal sValue As Object) As Boolean
            If sValue Is Nothing Then Return True
            If sValue.ToString = String.Empty Then Return True
            Try
                Dim d As Double = Convert.ToDouble(sValue)
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check Single
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCheckSingle(ByVal sValue As Object) As Boolean
            If sValue Is Nothing Then Return True
            If sValue.ToString() = String.Empty Then Return True
            Try
                Dim s As Single = Convert.ToSingle(sValue)
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check DateTime
        ''' </summary>
        ''' <param name="sValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewCheckDateTime(ByVal sValue As Object) As Boolean
            If sValue Is Nothing Then Return True
            If sValue.ToString = String.Empty Then Return True
            Try
                Dim d As Date = Convert.ToDateTime(sValue)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Zero Pad
        ''' </summary>
        ''' <param name="m"></param>
        ''' <param name="t"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function ewZeroPad(ByVal m As Integer, ByVal t As Integer) As String
            Return New String("0"c, t - Convert.ToString(m).Length) + Convert.ToString(m)
        End Function

        ''' <summary>
        ''' Lookup Boolean (Checkbox)
        ''' </summary>
        ''' <param name="s"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewBooleanLookup(ByVal s As Object) As Boolean
            If s Is Nothing OrElse DirectCast(s, Boolean) = False Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Lookup Boolean (Radio)
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="v1"></param>
        ''' <param name="v2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewBooleanLookup(ByVal s As Object, ByVal v1 As String, ByVal v2 As String) As String
            If s Is Nothing OrElse DirectCast(s, Boolean) = False Then
                Return v2
            Else
                Return v1
            End If
        End Function

        ''' <summary>
        ''' Average Value
        ''' </summary>
        ''' <param name="cnt"></param>
        ''' <param name="tot"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ewAverage(ByVal cnt As Integer, ByVal tot As Double) As String
            Dim val As Double
            If cnt > 0 Then
                val = tot / cnt
            Else
                val = tot
            End If
            Return String.Format("{0}", val)
        End Function


    End Class

    Public Class Email

#Region "Enumerations"
        Enum EmailFormat
            Text
            HTML
        End Enum
#End Region

#Region "Properties"
        Private _Subject As String
        Public ReadOnly Property Subject() As String
            Get
                Return _Subject
            End Get
        End Property

        Private _From As String
        Public ReadOnly Property From() As String
            Get
                Return _From
            End Get
        End Property

        Private _To As String
        Public ReadOnly Property [To]() As String
            Get
                Return _To
            End Get
        End Property

        Private _Cc As String
        Public ReadOnly Property Cc() As String
            Get
                Return _Cc
            End Get
        End Property

        Private _Bcc As String
        Public ReadOnly Property Bcc() As String
            Get
                Return _Bcc
            End Get
        End Property

        Private _Format As String
        Public ReadOnly Property Format() As String
            Get
                Return _Format
            End Get
        End Property

        Private _EmailContent As String
        Public ReadOnly Property EmailContent() As String
            Get
                Return _EmailContent
            End Get
        End Property
#End Region

        ''' <summary>
        ''' Load Email
        ''' </summary>
        ''' <param name="Content"></param>
        ''' <remarks></remarks>
        Public Sub LoadEmail(ByVal Content As String)
            Dim sWrk As String = Content
            Dim sHeader As String, sName As String, sValue As String
            Dim arrHeader As String()
            Const CrLf As String = "" & Chr(13) & "" & Chr(10) & ""
            If sWrk.Length > 0 Then
                ' Locate Header & Mail Content
                'int i = InStr(sWrk, vbCrLf&vbCrLf);
                Dim i As Integer = sWrk.IndexOf(CrLf + CrLf)
                If i > 0 Then
                    sHeader = sWrk.Substring(0, i)
                    _EmailContent = sWrk.Substring(i + 4)
                    arrHeader = sHeader.Split(New String() {CrLf}, StringSplitOptions.None)
                    For j As Integer = 0 To arrHeader.Length - 1
                        i = arrHeader(j).IndexOf(":")
                        If i > 0 Then
                            sName = arrHeader(j).Substring(0, i).Trim()
                            sValue = arrHeader(j).Substring(i + 1).Trim()
                            Select Case sName.ToLower()
                                Case "subject"
                                    _Subject = sValue
                                    Exit Select
                                Case "from"
                                    _From = sValue
                                    Exit Select
                                Case "to"
                                    _To = sValue
                                    Exit Select
                                Case "cc"
                                    _Cc = sValue
                                    Exit Select
                                Case "bcc"
                                    _Bcc = sValue
                                    Exit Select
                                Case "format"
                                    _Format = sValue
                                    Exit Select
                            End Select
                        End If
                    Next
                End If
            End If
        End Sub

        ''' <summary>
        ''' Send electronic mail to a Simple Mail Transfer Protocol (SMTP) server for delivery
        ''' </summary>
        ''' <param name="sTo">The e-mail address of the recipient</param>
        ''' <param name="sSubject">The subject of the message</param>
        ''' <param name="sMail">The message text</param>
        Public Overloads Function SendEmail(ByVal sTo As String, ByVal sSubject As String, ByVal sMail As String) As String
            Dim sFrom As String = String.Empty
            Dim sCc As String = String.Empty
            Dim sBcc As String = String.Empty
            Dim eFormat As EmailFormat = EmailFormat.Text
            Dim sServer As String = String.Empty
            Dim iPort As Integer = 0
            Dim sMessage As String = SendEmailMessage(sFrom, sTo, sCc, sBcc, sSubject, sMail, eFormat, sServer, iPort)
            Return sMessage
        End Function

        ''' <summary>
        ''' Send electronic mail to a Simple Mail Transfer Protocol (SMTP) server for delivery
        ''' </summary>
        ''' <param name="sTo">The e-mail address of the recipient</param>
        ''' <param name="sSubject">The subject of the message</param>
        ''' <param name="sMail">The message text</param>
        ''' <param name="sFormat">The format of the message</param>
        Public Overloads Function SendEmail(ByVal sTo As String, ByVal sSubject As String, ByVal sMail As String, ByVal sFormat As String) As String
            Dim sFrom As String = String.Empty
            Dim sCc As String = String.Empty
            Dim sBcc As String = String.Empty
            Dim eFormat As EmailFormat
            If sFormat.ToLower() <> "html" Then
                eFormat = EmailFormat.Text
            Else
                eFormat = EmailFormat.HTML
            End If
            Dim sServer As String = String.Empty
            Dim iPort As Integer = 0
            Dim sMessage As String = SendEmailMessage(sFrom, sTo, sCc, sBcc, sSubject, sMail, eFormat, sServer, iPort)
            Return sMessage
        End Function

        ''' <summary>
        ''' Send electronic mail to a Simple Mail Transfer Protocol (SMTP) server for delivery
        ''' </summary>
        ''' <param name="sTo">The e-mail address of the recipient</param>
        ''' <param name="sCc">The e-mail address of the carbon copies (Cc) recipient</param>
        ''' <param name="sBcc">The e-mail address of the blind carbon copies (BCC) recipient</param>
        ''' <param name="sSubject">The subject of the message</param>
        ''' <param name="sMail">The message text</param>
        ''' <param name="sFormat">The format of the message</param>
        Public Overloads Function SendEmail(ByVal sTo As String, ByVal sCc As String, ByVal sBcc As String, ByVal sSubject As String, ByVal sMail As String, ByVal sFormat As String) As String
            Dim sFrom As String = String.Empty
            Dim eFormat As EmailFormat
            If sFormat.ToLower() <> "html" Then
                eFormat = EmailFormat.Text
            Else
                eFormat = EmailFormat.HTML
            End If
            Dim sServer As String = String.Empty
            Dim iPort As Integer = 0
            Dim sMessage As String = SendEmailMessage(sFrom, sTo, sCc, sBcc, sSubject, sMail, eFormat, sServer, iPort)
            Return sMessage
        End Function


        ''' <summary>
        ''' Send electronic mail to a Simple Mail Transfer Protocol (SMTP) server for delivery
        ''' </summary>
        ''' <param name="sFrom">The e-mail address of the sender</param>
        ''' <param name="sTo">The e-mail address of the recipient</param>
        ''' <param name="sCc">The e-mail address of the carbon copies (Cc) recipient</param>
        ''' <param name="sBcc">The e-mail address of the blind carbon copies (BCC) recipient</param>
        ''' <param name="sSubject">The subject of the message</param>
        ''' <param name="sMail">The message text</param>
        ''' <param name="eFormat">The format of the message, TEXT</param>
        ''' <param name="sServer">SMTP Server name</param>
        ''' <param name="iPort">SMTP Server port</param>
        ''' <returns>Return SMTP exception message</returns>
        Private Function SendEmailMessage(ByVal sFrom As String, ByVal sTo As String, ByVal sCc As String, ByVal sBcc As String, ByVal sSubject As String, ByVal sMail As String, ByVal eFormat As EmailFormat, ByVal sServer As String, ByVal iPort As Integer) As String
            Dim arCc As String = String.Empty
            Dim arBcc As String = String.Empty
            Dim sMessage As String = String.Empty
            Dim mail As MailMessage = New MailMessage()
            If sFrom <> String.Empty Then mail.From = New MailAddress(sFrom) ' If there is a from email address, set it to the e-mail sender
            If sTo <> String.Empty Then mail.[To].Add(sTo) 'If there is a to email address, set it to the e-mail recipient
            If sCc <> String.Empty Then mail.CC.Add(sCc) ' If there is a Cc email address, set it to the e-mail carbon copies (Cc) recipient
            If sBcc <> String.Empty Then mail.Bcc.Add(sBcc) ' If there is a BCC email address, set it to the e-mail blind carbo copies (BCC) recipient
            mail.Subject = sSubject ' Set the e-mail subject
            mail.SubjectEncoding = Encoding.GetEncoding(My.Settings.EmailEncoding) ' Set the e-mail subject encoding to send non-ascii content: iso-8859-1; utf-8;
            mail.Body = sMail ' Set the e-mail body
            mail.BodyEncoding = Encoding.GetEncoding(My.Settings.EmailEncoding) ' Set the e-mail body encoding to send non-ascii content: iso-8859-1; utf-8;
            If eFormat = EmailFormat.HTML Then ' Set the e-mail body to html or plain text
                mail.IsBodyHtml = True
            Else
                mail.IsBodyHtml = False
            End If
            Dim smtp As SmtpClient = New SmtpClient() ' If there is a mailserver name, set it to the SMTP Host
            If sServer <> String.Empty Then smtp.Host = sServer
            If iPort > 0 Then smtp.Port = iPort ' If there is a server port, set it to SMTP Port
            Try ' Send the e-mail message
                smtp.Send(mail)
            Catch e As System.Net.Mail.SmtpException
                sMessage = "Message: " + e.Message + ", Status Code: " + e.StatusCode.ToString()
            Finally
                smtp = Nothing
            End Try
            Return sMessage ' Return SMTP exception
        End Function

        ''' <summary>
        ''' Get Mail Server (SMTP).
        ''' </summary>
        ''' <return>Mail Server (SMTP) string.</return>
        Public Function GetMailServerName() As String
            Try
                Dim smtp As System.Net.Configuration.SmtpSection
                smtp = ConfigurationManager.GetSection("system.net/mailSettings/smtp")
                Return smtp.Network.Host
            Catch ex As Exception
                Return "Message: " + ex.Message.ToString
            End Try
        End Function

        ''' <summary>
        ''' Get Mail Sender (FROM).
        ''' </summary>
        ''' <return>Mail Sender (FROM) string.</return>
        Public Function GetMailSenderAddress() As String
            Try
                Dim smtp As System.Net.Configuration.SmtpSection
                smtp = ConfigurationManager.GetSection("system.net/mailSettings/smtp")
                Return smtp.From
            Catch ex As Exception
                Return "Message: " + ex.Message.ToString
            End Try
        End Function

        ''' <summary>
        ''' E-mail Address Parser
        ''' </summary>
        Public Shared Function EmailParser(ByVal Email As String) As String
            Dim sTmp As String = Email.Trim
            If (Not String.IsNullOrEmpty(sTmp)) Then
                If sTmp.Contains("@") Then
                    If sTmp.IndexOf("@", sTmp.IndexOf("@", 0) + 1) > 0 Then ' For multiple email address
                        If sTmp.Contains("|") Then sTmp = EmailParserEx(sTmp, "|")
                        If sTmp.Contains("/") Then sTmp = EmailParserEx(sTmp, "/")
                        If sTmp.Contains(":") Then sTmp = EmailParserEx(sTmp, ":")
                        If sTmp.Contains(";") Then sTmp = EmailParserEx(sTmp, ";")
                        sTmp = EmailParserEx(sTmp, ",")
                    End If
                    If sTmp.Contains("ñ") Then sTmp = sTmp.Replace("ñ", "n") ' Replace invalid characters
                    If sTmp.Contains("Ñ") Then sTmp = sTmp.Replace("Ñ", "N")
                End If
                sTmp = ParseLastSeparator(sTmp)
            End If
            Return sTmp
        End Function

        ''' <summary>
        ''' E-mail Address Remove Last Separator
        ''' </summary>
        Public Shared Function ParseLastSeparator(ByVal Email As String) As String
            Dim sTmp As String = Email.Trim
            If (Not String.IsNullOrEmpty(sTmp)) Then
                If sTmp.Contains(";") Then sTmp = sTmp.Replace(";", ",")
                If sTmp.EndsWith(",") Then sTmp = sTmp.Substring(0, sTmp.Length - 1)
                End If
            Return sTmp
        End Function

        ''' <summary>
        ''' E-mail Address Parser
        ''' </summary>
        Private Shared Function EmailParserEx(ByVal Email As String, ByVal Separator As String) As String
            Dim sTmp As String = Email.Trim
            If (Not String.IsNullOrEmpty(sTmp)) And (Not String.IsNullOrEmpty(Separator)) Then
                sTmp = sTmp.Replace(Separator & " ", ",")
                sTmp = sTmp.Replace(" " & Separator, ",")
                sTmp = sTmp.Replace(Separator, ",")
            End If
            Return sTmp
        End Function

        ''' <summary>
        ''' Get E-mail Address Domain
        ''' </summary>
        ''' <param name="email"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetEmailDomain(ByVal Email As String) As String
            Dim wrkstr As String = Email.Trim
            If (Not String.IsNullOrEmpty(wrkstr)) Then
                If wrkstr.Contains("@") Then
                    Return wrkstr.Substring(wrkstr.IndexOf("@", 0) + 1)
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End Function


        'Private Function SendEmailMessage()
        '    **********************************************************************
        '    If sTo <> String.Empty Then
        '        If IsEmailAddress(sTo) Then
        '            mail.[To].Add(sTo)
        '        Else
        '            Dim sArrTo As String = EmailParser(sTo)
        '            Dim arrTo() As String = sArrTo.Split(";")
        '            For Each aTo As String In arrTo
        '                mail.[To].Add(aTo)
        '            Next
        '        End If
        '    End If
        '    **********************************************************************
        '    If sCc <> String.Empty Then
        '        If IsEmailAddress(sCc) Then
        '            mail.CC.Add(sCc)
        '        Else
        '            Dim sArrCc As String = EmailParser(sCc)
        '            Dim arrCc() As String = sArrCc.Split(";")
        '            For Each aCc As String In arrCc
        '                mail.CC.Add(aCc)
        '            Next
        '        End If
        '    End If
        '    **********************************************************************
        '    If sBcc <> String.Empty Then
        '        If IsEmailAddress(sBcc) Then
        '            mail.Bcc.Add(sBcc)
        '        Else
        '            Dim sArrBcc As String = EmailParser(sBcc)
        '            Dim arrBcc() As String = sArrBcc.Split(";")
        '            For Each aBcc As String In arrBcc
        '                mail.Bcc.Add(aBcc)
        '            Next
        '        End If
        '    End If
        '    **********************************************************************
        '    This is the code for getting the SMTP Host fron the web.config
        '    Dim smtp As System.Net.Configuration.SmtpSection
        '    smtp = ConfigurationManager.GetSection("system.net/mailSettings/smtp")
        '    **********************************************************************
        'End Function
    End Class

End Module