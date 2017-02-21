Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Validators
#Region "Enumerations"
    Private Enum RegExpPattern
        Address
        CreditCard
        Currency
        EmailAddress
        EmailAddressRFC
        GUID
        IntegerValue
        Name
        Nickname
        Password
        SecretAnswer
        TrackingDHL
        TrackingEPS
        TrackingFedEx
        TrackingUPS
        TrackingUSPS
        URLAddress
        USDate
        USPhone
        USSSN
        USZipCode
        Username
    End Enum
#End Region

    ''' <summary>
    ''' Is Email Address
    ''' </summary>
    ''' <param name="EmailAddress">Email Address to validate</param>
    Public Shared Function IsEmail(ByVal EmailAddress As String) As Boolean
        If EmailAddress Is Nothing OrElse EmailAddress.Length = 0 Then
            Return False
        Else
            If RegExpValidator(EmailAddress, RegExpPattern.EmailAddressRFC) = True Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Is Error Exception Message
    ''' </summary>
    ''' <param name="ErrorMessage">Error Message</param>
    ''' <remarks>TODO: Create regex for function</remarks>
    Public Shared Function IsErrorMessage(ByVal ErrorMessage As String) As Boolean
        If ErrorMessage Is Nothing OrElse ErrorMessage.Length = 0 Then
            Return False
        Else
            If Left(ErrorMessage, 6) = "ERROR:" Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Is GUID
    ''' </summary>
    ''' <param name="GUID"></param>
    Public Shared Function IsGUID(ByVal GUID As String) As Boolean
        If GUID Is Nothing OrElse GUID.Length = 0 Then
            Return False
        Else
            If RegExpValidator(GUID, RegExpPattern.GUID) = True Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Is Nickname
    ''' </summary>
    ''' <param name="Nickname"></param>
    Public Shared Function IsNickname(ByVal Nickname As String) As Boolean
        If Nickname Is Nothing OrElse Nickname.Length = 0 Then
            Return False
        Else
            If RegExpValidator(Nickname, RegExpPattern.Nickname) = True Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Is Numeric Key
    ''' </summary>
    ''' <param name="NumericValue"></param>
    Public Shared Function IsNumericKey(ByVal NumericValue As String) As Boolean
        If NumericValue Is Nothing OrElse NumericValue.Length = 0 Then
            Return False
        Else
            If RegExpValidator(NumericValue, RegExpPattern.IntegerValue) = True Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Is Password
    ''' </summary>
    ''' <param name="Password"></param>
    Public Shared Function IsPassword(ByVal Password As String) As Boolean
        If Password Is Nothing OrElse Password.Length = 0 Then
            Return False
        Else
            If RegExpValidator(Password, RegExpPattern.Password) = True Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Is Phone Number
    ''' </summary>
    ''' <param name="PhoneNumber">Phone Number</param>
    Public Shared Function IsPhoneNumber(ByVal PhoneNumber As String) As Boolean
        If PhoneNumber Is Nothing OrElse PhoneNumber.Length = 0 Then
            Return False
        Else
            If RegExpValidator(PhoneNumber, RegExpPattern.USPhone) = True Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Is Secret Answer
    ''' </summary>
    ''' <param name="SecretAnswer">Secret Answer</param>
    Public Shared Function IsSecretAnswer(ByVal SecretAnswer As String) As Boolean
        If SecretAnswer Is Nothing OrElse SecretAnswer.Length = 0 Then
            Return False
        Else
            If RegExpValidator(SecretAnswer, RegExpPattern.SecretAnswer) = True Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Class Tracking
        ''' <summary>
        ''' Is DHL Tracking Number
        ''' </summary>
        ''' <param name="TrackingNumber">Tracking Number</param>
        Public Shared Function IsDHLTrackingNumber(ByVal TrackingNumber As String) As Boolean
            If TrackingNumber Is Nothing OrElse TrackingNumber.Length = 0 Then
                Return False
            Else
                If RegExpValidator(TrackingNumber, RegExpPattern.TrackingDHL) = True Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        ''' <summary>
        ''' Is EPS Air WayBill Number or Tracking Number
        ''' </summary>
        ''' <param name="TrackingNumber">Tracking Number</param>
        Public Shared Function IsEPSTrackingNumber(ByVal TrackingNumber As String) As Boolean
            If TrackingNumber Is Nothing OrElse TrackingNumber.Length = 0 Then
                Return False
            Else
                If RegExpValidator(TrackingNumber, RegExpPattern.TrackingEPS) = True Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        ''' <summary>
        ''' Is FedEx Tracking Number
        ''' </summary>
        ''' <param name="TrackingNumber">Tracking Number</param>
        Public Shared Function IsFedExTrackingNumber(ByVal TrackingNumber As String) As Boolean
            If TrackingNumber Is Nothing OrElse TrackingNumber.Length = 0 Then
                Return False
            Else
                If RegExpValidator(TrackingNumber, RegExpPattern.TrackingFedEx) = True Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        ''' <summary>
        ''' Is UPS Tracking Number
        ''' </summary>
        ''' <param name="TrackingNumber">Tracking Number</param>
        Public Shared Function IsUPSTrackingNumber(ByVal TrackingNumber As String) As Boolean
            If TrackingNumber Is Nothing OrElse TrackingNumber.Length = 0 Then
                Return False
            Else
                If RegExpValidator(TrackingNumber, RegExpPattern.TrackingUPS) = True Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        ''' <summary>
        ''' Is USPS Tracking Number
        ''' </summary>
        ''' <param name="TrackingNumber">Tracking Number</param>
        Public Shared Function IsUSPSTrackingNumber(ByVal TrackingNumber As String) As Boolean
            If TrackingNumber Is Nothing OrElse TrackingNumber.Length = 0 Then
                Return False
            Else
                If RegExpValidator(TrackingNumber, RegExpPattern.TrackingUSPS) = True Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        ''' <summary>
        ''' Is Tracking Number
        ''' </summary>
        ''' <param name="TrackingNumber">Tracking Number</param>
        Public Shared Function IsTrackingNumber(ByVal TrackingNumber As String) As Boolean
            If TrackingNumber Is Nothing OrElse TrackingNumber.Length = 0 Then
                Return False
            ElseIf IsUPSTrackingNumber(TrackingNumber) Then
                Return True
            ElseIf IsUSPSTrackingNumber(TrackingNumber) Then
                Return True
            ElseIf IsFedExTrackingNumber(TrackingNumber) Then
                Return True
            ElseIf IsDHLTrackingNumber(TrackingNumber) Then
                Return True
            ElseIf IsEPSTrackingNumber(TrackingNumber) Then
                Return True
            Else
                Return False
            End If
        End Function

    End Class

    ''' <summary>
    ''' Regular Expression Validator
    ''' </summary>
    ''' <param name="wrkstr">Working string</param>
    ''' <param name="pattern">Regular expression pattern to match</param>
    Private Shared Function RegExpValidator(ByVal wrkstr As String, ByVal pattern As RegExpPattern) As Boolean
        Dim sPattern As String = ""

        Select Case pattern
            Case RegExpPattern.Address
                sPattern = "^[a-zA-Z0-9\s.\-]+$"
            Case RegExpPattern.CreditCard
                sPattern = "^((4\d{3})|(5[1-5]\d{2})|(6011))-?\d{4}-?\d{4}-?\d{4}|3[4,7]\d{13}$"
            Case RegExpPattern.Currency
                sPattern = "^\$?([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])?$"
            Case RegExpPattern.EmailAddress
                ' sPattern = "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
                sPattern = "^([a-zA-Z0-9_\-\.\47]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,7}|[0-9]{1,3})(\]?)$"
            Case RegExpPattern.EmailAddressRFC
                sPattern = "^(?:[\w\!\#\$\%\&\'\*\+\-\/\=\?\^\`\{\|\}\~]+\.)*[\w\!\#\$\%\&\'\*\+\-\/\=\?\^\`\{\|\}\~]+@(?:(?:(?:[a-zA-Z0-9](?:[a-zA-Z0-9\-](?!\.)){0,61}[a-zA-Z0-9]?\.)+[a-zA-Z0-9](?:[a-zA-Z0-9\-](?!$)){0,61}[a-zA-Z0-9]?)|(?:\[(?:(?:[01]?\d{1,2}|2[0-4]\d|25[0-5])\.){3}(?:[01]?\d{1,2}|2[0-4]\d|25[0-5])\]))$"
            Case RegExpPattern.GUID
                sPattern = "^([A-Fa-f0-9]{32}|[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}|\{[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}\})$"
            Case RegExpPattern.IntegerValue
                sPattern = "^\d{1,38}$"
            Case RegExpPattern.Name
                sPattern = "^[a-zA-Z]+(([\'\,\.\- ][a-zA-Z ])?[a-zA-Z]*)*$"
            Case RegExpPattern.Nickname
                sPattern = "^([a-zA-Z0-9._-]{4,32})$"
            Case RegExpPattern.Password
                sPattern = "^([a-zA-Z0-9]{4,20})$"
            Case RegExpPattern.SecretAnswer
                sPattern = "^([a-zA-Z0-9 ._-]{3,32})$"
            Case RegExpPattern.TrackingDHL
                ' 9 to 11 numbers (000000000) 
                ' 1 letter and 9 to 11 numbers (B000000000) 
                sPattern = "\b([0-9]{9,11}|[A-Z]\d{9,11})\b"
            Case RegExpPattern.TrackingEPS
                sPattern = "\b(MIA[0-9]*|[0-9]*LLC)\b"
            Case RegExpPattern.TrackingFedEx
                ' 12 numbers (0000 0000 0000 0000)
                ' 15 numbers (0000 0000 0000 0000 000)
                sPattern = "\b(\d\d\d\d ?\d\d\d\d ?\d\d\d\d|\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d)\b"
            Case RegExpPattern.TrackingUPS
                ' 1Z numbers and letters
                ' 1 letter and 10 numbers (A000 0000 000)
                sPattern = "\b(1Z ?[0-9A-Z]{3} ?[0-9A-Z]{3} ?[0-9A-Z]{2} ?[0-9A-Z]{4} ?[0-9A-Z]{3} ?[0-9A-Z]|[A-Z]\d\d\d ?\d\d\d\d ?\d\d\d)\b"
            Case RegExpPattern.TrackingUSPS
                ' 22 numbers (0000 0000 0000 0000 0000 00)
                ' 20 numbers (0000 0000 0000 0000 0000)
                ' 4 letters and 9 numbers (AA 000 000 000 AA) 
                ' 1 letter and 11 numbers (H000 0000 0000)
                sPattern = "\b(\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d|\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d ?\d\d\d\d|[A-Z]{2}\ ?\d\d\d ?\d\d\d ?\d\d\d ?[A-Z]{2}|H\d\d\d ?\d\d\d\d ?\d\d\d\d)\b"
            Case RegExpPattern.URLAddress
                sPattern = "(http|https|ftp)\://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?"
            Case RegExpPattern.USDate
                sPattern = "(((0[13578]|10|12)([-./])(0[1-9]|[12][0-9]|3[01])([-./])(\d{4}))|((0[469]|11)([-./])([0][1-9]|[12][0-9]|30)([-./])(\d{4}))|((2)([-./])(0[1-9]|1[0-9]|2[0-8])([-./])(\d{4}))|((2)(\.|-|\/)(29)([-./])([02468][048]00))|((2)([-./])(29)([-./])([13579][26]00))|((2)([-./])(29)([-./])([0-9][0-9][0][48]))|((2)([-./])(29)([-./])([0-9][0-9][2468][048]))|((2)([-./])(29)([-./])([0-9][0-9][13579][26])))"
            Case RegExpPattern.USPhone
                sPattern = "((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4}"
            Case RegExpPattern.USSSN
                sPattern = "\d{3}-\d{2}-\d{4}"
            Case RegExpPattern.USZipCode
                sPattern = "^(\d{5}-\d{4}|\d{5})$|^([a-zA-Z]\d[a-zA-Z] \d[a-zA-Z]\d)$"
            Case RegExpPattern.Username
                sPattern = "^([a-zA-Z0-9._-]{4,32})$"
        End Select

        If sPattern Is Nothing OrElse sPattern.Length = 0 Then
            Return False
        Else
            Dim regexp As Regex = New Regex(sPattern)
            Dim m As Match = regexp.Match(wrkstr)
            If m.Success Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

End Class