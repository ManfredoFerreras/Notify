Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class db

    Public Sub New()
    End Sub

#Region "Enumerations"
    Public Enum DataType
        eDecimal
        eString
        eInteger
        eDate
        eDateTime
        eBoolean
    End Enum

    Public Enum Direction
        eInput
        eOutput
        eInputOutput
        eReturn
    End Enum
#End Region

#Region "Execute Scalar"

    ''' <summary>
    ''' Execute SQL and return first value of first row
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ewExecuteScalar(ByVal Sql As String) As Object
        Try
            Return ewExecuteScalarEx(Sql, GetConnectionString())
        Catch
            Throw
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Execute SQL and return first value of first row
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <param name="ConnStr">DB Connection String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ewExecuteScalar(ByVal Sql As String, _
                                                     ByVal ConnStr As String) As Object
        Try
            Return ewExecuteScalarEx(Sql, ConnStr)
        Catch
            Throw
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Execute SQL and return first value of first row
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <param name="ConnStr">DB Connection String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function ewExecuteScalarEx(ByVal Sql As String, _
                                              ByVal ConnStr As String) As Object

        ' Create a new SqlConnection object
        Dim oConn As SqlConnection = New SqlConnection(ConnStr)

        Try
            oConn.Open()

            ' Create a new SqlCommand object
            Dim oCmd As New SqlCommand

            With oCmd
                .Connection = oConn
                .CommandType = CommandType.Text
                .CommandText = Sql
            End With

            Return oCmd.ExecuteScalar
        Catch
            Throw
            Return Nothing
        Finally
            oConn.Close()
        End Try

    End Function

#End Region

#Region "Execute NonQuery"

    ''' <summary>
    ''' Execute SQL and return the number of rows affected
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ewExecute(ByVal Sql As String) As Integer
        Try
            Return ewExecuteEx(Sql, GetConnectionString())
        Catch
            Throw
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Execute SQL and return the number of rows affected
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <param name="ConnStr">DB Connection String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ewExecute(ByVal Sql As String, _
                                               ByVal ConnStr As String) As Integer
        Try
            Return ewExecuteEx(Sql, ConnStr)
        Catch
            Throw
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Execute SQL and return the number of rows affected
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <param name="ConnStr">DB Connection String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function ewExecuteEx(ByVal Sql As String, _
                                        ByVal ConnStr As String) As Integer

        ' Create a new SqlConnection object
        Dim oConn As SqlConnection = New SqlConnection(ConnStr)

        Try
            oConn.Open()

            ' Create a new SqlCommand object
            Dim oCmd As New SqlCommand

            With oCmd
                .Connection = oConn
                .CommandType = CommandType.Text
                .CommandText = Sql
            End With

            Return oCmd.ExecuteNonQuery
        Catch
            Throw
            Return -1
        Finally
            oConn.Close()
        End Try

    End Function

#End Region

#Region "Get DataSet"

    ''' <summary>
    ''' Get dataset
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ewGetDataSet(ByVal Sql As String) As DataSet
        Try
            Return ewGetDataSetEx(Sql, GetConnectionString())
        Catch
            Throw
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Get dataset
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <param name="ConnStr">DB Connection String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ewGetDataSet(ByVal Sql As String, _
                                                  ByVal ConnStr As String) As DataSet
        Try
            Return ewGetDataSetEx(Sql, ConnStr)
        Catch
            Throw
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Get dataset
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <param name="ConnStr">DB Connection String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function ewGetDataSetEx(ByVal Sql As String, _
                                           ByVal ConnStr As String) As DataSet

        ' Create a new SqlConnection object
        Dim oConn As SqlConnection = New SqlConnection(ConnStr)

        Try
            ' Create a new DataAdapter using the Connection Object and SQL statement
            Dim oAdapter As New SqlDataAdapter(Sql, oConn)

            ' Create a new DataSet object to fill with data
            Dim oDs As DataSet = New DataSet

            ' Fill the DataSet with data from the DataAdapter object
            oAdapter.Fill(oDs)

            ' Return DataSet
            Return oDs
        Catch
            Throw
            Return Nothing
        Finally
            oConn.Close()
        End Try

    End Function

#End Region

#Region "Add SqlParameter"

    ''' <summary>
    ''' Add Stored Procedure String Parameter
    ''' </summary>
    ''' <param name="Command">A System.Data.SqlClient.SqlCommand object</param>
    ''' <param name="ParameterName">The name of the parameter to map</param>
    ''' <param name="Value">An System.Object that is the value of the System.Data.SqlClient.SqlParameter</param>
    ''' <param name="Size">The length of the parameter</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub AddParameter(ByVal Command As SqlCommand, _
                                             ByVal ParameterName As String, _
                                             ByVal Value As Object, _
                                             ByVal Size As Integer)

        AddParameterEx(Command, ParameterName, DataType.eString, Direction.eInput, Value, Size)

    End Sub

    ''' <summary>
    ''' Add Stored Procedure Parameter
    ''' </summary>
    ''' <param name="Command">A System.Data.SqlClient.SqlCommand object</param>
    ''' <param name="ParameterName">The name of the parameter to map</param>
    ''' <param name="DataType">One of the System.Data.DbType values</param>
    ''' <param name="Direction">One of the System.Data.ParameterDirection values</param>
    ''' <param name="Value">An System.Object that is the value of the System.Data.SqlClient.SqlParameter</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub AddParameter(ByVal Command As SqlCommand, _
                                             ByVal ParameterName As String, _
                                             ByVal DataType As DataType, _
                                             ByVal Direction As Direction, _
                                             ByVal Value As Object)

        AddParameterEx(Command, ParameterName, DataType, Direction, Value, 0)

    End Sub

    ''' <summary>
    ''' Add Stored Procedure Parameter
    ''' </summary>
    ''' <param name="Command">A System.Data.SqlClient.SqlCommand object</param>
    ''' <param name="ParameterName">The name of the parameter to map</param>
    ''' <param name="DataType">One of the System.Data.DbType values</param>
    ''' <param name="Direction">One of the System.Data.ParameterDirection values</param>
    ''' <param name="Value">An System.Object that is the value of the System.Data.SqlClient.SqlParameter</param>
    ''' <param name="Size">The length of the parameter</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub AddParameter(ByVal Command As SqlCommand, _
                                             ByVal ParameterName As String, _
                                             ByVal DataType As DataType, _
                                             ByVal Direction As Direction, _
                                             ByVal Value As Object, _
                                             ByVal Size As Integer)

        AddParameterEx(Command, ParameterName, DataType, Direction, Value, Size)

    End Sub

    ''' <summary>
    ''' Add Stored Procedure Parameter
    ''' </summary>
    ''' <param name="Command">A System.Data.SqlClient.SqlCommand object</param>
    ''' <param name="ParameterName">The name of the parameter to map</param>
    ''' <param name="DataType">One of the System.Data.DbType values</param>
    ''' <param name="Direction">One of the System.Data.ParameterDirection values</param>
    ''' <param name="Value">An System.Object that is the value of the System.Data.SqlClient.SqlParameter</param>
    ''' <param name="Size">The length of the parameter</param>
    ''' <remarks></remarks>
    Private Shared Sub AddParameterEx(ByVal Command As SqlCommand, _
                                      ByVal ParameterName As String, _
                                      ByVal DataType As DataType, _
                                      ByVal Direction As Direction, _
                                      ByVal Value As Object, _
                                      ByVal Size As Integer)
        ' Add params
        Dim SqlParam As New SqlParameter

        With SqlParam
            ' Parameter name
            .ParameterName = ParameterName

            ' Parameter data type
            Select Case DataType

                Case DataType.eString
                    .DbType = DbType.String
                    If Size > 0 Then
                        .Size = Size
                    End If
                    If IsNothing(Value) Then
                        .Value = System.DBNull.Value
                    ElseIf String.IsNullOrEmpty(Value.ToString) Then
                        .Value = System.DBNull.Value
                    Else
                        .Value = CType(Value, String)
                    End If

                Case DataType.eInteger
                    .DbType = DbType.Int32
                    If IsNothing(Value) Then
                        .Value = System.DBNull.Value
                    Else
                        .Value = CType(Value, Integer)
                    End If

                Case DataType.eDecimal
                    .DbType = DbType.Decimal
                    If IsNothing(Value) Then
                        .Value = System.DBNull.Value
                    Else
                        .Value = CType(Value, Decimal)
                    End If

                Case DataType.eBoolean
                    .DbType = DbType.Boolean
                    If IsNothing(Value) Then
                        .Value = System.DBNull.Value
                    Else
                        .Value = CType(Value, Boolean)
                    End If

                Case DataType.eDate
                    .DbType = DbType.Date
                    If IsNothing(Value) Then
                        .Value = System.DBNull.Value
                    Else
                        .Value = CType(Value, Date)
                    End If

                Case DataType.eDateTime
                    .DbType = DbType.DateTime
                    If IsNothing(Value) Then
                        .Value = System.DBNull.Value
                    Else
                        .Value = CType(Value, DateTime)
                    End If

            End Select

            ' Parameter direction
            Select Case Direction
                Case Direction.eInput
                    .Direction = ParameterDirection.Input
                Case Direction.eOutput
                    .Direction = ParameterDirection.Output
                Case Direction.eInputOutput
                    .Direction = ParameterDirection.InputOutput
                Case Direction.eReturn
                    .Direction = ParameterDirection.ReturnValue
            End Select

        End With

        ' Add command parameter
        Command.Parameters.Add(SqlParam)
    End Sub

#End Region

#Region "Other Methods"

    ''' <summary>
    ''' Verify Connection to DataBase
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function VerifyConnection(ByVal Sql As String) As Boolean
        Try
            ewExecuteScalar(Sql)
            Return True
        Catch
            Throw
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Verify Connection to DataBase
    ''' </summary>
    ''' <param name="Sql">SQL Query String</param>
    ''' <param name="ConnStr">DB Connection String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function VerifyConnection(ByVal Sql As String, _
                                                      ByVal ConnStr As String) As Boolean
        Try
            ewExecuteScalar(Sql, ConnStr)
            Return True
        Catch
            Throw
            Return False
        End Try
    End Function

#End Region

#Region "Connection String"

    ''' <summary>
    ''' Gets the Connection String
    ''' </summary>
    ''' <returns>Database connection string</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetConnectionString() As String
        Return GetConnectionStringEx("ApplicationConnectionString")
    End Function

    ''' <summary>
    ''' Get the Connection String by Connection String Name
    ''' </summary>
    ''' <param name="ConnName"></param>
    ''' <returns>Database connection string</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetConnectionString(ByVal ConnName As String) As String
        Return GetConnectionStringEx(ConnName)
    End Function

    ''' <summary>
    ''' Get the Connection String by Connection String Name
    ''' </summary>
    ''' <param name="ConnName"></param>
    ''' <returns>Database connection string</returns>
    ''' <remarks></remarks>
    Private Shared Function GetConnectionStringEx(ByVal ConnName As String) As String
        Dim sTmp As String = String.Empty
        Try
            sTmp = ConfigurationManager.ConnectionStrings(ConnName).ToString
            Return sTmp
        Catch ex As Exception
            Return "ERROR: GetConnectionStringEx() - " & ex.Message.ToString
        End Try
    End Function

#End Region

#Region "Adjust SQL String"

    ''' <summary>
    ''' Adjust SQL
    ''' </summary>
    ''' <param name="str">SQL String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function AdjustSql(ByVal str As String, ByVal len As Integer) As String

        Dim sWrk As String = str

        If len > 0 Then
            If sWrk.Length > len Then
                sWrk = sWrk.Substring(0, len)
            End If
        End If

        If Not String.IsNullOrEmpty(sWrk) Then
            sWrk = AdjustSql(sWrk)
        End If

        Return sWrk

    End Function

    ''' <summary>
    ''' Adjust SQL
    ''' </summary>
    ''' <param name="str">SQL String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function AdjustSql(ByVal str As String) As String

        Dim sWrk As String = str

        If Not String.IsNullOrEmpty(sWrk) Then
            sWrk = sWrk.Replace("'", "''")
            ' sWrk = sWrk.Replace("[", "[[]")
        End If

        Return sWrk

    End Function

    ''' <summary>
    ''' Adjust and Replace SQL
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="len"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ReplaceSql(ByVal str As String, ByVal len As Integer) As String

        Dim sWrk As String = str

        If len > 0 Then
            If sWrk.Length > len Then
                sWrk = sWrk.Substring(0, len)
            End If
        End If

        If Not String.IsNullOrEmpty(sWrk) Then
            sWrk = ReplaceSql(sWrk)
        End If

        Return sWrk

    End Function

    ''' <summary>
    ''' Adjust and Replace SQL
    ''' </summary>
    ''' <param name="str">SQL String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ReplaceSql(ByVal str As String) As String
        Dim sWrk As String = str.Replace("'", "´")
        Return sWrk
    End Function

#End Region

#Region "Converters"

    ''' <summary>
    ''' Get Date Time for SQL
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewDateTimeSql() As String

        Return Utilities.DataFormat.ewDateTimeFormat(14, "-", Date.Now)

    End Function

    ''' <summary>
    ''' Converts the string representation of a number in a specified base to an equivalent 16-bit signed integer, or null
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToIntegerNullable(ByVal value As Object) As System.Nullable(Of Integer)

        Dim returnValue As Integer

        If Integer.TryParse(value.ToString(), returnValue) Then
            Return returnValue
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to Integer, or 0
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToInteger(ByVal value As Object) As Integer

        Dim returnValue As Integer

        If Integer.TryParse(value.ToString(), returnValue) Then
            Return returnValue
        Else
            Return 0
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to string, return string empty if is blank
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToString(ByVal value As Object) As String

        Dim returnValue As String
        If Not String.IsNullOrEmpty(value.ToString.Trim) Then
            returnValue = value.ToString.Trim
            Return returnValue
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to string, return nothing if is blank
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToStringNullable(ByVal value As Object) As String

        Dim returnValue As String
        If Not String.IsNullOrEmpty(value.ToString.Trim) Then
            returnValue = value.ToString.Trim
            Return returnValue
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to string upper, return string empty if is blank
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToStringUpper(ByVal value As Object) As String

        Dim returnValue As String
        If Not String.IsNullOrEmpty(value.ToString.Trim) Then
            returnValue = value.ToString.ToUpper.Trim
            Return returnValue
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to string lower, return string empty if is blank
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToStringLower(ByVal value As Object) As String

        Dim returnValue As String
        If Not String.IsNullOrEmpty(value.ToString.Trim) Then
            returnValue = value.ToString.ToLower.Trim
            Return returnValue
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to an equivalent Boolean value, or null.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToBool(ByVal value As Object) As System.Nullable(Of Boolean)

        Dim returnValue As Boolean

        If Boolean.TryParse(value.ToString(), returnValue) Then
            Return returnValue
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to DateTime value, or null.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToDateTime(ByVal value As Object) As System.Nullable(Of DateTime)

        If value IsNot Nothing Then
            Dim returnValue As DateTime
            If DateTime.TryParse(value.ToString(), returnValue) Then
                Return returnValue
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to a decimal number, or null
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToDecimalNullable(ByVal value As Object) As Decimal

        Dim returnValue As Decimal
        If Decimal.TryParse(value.ToString(), returnValue) Then
            Return returnValue
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to a decimal number, or 0
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToDecimal(ByVal value As Object) As Decimal

        Dim returnValue As Decimal
        If Decimal.TryParse(value.ToString(), returnValue) Then
            Return returnValue
        Else
            Return 0
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to a double-precision floating point number, or null.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToDouble(ByVal value As Object) As System.Nullable(Of Double)

        Dim returnValue As Double
        If Double.TryParse(value.ToString(), returnValue) Then
            Return returnValue
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to a single-precision floating point number, or null.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToSingle(ByVal value As Object) As System.Nullable(Of Single)

        Dim returnValue As Single
        If Single.TryParse(value.ToString(), returnValue) Then
            Return returnValue
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Converts a specified value to a GUID, or null.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ewToGUID(ByVal value As Object) As System.Nullable(Of Guid)

        If value IsNot Nothing Then
            If Validators.IsGUID(value.ToString()) Then
                Dim returnValue As New Guid(value.ToString())
                Return returnValue
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

#End Region

End Class