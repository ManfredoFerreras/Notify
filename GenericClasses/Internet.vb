Imports System
Imports System.Net
Imports System.Text

Public Class Internet

    ''' <summary>
    ''' Check for an Internet connection
    ''' </summary>
    ''' <param name="BaseURL">Base URL Address</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckConnection(ByVal BaseURL As String) As Boolean
        Dim ByteResponse As [Byte]()

        ' Instantiate a web client
        Dim client As WebClient = New WebClient

        Try
            ' Get data from server as byte()
            ByteResponse = client.DownloadData(BaseURL)

            ' Read the response
            Dim response As String = Encoding.ASCII.GetString(ByteResponse).ToString

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

End Class
