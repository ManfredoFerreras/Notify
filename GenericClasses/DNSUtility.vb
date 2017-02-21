Imports System
Imports System.Net

Public Class DNSUtility

    Public Shared Function GetIPAddress(ByVal HostName As String) As String
        Dim wrkstr As String = String.Empty
        Dim sReturn As String = String.Empty

        ' Getting Ip address of local machine
        ' First get the host name of local machine
        If String.IsNullOrEmpty(HostName.Trim) Then
            wrkstr = Dns.GetHostName
        Else
            wrkstr = HostName.Trim
        End If

        ' The using host name, get the IP address list
        Dim hostEntry As IPHostEntry
        Dim addrsList() As IPAddress

        Try
            hostEntry = Dns.GetHostEntry(wrkstr)
            addrsList = hostEntry.AddressList
            For i As Integer = 0 To addrsList.GetUpperBound(0)
                sReturn += addrsList(i).ToString
            Next

        Catch ex As Exception
            ' Do nothing
        End Try

        Return sReturn
    End Function

End Class
