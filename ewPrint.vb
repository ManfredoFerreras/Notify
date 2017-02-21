Public Class ewPrint

    ''' <summary>
    ''' Writes a single line message
    ''' </summary>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    Public Shared Sub WriteLine(ByVal Message As String)
        WriteLineEx(Message, False)
    End Sub

    ''' <summary>
    ''' Writes a doble line message
    ''' </summary>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    Public Shared Sub WriteDobleLine(ByVal Message As String)
        WriteLineEx(Message, True)
    End Sub

    ''' <summary>
    ''' Write a message to the Immediate Window and to the Console
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <param name="AddSpace"></param>
    ''' <remarks></remarks>
    Private Shared Sub WriteLineEx(ByVal Message As String, ByVal AddSpace As Boolean)

        ' Write message to the Immediate Window
        Debug.WriteLine(Message)

        ' Write message to the Console
        Console.WriteLine(Message)

        ' Add line spacing
        If AddSpace = True Then
            Debug.WriteLine("")
            Console.WriteLine("")
        End If

    End Sub

End Class
