Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class ewEmailQueue

    ''' <summary>
    ''' Add Email Message To Queue
    ''' </summary>
    ''' <param name="Format"></param>
    ''' <param name="EmailFrom"></param>
    ''' <param name="EmailTo"></param>
    ''' <param name="EmailSubject"></param>
    ''' <param name="EmailBody"></param>
    ''' <param name="EPS"></param>
    ''' <param name="ApplicationID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddToQueue(ByVal Format As String, _
                                      ByVal EmailFrom As String, _
                                      ByVal EmailTo As String, _
                                      ByVal EmailSubject As String, _
                                      ByVal EmailBody As String, _
                                      ByVal EPS As String, _
                                      ByVal ApplicationID As Integer) As Integer

        Dim nReturn As Integer = 0

        Dim oConn As New SqlConnection(db.GetConnectionString())
        Dim oCmd As New SqlCommand

        With oCmd
            .Connection = oConn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "EmailQueue.dbo.proc_tbl_EmailQueueSend"

            '@EmailQueueID int
            db.AddParameter(oCmd, "@EmailQueueID", db.DataType.eInteger, db.Direction.eOutput, 0)

            '@Format nchar(4)
            db.AddParameter(oCmd, "@Format", Format, 4)

            '@EmailFrom nvarchar(255)
            db.AddParameter(oCmd, "@EmailFrom", EmailFrom, 255)

            '@EmailTo nvarchar(255)
            db.AddParameter(oCmd, "@EmailTo", EmailTo, 255)

            '@EmailSubject nvarchar(255)
            db.AddParameter(oCmd, "@EmailSubject", EmailSubject, 255)

            '@EmailBody ntext
            db.AddParameter(oCmd, "@EmailBody", EmailBody, 0)

            '@EPS char(12)
            db.AddParameter(oCmd, "@EPS", EPS, 12)

            '@ApplicationID int
            db.AddParameter(oCmd, "@ApplicationID", db.DataType.eInteger, db.Direction.eInput, ApplicationID)

        End With

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()

            ' To get a procedure's return value
            nReturn = CType(oCmd.Parameters("@EmailQueueID").Value, Integer)
        Catch ex As Exception
            nReturn = -1
        Finally
            oConn.Close()
        End Try

        Return nReturn

    End Function

End Class
