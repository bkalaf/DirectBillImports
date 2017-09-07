Imports System.Configuration
Class Hartford
    Dim sBatchId As String = ""
    Dim sMessage As String = ""
    Dim sResult As String = ""
    Dim sFileName As String = ConfigurationSettings.AppSettings("FileName")
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")
    Dim conn As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("CIS"))

    Dim HartfordData As New DirectBillImports.clsHartfordData
    Dim MailResults As New DirectBillImports.clsMail
    Dim errorWriter As New DirectBillImports.clsTextWriterTraceListener
    Dim InvoiceBatch As New DirectBillImports.Aim

    Public Function ProcessHartford() As String
        With errorWriter
            .CreateErrorWriter("Hartford")
            .AddListerners()
        End With
        If MsgBoxResult.Ok Then
            Try
                With HartfordData

                    Debug.WriteLine("Begin Hartford Excel import")
                    .ClearDataFromHartfordTable()
                    sResult = .ImportFileToHartfordTable()
                    If Not sResult.Contains("Error:") Then
                        Debug.WriteLine("Importing Hartford Rows into staging table")
                        sBatchId = .StageHartford
                    Else
                        Return sResult
                    End If

                    If Not sBatchId.Contains("Error: ") Then
                        sMessage = InvoiceBatch.InvoiceBatch(sBatchId, sFileName, conn, "Hartford")
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sMessage.Contains("Error") Then
                        'MailResults.EmailResults("Hartford Direct Deposit", "Batch Number: " & sBatchId, sTo) 'add later
                    Else
                        'MailResults.EmailResults("Error From CIS Aim Invoicing", sMessage, sErrorTo) 'add later
                        sBatchId = sMessage
                    End If
                End With
                Debug.WriteLine("Finished import for Hartford direct bill.")
                errorWriter.CloseWriter()
            Catch ex As Exception
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                'MailResults.EmailResults("Hartford Direct bill import error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo) 'add later
            End Try
        Else
            Debug.WriteLine("Closed by user.")
            conn.Close()
            conn.Dispose()
            errorWriter.CloseWriter()
            Return sBatchId
            Exit Function
        End If
        Return sBatchId
    End Function

End Class