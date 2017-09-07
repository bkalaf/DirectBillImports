Imports System.Configuration

Public Class Travelers
    Dim sBatchId As String = ""
    Dim sMessage As String = ""
    Dim sResult As String = ""
    Dim sFileName As String = ConfigurationSettings.AppSettings("FileName") '"\\spartacus\sys1\InvoicingStatus.txt"
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")
    Dim conn As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("CIS"))

    Dim TravelersData As New DirectBillImports.clsTravelersData
    Dim MailResults As New DirectBillImports.clsMail
    Dim errorWriter As New DirectBillImports.clsTextWriterTraceListener
    Dim InvoiceBatch As New DirectBillImports.Aim

    Function ProcessTrav() As String
        With errorWriter
            .CreateErrorWriter("Travelers")
            .AddListerners()
        End With
        If MsgBoxResult.Ok Then
            Try
                Debug.WriteLine("Begin Travelers import")
                With TravelersData
                    .ClearDataFromTravelersTable()
                    sResult = .ImportFileToTravelersTable()
                    If Not sResult.Contains("Error: ") Then
                        sBatchId = .StageTravelers
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sBatchId.Contains("Error: ") Then
                        sMessage = InvoiceBatch.InvoiceBatch(sBatchId, sFileName, conn, "Travelers")
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sMessage.Contains("Error") Then
                        MailResults.EmailResults("Travelers Direct Deposit", "Batch Number: " & sBatchId, sTo)
                    Else
                        MailResults.EmailResults("Error From CIS Aim Invoicing", sMessage, sErrorTo)
                        sBatchId = sMessage
                    End If
                End With
                Debug.WriteLine("Finished import for Travelers direct bill.")
                errorWriter.CloseWriter()
            Catch ex As Exception
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                MailResults.EmailResults("Travelers Direct bill import error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
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
