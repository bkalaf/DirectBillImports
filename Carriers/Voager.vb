Imports System.Configuration

Public Class Voager
    Dim sBatchId As String = ""
    Dim sMessage As String = ""
    Dim sResult As String = ""
    Dim conn As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("CIS"))
    Dim sFileName As String = ConfigurationSettings.AppSettings("FileName")
    Dim errorWriter As New DirectBillImports.clsTextWriterTraceListener
    Dim InvoiceBatch As New DirectBillImports.Aim
    Dim VoyagerData As New clsVoyagerData
    Dim mailresults As New DirectBillImports.clsMail
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Public Function ProcessVoyager() As String
        With errorWriter
            .CreateErrorWriter("Voyager")
            .AddListerners()
        End With
        If MsgBoxResult.Ok Then
            Try
                Debug.WriteLine("Begin Voyager import")
                With VoyagerData
                    sResult = .ImportFileToVoyagerTable()
                    If Not sResult.Contains("Error: ") Then
                        .PreProcess()
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sResult.Contains("Error: ") Then
                        sBatchId = .StageVoyager
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sBatchId.Contains("Error: ") Then
                        sMessage = InvoiceBatch.InvoiceBatch(sBatchId, sFileName, conn, "Voyager")
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sMessage.Contains("Error") Then
                        mailresults.EmailResults("Voyager Direct Deposit", "Batch Number: " & sBatchId, sTo)
                    Else
                        mailresults.EmailResults("Error From CIS Aim Invoicing", sMessage, sErrorTo)
                        sBatchId = sMessage
                    End If
                End With
                Debug.WriteLine("Finished import for Voyager direct bill.")
                errorWriter.CloseWriter()
            Catch ex As Exception
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                mailresults.EmailResults("Voyager Direct bill import error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            End Try
        Else
            Debug.WriteLine("Closed by user.")
            conn.Close()
            conn.Dispose()
            errorWriter.CloseWriter()
            Return Nothing
            Exit Function
        End If
        Return sBatchId
    End Function

End Class