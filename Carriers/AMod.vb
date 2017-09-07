Imports System.Data
Imports System.Text
Imports System.Net.Mail
Imports System.Configuration


Public Class AMod
    Dim sBatchId As String = ""
    Dim sMessage As String = ""
    Dim sResult As String = ""
    Dim sFileName As String = ConfigurationSettings.AppSettings("FileName")
    Dim sCon As String = ConfigurationSettings.AppSettings("cis")
    Dim conn As New SqlClient.SqlConnection(sCon)

    Dim AMODData As New DirectBillImports.clsAmodData
    Dim MailResults As New DirectBillImports.clsMail
    Dim errorWriter As New DirectBillImports.clsTextWriterTraceListener
    Dim InvoiceBatch As DirectBillImports.Aim
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Sub Main()

    End Sub
    Public Function ProcessAMod() As String
        With errorWriter
            .CreateErrorWriter("AMod")
            .AddListerners()
        End With
        If MsgBoxResult.Ok = True Then
            Try
                Debug.WriteLine("Begin AMOD import")
                With AMODData
                    sResult = .ImportFileToAMODTable()
                    If Not sResult.Contains("Error: ") Then
                        sBatchId = .StageAMOD
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sBatchId.Contains("Error: ") Then
                        sMessage = InvoiceBatch.InvoiceBatch(sBatchId, sFileName, conn, "AMod")
                    Else
                        Return Nothing
                        Exit Function
                    End If
                    If Not sMessage.Contains("Error") Then
                        MailResults.EmailResults("AMOD Direct Deposit", "Batch Number: " & sBatchId & vbCrLf & sMessage, sTo)
                    Else
                        MailResults.EmailResults("Error From CIS Aim Invoicing", sMessage, sErrorTo)
                    End If
                End With
                Debug.WriteLine("Finished import for AMOD direct bill.")
                errorWriter.CloseWriter()
            Catch ex As Exception
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                MailResults.EmailResults("AMOD Direct bill import error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
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