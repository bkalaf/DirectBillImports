Imports System.IO
Imports System.Configuration

Public Class Safeway
    Dim sBatchId As String = ""
    Dim sMessage As String = ""
    Dim sResult As String = ""
    Dim sFileName As String = ConfigurationSettings.AppSettings("FileName")
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Dim clsSafewayData As New clsSafewayData
    Dim MailResults As New clsMail
    Dim ErrorWriter As New DirectBillImports.clsTextWriterTraceListener

    Public Function ProcessSafeway()
        With ErrorWriter
            .CreateErrorWriter("Safeway")
            .AddListerners()
        End With
        If MsgBoxResult.Ok Then
            Try
                Debug.WriteLine("Beginning Safeway import process")
                sBatchId = BeginProcessing()
                Debug.WriteLine("Ending Safeway import process")
                MsgBox(sMessage & "Finished import, batch number is: " & sBatchId)
            Catch ex As Exception
                MsgBox("Error: " & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                MailResults.EmailResults("Safeway Import Error Batch Id: " & sBatchId, ex.Message + ex.StackTrace, sErrorTo)
                Return Nothing
            Finally
                ErrorWriter.CloseWriter()
                MsgBox("Finished Import Process", MsgBoxStyle.OkOnly, "Finished")
            End Try
        Else
            ErrorWriter.CloseWriter()
            Return Nothing
            Exit Function
        End If
        Return sBatchId
    End Function

    Public Function BeginProcessing() As String
        Debug.WriteLine("Running DTS package")
        clsSafewayData.ClearDataFromSFWYTable()
        sResult = ImportFileToAutoInvoice()  'Saves to SFWY table in AutoInvoicing on El-Cid
        If Not sResult.Contains("Error:") Then
            Debug.WriteLine("Importing Safeway file into staging tables on SQL2008R2\SIU in CIS")
            sBatchId = StageBatch() 'Imports data from AutoInvoicing SFWY to staging tables in CIS database
        Else
            Return sResult
        End If
        If Not sResult.Contains("Error:") Then
            Debug.WriteLine("Invoicing Batch into CIS invoiceheader and invoicedetail tables.")
            sResult = InvoiceBatch() 'Imports data from AutoInvoicing SFWY to staging tables in CIS database
        Else
            Return sResult
            Exit Function
        End If
        If Not sResult.Contains("Error") Then
            Debug.WriteLine(sBatchId)
            MailResults.EmailResults("Safeway Directbill Import", sBatchId, sTo)
        Else
            Debug.WriteLine(sResult)
        End If
        Debug.WriteLine("Invoicing Safeway records into AIM")
        Return sBatchId
    End Function

    Public Function ImportFileToAutoInvoice() As String
        Try
            Dim clsImport As New clsSafewayData
            clsImport.ImportFileToSFWYTable()
            Debug.WriteLine("Importing file.")
            Return "ImportFile Done"
        Catch ex As Exception
            MailResults.EmailResults("Safeway Import Error: ", ex.Message + ex.StackTrace, sErrorTo)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function StageBatch() As String
        Try
            Dim clsStage As New clsSafewayData
            StageBatch = clsStage.ImportToStaging()
            Debug.WriteLine("Saving to staging tables has been completed.")
            Return StageBatch
        Catch ex As Exception
            MailResults.EmailResults("Safeway Import Error: ", ex.Message + ex.StackTrace, sErrorTo)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function InvoiceBatch() As String
        Dim invBatch As New DirectBillImports.Aim
        Dim conn As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("CIS"))
        Try
            If sResult <> "" Then
                If File.Exists(sFileName) Then File.Delete(sFileName)
                sMessage = invBatch.InvoiceBatch(sBatchId, sFileName, conn, "Safeway")
                If Len(Trim(sMessage)) > 0 Then
                    MsgBox(sMessage)
                End If
                conn.Close()
            End If
            Return sMessage
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message + ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function
End Class