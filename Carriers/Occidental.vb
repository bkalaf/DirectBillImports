Imports System.IO
Imports System.Configuration

Public Class Occidental
    Dim sBatchId As String = ""
    Dim sMessage As String = ""
    Dim sResult As String = ""
    Dim sFileName As String = ConfigurationSettings.AppSettings("FileName")
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Dim clsOccidentalData As New ClsOccidentalData
    Dim MailResults As New clsMail
    Dim ErrorWriter As New DirectBillImports.clsTextWriterTraceListener

    Public Sub ProcessOccidental()
        With ErrorWriter
            .CreateErrorWriter("Occidental")
            .AddListerners()
        End With
        If MsgBoxResult.Ok Then
            Try
                Debug.WriteLine("Beginning Occidental import process")
                sBatchId = BeginProcessing()
                Debug.WriteLine("Ending Occidental import process")
                MsgBox(sMessage & "Finished import, batch number is: " & sBatchId)
            Catch ex As Exception
                MsgBox("Error: " & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                MailResults.EmailResults("Occidental Import Error: ", ex.Message + ex.StackTrace, sErrorTo)
            Finally
                ErrorWriter.CloseWriter()
                MsgBox("Finished Import Process", MsgBoxStyle.OkOnly, "Finished")
            End Try
        Else
            ErrorWriter.CloseWriter()
            Exit Sub
        End If
    End Sub

    Public Function BeginProcessing() As String
        Debug.WriteLine("Running DTS package")
        clsOccidentalData.ClearDataFromOccTable()
        sResult = ImportFileToAutoInvoice()  'Saves to Occ table in AutoInvoicing on El-Cid
        If Not sResult.Contains("Error:") Then
            Debug.WriteLine("Importing Occidental file into staging tables on SQL2008R2\SIU in CIS")
            sBatchId = StageBatch() 'Imports data from AutoInvoicing Occ to staging tables in CIS database
        Else
            Return sResult
        End If

        If Not sResult.Contains("Error:") Then
            Debug.WriteLine("Invoicing Batch into CIS invoiceheader and invoicedetail tables.")
            sResult = InvoiceBatch() 'Imports data from AutoInvoicing Occ to staging tables in CIS database
        Else
            Return sResult
            Exit Function
        End If

        If Not sResult.Contains("Error") Then
            Debug.WriteLine(sBatchId)
            MailResults.EmailResults("Occidental Directbill Import", sBatchId, sTo)
        Else
            Debug.WriteLine(sResult)
            sBatchId = sResult
        End If

        Return sBatchId

    End Function

    Public Function ImportFileToAutoInvoice() As String
        Try
            Debug.WriteLine("Importing file.")
            Dim clsImport As New ClsOccidentalData
            clsImport.ImportFileToOccPremiumTable()
            clsImport.ImportFileToOccRCPTable()
            clsImport.PreProcess()
            Return "ImportFile Done"
        Catch ex As Exception
            MailResults.EmailResults("Occidental Import Error: ", ex.Message + ex.StackTrace, sErrorTo)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function StageBatch() As String
        Try
            Dim clsStage As New ClsOccidentalData
            Return clsStage.ImportToStaging()
            Debug.WriteLine("Save to staging tables complete.")
        Catch ex As Exception
            MailResults.EmailResults("Occidental Import Error: ", ex.Message + ex.StackTrace, sErrorTo)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function InvoiceBatch() As String
        Dim invBatch As New DirectBillImports.Aim
        Dim conn As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("CIS"))
        Try
            If sResult <> "" Then
                If File.Exists(sFileName) Then File.Delete(sFileName)
                sMessage = invBatch.InvoiceBatch(sBatchId, sFileName, conn, "Occidental")
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