Imports System.Configuration

Public Class Geovera
    Dim sName As String = ""
    Dim sPath As String = ConfigurationSettings.AppSettings("sPath")
    Dim sBatchId As String = ""
    Dim sMessage As String = ""
    Dim sResults As String = ""
    Dim sFileName As String = ConfigurationSettings.AppSettings("FileName") + "USFG\"
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Dim conn As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("CIS"))

    Dim GeoveraData As New DirectBillImports.clsGeoveraData
    Dim MailResults As New DirectBillImports.clsMail
    Dim errorWriter As New DirectBillImports.clsTextWriterTraceListener

    Public Function ProcessGeovera(ByVal FileSelect As String)
        Dim sFilePath As String = sPath
        With errorWriter
            .CreateErrorWriter("Geovera")
            .AddListerners()
        End With
        If MsgBoxResult.Yes Then
            Try
                With GeoveraData
                    Debug.WriteLine("Clearing old data from temp tables")
                    .ClearUSFG()
                    Debug.WriteLine("Importing current data to temp tables")
                    If FileSelect = "WholeThingNormal.xls" Then
                        .ImportFileToGeovera()
                    Else
                        .ImportFileToGeoveraFFB()
                    End If
                    Debug.WriteLine("Staging data.")
                    If Not sResults.Contains("Error") Then
                        If FileSelect = "WholeThingNormal.xls" Then
                            sBatchId = .StageGeoveraNormal()
                        ElseIf FileSelect = "WholeThingFFB.xls" Then
                            sBatchId = .StageGeoveraFFB()
                        Else
                            Return "Error with " & sBatchId
                            Exit Function
                        End If
                    Else
                        Return sResults
                        Exit Function
                    End If
                    Debug.WriteLine("Begining invoice of USFG")
                    Dim InvoiceBatch As New DirectBillImports.Aim
                    If Not sResults.Contains("Error") Then
                        sMessage = InvoiceBatch.InvoiceBatch(sBatchId, sFileName, conn, "Geovera")
                    Else
                        Debug.WriteLine(sResults)
                        MailResults.EmailResults("Error - Geovera Direct Bill Imports", sResults, sErrorTo)
                        Return "Error, contact I.T."
                        Exit Function
                    End If
                    If Not sMessage.Contains("Error") Then
                        MailResults.EmailResults("Geovera Direct Bill", "Batch Number: " & sBatchId & vbCrLf & sMessage, sTo)
                    Else
                        MsgBox("Error with import, contact IT", MsgBoxStyle.OkOnly, "Error")
                        MailResults.EmailResults("Error From CIS Aim Invoicing", sMessage, sErrorTo)
                    End If
                End With
            Catch ex As Exception
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            End Try
        Else
            Debug.WriteLine("Closed by user")
            conn.Close()
            conn.Dispose()
            errorWriter.CloseWriter()
        End If
        conn.Close()
        conn.Dispose()
        errorWriter.CloseWriter()
        Return sBatchId
    End Function

End Class