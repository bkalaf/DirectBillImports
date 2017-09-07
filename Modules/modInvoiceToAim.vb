Imports System
Imports System.IO
Imports System.Data.SqlClient

Module modInvoiceToAim
    Public Sub StartInvoice(ByVal BatNbr As String, ByVal sFileName As String, ByVal conn As SqlConnection)
        Dim myAIM As New DirectBillImports.Aim
        If File.Exists(sFileName) Then File.Delete(sFileName)
        conn.Open()
        Dim sMessage As String = myAIM.InvoiceBatch(BatNbr, sFileName, conn, "modInvoiceToAim")
        If Len(Trim(sMessage)) > 0 Then
            MsgBox(sMessage)
        End If
        conn.Close()
        MsgBox("done")
        BatNbr = ""
    End Sub
End Module