Imports System
Imports System.IO
Imports System.Text
Imports System.Configuration

Public Class DirectBillImportForm
    Dim GeoveraPath As String = ConfigurationSettings.AppSettings("GeoveraPath")
    Dim sVal As String = ConfigurationSettings.AppSettings("Carriers")
    Dim sKillExcel As String = ConfigurationSettings.AppSettings("Excel")
    Dim FileSaveLoc As String = ConfigurationSettings.AppSettings("FileSave")
    Private sBatchId As String = ""
    Dim gArchive As New clsMove
    Dim sPeriod As String = ""
    Dim iCount As Integer = 0


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim tmparray() As String = sVal.Split(",")
            Dim tmpstr As String
            For Each tmpstr In tmparray
                cbCarrier.Items.Add(tmpstr)
            Next
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.OkOnly, "Error")
            Close()
        End Try
    End Sub

    Private Sub cbCarrier_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCarrier.SelectedIndexChanged
        cbFileSelect.Items.Clear()
        Dim CarrierHistory As New clsDataCommon
        CarrierHistory.DirectBillHistoryInsert(cbCarrier.Text)
        If cbCarrier.Text = "Geovera" Then btnMove.Enabled = True Else btnMove.Enabled = False
        If cbCarrier.Text = "Geovera" Then btnMakeWhole.Enabled = True Else btnMakeWhole.Enabled = False
        If cbCarrier.Text = "Geovera" Then btnGArchive.Enabled = True Else btnGArchive.Enabled = False
        If cbCarrier.Text = "Geovera" Then cbFileSelect.Enabled = True
        If cbCarrier.Text = "Geovera" Then
            Dim info As New IO.DirectoryInfo(GeoveraPath)
            For Each File In info.GetFiles
                Dim sName As String = File.Name
                cbFileSelect.Items.Add(sName)
            Next
            sPeriod = InputBox("yyyymm", "Date", "")
            iCount = InputBox("Number of business days:", "Days")
        End If
    End Sub

    Private Sub btnMakeWhole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMakeWhole.Click
        Dim sType As String = cbFileSelect.SelectedItem.ToString
        sType = Replace(sType.Substring(10), ".xls", "")
        modMakeWhole.GeoveraMakeWholeExcel()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        If cbCarrier.Text = "Geovera" Then
            Dim GeoveraStaging As New Geovera
            sBatchId = GeoveraStaging.ProcessGeovera(cbFileSelect.Text)
            gArchive.ArchiveFiles()
        End If
        If cbCarrier.Text = "Hartford" Then
            Dim HartfordStaging As New Hartford
            sBatchId = HartfordStaging.ProcessHartford()
        End If
        If cbCarrier.Text = "Occidental" Then
            Dim OccidentalProcess As New Occidental
            sBatchId = OccidentalProcess.BeginProcessing()
        End If
        If cbCarrier.Text = "Safeway" Then
            Dim SafeayStaging As New Safeway
            sBatchId = SafeayStaging.ProcessSafeway
        End If
        If cbCarrier.Text = "Travelers" Then
            Dim TravelersProcessing As New Travelers
            sBatchId = TravelersProcessing.ProcessTrav()
        End If
        If cbCarrier.Text = "Voyager" Then
            Dim VoyagerProcessing As New Voager
            sBatchId = VoyagerProcessing.ProcessVoyager()
        End If
        If cbCarrier.Text = "Capitol" Then
            Dim CapitolStaging As New Capitol
            sBatchId = CapitolStaging.ProcessCapitol()
        End If
        If cbCarrier.Text = "NICO" Then
            Dim NICOStaging As New NICO
            sBatchId = NICOStaging.ProcessNICO()
        End If
        If cbCarrier.Text = "AMTrust" Then
            Dim AMTrustStaging As New AMtrust
            sBatchId = AMTrustStaging.ProcessAMTrust()
        End If
        If cbCarrier.Text = "SafewayReconciliation" Then
            Dim SafewayReconciliationStaging As New SafewayReconciliation
            sBatchId = SafewayReconciliationStaging.ProcessSafewayReconciliation()
        End If
       
        Dim GetNbrs As New clsDataCommon
        Dim sqlReader As SqlClient.SqlDataReader = Nothing
        sqlReader = GetNbrs.GetNumbers(sBatchId)
        tbBatId.Text = sBatchId
        While sqlReader.Read
            If Not IsDBNull(sqlReader.Item("Premium")) Then
                tbPremium.Text = sqlReader.Item("Premium")
            End If

            If Not IsDBNull(sqlReader.Item("Total Agent Commission")) Then
                tbAComm.Text = sqlReader.Item("Total Agent Commission")
            End If

            If Not IsDBNull(sqlReader.Item("Total Gross Commission")) Then
                tbGComm.Text = sqlReader.Item("Total Gross Commission")
            End If

            If Not IsDBNull(sqlReader.Item("Total Payable")) Then
                tbTPay.Text = sqlReader.Item("Total Payable")
            End If

            If Not IsDBNull(sqlReader.Item("Amount")) Then
                tbAmount.Text = sqlReader.Item("Amount")
            End If

            If Not IsDBNull(sqlReader.Item("Revenue Amount")) Then
                tbRevAmt.Text = sqlReader.Item("Revenue Amount")
            End If

            If Not IsDBNull(sqlReader.Item("AP Amount")) Then
                tbAPAmt.Text = sqlReader.Item("AP Amount")
            End If
        End While
        GetNbrs.ConnClose()
        sqlReader.Close()
        If String.IsNullOrWhiteSpace(sBatchId) Then
            tbPolicyCount.Text = "0"
        Else
            tbPolicyCount.Text = GetNbrs.GetPolicyCount(sBatchId)
        End If
    End Sub

    Private Sub cbFileSelect_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFileSelect.SelectedIndexChanged
        Dim i As Integer = 0
        If cbFileSelect.Text = "WholeThingNormal.xls" Then
            Dim iInfo As New IO.DirectoryInfo(ConfigurationSettings.AppSettings("GeoveraUSFGPath") & "Normal\" + sPeriod + "\")
            For Each File In iInfo.GetFiles
                i = i + 1
            Next
        Else
            Dim iInfo As New IO.DirectoryInfo(ConfigurationSettings.AppSettings("GeoveraUSFGPath") & "FFB\" + sPeriod + "\")
            For Each File In iInfo.GetFiles
                i = i + 1
            Next
        End If
        If iCount > i Then
            MsgBox("To many files. Please have I.T. remove the " & iCount - i & " extra file(s).", MsgBoxStyle.OkOnly, "Warning!")
        ElseIf iCount < i Then
            MsgBox("Missing " & i - iCount & " file(s) from Geovera. Please contact I.T. to get the remaining files.", MsgBoxStyle.OkOnly, "Warning!")
        ElseIf iCount = i Then
            btnMakeWhole.Enabled = True
        End If
    End Sub

    Private Sub btnMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMove.Click
        Dim NewMove As New clsMove
        NewMove.MoveFiles()
        cbFileSelect.Enabled = True
        Dim info As New IO.DirectoryInfo(GeoveraPath)
        For Each File In info.GetFiles
            Dim sName As String = File.Name
            cbFileSelect.Items.Add(sName)
        Next
        MsgBox("Done moving files")
    End Sub

    'Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGArchive.Click
    '    Dim mydocpath As String = FileSaveLoc
    '    Dim sb As New StringBuilder()
    '    mydocpath = mydocpath & "\" & Date.Now.ToString("yyyy") & "\"
    '    If Not Directory.Exists(mydocpath) Then Directory.CreateDirectory(mydocpath)
    '    mydocpath = mydocpath & "\" & Date.Now.ToString("MM") & "\"
    '    If Not Directory.Exists(mydocpath) Then Directory.CreateDirectory(mydocpath)
    '    mydocpath = mydocpath & "\" & cbCarrier.SelectedItem
    '    If Not Directory.Exists(mydocpath) Then Directory.CreateDirectory(mydocpath)
    '    Dim outfile As New StreamWriter(mydocpath & "\" & Date.Now.ToString("MMddyyyy") & ".txt")
    '    sb.AppendLine("Carrier: " & cbCarrier.SelectedItem.ToString & vbTab & vbTab & "Batch Number: " & tbBatId.Text)
    '    sb.AppendLine("Import Count: " & tbPolicyCount.Text)
    '    sb.AppendLine("Premium: " & tbPremium.Text & vbTab & vbTab & "Total Amount: " & tbAmount.Text)
    '    sb.AppendLine("Total Agent Commission:  " & vbTab & tbAComm.Text)
    '    sb.AppendLine("Total Gross Commission:      -" & vbTab & tbGComm.Text)
    '    sb.AppendLine(vbTab & vbTab & vbTab & vbTab & "________________")
    '    sb.AppendLine("Total Payable:" & vbTab & vbTab & vbTab & tbTPay.Text)
    '    sb.AppendLine("Revenue Amount:" & vbTab & vbTab & tbRevAmt.Text)
    '    sb.AppendLine("AP Amount:" & vbTab & vbTab & tbAPAmt.Text)
    '    Using outfile
    '        outfile.Write(sb.ToString())
    '    End Using
    'End Sub

End Class