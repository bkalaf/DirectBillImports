Imports System
Imports System.IO
Imports System.Data.OleDb
Imports System.Configuration
Imports DirectBillImports.clsMail

Module modMakeWhole
    Dim sTo As String = "aholbrook@siuins.com, lcox@siuins.com"

    Private Function MakePeriod(ByVal sToFolder As String) As String
        Dim sAnswer As String = ""
        Dim sFirstDay As Date = CDate(Mid(sToFolder, 5, 2) & "/01/" & Mid(sToFolder, 1, 4))
        Dim sLastMonthFirstDay As Date = DateAdd(DateInterval.Month, -1, sFirstDay)
        sAnswer = IIf(Month(sLastMonthFirstDay) < 10, "0", "") & CStr(Month(sLastMonthFirstDay)) & Mid(CStr(Year(sLastMonthFirstDay)), 3, 2)
        Return sAnswer
    End Function

    Public Sub GeoveraMakeWholeExcel()
        Dim EachFile As StreamReader

        Dim myfiles As String()
        Dim FileName As String
        Dim sPeriod As String = InputBox("Enter Period")
        Dim sType As String = InputBox("Normal or FFB?")
        Dim sPath As String = ConfigurationSettings.AppSettings("GeoveraUSFGPath") & sType & "\"
        Dim sFileName As String = sPath & "\WholeThing" & sPeriod & sType & ".csv"
        If File.Exists(sFileName) Then File.Delete(sFileName)

        Dim WholeThing As StreamWriter = File.CreateText(sFileName)

        myfiles = Directory.GetFiles(ConfigurationSettings.AppSettings("GeoveraUSFGPath") & sType & "\" & sPeriod & "\")
        WholeThing.WriteLine("DatePosted,MarketingRep,AgentID,PolicyNbr,InsuredName,TTCode,TTExtended,EffDate,ExpDate,WrittenPremium,NetPremium,StateTax,StampFee,CPICFee,UW_Fee,PolicyFee,FHCFFEE,SrvcCharge,Refund,WriteOff,AgentRetained,CashEntered,PendingRefund,BeginningBalance,EndingBalance,Address,City,State,Zip5,Zip4,,Coverage,Cov A,AP Ded,Wind Ded,,InspectionFee")

        Dim MyCounter As Integer
        For Each FileName In myfiles
            MyCounter += 1
            Debug.WriteLine(FileName)
            EachFile = File.OpenText(FileName)
            Dim inLine As String
            Dim Counter As Integer = 0
            Do While EachFile.Peek <> -1
                Counter = Counter + 1
                inLine = EachFile.ReadLine
                If Counter > 1 Then
                    WholeThing.WriteLine(inLine)
                End If
            Loop
            EachFile.Close()
        Next
        WholeThing.Close()
        MsgBox(MyCounter)
        Try
            Dim mailFile As New clsMail
            With mailFile
                Dim sBody As String = "File attached."
                .EmailFiles("Geovera " & sType, sBody, sTo, sFileName)
            End With
        Catch ex As Exception
            Debug.WriteLine(ex)
        End Try
    End Sub
End Module
