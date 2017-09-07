Imports System.Data.SqlClient
Imports System.IO
Imports DirectBillImports.clsDataCommon

Public Class Aim

#Region "Public Methods"

    Public Function InvoiceItem(ByRef pHeader As clsHeader, ByVal pBatNbr As String, ByRef conn As SqlConnection, ByVal carrierName As String) As String
        Try
            Dim sMessage As String = ""
            sMessage = pHeader.Save(conn)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If

            'create new submission if dates does not exist for hartford and capitol (getHistoryForEffectiveDate should be used only for Hartford, Capitol and other Sun direct bill etc)
            If carrierName = "Hartford" Or carrierName = "Capitol" Or carrierName = "AMTrust" Then
                sMessage = GetHistoryForEffectiveDate(conn, pHeader)
            ElseIf carrierName = "NICO" Then
                sMessage = GetHistoryFor922(conn, pHeader)
            Else
                sMessage = GetHistory(conn, pHeader)
            End If
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If

            'no tax for hartford and capitol
            If carrierName = "Hartford" Or carrierName = "Capitol" Or carrierName = "AMTrust" Then
                pHeader.Version.Taxed = "N"

                If Not pHeader.SubmissionExists Then
                    sMessage = CreateSubmissionForHartfordCapitol(conn, pHeader)
                    If Len(Trim(sMessage)) > 0 Then
                        Return sMessage
                    End If
                End If
            ElseIf carrierName = "NICO" Then
                pHeader.Version.Taxed = "N"
            Else
                pHeader.Version.Taxed = "Y"

                If Not pHeader.SubmissionExists Then
                    sMessage = CreateSubmission(conn, pHeader)
                    If Len(Trim(sMessage)) > 0 Then
                        Return sMessage
                    End If
                End If
            End If

            sMessage = ImportInvoiceHeader(conn, pHeader, pBatNbr, carrierName)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
MyPause:
            sMessage = ImportInvoiceDetails(conn, pHeader, "")
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If

            'If carrierName <> "NICO" Then

            'End If

            sMessage = SetSubmission(conn, pHeader)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
           
            If pHeader.TeamID = "901" Then
                sMessage = FixBad901DetailCommission(conn, pHeader.InvoiceKey_PK)
                If Len(Trim(sMessage)) > 0 Then
                    Return sMessage
                End If
            End If
            sMessage = Me.FinalCleanup(conn, pHeader.InvoiceKey_PK)

            'If carrierName = "Hartford" Or carrierName = "Capitol" Or carrierName = "AMTrust" Then
            '    sMessage = Me.FinalCleanupForHartfordCapitol(conn, carrierName, pHeader.PolicyNumber, pHeader.InvoiceKey_PK)
            'End If

            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function InvoiceItem(ByRef pHeader As clsHeader, ByVal pBatNbr As String) As String
        Try
            Dim sMessage As String = ""
            'TODO : it is to be changed for production from T to P
            Dim conn As New SqlConnection(Common.AIMConnectionString("T"))
            conn.Open()
            sMessage = pHeader.Save(conn)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            sMessage = GetHistory(conn, pHeader)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            If Not pHeader.SubmissionExists Then
                sMessage = CreateSubmission(conn, pHeader)
                If Len(Trim(sMessage)) > 0 Then
                    Return sMessage
                End If
            End If
            sMessage = ImportInvoiceHeader(conn, pHeader, pBatNbr, "")
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            sMessage = ImportInvoiceDetails(conn, pHeader, "")
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            sMessage = SetSubmission(conn, pHeader)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            If pHeader.TeamID = "901" Then
                sMessage = FixBad901DetailCommission(conn, pHeader.InvoiceKey_PK)
                If Len(Trim(sMessage)) > 0 Then
                    Return sMessage
                End If
            End If
            sMessage = Me.FinalCleanup(conn, pHeader.InvoiceKey_PK)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            conn.Close()
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function InvoiceBatch(ByVal pBatNbr As String, ByVal sfilename As String, ByRef conn As SqlConnection, ByVal carrierName As String) As String
        Try
            Dim sBatNbr As String = pBatNbr
            Dim comm As New SqlCommand("SIU_p_ListStagingBatch", conn)
            conn.Open()
            Dim rs As SqlDataReader
            Dim myH As clsHeader
            Dim sMessage As String
            'conn.Open()
            If BatchExists(conn, sBatNbr) Then
                Return "Batch Exists"
            End If
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@BatNbr", sBatNbr)
                rs = .ExecuteReader
                .Dispose()
            End With
            Dim Headers As New Collection
            Do While rs.Read
                myH = New clsHeader(rs("policynumber"), rs("trantype"), rs("effdate"), rs("expdate"), rs("agencyid"), rs("companyid"), rs("coverageid"),
                                    rs("productid"), rs("teamid"), rs("insuredname"), rs("insuredaddress1"), rs("insuredaddress2"), rs("insuredcity"),
                                    rs("insuredstate"), rs("insuredzip"), rs("HeaderID"), rs("coveragea"), rs("apdeductible"), rs("winddeductible"),
                                    rs("coverage"), rs("inceptiondate"))
                Headers.Add(myH)
            Loop
            rs.Close()
            Dim HeaderCount As New clsDataCommon
            'MsgBox(Headers.Count)
            Dim iCounter As Integer = 0
            Debug.WriteLine("Real Deal - Invoice Inserts")
            For Each myH In Headers
                iCounter += 1
                Debug.WriteLine("Insert: " & iCounter)
                Debug.WriteLine("Load Details for Header Id: " & myH.HeaderID.ToString())
                myH.LoadDetails(conn, myH.HeaderID)
                Debug.WriteLine("Invoice Item: " + myH.HeaderID.ToString())
                sMessage = InvoiceItem(myH, sBatNbr, conn, carrierName)
                Debug.WriteLine("Invoice Item completed")
                'If Len(Trim(sMessage)) > 0 Then
                '    MsgBox(sMessage)
                'End If
            Next
            'Taking it out of InvoiceItem as committ has not yet happened
            For Each myH In Headers
                If carrierName = "Hartford" Or carrierName = "Capitol" Or carrierName = "AMTrust" Then
                    sMessage = Me.FinalCleanupForHartfordCapitol(conn, carrierName, myH.PolicyNumber, myH.InvoiceKey_PK)
                End If
            Next
            conn.Close()
            Return ""
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return ex.Message
        End Try
    End Function
#End Region

#Region "Private Methods"

    Private Function BatchExists(ByRef conn As SqlConnection, ByVal sBatNbr As String) As Boolean
        Dim comm As New SqlCommand("select count(*) as myCount from invoiceheader where description = '" & sBatNbr & "'", conn)
        Dim rs As SqlDataReader
        With comm
            .CommandTimeout = 0
            rs = .ExecuteReader
            .Dispose()
        End With
        rs.Read()
        BatchExists = rs("MyCount") > 0
        rs.Close()
    End Function

    Private Function ImportInvoiceHeader(ByRef conn As SqlConnection, ByRef pHeader As clsHeader, ByVal pBatNbr As String, ByVal carrierName As String) As String
        Dim sProgress As String = ""
        Try
            If pHeader.HeaderID = 2001 Then Stop
            Dim InvoiceID As String
            Dim InvoiceKey_PK As Integer
            Dim ReferenceID As Integer
            Dim TaxCode1 As String = "SLT"
            Dim TaxCode2 As String = "SOF"
            Dim InvoiceDate As Date = Date.Now
            Dim PaidByStatement As String = Common.GetPaidByStatement(conn, pHeader.Quote.ProducerID)
            Dim sMessage As String = ""
            sProgress = "Getting InvoiceKey"
            InvoiceKey_PK = Common.GetKeyField(conn, "InvoiceID")
            InvoiceID = Right("00000" & CStr(InvoiceKey_PK), 8)
            pHeader.InvoiceKey_PK = InvoiceKey_PK
            pHeader.InvoiceID = InvoiceID
            sProgress = "Getting ReferenceID"
            ReferenceID = Common.GetKeyField(conn, "ReferenceID")
            Dim InvoiceKey As Integer = ReferenceID
            Dim comm As New SqlCommand("siu_p_insertinvoiceheader", conn)
            sProgress = "Setting Parameters for Insert"
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@producerID", pHeader.AgencyID)
                .Parameters.AddWithValue("@BillToCode", pHeader.AgencyID)
                .Parameters.AddWithValue("@InvoiceID", InvoiceID)
                .Parameters.AddWithValue("@DefaultPayableID", pHeader.CompanyID)
                .Parameters.AddWithValue("@Effective", pHeader.EffDate)
                .Parameters.AddWithValue("@InvoiceKey_PK", InvoiceKey_PK)
                .Parameters.AddWithValue("@InvoiceTypeID", pHeader.TranType)
                .Parameters.AddWithValue("@Premium", pHeader.Version.Premium)
                .Parameters.AddWithValue("@Non_Premium", pHeader.Version.Non_Premium)
                .Parameters.AddWithValue("@Misc_Premium", pHeader.Version.Misc_Premium)
                .Parameters.AddWithValue("@NonTax_Premium", pHeader.Version.NonTax_Premium)
                .Parameters.AddWithValue("@Tax1", pHeader.Version.Tax1)
                .Parameters.AddWithValue("@Tax2", pHeader.Version.Tax2)
                .Parameters.AddWithValue("@Tax3", pHeader.Version.Tax3)
                .Parameters.AddWithValue("@Tax4", pHeader.Version.Tax4)
                .Parameters.AddWithValue("@Taxed", pHeader.Version.Taxed)
                .Parameters.AddWithValue("@Description", pBatNbr)
                .Parameters.AddWithValue("@BillingType", "AB")
                .Parameters.AddWithValue("@PolicyKey_FK", pHeader.Policy.PolicyKey_PK)
                .Parameters.AddWithValue("@QuoteID", pHeader.Quote.QuoteID)
                .Parameters.AddWithValue("@PolicyID", pHeader.Policy.PolicyID)
                .Parameters.AddWithValue("@TaxState", pHeader.Quote.TaxState)
                .Parameters.AddWithValue("@DirectBillFlag", "N")
                .Parameters.AddWithValue("@DueDate", Date.Now)
                .Parameters.AddWithValue("@PayToDueDate", Date.Now)
                Dim GrossComm As Double
                Dim AgentComm As Double
                If carrierName = "Hartford" Or carrierName = "Capitol" Or carrierName = "NICO" Or carrierName = "AMTrust" Or carrierName = "SafewayReconciliation" Then
                    GrossComm = pHeader.Details.GetGrossCommissionFromDetails()
                    AgentComm = pHeader.Details.GetAgentCommissionFromDetails()
                Else
                    GrossComm = pHeader.Details.AverageCommission("G")
                    AgentComm = pHeader.Details.AverageCommission("A")
                End If
                .Parameters.AddWithValue("@GrossComm", GrossComm)
                .Parameters.AddWithValue("@AgentComm", AgentComm)
                .Parameters.AddWithValue("@StatusID", "P")
                .Parameters.AddWithValue("@InsuredID", pHeader.InsuredID)
                .Parameters.AddWithValue("@TeamID", pHeader.TeamID)
                .Parameters.AddWithValue("@InvoicedByID", pHeader.TeamID)
                .Parameters.AddWithValue("@DateCreated", Date.Now)
                .Parameters.AddWithValue("@ProductID", pHeader.ProductID)
                .Parameters.AddWithValue("@CoverageID", pHeader.CoverageID)
                .Parameters.AddWithValue("@FlagFinanced", IIf(Len(Trim(pHeader.Policy.FinanceCompanyID)) > 0, "Y", "N"))
                .Parameters.AddWithValue("@FinanceID", pHeader.Policy.FinanceCompanyID)
                .Parameters.AddWithValue("@MarketID", pHeader.Version.MarketID)
                .Parameters.AddWithValue("@AcctExec", pHeader.Quote.AcctExec)
                .Parameters.AddWithValue("@InstallmentFlag", "N")
                .Parameters.AddWithValue("@FlagRebill", "N")
                .Parameters.AddWithValue("@PaidByStatement", PaidByStatement)
                .Parameters.AddWithValue("@InvoiceDate", InvoiceDate)
                .Parameters.AddWithValue("@AccountingEffectiveDate", InvoiceDate)
                .Parameters.AddWithValue("@CompanyID", pHeader.CompanyID)
                .Parameters.AddWithValue("@DivisionID", pHeader.Quote.DivisionID)
                .Parameters.AddWithValue("@AcuityTargetCompanyID", pHeader.TeamID)
                .Parameters.AddWithValue("@ContractID", "")
                .Parameters.AddWithValue("@FlagMultiPremium", IIf(pHeader.Details.Multiple("P"), "Y", "N"))
                .Parameters.AddWithValue("@Expiration", pHeader.ExpDate)
                sProgress = "About to Execute"
                Debug.WriteLine("Inserting InvoiceHeader: " + pHeader.HeaderID.ToString())
                .ExecuteNonQuery()
                .Dispose()
            End With
            Return ""
        Catch ex As Exception
            MsgBox(pHeader.HeaderID)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "ImportInvoiceHeader " & CStr(pHeader.HeaderID) & "(" & sProgress & "): " & ex.Message
        End Try
    End Function

    Private Function GetHistory(ByRef conn As SqlConnection, ByRef pHeader As clsHeader) As String
        Try
            'conn.Open()
            Dim comm As New SqlCommand("SIU_p_GetSubmissionInfo", conn)
            Dim rs As SqlDataReader
            Dim QuoteID As String = ""
            Dim Effective As Date
            Dim PolicyKey_FK As Integer
            Dim InsuredID As String
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@PolicyID", pHeader.PolicyNumber)
                .Parameters.AddWithValue("@Effective", pHeader.EffDate)
                rs = .ExecuteReader
                .Dispose()
            End With
            Dim HistoryExists As Boolean = False
            If rs.Read Then
                QuoteID = rs("quoteid")
                Effective = rs("effective")
                PolicyKey_FK = rs("policykey_fk")
                InsuredID = rs("insuredid")
                pHeader.InsuredID = rs("InsuredID")
                pHeader.SubmissionExists = True
                HistoryExists = True
            End If
            rs.Close()
            Dim sMessage As String = ""
            If HistoryExists = True Then
                sMessage = pHeader.Quote.Load(conn, QuoteID)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If
                sMessage = pHeader.Version.Load(conn, QuoteID)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If
                sMessage = pHeader.Policy.Load(conn, PolicyKey_FK)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If

                'sMessage = pHeader.Quote.Load(conn, QuoteID)
                'If Len(Trim(sMessage)) > 0 Then
                '    GoTo _return
                'End If
            Else
                pHeader.SubmissionExists = False
            End If
_return:
            Return sMessage
        Catch ex As Exception
            Return ex.Message & ex.StackTrace

        End Try
    End Function

    Private Function GetHistoryForEffectiveDate(ByRef conn As SqlConnection, ByRef pHeader As clsHeader) As String
        Try
            'conn.Open()
            Dim comm As New SqlCommand("SIU_p_GetSubmissionInfo_WithEffectiveDate", conn)
            Dim rs As SqlDataReader
            Dim QuoteID As String = ""
            Dim Effective As Date
            Dim PolicyKey_FK As Integer
            Dim InsuredID As String
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@PolicyID", pHeader.PolicyNumber)
                .Parameters.AddWithValue("@Effective", pHeader.InceptionDate)
                rs = .ExecuteReader
                .Dispose()
            End With
            Dim HistoryExists As Boolean = False
            If rs.Read Then
                QuoteID = rs("quoteid")
                Effective = rs("effective")
                PolicyKey_FK = rs("policykey_fk")
                InsuredID = rs("insuredid")
                pHeader.InsuredID = rs("InsuredID")
                pHeader.SubmissionExists = True
                HistoryExists = True
            End If
            rs.Close()
            Dim sMessage As String = ""
            If HistoryExists = True Then
                sMessage = pHeader.Quote.Load(conn, QuoteID)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If
                sMessage = pHeader.Version.Load(conn, QuoteID)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If
                sMessage = pHeader.Policy.Load(conn, PolicyKey_FK)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If

                'sMessage = pHeader.Quote.Load(conn, QuoteID)
                'If Len(Trim(sMessage)) > 0 Then
                '    GoTo _return
                'End If
            Else
                pHeader.SubmissionExists = False
            End If
_return:
            Return sMessage
        Catch ex As Exception
            Return ex.Message & ex.StackTrace

        End Try
    End Function

    Private Function GetHistoryFor922(ByRef conn As SqlConnection, ByRef pHeader As clsHeader) As String
        Try
            'conn.Open()
            Dim comm As New SqlCommand("SIU_p_GetSubmissionInfo_For922", conn)
            Dim rs As SqlDataReader
            Dim QuoteID As String = ""
            Dim Effective As Date
            Dim PolicyKey_FK As Integer
            Dim InsuredID As String
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@PolicyID", pHeader.PolicyNumber)
                rs = .ExecuteReader
                .Dispose()
            End With
            Dim HistoryExists As Boolean = False
            If rs.Read Then
                QuoteID = rs("quoteid")
                Effective = rs("effective")
                PolicyKey_FK = rs("policykey_fk")
                InsuredID = rs("insuredid")
                pHeader.InsuredID = rs("InsuredID")
                pHeader.SubmissionExists = True
                HistoryExists = True
            End If
            rs.Close()
            Dim sMessage As String = ""
            If HistoryExists = True Then
                sMessage = pHeader.Quote.Load(conn, QuoteID)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If
                sMessage = pHeader.Version.Load(conn, QuoteID)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If
                sMessage = pHeader.Policy.Load(conn, PolicyKey_FK)
                If Len(Trim(sMessage)) > 0 Then
                    GoTo _return
                End If

            Else
                pHeader.SubmissionExists = False
            End If
_return:
            Return sMessage
        Catch ex As Exception
            Return ex.Message & ex.StackTrace

        End Try
    End Function

    Private Function ImportInvoiceDetails(ByRef conn As SqlConnection, ByRef pHeader As clsHeader, ByVal pBatNbr As String) As String
        Try
            Dim comm As SqlCommand

            Dim myDetail As clsDetail
            Dim iCounter As Integer = 0

            For Each myDetail In pHeader.Details
                iCounter += 1
                comm = New SqlCommand("SIU_p_InsertInvoiceDetail", conn)
                With comm
                    .CommandTimeout = 0
                    .CommandType = Data.CommandType.StoredProcedure
                    .Parameters.AddWithValue("@InvoiceKey_FK", pHeader.InvoiceKey_PK)
                    .Parameters.AddWithValue("@InvoiceDetailKey_PK", iCounter)
                    .Parameters.AddWithValue("@Description", myDetail.Description)
                    .Parameters.AddWithValue("@LineTypeID", myDetail.LineTypeID)
                    .Parameters.AddWithValue("@TransCd", myDetail.TransCD)
                    .Parameters.AddWithValue("@Amount", myDetail.Amount)
                    .Parameters.AddWithValue("@GrossComm", myDetail.GrossComm)
                    .Parameters.AddWithValue("@AgentComm", myDetail.AgentComm)
                    .Parameters.AddWithValue("@Revenue_Amt", myDetail.Revenue_Amt)
                    .Parameters.AddWithValue("@Expense_Amt", myDetail.Expense_Amt)
                    .Parameters.AddWithValue("@CollectedBy", myDetail.CollectedBy)
                    .Parameters.AddWithValue("@CoverageID", pHeader.CoverageID)
                    .Parameters.AddWithValue("@PayableID", myDetail.PayID)
                    .Parameters.AddWithValue("@ComputeAgtComm", "Y")
                    .Parameters.AddWithValue("@ComputeAgyComm", "Y")
                    .Parameters.AddWithValue("@TermPremium", 0)
                    .Parameters.AddWithValue("@MinimumPremium", 0)
                    .Parameters.AddWithValue("@MarketID", pHeader.CompanyID)
                    .Parameters.AddWithValue("@QuoteID", pHeader.Quote.QuoteID)
                    .Parameters.AddWithValue("@ContractID", "")
                    .Parameters.AddWithValue("@linekey_FK", iCounter - 1)
                    .ExecuteNonQuery()
                    .Dispose()
                End With
            Next
            'todo next:  add description, collected by to details
            '            update taxed in version if taxed
            '            look at next in spaimimportinvoicedetl
            Return ""
        Catch ex As Exception
            Return "Import Invoice Details: " & ex.Message
        End Try
    End Function

    Private Function SetSubmission(ByRef conn As SqlConnection, ByRef pHeader As clsHeader) As String
        Try
            Dim comm As New SqlCommand("SIU_p_UpdatePolicyAndVersion", conn)
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@QuoteID", pHeader.Quote.QuoteID)
                .Parameters.AddWithValue("@Taxed", IIf(pHeader.Details.Taxed, "Y", "N"))
                .Parameters.AddWithValue("@InvoiceKey_FK", pHeader.InvoiceKey_PK)
                .Parameters.AddWithValue("@CoverageA", pHeader.CoverageA)
                .Parameters.AddWithValue("@APDeductible", pHeader.APDeductible)
                .Parameters.AddWithValue("@WindDeductible", pHeader.WindDeductible)
                .ExecuteNonQuery()
                .Dispose()
            End With
            Return ""
        Catch ex As Exception
            Return "SetSubmission: " & ex.Message
        End Try
    End Function

    Private Function FixBad901DetailCommission(ByRef conn As SqlConnection, ByVal pInvoiceKey_PK As Integer) As String
        Try
            Dim comm As New SqlCommand("SIU_p_FixBad901DetailCommissionsBykey", conn)
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@INvoiceKey_PK", pInvoiceKey_PK)
                .ExecuteNonQuery()
                .Dispose()
            End With
            Return ""
        Catch ex As Exception
            Return "FixBad901DetailCommission: " & ex.Message
        End Try
    End Function

    Private Function FinalCleanup(ByRef conn As SqlConnection, ByVal pInvoiceKey_PK As Integer) As String
        Try
            Dim comm As New SqlCommand("SIU_p_UpdateImportedInvoice", conn)
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@INvoiceKey_PK", pInvoiceKey_PK)
                .ExecuteNonQuery()
                .Dispose()
            End With
            Return ""
        Catch ex As Exception
            Return "Final Cleanup: " & ex.Message
        End Try
    End Function

    Private Function FinalCleanupForHartfordCapitol(ByRef conn As SqlConnection, ByVal carrierName As String, ByVal policyNumber As String, ByVal pInvoiceKey_PK As Integer) As String
        Try
            Dim comm As New SqlCommand("SIU_p_UpdateImportedInvoiceHartfordCapitol", conn)
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@CarrierName", carrierName)
                .Parameters.AddWithValue("@PolicyId", policyNumber)
                .Parameters.AddWithValue("@INvoiceKey_PK", pInvoiceKey_PK)
                .ExecuteNonQuery()
                .Dispose()
            End With
            Return ""
        Catch ex As Exception
            Return "Final Cleanup: " & ex.Message
        End Try
    End Function

    Private Function CreateSubmission(ByRef conn As SqlConnection, ByRef pHeader As clsHeader) As String
        Try
            Dim iReferenceID As Integer
            Dim sMessage As String = ""
            iReferenceID = Common.GetKeyField(conn, "ReferenceID")
            pHeader.InsuredID = Common.GetNewInsuredID(conn)
            sMessage = InsertInsured(conn, pHeader)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            With pHeader.Quote
                .QuoteID = CStr(Common.GetKeyField(conn, "QuoteID"))
                .StatusID = "PIF"
                .InsuredID = pHeader.InsuredID
                .NamedInsured = pHeader.InsuredName
                .Address1 = pHeader.InsuredAddress1
                .Address2 = pHeader.InsuredAddress2
                .City = pHeader.InsuredCity
                .State = pHeader.InsuredState
                .Zip = pHeader.InsuredZip
                .ProducerID = pHeader.AgencyID
                .ReferenceID = iReferenceID
                .VersionCounter = "B"
                .Effective = pHeader.EffDate
                .Expiration = pHeader.ExpDate
                .CoverageID = pHeader.CoverageID
                .TaxState = pHeader.InsuredState
                .AcctExec = "SYS"
                .CsrID = "SYS"
                .CompanyID = pHeader.CompanyID
                .BndPremium = pHeader.Details.Sum("P")
                .Renewal = IIf(pHeader.TranType = "REN", "Y", "N")
                .ActivePolicyFlag = "Y"
                .ClaimsFlag = "N"
                .SuspenseFlag = "N"
                .OpenItem = "Y"
                .Received = Date.Now
                .VersionBound = "A"
                .PolicyID = pHeader.PolicyNumber
                .Quoted = Date.Now
                .TeamID = pHeader.TeamID
                .ProductID = pHeader.ProductID
                .SubmitTypeID = pHeader.TranType
                .BndMarketID = pHeader.CompanyID
                .PolicyInception = pHeader.EffDate
                sMessage = .Save(conn, pHeader.CoverageA, pHeader.APDeductible, pHeader.WindDeductible, pHeader.Coverage)
            End With
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            With pHeader.Version
                .QuoteID = pHeader.Quote.QuoteID
                .Quoted = Date.Now
                .Version = "A"
                .VerOriginal = "A"
                .VersionID = pHeader.Quote.QuoteID & "A"
                .AgentComm = pHeader.Details.AverageCommission("A")
                .GrossComm = pHeader.Details.AverageCommission("G")
                .MarketID = pHeader.CompanyID
                .CompanyID = pHeader.CompanyID
                .ProductID = pHeader.ProductID
                .Premium = pHeader.Details.Sum("P")
                .Non_Premium = pHeader.Details.Sum("F")
                .Taxed = IIf(pHeader.Details.Taxed, "Y", "N")
                .ProposedEffective = pHeader.EffDate
                .ProposedExpiration = pHeader.ExpDate
                .Financed = "N"
                .BoundFlag = "Y"
                .PolicyTerm = CStr(DateDiff(DateInterval.Month, pHeader.EffDate, pHeader.ExpDate)) & " Months"
                .FlagOverrideCalc = "N"
                sMessage = .Save(conn)
            End With
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            Dim iPolicyKey_PK As Integer = Common.GetKeyField(conn, "ReferenceID")
            With pHeader.Policy
                .QuoteID = pHeader.Quote.QuoteID
                .PolicyID = pHeader.PolicyNumber
                .Effective = pHeader.EffDate
                .Expiration = pHeader.ExpDate
                .Endorsement = 1
                .ActivePolicyFlag = "Y"
                .PolicyKey_PK = iPolicyKey_PK
                .Version = "0"
                .Inception = pHeader.EffDate
                .Term = DateDiff(DateInterval.Day, pHeader.EffDate, pHeader.ExpDate)
                .Bound = Date.Now
                .Invoiced = "Y"
                .CompanyID = pHeader.CompanyID
                .ProductID = pHeader.ProductID
                .InvoiceDate = Date.Now
                .PremiumWritten = pHeader.Details.Sum("P")
                .PremiumTerm = pHeader.Details.Sum("P")
                .FlagInspectionRequired = "N"
                .FlagOverrideServiceUW = "N"
                .DefaultBillingType = "AB"
                sMessage = .Save(conn)
            End With
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            sMessage = GetHistory(conn, pHeader)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            Return ""
        Catch ex As Exception
            Return "CreateSubmission: " & ex.Message
        End Try
    End Function

    Private Function CreateSubmissionForHartfordCapitol(ByRef conn As SqlConnection, ByRef pHeader As clsHeader) As String
        Try
            Dim iReferenceID As Integer
            Dim sMessage As String = ""
            iReferenceID = Common.GetKeyField(conn, "ReferenceID")
            pHeader.InsuredID = Common.GetNewInsuredID(conn)
            sMessage = InsertInsured(conn, pHeader)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            With pHeader.Quote
                .QuoteID = CStr(Common.GetKeyField(conn, "QuoteID"))
                .StatusID = "PIF"
                .InsuredID = pHeader.InsuredID
                .NamedInsured = pHeader.InsuredName
                .Address1 = pHeader.InsuredAddress1
                .Address2 = pHeader.InsuredAddress2
                .City = pHeader.InsuredCity
                .State = pHeader.InsuredState
                .Zip = pHeader.InsuredZip
                .ProducerID = pHeader.AgencyID
                .ReferenceID = iReferenceID
                .VersionCounter = "B"
                .Effective = pHeader.InceptionDate
                .Expiration = pHeader.ExpDate
                .CoverageID = pHeader.CoverageID
                .TaxState = pHeader.InsuredState
                .AcctExec = "SYS"
                .CsrID = "SYS"
                .CompanyID = pHeader.CompanyID
                .BndPremium = pHeader.Details.Sum("P")
                .Renewal = IIf(pHeader.TranType = "REN", "Y", "N")
                .ActivePolicyFlag = "Y"
                .ClaimsFlag = "N"
                .SuspenseFlag = "N"
                .OpenItem = "Y"
                .Received = Date.Now
                .VersionBound = "A"
                .PolicyID = pHeader.PolicyNumber
                .Quoted = Date.Now
                .TeamID = pHeader.TeamID
                .ProductID = pHeader.ProductID
                .SubmitTypeID = pHeader.TranType
                .BndMarketID = pHeader.CompanyID
                .PolicyInception = pHeader.InceptionDate
                sMessage = .Save(conn, pHeader.CoverageA, pHeader.APDeductible, pHeader.WindDeductible, pHeader.Coverage)
            End With
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            With pHeader.Version
                .QuoteID = pHeader.Quote.QuoteID
                .Quoted = Date.Now
                .Version = "A"
                .VerOriginal = "A"
                .VersionID = pHeader.Quote.QuoteID & "A"
                .AgentComm = pHeader.Details.AverageCommission("A")
                .GrossComm = pHeader.Details.AverageCommission("G")
                .MarketID = pHeader.CompanyID
                .CompanyID = pHeader.CompanyID
                .ProductID = pHeader.ProductID
                .Premium = pHeader.Details.Sum("P")
                .Non_Premium = pHeader.Details.Sum("F")
                .Taxed = IIf(pHeader.Details.Taxed, "Y", "N")
                .ProposedEffective = pHeader.InceptionDate
                .ProposedExpiration = pHeader.ExpDate
                .Financed = "N"
                .BoundFlag = "Y"
                .PolicyTerm = CStr(DateDiff(DateInterval.Month, pHeader.InceptionDate, pHeader.ExpDate)) & " Months"
                .FlagOverrideCalc = "N"
                .TerrorActPremium = pHeader.Details.GetTRIA()
                sMessage = .Save(conn)
            End With
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            Dim iPolicyKey_PK As Integer = Common.GetKeyField(conn, "ReferenceID")
            With pHeader.Policy
                .QuoteID = pHeader.Quote.QuoteID
                .PolicyID = pHeader.PolicyNumber
                .Effective = pHeader.InceptionDate
                .Expiration = pHeader.ExpDate
                .Endorsement = 1
                .ActivePolicyFlag = "Y"
                .PolicyKey_PK = iPolicyKey_PK
                .Version = "0"
                .Inception = pHeader.InceptionDate
                .Term = DateDiff(DateInterval.Day, pHeader.InceptionDate, pHeader.ExpDate)
                .Bound = Date.Now
                .Invoiced = "Y"
                .CompanyID = pHeader.CompanyID
                .ProductID = pHeader.ProductID
                .InvoiceDate = Date.Now
                .PremiumWritten = pHeader.Details.Sum("P")
                .PremiumTerm = pHeader.Details.Sum("P")
                .FlagInspectionRequired = "N"
                .FlagOverrideServiceUW = "N"
                .DefaultBillingType = "AB"
                sMessage = .Save(conn)
            End With
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            sMessage = GetHistory(conn, pHeader)
            If Len(Trim(sMessage)) > 0 Then
                Return sMessage
            End If
            Return ""
        Catch ex As Exception
            Return "CreateSubmission: " & ex.Message
        End Try
    End Function

    Private Function InsertInsured(ByRef conn As SqlConnection, ByRef pHeader As clsHeader) As String
        Try
            Dim iReferenceID As Integer = Common.GetKeyField(conn, "ReferenceID")
            Dim comm As New SqlCommand("siu_p_insertinsured", conn)
            With comm
                .CommandTimeout = 0
                .CommandType = Data.CommandType.StoredProcedure
                .Parameters.AddWithValue("@NamedInsured", pHeader.InsuredName)
                .Parameters.AddWithValue("@InsuredID", pHeader.InsuredID)
                .Parameters.AddWithValue("@InsuredKey_PK", iReferenceID)
                .Parameters.AddWithValue("@NameType", "SYS")
                .Parameters.AddWithValue("@DBAName", pHeader.InsuredName)
                .Parameters.AddWithValue("@Address1", pHeader.InsuredAddress1)
                .Parameters.AddWithValue("@Address2", pHeader.InsuredAddress2)
                .Parameters.AddWithValue("@City", pHeader.InsuredCity)
                .Parameters.AddWithValue("@State", pHeader.InsuredState)
                .Parameters.AddWithValue("@Zip", pHeader.InsuredZip)
                .Parameters.AddWithValue("@MailAddress1", pHeader.InsuredAddress1)
                .Parameters.AddWithValue("@MailAddress2", pHeader.InsuredAddress2)
                .Parameters.AddWithValue("@MailCity", pHeader.InsuredCity)
                .Parameters.AddWithValue("@MailState", pHeader.InsuredState)
                .Parameters.AddWithValue("@MailZip", pHeader.InsuredZip)
                .Parameters.AddWithValue("@Phone", "")
                .Parameters.AddWithValue("@Fax", "")
                .Parameters.AddWithValue("@EMail", "")
                .Parameters.AddWithValue("@ProducerID", pHeader.AgencyID)
                .Parameters.AddWithValue("@AcctExec", "SYS")
                .Parameters.AddWithValue("@DateOfBirth", "01/01/1900")
                .Parameters.AddWithValue("@SSN", "")
                .Parameters.AddWithValue("@BusinessStructureID", "")
                .Parameters.AddWithValue("@NCCI", vbNull)
                .Parameters.AddWithValue("@Employees", 0)
                .Parameters.AddWithValue("@SicID", vbNull)
                .Parameters.AddWithValue("@License", vbNull)
                .Parameters.AddWithValue("@WebSite", "")
                .Parameters.AddWithValue("@DateAdded", Date.Now)
                .Parameters.AddWithValue("@ParentKey_FK", vbNull)
                .Parameters.AddWithValue("@MapToID", vbNull)
                .Parameters.AddWithValue("@DirectBillFlag", "N")
                .Parameters.AddWithValue("@FlagProspect", "N")
                .ExecuteNonQuery()
            End With
            Return ""
        Catch ex As Exception
            Return "Insert Insured: " & ex.Message
        End Try
    End Function
#End Region

End Class