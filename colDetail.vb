Imports System.Data.SqlClient
Public Class colDetail
    Inherits CollectionBase
    Default Public ReadOnly Property Item(ByVal Index As Integer) As clsDetail
        Get
            Return CType(list.Item(Index), clsDetail)
        End Get
    End Property
    Public Function Add(ByVal pLineTypeID As String, ByVal pTransCD As String, ByVal pAmount As Double, ByVal pGrossComm As Double, ByVal pAgentComm As Double, ByVal pCollectedBy As String, ByVal pDescription As String, ByVal pPayID As String) As clsDetail
        Dim myDetail As New clsDetail
        With myDetail
            .AgentComm = pAgentComm
            .Amount = pAmount
            .GrossComm = pGrossComm
            .LineTypeID = pLineTypeID
            .TransCD = pTransCD
            .CollectedBy = pCollectedBy
            .Description = pDescription
            .PayID = pPayID
        End With
        list.Add(myDetail)
        Return myDetail
    End Function
    Public Function Taxed() As Boolean
        Dim myD As clsDetail
        For Each myD In list
            If UCase(Trim(myD.LineTypeID)) = "T" Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function Sum(ByVal LineTypeID As String) As Double
        Dim dAnswer As Double = 0
        Dim myD As clsDetail
        For Each myD In list
            If UCase(Trim(myD.LineTypeID)) = UCase(Trim(LineTypeID)) And myD.TransCD <> "DBC" Then
                dAnswer += myD.Amount
            End If
        Next
        Return dAnswer
    End Function
    Public Function HTMLSummary() As String
        Dim sAnswer As String = ""
        Dim myD As clsDetail
        Dim iCounter As Integer = 0

        For Each myD In list
            iCounter += 1
            sAnswer += myD.HTML(iCounter) & vbCrLf & "<hr>"
        Next
        Return sAnswer
    End Function
    Public Function AverageCommission(ByVal pAgentOrGross As String) As Double
        Dim myD As clsDetail
        Dim dCommAmount As Double
        Dim dSum As Double
        For Each myD In list
            If UCase(Trim(myD.TransCD)) <> "DBC" And UCase(Trim(myD.LineTypeID)) = "P" Then
                dCommAmount += IIf(UCase(Trim(pAgentOrGross)) = "A", myD.Expense_Amt, myD.Revenue_Amt)
            End If
        Next
        dSum = Sum("P")
        Return (dCommAmount / dSum) * 100
    End Function
    Public Function GetTRIA() As Double
        Dim myD As clsDetail
        Dim TRIA As Double
        For Each myD In List
            If UCase(Trim(myD.TransCD)) = "TRE" Then
                TRIA = myD.Amount
            End If
        Next
        Return TRIA
    End Function
    Public Function GetGrossCommissionFromDetails() As Double
        Dim myD As clsDetail
        Dim dcomm As Double = 0
        For Each myD In List
            If dcomm < myD.GrossComm Then
                dcomm = myD.GrossComm
            End If
        Next
        Return dcomm
    End Function
    Public Function GetAgentCommissionFromDetails() As Double
        Dim myD As clsDetail
        Dim dcomm As Double = 0
        For Each myD In List
            If dcomm < myD.AgentComm Then
                dcomm = myD.AgentComm
            End If
        Next
        Return dcomm
    End Function
    Public Function Multiple(ByVal pLineTypeID As String) As Boolean
        Dim iCount As Integer = 0
        Dim myD As clsDetail
        For Each myD In list
            If UCase(Trim(pLineTypeID)) = UCase(Trim(myD.LineTypeID)) Then
                iCount = iCount + 1
            End If
        Next
        Return iCount > 1
    End Function
    Friend Function Save(ByRef conn As sqlconnection, ByVal pHeaderID As Integer) As String
        Try
            Dim iCounter As Integer
            Dim sMessage As String = ""
            For iCounter = 0 To list.Count - 1
                With Item(iCounter)
                    sMessage = .Save(conn, pHeaderID, iCounter)
                    If Len(Trim(sMessage)) > 0 Then Return sMessage
                End With
            Next
            Return ""
        Catch ex As Exception
            Return "colDetail.Save: " & ex.Message
        End Try
    End Function
End Class
