' -------------------------------------------------------------------------------
' MUSTER.DataAccess.FinancialCommitmentDB
' Provides the means for marshalling Financial Activity state to/from the repository
' 
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB      06/24/2005   Original class definition

' 2.0    Thomas Franey  2/24/2009   Added Install Setup To Put and Get data Functions
' 
' Function                  Description
' -------------------------------------------------------------------------------    
' 
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class FinancialCommitmentDB


#Region "Private Member Variables"
        Private _strConn As Object
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
#End Region

#Region "Exposed Methods"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            Try
                If MusterXCEP Is Nothing Then
                    MusterException = New MUSTER.Exceptions.MusterExceptions
                Else
                    MusterException = MusterXCEP
                End If
                If strDBConn = String.Empty Then
                    Dim oCnn As New ConnectionSettings
                    _strConn = oCnn.cnString
                    oCnn = Nothing
                Else
                    _strConn = strDBConn
                End If
            Catch ex As Exception
                Throw ex
            End Try

        End Sub

        ' Operation to return an INFO object by sending in the ID
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.FinancialCommitmentInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialCommitmentInfo
                End If
                strSQL = "spGetFinancialCommitment"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@CommitmentID").Value = nVal
                Params("@EventID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialCommitmentInfo(drSet.Item("CommitmentID"), _
                                                                    drSet.Item("Fin_Event_ID"), _
                                                                    drSet.Item("FundingType"), _
                                                                    drSet.Item("PONumber"), _
                                                                    drSet.Item("NewPONumber"), _
                                                                    drSet.Item("RollOver"), _
                                                                    drSet.Item("ZeroOut"), _
                                                        AltIsDBNull(drSet.Item("ApprovedDate"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("SOWDate"), "1/1/0001"), _
                                                                    drSet.Item("ContractType"), _
                                                                    drSet.Item("ActivityType"), _
                                                                    drSet.Item("ReimbursementCondition"), _
                                                        AltIsDBNull(drSet.Item("Due_Date"), "1/1/0001"), _
                                                                    drSet.Item("ThirdPartyPayment"), _
                                                                    drSet.Item("ThirdPartyPayee"), _
                                                        AltIsDBNull(drSet.Item("DueDateStatement"), "1/1/0001"), _
                                                                    drSet.Item("Case_Letter"), _
                                                                    drSet.Item("ERACServices"), _
                                                                    drSet.Item("LaboratoryServices"), _
                                                                    drSet.Item("FixedFee"), _
                                                                    drSet.Item("NumberofEvents"), _
                                                                    drSet.Item("WellAbandonment"), _
                                                                    drSet.Item("FreeProductRecovery"), _
                                                                    drSet.Item("VacuumContServices"), _
                                                                    drSet.Item("VacuumContServicesCnt"), _
                                                                    drSet.Item("PTTTesting"), _
                                                                    drSet.Item("ERACVacuum"), _
                                                                    drSet.Item("ERACVacuumCnt"), _
                                                                    drSet.Item("ERACSampling"), _
                                                                    drSet.Item("IRACServicesEstimate"), _
                                                                    drSet.Item("SubContractorSvcs"), _
                                                                    drSet.Item("ORCContractorSvcs"), _
                                                                    drSet.Item("REMContractorSvcs"), _
                                                                    drSet.Item("PreInstallSetup"), _
                                                                    drSet.Item("MonthlySystemUse"), _
                                                                    drSet.Item("MonthlySystemUseCnt"), _
                                                                    drSet.Item("MonthlyOMSampling"), _
                                                                    drSet.Item("MonthlyOMSamplingCnt"), _
                                                                    drSet.Item("TriAnnualOMSampling"), _
                                                                    drSet.Item("TriAnnualOMSamplingCnt"), _
                                                                    drSet.Item("EstimateTriAnnualLab"), _
                                                                    drSet.Item("EstimateTriAnnualLabCnt"), _
                                                                    drSet.Item("EstimateUtilities"), _
                                                                    drSet.Item("EstimateUtilitiesCnt"), _
                                                                    drSet.Item("ThirdPartySettlement"), _
                                                                    drSet.Item("Comments"), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted"), _
                                                                    drSet.Item("CostRecovery"), _
                                                                    drSet.Item("Markup"), drSet.Item("InstallSetup"), drSet.Item("ReimburseERAC"))

                Else

                    Return New MUSTER.Info.FinancialCommitmentInfo
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        ' Operation to return an INFO object by sending in the Name
        Public Function DBGetByFinancialEvent(ByVal nVal As Integer) As MUSTER.Info.FinancialCommitmentCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim colText As New MUSTER.Info.FinancialCommitmentCollection

            Try

                strSQL = "spGetFinancialCommitment"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@CommitmentID").Value = 0
                Params("@EventID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Dim otmpObject As New MUSTER.Info.FinancialCommitmentInfo(drSet.Item("CommitmentID"), _
                                                                    drSet.Item("Fin_Event_ID"), _
                                                                    drSet.Item("FundingType"), _
                                                                    drSet.Item("PONumber"), _
                                                                    drSet.Item("NewPONumber"), _
                                                                    drSet.Item("RollOver"), _
                                                                    drSet.Item("ZeroOut"), _
                                                        AltIsDBNull(drSet.Item("ApprovedDate"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("SOWDate"), "1/1/0001"), _
                                                                    drSet.Item("ContractType"), _
                                                                    drSet.Item("ActivityType"), _
                                                                    drSet.Item("ReimbursementCondition"), _
                                                        AltIsDBNull(drSet.Item("Due_Date"), "1/1/0001"), _
                                                                    drSet.Item("ThirdPartyPayment"), _
                                                                    drSet.Item("ThirdPartyPayee"), _
                                                        AltIsDBNull(drSet.Item("DueDateStatement"), "1/1/0001"), _
                                                                    drSet.Item("Case_Letter"), _
                                                                    drSet.Item("ERACServices"), _
                                                                    drSet.Item("LaboratoryServices"), _
                                                                    drSet.Item("FixedFee"), _
                                                                    drSet.Item("NumberofEvents"), _
                                                                    drSet.Item("WellAbandonment"), _
                                                                    drSet.Item("FreeProductRecovery"), _
                                                                    drSet.Item("VacuumContServices"), _
                                                                    drSet.Item("VacuumContServicesCnt"), _
                                                                    drSet.Item("PTTTesting"), _
                                                                    drSet.Item("ERACVacuum"), _
                                                                    drSet.Item("ERACVacuumCnt"), _
                                                                    drSet.Item("ERACSampling"), _
                                                                    drSet.Item("IRACServicesEstimate"), _
                                                                    drSet.Item("SubContractorSvcs"), _
                                                                    drSet.Item("ORCContractorSvcs"), _
                                                                    drSet.Item("REMContractorSvcs"), _
                                                                    drSet.Item("PreInstallSetup"), _
                                                                    drSet.Item("MonthlySystemUse"), _
                                                                    drSet.Item("MonthlySystemUseCnt"), _
                                                                    drSet.Item("MonthlyOMSampling"), _
                                                                    drSet.Item("MonthlyOMSamplingCnt"), _
                                                                    drSet.Item("TriAnnualOMSampling"), _
                                                                    drSet.Item("TriAnnualOMSamplingCnt"), _
                                                                    drSet.Item("EstimateTriAnnualLab"), _
                                                                    drSet.Item("EstimateTriAnnualLabCnt"), _
                                                                    drSet.Item("EstimateUtilities"), _
                                                                    drSet.Item("EstimateUtilitiesCnt"), _
                                                                    drSet.Item("ThirdPartySettlement"), _
                                                                    drSet.Item("Comments"), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted"), _
                                                                    drSet.Item("CostRecovery"), _
                                                                    drSet.Item("Markup"), drSet.Item("InstallSetup"), drSet.Item("ReimburseERAC"))

                    colText.Add(otmpObject)
                End If
                Return colText
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function

        ' Operation to send the INFO object to the repository
        Public Sub Put(ByRef oFinComInfo As MUSTER.Info.FinancialCommitmentInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.FinancialCommitment, Integer))) Then
                    returnVal = "You do not have rights to save Financial Commitment."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFinancialCommitment")

                With oFinComInfo
                    If .CommitmentID = 0 Or .CommitmentID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                    Else
                        Params(0).Value = .CommitmentID
                    End If
                    Params(1).Value = .Fin_Event_ID
                    Params(2).Value = .FundingType
                    Params(3).Value = IIf(IsNothing(.PONumber), "", .PONumber)
                    Params(4).Value = IIf(IsNothing(.NewPONumber), "", .NewPONumber)
                    Params(5).Value = .RollOver
                    Params(6).Value = .ZeroOut
                    Params(7).Value = IIFIsDateNull(.ApprovedDate, DBNull.Value)
                    Params(8).Value = IIFIsDateNull(.SOWDate, DBNull.Value)
                    Params(9).Value = .ContractType
                    Params(10).Value = .ActivityType
                    Params(11).Value = IIf(IsNothing(.ReimbursementCondition), "", .ReimbursementCondition)
                    Params(12).Value = IIFIsDateNull(.DueDate, DBNull.Value)
                    Params(13).Value = .ThirdPartyPayment
                    Params(14).Value = IIf(IsNothing(.ThirdPartyPayee), "", .ThirdPartyPayee)
                    Params(15).Value = .DueDateStatement
                    Params(16).Value = .Case_Letter
                    Params(17).Value = .ERACServices
                    Params(18).Value = .LaboratoryServices
                    Params(19).Value = .FixedFee
                    Params(20).Value = .NumberofEvents
                    Params(21).Value = .WellAbandonment
                    Params(22).Value = .FreeProductRecovery
                    Params(23).Value = .VacuumContServices
                    Params(24).Value = .VacuumContServicesCnt
                    Params(25).Value = .PTTTesting
                    Params(26).Value = .ERACVacuum
                    Params(27).Value = .ERACVacuumCnt
                    Params(28).Value = .ERACSampling
                    Params(29).Value = .IRACServicesEstimate
                    Params(30).Value = .SubContractorSvcs
                    Params(31).Value = .ORCContractorSvcs
                    Params(32).Value = .REMContractorSvcs
                    Params(33).Value = .PreInstallSetup
                    Params(34).Value = .MonthlySystemUse
                    Params(35).Value = .MonthlySystemUseCnt
                    Params(36).Value = .MonthlyOMSampling
                    Params(37).Value = .MonthlyOMSamplingCnt
                    Params(38).Value = .TriAnnualOMSampling
                    Params(39).Value = .TriAnnualOMSamplingCnt
                    Params(40).Value = .EstimateTriAnnualLab
                    Params(41).Value = .EstimateTriAnnualLabCnt
                    Params(42).Value = .EstimateUtilities
                    Params(43).Value = .EstimateUtilitiesCnt
                    Params(44).Value = .ThirdPartySettlement
                    Params(45).Value = IIf(IsNothing(.Comments), "", .Comments)
                    Params(46).Value = .Deleted
                    Params(47).Value = .CostRecovery
                    Params(48).Value = .Markup

                    If .CommitmentID <= 0 Then
                        Params(49).Value = .CreatedBy
                    Else
                        Params(49).Value = .ModifiedBy
                    End If

                    Params(50).Value = .InstallSetup
                    Params(51).Value = .ReimburseERAC

                End With

                'IIFIsDateNull(oLustEvent.EventEnded, DBNull.Value)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFinancialCommitment", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oFinComInfo.CommitmentID Then
                    oFinComInfo.CommitmentID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region

    End Class
End Namespace
