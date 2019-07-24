Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    ' MUSTER.DataAccess.FeeLateFeeDB
    ' Provides the means for marshalling Late Fee state to/from the repository
    ' 
    ' Copyright (C) 2004, 2005 CIBER, Inc.
    ' All rights reserved.
    ' 
    ' Release   Initials    Date        Description
    ' 1.0         AB      12/05/05    Original class definition
    ' 
    ' 
    ' Function                  Description
    ' -------------------------------------------------------------------------------    

    Public Class FeeLateFeeDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

#Region "Exposed Operations"
        ' Operation to return an INFO object by sending in the Fee Basis Fiscal Year
        Public Function DBGetByid(ByVal LateFeeCertificationID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeLateFeeInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFees_LateFeeCertification"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@LATE_CERT_ID").Value = LateFeeCertificationID
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return New MUSTER.Info.FeeLateFeeInfo(drSet.Item("LATE_CERT_ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        drSet.Item("FISCAL_YEAR"), _
                                                        drSet.Item("CHARGES"), _
                                                        AltIsDBNull(drSet.Item("INVOICE_NUMBER"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("CERT_LETTER_NUMBER"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("WAIVE_APPROVAL_REC"), 0), _
                                                        AltIsDBNull(drSet.Item("WAIVE_APPROVAL_STATUS"), 0), _
                                                        AltIsDBNull(drSet.Item("WAIVE_REASON"), 0), _
                                                        AltIsDBNull(drSet.Item("PROCESS_CERTIFICATION"), 0), _
                                                        AltIsDBNull(drSet.Item("PROCESS_WAIVER"), 0), _
                                                        AltIsDBNull(drSet.Item("WAIVE_FINALIZED"), CDate("01/01/0001")), _
                                                        drSet.Item("DELETED"))
                    End While
                Else
                    Return New MUSTER.Info.FeeLateFeeInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        ' Operation to return an INFO object by sending in the Fee Basis Fiscal Year
        Public Function DBGetByInvoiceNumber(ByVal InvoiceNumber As String, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeLateFeeInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFees_LateFeeCertification_ByInvoiceNumber"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INVOICE_NUMBER").Value = InvoiceNumber
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return New MUSTER.Info.FeeLateFeeInfo(drSet.Item("LATE_CERT_ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        drSet.Item("FISCAL_YEAR"), _
                                                        drSet.Item("CHARGES"), _
                                                        AltIsDBNull(drSet.Item("INVOICE_NUMBER"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("CERT_LETTER_NUMBER"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("WAIVE_APPROVAL_REC"), 0), _
                                                        AltIsDBNull(drSet.Item("WAIVE_APPROVAL_STATUS"), 0), _
                                                        AltIsDBNull(drSet.Item("WAIVE_REASON"), 0), _
                                                        AltIsDBNull(drSet.Item("PROCESS_CERTIFICATION"), 0), _
                                                        AltIsDBNull(drSet.Item("PROCESS_WAIVER"), 0), _
                                                        AltIsDBNull(drSet.Item("WAIVE_FINALIZED"), CDate("01/01/0001")), _
                                                        drSet.Item("DELETED"))
                    End While
                Else
                    Return New MUSTER.Info.FeeLateFeeInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        ' Operation to return a dataset object by sending in the SQL String used to produce it
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


        Public Function DBGetIfCertIsValid(ByVal ownerID As String, ByVal cert As String, ByVal Id As String, ByRef Results As Integer) As Integer

            Try

                Dim Params(11) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "CheckForValidCertNum")
                Params(0).Value = ownerID
                Params(1).Value = cert
                Params(2).Value = Id
                Params(3).Direction = ParameterDirection.Output
                Params(3).Value = 0

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "CheckForValidCertNum", Params)

                Results = Params(3).Value

                Return Results

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Operation to return a dataset object by sending in the SQL String used to produce it
        Public Function DBExecNonQuery(ByVal strSQL As String) As Boolean
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteNonQuery(_strConn, CommandType.Text, strSQL)
                Return True
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Operation to send the INFO object to the repository
        Public Sub put(ByRef oFeeLateFeeInfo As MUSTER.Info.FeeLateFeeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim tmpdate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Fees, Integer))) Then
                    returnVal = "You do not have rights to save Late Fee Information."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(11) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFees_LateFeeCertification")
                Params(0).Value = oFeeLateFeeInfo.ID
                Params(1).Value = oFeeLateFeeInfo.FiscalYear
                Params(2).Value = oFeeLateFeeInfo.InvoiceNumber
                Params(3).Value = oFeeLateFeeInfo.LateCharges
                Params(4).Value = oFeeLateFeeInfo.CertLetterNumber
                Params(5).Value = oFeeLateFeeInfo.WaiveApprovalRecommendation
                Params(6).Value = oFeeLateFeeInfo.WaiveApprovalStatus
                Params(7).Value = oFeeLateFeeInfo.WaiveReason
                Params(8).Value = oFeeLateFeeInfo.ProcessCertification
                Params(9).Value = oFeeLateFeeInfo.ProcessWaiver
                If Date.Compare(oFeeLateFeeInfo.WaiverFinalizedOn, CDate("01/01/0001")) = 0 Then
                    Params(10).Value = DBNull.Value
                Else
                    Params(10).Value = oFeeLateFeeInfo.WaiverFinalizedOn
                End If
                'Params(10).Value = IIf(oFeeLateFeeInfo.WaiverFinalizedOn = tmpdate, DBNull.Value, oFeeLateFeeInfo.WaiverFinalizedOn)
                Params(11).Value = oFeeLateFeeInfo.Deleted

                If oFeeLateFeeInfo.ID <= 0 Then
                    Params(12).Value = oFeeLateFeeInfo.CreatedBy
                Else
                    Params(12).Value = oFeeLateFeeInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFees_LateFeeCertification", Params)
                If Params(0).Value <> 0 Then
                    oFeeLateFeeInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub


        Public Function putLatFees(ByVal excuse As String, ByVal userID As String) As Integer
            Dim tmpdate As Date
            Dim retVal As Integer = -1
            Try


                Dim Params(8) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSYSPROPERTYMASTER")
                Params(0).Value = retVal
                Params(0).Direction = ParameterDirection.Output


                Params(1).Value = 149
                Params(2).Value = excuse
                Params(3).Value = "User Custom Excuse"
                Params(4).Value = Nothing

                Params(5).Value = 0
                Params(6).Value = "YES"
                Params(7).Value = userID


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSYSPROPERTYMASTER", Params)
                If Params(0).Value <> 0 Then
                    retVal = Params(0).Value
                End If

                Return retVal

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function putRegtag(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal ownerID As Integer, ByVal processed As Boolean, ByVal year As Integer, ByVal certNumber As String, ByVal facilityID As Integer, Optional ByVal ID As Integer = -1) As Integer
            Dim tmpdate As Date
            Dim retVal As Integer = -1
            Try


                Dim Params(5) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutRedTag")


                Params(0).Value = ownerID
                Params(1).Value = year
                Params(2).Value = certNumber
                Params(3).Value = processed
                Params(4).Value = facilityID
                Params(5).Direction = ParameterDirection.InputOutput
                Params(5).Value = DBNull.Value


                If ID > -1 Then
                    Params(5).Value = ID
                End If



                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutRedTag", Params)

                Return Params(5).Value

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region

    End Class
End Namespace