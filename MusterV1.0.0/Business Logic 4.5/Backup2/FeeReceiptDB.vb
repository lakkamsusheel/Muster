
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    ' MUSTER.DataAccess.FeeReceiptDB
    ' Provides the means for marshalling Fee Invoice state to/from the repository
    ' 
    ' Copyright (C) 2004, 2005 CIBER, Inc.
    ' All rights reserved.
    ' 
    ' Release   Initials    Date        Description
    ' 1.0         AB      09/21/05    Original class definition
    ' 
    ' 
    ' Function                  Description
    ' -------------------------------------------------------------------------------    

    Public Class FeeReceiptDB
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
        ' Operation to return an INFO object by sending in the Fee Invoice Fiscal Year
        Public Function DBGetByFiscalYear(Optional ByVal FiscalYear As Int32 = 2005, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeReceiptInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeeReceipt"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FeeReceiptID").Value = DBNull.Value
                Params("@FiscalYear").Value = FiscalYear
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return New MUSTER.Info.FeeReceiptInfo(drSet.Item("RECPT_ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("SFY"), ""), _
                                                        AltIsDBNull(drSet.Item("RETURN_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("CHECK_TRANS_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_NUMBER"), ""), _
                                                        AltIsDBNull(drSet.Item("CHECK_NUMBER"), ""), _
                                                        AltIsDBNull(drSet.Item("MISAPPLY_FLAG"), ""), _
                                                        AltIsDBNull(drSet.Item("MISAPPLY_REASON"), ""), _
                                                        AltIsDBNull(drSet.Item("ISSUING_COMPANY"), ""), _
                                                        AltIsDBNull(drSet.Item("ITEM_SEQ_NUMBER"), 0), _
                                                        AltIsDBNull(drSet.Item("AMT_RECEIVED"), 0), _
                                                        AltIsDBNull(drSet.Item("RECEIPT_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("OVERPAYMENT_REASON"), ""), _
                                                        drSet.Item("DELETED"))
                    End While
                Else
                    Return New MUSTER.Info.FeeReceiptInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        ' Operation to return an INFO object by sending in the ID
        Public Function DBGetByID(Optional ByVal ReceiptID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeReceiptInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeeReceipt"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FeeReceiptID").Value = ReceiptID
                Params("@FiscalYear").Value = DBNull.Value
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FeeReceiptInfo(drSet.Item("RECPT_ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("SFY"), ""), _
                                                        AltIsDBNull(drSet.Item("RETURN_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("CHECK_TRANS_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_NUMBER"), ""), _
                                                        AltIsDBNull(drSet.Item("CHECK_NUMBER"), ""), _
                                                        AltIsDBNull(drSet.Item("MISAPPLY_FLAG"), ""), _
                                                        AltIsDBNull(drSet.Item("MISAPPLY_REASON"), ""), _
                                                        AltIsDBNull(drSet.Item("ISSUING_COMPANY"), ""), _
                                                        AltIsDBNull(drSet.Item("ITEM_SEQ_NUMBER"), 0), _
                                                        AltIsDBNull(drSet.Item("AMT_RECEIVED"), 0), _
                                                        AltIsDBNull(drSet.Item("RECEIPT_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("OVERPAYMENT_REASON"), ""), _
                                                        drSet.Item("DELETED"))

                Else
                    Return New MUSTER.Info.FeeReceiptInfo
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
        Public Sub put(ByRef oFeeReceiptInfo As MUSTER.Info.FeeReceiptInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim tmpdate As Date
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Fees, Integer))) Then
                    returnVal = "You do not have rights to save Fee Receipt."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(17) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFeeReceipt")
                Params(0).Value = oFeeReceiptInfo.ID
                Params(1).Value = oFeeReceiptInfo.FiscalYear
                Params(2).Value = oFeeReceiptInfo.ReturnType
                Params(3).Value = oFeeReceiptInfo.CheckTransID
                Params(4).Value = oFeeReceiptInfo.OwnerID
                Params(5).Value = oFeeReceiptInfo.FacilityID
                Params(6).Value = oFeeReceiptInfo.InvoiceNumber
                Params(7).Value = oFeeReceiptInfo.CheckNumber
                Params(8).Value = oFeeReceiptInfo.MisapplyFlag
                Params(9).Value = oFeeReceiptInfo.MisapplyReason
                Params(10).Value = oFeeReceiptInfo.IssuingCompany
                Params(11).Value = oFeeReceiptInfo.SequenceNumber
                Params(12).Value = oFeeReceiptInfo.AmountReceived
                Params(13).Value = IIf(oFeeReceiptInfo.ReceiptDate = tmpdate, DBNull.Value, oFeeReceiptInfo.ReceiptDate)
                Params(14).Value = oFeeReceiptInfo.Deleted
                Params(15).Value = oFeeReceiptInfo.OverpaymentReason

                If oFeeReceiptInfo.ID <= 0 Then
                    Params(16).Value = oFeeReceiptInfo.CreatedBy
                Else
                    Params(16).Value = oFeeReceiptInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFeeReceipt", Params)
                If Params(0).Value <> 0 Then
                    oFeeReceiptInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region

    End Class
End Namespace