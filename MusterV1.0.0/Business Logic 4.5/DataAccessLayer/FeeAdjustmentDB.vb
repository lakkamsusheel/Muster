
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    ' MUSTER.DataAccess.FeeAdjustmentDB
    ' Provides the means for marshalling Fee Basis state to/from the repository
    ' 
    ' Copyright (C) 2004, 2005 CIBER, Inc.
    ' All rights reserved.
    ' 
    ' Release   Initials    Date        Description
    ' 1.0         JVC      06/14/05    Original class definition
    ' 
    ' 
    ' Function                  Description
    ' -------------------------------------------------------------------------------    

    Public Class FeeAdjustmentDB
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

        Public Function DBGetByID(ByVal AdjustmentID As Int64) As MUSTER.Info.FeeAdjustmentInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeeAdjustment"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FEES_ADJ_ID").Value = AdjustmentID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()

                    Return New MUSTER.Info.FeeAdjustmentInfo(drSet.Item("FEES_ADJ_ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        drSet.Item("DELETED"), _
                                                        AltIsDBNull(drSet.Item("Owner_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("Credit_Code"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("SFY"), 0), _
                                                        AltIsDBNull(drSet.Item("Facility_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("Inv_Number"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("Item_SEQ_Number"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("Amount"), 0), _
                                                        AltIsDBNull(drSet.Item("DATE_Applied"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("Check_Number"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("Reason"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("Returned_From_BP2K"), 0))

                Else
                    Return New MUSTER.Info.FeeAdjustmentInfo
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
        Public Sub put(ByRef oFeeAdjustmentInfo As MUSTER.Info.FeeAdjustmentInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim tmpdate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Fees, Integer))) Then
                    returnVal = "You do not have rights to save Fees Adjustment."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(13) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFeeAdjustment")
                Params(0).Value = oFeeAdjustmentInfo.ID
                Params(1).Value = oFeeAdjustmentInfo.Deleted
                Params(2).Value = oFeeAdjustmentInfo.OwnerID
                Params(3).Value = oFeeAdjustmentInfo.CreditCode
                Params(4).Value = oFeeAdjustmentInfo.FiscalYear
                Params(5).Value = oFeeAdjustmentInfo.FacilityID
                Params(6).Value = oFeeAdjustmentInfo.InvoiceNumber
                Params(7).Value = oFeeAdjustmentInfo.ItemSeqNumber
                Params(8).Value = oFeeAdjustmentInfo.Amount
                Params(9).Value = IIf(oFeeAdjustmentInfo.Applied = tmpdate, DBNull.Value, oFeeAdjustmentInfo.Applied)
                Params(10).Value = oFeeAdjustmentInfo.CheckNumber
                Params(11).Value = oFeeAdjustmentInfo.Reason

                If oFeeAdjustmentInfo.ID <= 0 Then
                    Params(12).Value = oFeeAdjustmentInfo.CreatedBy
                Else
                    Params(12).Value = oFeeAdjustmentInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFeeAdjustment", Params)
                If Params(0).Value <> 0 Then
                    oFeeAdjustmentInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region

    End Class
End Namespace