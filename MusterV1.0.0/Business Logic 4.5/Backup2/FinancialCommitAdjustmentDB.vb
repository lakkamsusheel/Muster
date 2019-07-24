' -------------------------------------------------------------------------------
' MUSTER.DataAccess.FinancialCommitmentDB
' Provides the means for marshalling Financial Activity state to/from the repository
' 
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB      06/24/2005   Original class definition
' 
' 
' Function                  Description
' -------------------------------------------------------------------------------    
' 
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class FinancialCommitAdjustmentDB



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
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.FinancialCommitAdjustmentInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialCommitAdjustmentInfo
                End If
                strSQL = "spGetFinancialCommitAdjustment"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Com_Adj_ID").Value = nVal
                Params("@CommitmentID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialCommitAdjustmentInfo(drSet.Item("Com_Adj_ID"), _
                                                                    drSet.Item("Commitment_ID"), _
                                                                    AltIsDBNull(drSet.Item("Adjust_Date"), "1/1/0001"), _
                                                                    drSet.Item("Adjust_Type"), _
                                                                    drSet.Item("Adjust_Amount"), _
                                                                    AltIsDBNull(drSet.Item("Director_App_Req"), 0), _
                                                                    drSet.Item("Fin_App_Req"), _
                                                                    drSet.Item("Approved"), _
                                                                    drSet.Item("Comments"), _
                                                        AltIsDBNull(drSet.Item("CREATE_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted") _
                    )

                Else

                    Return New MUSTER.Info.FinancialCommitAdjustmentInfo
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
        Public Function DBGetByCommitment(ByVal nVal As Integer) As MUSTER.Info.FinancialCommitAdjustmentCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim colText As New MUSTER.Info.FinancialCommitAdjustmentCollection

            Try

                strSQL = "spGetFinancialCommitAdjustment"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Com_Adj_ID").Value = 0
                Params("@CommitmentID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Dim otmpObject As New MUSTER.Info.FinancialCommitAdjustmentInfo(drSet.Item("Com_Adj_ID"), _
                                                                    drSet.Item("Commitment_ID"), _
                                                                    AltIsDBNull(drSet.Item("Adjust_Date"), "1/1/0001"), _
                                                                    drSet.Item("Adjust_Type"), _
                                                                    drSet.Item("Adjust_Amount"), _
                                                                    AltIsDBNull(drSet.Item("Director_App_Req"), 0), _
                                                                    drSet.Item("Fin_App_Req"), _
                                                                    drSet.Item("Approved"), _
                                                                    drSet.Item("Comments"), _
                                                        AltIsDBNull(drSet.Item("CREATE_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted") _
                    )

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
        Public Sub Put(ByRef oFinComAdjInfo As MUSTER.Info.FinancialCommitAdjustmentInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.FinancialCommitment, Integer))) Then
                    returnVal = "You do not have rights to save Financial Commit Adjustment."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFinancialCommitAdjustment")

                With oFinComAdjInfo
                    If .CommitAdjustmentID = 0 Or .CommitAdjustmentID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                    Else
                        Params(0).Value = .CommitAdjustmentID
                    End If
                    Params(1).Value = .CommitmentID
                    Params(2).Value = IIFIsDateNull(.AdjustDate, DBNull.Value)
                    Params(3).Value = .AdjustType
                    Params(4).Value = .AdjustMoney
                    Params(5).Value = .DirectorApprovalReq
                    Params(6).Value = .FinancialApprovalReq
                    Params(7).Value = .Approved
                    Params(8).Value = .Comments
                    Params(9).Value = .Deleted
                    If .CommitAdjustmentID <= 0 Then
                        Params(10).Value = .CreatedBy
                    Else
                        Params(10).Value = .ModifiedBy
                    End If
                End With

                'IIFIsDateNull(oLustEvent.EventEnded, DBNull.Value)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFinancialCommitAdjustment", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oFinComAdjInfo.CommitAdjustmentID Then
                    oFinComAdjInfo.CommitAdjustmentID = Params(0).Value
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

