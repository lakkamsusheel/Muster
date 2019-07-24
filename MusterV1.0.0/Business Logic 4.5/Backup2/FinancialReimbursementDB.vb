' -------------------------------------------------------------------------------
' MUSTER.DataAccess.FinancialReimbursementDB
' Provides the means for marshalling Financial Activity state to/from the repository
' 
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB        06/24/2005   Original class definition
' 2.0   Thomas Franey   02/25/2009   Added Comment Data field to retreive from database
' 
' Function                  Description
' -------------------------------------------------------------------------------    
' 
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess

    Public Class FinancialReimbursementDB

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
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.FinancialReimbursementInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialReimbursementInfo
                End If
                strSQL = "spGetFinancialReimbursement"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@REIMBURSEMENT_ID").Value = nVal
                Params("@COMMITMENT_ID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialReimbursementInfo(drSet.Item("REIMBURSEMENT_ID"), _
                                                                    drSet.Item("FINANCIAL_EVENTID"), _
                                                                    drSet.Item("COMMITMENT_ID"), _
                                                                    drSet.Item("PAYMENT_NUMBER"), _
                                                        AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("PAYMENT_DATE"), "1/1/0001"), _
                                                                    drSet.Item("REQUESTED_AMOUNT"), _
                                                                    drSet.Item("INCOMPLETE_REASON"), _
                                                                    drSet.Item("INCOMPLETE"), _
                                                                    drSet.Item("INCOMPLETE_OTHER"), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted"), _
                                                        AltIsDBNull(drSet.Item("PONumber"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("Comment"), String.Empty))

                Else

                    Return New MUSTER.Info.FinancialReimbursementInfo
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
        Public Function DBGetByCommitment(ByVal nVal As Integer) As MUSTER.Info.FinancialReimbursementCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim colText As New MUSTER.Info.FinancialReimbursementCollection

            Try

                strSQL = "spGetFinancialReimbursement"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@REIMBURSEMENT_ID").Value = 0
                Params("@COMMITMENT_ID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Dim otmpObject As New MUSTER.Info.FinancialReimbursementInfo(drSet.Item("REIMBURSEMENT_ID"), _
                                                                    drSet.Item("FINANCIAL_EVENTID"), _
                                                                    drSet.Item("COMMITMENT_ID"), _
                                                                    drSet.Item("PAYMENT_NUMBER"), _
                                                        AltIsDBNull(drSet.Item("RECEIVED_DATE"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("PAYMENT_DATE"), "1/1/0001"), _
                                                                    drSet.Item("REQUESTED_AMOUNT"), _
                                                                    drSet.Item("INCOMPLETE_REASON"), _
                                                                    drSet.Item("INCOMPLETE"), _
                                                                    drSet.Item("INCOMPLETE_OTHER"), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted"), _
                                                        AltIsDBNull(drSet.Item("PONumber"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("Comment"), String.Empty))

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
        Public Sub Put(ByRef oFinInvoiceInfo As MUSTER.Info.FinancialReimbursementInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim tmpDate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.FinancialReimbursement, Integer))) Then
                    returnVal = "You do not have rights to save Financial Reimbursement."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFinancialReimbursement")

                With oFinInvoiceInfo
                    If .ID = 0 Or .ID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                    Else
                        Params(0).Value = .ID
                    End If
                    Params(1).Value = .FinancialEventID
                    Params(2).Value = .CommitmentID
                    Params(3).Value = .PaymentNumber
                    Params(4).Value = IIf(.ReceivedDate = tmpDate, DBNull.Value, .ReceivedDate)
                    Params(5).Value = IIf(.PaymentDate = tmpDate, DBNull.Value, .PaymentDate)
                    Params(6).Value = .RequestedAmount
                    Params(7).Value = .Incomplete
                    Params(8).Value = IIf(.IncompleteReason Is Nothing, String.Empty, .IncompleteReason)
                    Params(9).Value = IIf(.IncompleteOther Is Nothing, String.Empty, .IncompleteOther)
                    Params(10).Value = .Deleted
                    If .ID <= 0 Then
                        Params(11).Value = .CreatedBy
                    Else
                        Params(11).Value = .ModifiedBy
                    End If
                    Params(12).Value = .PONumber
                    Params(13).Value = .Comment
                End With

                'IIFIsDateNull(oLustEvent.EventEnded, DBNull.Value)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFinancialReimbursement", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oFinInvoiceInfo.ID Then
                    oFinInvoiceInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        ' Operation to return an INFO object by sending in the ID
        Public Function DBProcessReimbursementNotification(ByVal nVal As Int64, ByVal dteStart As Date, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim strSQL As String
            Dim Params As Collection

            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.FinancialReimbursement, Integer))) Then
                    returnVal = "You do not have rights to process Financial Reimbursement Notification."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If



                strSQL = "spProcessReimbursementInvoice"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@REIMBURSEMENTID").Value = nVal
                Params("@CloseDate").Value = dteStart
                Params("@UserID").Value = UserID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally

            End Try
        End Function

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

