' -------------------------------------------------------------------------------
' MUSTER.DataAccess.FinancialInvoiceDB
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
    Public Class FinancialInvoiceDB


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
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.FinancialInvoiceInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialInvoiceInfo
                End If
                strSQL = "spGetFinancialInvoice"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INVOICES_ID").Value = nVal
                Params("@REIMBURSEMENT_ID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialInvoiceInfo(drSet.Item("INVOICES_ID"), _
                                                                    drSet.Item("REIMBURSEMENT_ID"), _
                                                                    drSet.Item("PAYMENT_SEQ"), _
                                                                    drSet.Item("VENDOR_INV_NUMBER"), _
                                                                    drSet.Item("INVOICED_AMOUNT"), _
                                                                    drSet.Item("PAID"), _
                                                                    drSet.Item("DEDUCTION_REASON"), _
                                                                    drSet.Item("ON_HOLD"), _
                                                                    drSet.Item("FINAL"), _
                                                                    drSet.Item("COMMENT"), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted"), _
                                                        AltIsDBNull(drSet.Item("PONumber"), ""))

                Else

                    Return New MUSTER.Info.FinancialInvoiceInfo
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
        Public Function DBGetByReimbursement(ByVal nVal As Integer) As MUSTER.Info.FinancialInvoiceCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim colText As New MUSTER.Info.FinancialInvoiceCollection

            Try

                strSQL = "spGetFinancialInvoice"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INVOICES_ID").Value = 0
                Params("@REIMBURSEMENT_ID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    Dim otmpObject As New MUSTER.Info.FinancialInvoiceInfo(drSet.Item("INVOICES_ID"), _
                                                                    drSet.Item("REIMBURSEMENT_ID"), _
                                                                    drSet.Item("PAYMENT_SEQ"), _
                                                                    drSet.Item("VENDOR_INV_NUMBER"), _
                                                                    drSet.Item("INVOICED_AMOUNT"), _
                                                                    drSet.Item("PAID"), _
                                                                    drSet.Item("DEDUCTION_REASON"), _
                                                                    drSet.Item("ON_HOLD"), _
                                                                    drSet.Item("FINAL"), _
                                                                    drSet.Item("COMMENT"), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                                                                    drSet.Item("Deleted"), _
                                                        AltIsDBNull(drSet.Item("PONumber"), ""))

                    colText.Add(otmpObject)
                End While
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
        Public Sub Put(ByRef oFinInvoiceInfo As MUSTER.Info.FinancialInvoiceInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.FinancialInvoice, Integer))) Then
                    returnVal = "You do not have rights to save Financial Invoice."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFinancialInvoice")

                With oFinInvoiceInfo
                    If .ID = 0 Or .ID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                    Else
                        Params(0).Value = .ID
                    End If
                    Params(1).Value = .ReimbursementID
                    Params(2).Value = .PaymentSequence
                    Params(3).Value = .VendorInvoice
                    Params(4).Value = .InvoicedAmount
                    Params(5).Value = .PaidAmount
                    Params(6).Value = .DeductionReason
                    Params(7).Value = .OnHold
                    Params(8).Value = .Final
                    Params(9).Value = .Comment
                    Params(10).Value = .Deleted
                    If .ID <= 0 Then
                        Params(11).Value = .CreatedBy
                    Else
                        Params(11).Value = .ModifiedBy
                    End If
                    Params(12).Value = .PONumber
                End With

                'IIFIsDateNull(oLustEvent.EventEnded, DBNull.Value)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFinancialInvoice", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oFinInvoiceInfo.ID Then
                    oFinInvoiceInfo.ID = Params(0).Value
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

