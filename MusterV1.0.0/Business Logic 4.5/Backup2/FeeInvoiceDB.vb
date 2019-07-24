Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    ' MUSTER.DataAccess.FeeInvoiceDB
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

    Public Class FeeInvoiceDB
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
        Public Function DBGetByFiscalYear(Optional ByVal FiscalYear As Int32 = 2005, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeInvoiceInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeeInvoice"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FeeInvoiceID").Value = DBNull.Value
                Params("@FiscalYear").Value = FiscalYear
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return New MUSTER.Info.FeeInvoiceInfo(drSet.Item("INV__ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("REC_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("FEE_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("ADVICE_ID"), ""), _
                                                        AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_AMT"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_LINE_AMT"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_NUMBER"), ""), _
                                                        AltIsDBNull(drSet.Item("SFY"), ""), _
                                                        AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("DESCRIPTION"), ""), _
                                                        AltIsDBNull(drSet.Item("ITEM_SEQ_NUMBER"), 0), _
                                                        AltIsDBNull(drSet.Item("UNIT_PRICE"), 0), _
                                                        AltIsDBNull(drSet.Item("QUANTITY"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("CREDIT_APPLY_TO"), ""), _
                                                        AltIsDBNull(drSet.Item("TYPE_GENERATION"), 0), _
                                                        AltIsDBNull(drSet.Item("PROCESSED"), 0), _
                                                        AltIsDBNull(drSet.Item("INVOICE_TYPE"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ADDR1"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ADDR2"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_CITY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_STATE"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ZIP"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_NAME"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("CHECK_TRANS_ID"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ORIG_CREDIT_MEMO"), String.Empty), _
                                                        drSet.Item("DELETED"))
                    End While
                Else
                    Return New MUSTER.Info.FeeInvoiceInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        ' Operation to return an INFO object by sending in the ID
        Public Function DBGetByID(Optional ByVal InvoiceID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeInvoiceInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeeInvoice"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FeeInvoiceID").Value = InvoiceID
                Params("@FiscalYear").Value = DBNull.Value
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FeeInvoiceInfo(drSet.Item("INV__ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("REC_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("FEE_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("ADVICE_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_AMT"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_LINE_AMT"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_NUMBER"), ""), _
                                                        AltIsDBNull(drSet.Item("SFY"), ""), _
                                                        AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("DESCRIPTION"), ""), _
                                                        AltIsDBNull(drSet.Item("ITEM_SEQ_NUMBER"), 0), _
                                                        AltIsDBNull(drSet.Item("UNIT_PRICE"), 0), _
                                                        AltIsDBNull(drSet.Item("QUANTITY"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("CREDIT_APPLY_TO"), ""), _
                                                        AltIsDBNull(drSet.Item("TYPE_GENERATION"), 0), _
                                                        AltIsDBNull(drSet.Item("PROCESSED"), 0), _
                                                        AltIsDBNull(drSet.Item("INVOICE_TYPE"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ADDR1"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ADDR2"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_CITY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_STATE"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ZIP"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_NAME"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("CHECK_TRANS_ID"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ORIG_CREDIT_MEMO"), String.Empty), _
                                                        drSet.Item("DELETED"))
                Else
                    Return New MUSTER.Info.FeeInvoiceInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        ' Operation to return a collection of line item invoices for a given header invoice id
        Public Function DBGetLineItemsByHeaderInvoiceID(ByVal invoiceID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeInvoiceCollection
            Dim colLineItems As New MUSTER.Info.FeeInvoiceCollection
            Dim lineItemInfo As MUSTER.Info.FeeInvoiceInfo
            Dim strSP As String
            Dim Params As Collection
            Dim drSet As SqlDataReader

            If invoiceID <= 0 Then Return colLineItems

            Try

                strSP = "spGetFeeLineInvoices"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSP)
                Params("@ADVIC_INV__ID").Value = invoiceID
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSP, Params)
                While drSet.Read
                    lineItemInfo = New MUSTER.Info.FeeInvoiceInfo(drSet.Item("INV__ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("REC_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("FEE_TYPE"), ""), _
                                                        AltIsDBNull(drSet.Item("ADVICE_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_AMT"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_LINE_AMT"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_NUMBER"), ""), _
                                                        AltIsDBNull(drSet.Item("SFY"), ""), _
                                                        AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("DESCRIPTION"), ""), _
                                                        AltIsDBNull(drSet.Item("ITEM_SEQ_NUMBER"), 0), _
                                                        AltIsDBNull(drSet.Item("UNIT_PRICE"), 0), _
                                                        AltIsDBNull(drSet.Item("QUANTITY"), 0), _
                                                        AltIsDBNull(drSet.Item("INV_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("CREDIT_APPLY_TO"), ""), _
                                                        AltIsDBNull(drSet.Item("TYPE_GENERATION"), 0), _
                                                        AltIsDBNull(drSet.Item("PROCESSED"), 0), _
                                                        AltIsDBNull(drSet.Item("INVOICE_TYPE"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ADDR1"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ADDR2"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_CITY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_STATE"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_ZIP"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ISSUE_NAME"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("CHECK_TRANS_ID"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("ORIG_CREDIT_MEMO"), String.Empty), _
                                                        drSet.Item("DELETED"))
                    colLineItems.Add(lineItemInfo)
                End While
                Return colLineItems
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
        Public Sub Put(ByRef oFeeInvoiceInfo As MUSTER.Info.FeeInvoiceInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolAddRegActivityFee As Boolean = False)
            Dim tmpdate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Fees, Integer))) Then
                    returnVal = "You do not have rights to save Fees Invoice."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(20) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFeesInvoices")
                Params(0).Value = oFeeInvoiceInfo.ID
                Params(1).Value = oFeeInvoiceInfo.RecType
                Params(2).Value = oFeeInvoiceInfo.FeeType
                Params(3).Value = oFeeInvoiceInfo.InvoiceAdviceID
                Params(4).Value = oFeeInvoiceInfo.OwnerID
                Params(5).Value = oFeeInvoiceInfo.InvoiceAmount
                Params(6).Value = oFeeInvoiceInfo.InvoiceLineAmount
                Params(7).Value = oFeeInvoiceInfo.WarrantNumber
                Params(8).Value = oFeeInvoiceInfo.FiscalYear
                Params(9).Value = oFeeInvoiceInfo.FacilityID
                Params(10).Value = oFeeInvoiceInfo.Description
                Params(11).Value = oFeeInvoiceInfo.SequenceNumber
                Params(12).Value = oFeeInvoiceInfo.UnitPrice
                Params(13).Value = oFeeInvoiceInfo.Quantity
                Params(14).Value = IIf(oFeeInvoiceInfo.WarrantDate = tmpdate, DBNull.Value, oFeeInvoiceInfo.WarrantDate)
                Params(15).Value = IIf(oFeeInvoiceInfo.DueDate = tmpdate, DBNull.Value, oFeeInvoiceInfo.DueDate)
                Params(16).Value = oFeeInvoiceInfo.CreditApplyTo
                Params(17).Value = oFeeInvoiceInfo.TypeGeneration
                Params(18).Value = oFeeInvoiceInfo.Processed
                Params(19).Value = oFeeInvoiceInfo.Deleted
                Params(20).Value = oFeeInvoiceInfo.InvoiceType
                Params(21).Value = oFeeInvoiceInfo.IssueAddr1
                Params(22).Value = oFeeInvoiceInfo.IssueAddr2
                Params(23).Value = oFeeInvoiceInfo.IssueCity
                Params(24).Value = oFeeInvoiceInfo.IssueState
                Params(25).Value = oFeeInvoiceInfo.IssueZip
                Params(26).Value = oFeeInvoiceInfo.IssueName
                Params(27).Value = oFeeInvoiceInfo.CheckTransID
                Params(28).Value = oFeeInvoiceInfo.CheckNumber

                If oFeeInvoiceInfo.ID <= 0 Then
                    Params(29).Value = oFeeInvoiceInfo.CreatedBy
                Else
                    Params(29).Value = oFeeInvoiceInfo.ModifiedBy
                End If
                Params(30).Value = bolAddRegActivityFee

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFeesInvoices", Params)
                If Params(0).Value <> 0 Then
                    oFeeInvoiceInfo.ID = Params(0).Value
                End If
                If Params(3).Value <> "" Then
                    oFeeInvoiceInfo.InvoiceAdviceID = Params(3).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Operation to send the INFO object to the repository
        Public Sub GenerateDebitMemo(ByRef CreditMemoID As Int64, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Fees, Integer))) Then
                    returnVal = "You do not have rights to Generate Debit Memo."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(2) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spFees_GenerateDebitMemo")
                Params(0).Value = CreditMemoID
                Params(1).Value = UserID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spFees_GenerateDebitMemo", Params)

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region

    End Class
End Namespace
