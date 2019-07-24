Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    ' MUSTER.DataAccess.PendFeeLineDB
    ' Provides the means for marshalling Pending Fee Line state to/from the repository
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

    Public Class PendFeeLineDB
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
        Public Function DBGetByFiscalYear(Optional ByVal FiscalYear As Int32 = 2005, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.PendFeeLineInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeesPendingInvoices"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INV_TEMP_ID").Value = DBNull.Value
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1
                'Params("@Group_ID").Value = nVal
                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS WHERE GROUP_ID = '" & nVal & "'")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return New MUSTER.Info.PendFeeLineInfo(drSet.Item("INV_TEMP_ID"), _
                                                                AltIsDBNull(drSet.Item("INV_ADV_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("ITEM_SEQ_NUMBER"), 0), _
                                                                String.Empty, _
                                                                AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("SFY"), String.Empty), _
                                                                Now(), _
                                                                AltIsDBNull(drSet.Item("QUANTITY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("UNIT_PRICE"), 0), _
                                                                AltIsDBNull(drSet.Item("INV_LINE_AMT"), 0), _
                                                                AltIsDBNull(drSet.Item("FEE_TYPE"), String.Empty), _
                                                                0, _
                                                                AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("DESCRIPTION"), String.Empty))
                    End While
                Else
                    Return New MUSTER.Info.PendFeeLineInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        ' Operation to return an INFO object by sending in the ID
        Public Function DBGetByID(Optional ByVal ComID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.PendFeeLineInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeesPendingInvoices"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INV_TEMP_ID").Value = ComID
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1
                'Params("@Group_ID").Value = nVal

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS WHERE GROUP_ID = '" & nVal & "'")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.PendFeeLineInfo(drSet.Item("INV_TEMP_ID"), _
                                        AltIsDBNull(drSet.Item("INV_ADV_ID"), 0), _
                                        AltIsDBNull(drSet.Item("ITEM_SEQ_NUMBER"), 0), _
                                        String.Empty, _
                                        AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                                        AltIsDBNull(drSet.Item("OWNER_ID"), 0), _
                                        AltIsDBNull(drSet.Item("SFY"), String.Empty), _
                                        Now(), _
                                        AltIsDBNull(drSet.Item("QUANTITY"), String.Empty), _
                                        AltIsDBNull(drSet.Item("UNIT_PRICE"), 0), _
                                        AltIsDBNull(drSet.Item("INV_LINE_AMT"), 0), _
                                        AltIsDBNull(drSet.Item("FEE_TYPE"), String.Empty), _
                                        0, _
                                        AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DESCRIPTION"), String.Empty))

                Else
                    Return New MUSTER.Info.PendFeeLineInfo
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
        ' Operation to send the INFO object to the repository
        Public Sub put(ByRef oPendFeeLineInfo As MUSTER.Info.PendFeeLineInfo)
            Try
                Dim Params(20) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFeesPendingInvoices")
                'Params(0).Value = oPendFeeLineInfo.ID
                Params(0).Value = oPendFeeLineInfo.ID
                Params(1).Value = String.Empty
                Params(2).Value = oPendFeeLineInfo.FeeType
                Params(3).Value = oPendFeeLineInfo.InvoiceAdviceId
                Params(4).Value = oPendFeeLineInfo.OwnerId
                Params(5).Value = 0
                Params(6).Value = oPendFeeLineInfo.InvoiceLineAmount
                Params(7).Value = oPendFeeLineInfo.InvoiceNumber
                Params(8).Value = oPendFeeLineInfo.FiscalYear
                Params(9).Value = oPendFeeLineInfo.FacilityId
                Params(10).Value = oPendFeeLineInfo.Description
                Params(11).Value = oPendFeeLineInfo.ItemSequenceNumber
                Params(12).Value = oPendFeeLineInfo.UnitPrice
                Params(13).Value = oPendFeeLineInfo.Quantity
                Params(14).Value = oPendFeeLineInfo.InvoiceDate
                Params(15).Value = oPendFeeLineInfo.DueDate
                Params(16).Value = String.Empty
                Params(17).Value = False
                Params(18).Value = False
                Params(19).Value = oPendFeeLineInfo.Deleted

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFeesPendingInvoices", Params)
                If Params(0).Value <> 0 Then
                    oPendFeeLineInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region

    End Class
End Namespace
