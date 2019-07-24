Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    ' MUSTER.DataAccess.FeeBasisDB
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

    Public Class FeeBasisDB
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
        Public Function DBGetFiscalYear(ByVal ForDate As DateTime) As Int16
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim FiscalYear As Int16
            Try


                strSQL = "Select dbo.udfGetFiscalYear('" & ForDate.Date & "') as FiscalYear"

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return drSet.Item("FiscalYear")
                    End While
                Else
                    FiscalYear = DatePart(DateInterval.Year, ForDate.Date)
                    If DatePart(DateInterval.Month, ForDate.Date) > 6 Then
                        FiscalYear += 1
                    End If
                    Return FiscalYear
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function


        Public Function DBGetFiscalYearForFees(Optional ByVal ForDate As DateTime = #1/1/1900#) As Int16
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim FiscalYear As Int16

            Try

                If ForDate = #1/1/1900# Then
                    ForDate = New Date(Now.Year, Now.Month, Now.Day)
                End If


                strSQL = String.Format("Select dbo.udfGetFeeFiscalYear('{0}') as FiscalYear", ForDate.Date)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return drSet.Item("FiscalYear")
                    End While
                Else
                    FiscalYear = Me.DBGetFiscalYear(Date.Now)

                    Return FiscalYear
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        ' Operation to return an INFO object by sending in the Fee Basis Fiscal Year
        Public Function DBGetByFiscalYear(Optional ByVal FiscalYear As Int32 = 2005, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeBasisInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeesBasis"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FEES_BASIS_ID").Value = DBNull.Value
                Params("@FISCAL_YEAR").Value = FiscalYear
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1
                'Params("@Group_ID").Value = nVal
                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS WHERE GROUP_ID = '" & nVal & "'")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        Return New MUSTER.Info.FeeBasisInfo(drSet.Item("FEES_BASIS_ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        drSet.Item("DELETED"), _
                                                        AltIsDBNull(drSet.Item("FISCAL_YEAR"), 0), _
                                                        AltIsDBNull(drSet.Item("STARTING_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("ENDING_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("EARLY_GRACE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("LATE_GRACE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("BASE_FEE"), 0), _
                                                        AltIsDBNull(drSet.Item("BASE_UNIT"), 0), _
                                                        AltIsDBNull(drSet.Item("LATE_TVOE"), 0), _
                                                        AltIsDBNull(drSet.Item("LATE_TYPE"), 0), _
                                                        AltIsDBNull(drSet.Item("PERIOD"), 0), _
                                                        AltIsDBNull(drSet.Item("INVOICE_GEN_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("INVOICE_APP_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("INVOICE_GEN_TIME"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("INVOICE_APP_TIME"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("GENERATED"), 0), _
                                                        AltIsDBNull(drSet.Item("DESCRIPTION"), String.Empty))
                    End While
                Else
                    Return New MUSTER.Info.FeeBasisInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        ' Operation to return an INFO object by sending in the ID
        Public Function DBGetByID(Optional ByVal ComID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeBasisInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetFeesBasis"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FEES_BASIS_ID").Value = ComID
                Params("@FISCAL_YEAR").Value = DBNull.Value
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1
                'Params("@Group_ID").Value = nVal

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS WHERE GROUP_ID = '" & nVal & "'")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FeeBasisInfo(drSet.Item("FEES_BASIS_ID"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        drSet.Item("DELETED"), _
                                                        AltIsDBNull(drSet.Item("FISCAL_YEAR"), 0), _
                                                        AltIsDBNull(drSet.Item("STARTING_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("ENDING_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("EARLY_GRACE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("LATE_GRACE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("BASE_FEE"), 0), _
                                                        AltIsDBNull(drSet.Item("BASE_UNIT"), 0), _
                                                        AltIsDBNull(drSet.Item("LATE_TVOE"), 0), _
                                                        AltIsDBNull(drSet.Item("LATE_TYPE"), 0), _
                                                        AltIsDBNull(drSet.Item("PERIOD"), 0), _
                                                        AltIsDBNull(drSet.Item("INVOICE_GEN_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("INVOICE_APP_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("INVOICE_GEN_TIME"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("INVOICE_APP_TIME"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("GENERATED"), 0), _
                                                        AltIsDBNull(drSet.Item("DESCRIPTION"), String.Empty))

                Else
                    Return New MUSTER.Info.FeeBasisInfo
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
            Dim dsData As Int64
            Try
                dsData = SqlHelper.ExecuteNonQuery(_strConn, CommandType.Text, strSQL)
                Return True
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Operation to send the INFO object to the repository
        Public Sub put(ByRef oFeeBasisInfo As MUSTER.Info.FeeBasisInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim tmpdate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Fees, Integer))) Then
                    returnVal = "You do not have rights to save Fee Basis."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(18) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFeesBasis")
                Params(0).Value = oFeeBasisInfo.ID
                Params(1).Value = oFeeBasisInfo.FiscalYear
                Params(2).Value = oFeeBasisInfo.BaseFee
                Params(3).Value = oFeeBasisInfo.BaseUnit
                Params(4).Value = oFeeBasisInfo.LateFee '0 'LATE_TVOE()
                Params(5).Value = oFeeBasisInfo.LateType
                Params(6).Value = oFeeBasisInfo.LatePeriod 'DBNull.Value 'PERIOD
                Params(7).Value = IIf(IsDBNull(oFeeBasisInfo.Generated), False, oFeeBasisInfo.Generated)
                Params(8).Value = oFeeBasisInfo.GenerateDate 'DBNull.Value 'INVOICE_GEN_DATE()
                Params(9).Value = IIf(oFeeBasisInfo.ApprovedDate = tmpdate, DBNull.Value, oFeeBasisInfo.ApprovedDate) 'DBNull.Value 'INVOICE_APP_DATE()
                Params(10).Value = oFeeBasisInfo.PeriodStart
                Params(11).Value = oFeeBasisInfo.PeriodEnd
                Params(12).Value = oFeeBasisInfo.EarlyGrace
                Params(13).Value = oFeeBasisInfo.LateGrace
                Params(14).Value = oFeeBasisInfo.Description
                Params(15).Value = oFeeBasisInfo.Deleted
                Params(16).Value = IIf(oFeeBasisInfo.GenerateTime = tmpdate, DBNull.Value, oFeeBasisInfo.GenerateTime)
                'Params(16).Value = oFeeBasisInfo.GenerateTime 'DBNull.Value 'INVOICE_GEN_DATE()
                Params(17).Value = IIf(oFeeBasisInfo.ApprovedTime = tmpdate, DBNull.Value, oFeeBasisInfo.ApprovedTime) 'DBNull.Value 'INVOICE_APP_DATE()

                If oFeeBasisInfo.ID <= 0 Then
                    Params(18).Value = oFeeBasisInfo.CreatedBy
                Else
                    Params(18).Value = oFeeBasisInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFeesBasis", Params)
                If Params(0).Value <> 0 Then
                    oFeeBasisInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region

    End Class
End Namespace
