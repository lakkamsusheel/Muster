
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    ' -------------------------------------------------------------------------------
    ' MUSTER.DataAccess.FinancialTextDB
    ' Provides the means for marshalling Financial Text state to/from the repository
    ' 
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    ' 
    ' Release   Initials    Date        Description
    ' 1.0         JVC      06/08/2005   Original class definition
    ' 
    ' 
    ' Function                  Description
    ' -------------------------------------------------------------------------------    
    ' 
    Public Class FinancialTextDB
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
        Public Function DBGetByID(ByVal nVal As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FinancialTextInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialTextInfo
                End If
                strSQL = "spGetText"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ID").Value = nVal
                Params("@REASON_TYPE").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialTextInfo(drSet.Item("Text_ID"), _
                            AltIsDBNull(drSet.Item("Reason_Type"), 0), _
                            AltIsDBNull(drSet.Item("Text_Name"), ""), _
                            AltIsDBNull(drSet.Item("Financial_Text"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0) _
                    )

                Else

                    Return New MUSTER.Info.FinancialTextInfo
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
        Public Function DBGetByReasonType(ByVal ReasonType As String, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FinancialTextCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim colText As New MUSTER.Info.FinancialTextCollection

            Try
                If ReasonType = "" Then
                    Return colText
                End If
                strSQL = "spGetText"
                strVal = ReasonType

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ID").Value = DBNull.Value
                Params("@REASON_TYPE").Value = ReasonType

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Dim otmpObject As New MUSTER.Info.FinancialTextInfo(drSet.Item("Text_ID"), _
                            AltIsDBNull(drSet.Item("Reason_Type"), 0), _
                            AltIsDBNull(drSet.Item("Text_Name"), ""), _
                            AltIsDBNull(drSet.Item("Financial_Text"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0) _
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
        Public Sub Put(ByRef oFinTextInfo As MUSTER.Info.FinancialTextInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Financial, Integer))) Then
                    returnVal = "You do not have rights to save Financial Text."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSYSText")

                With oFinTextInfo
                    If .ID = 0 Or .ID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                    Else
                        Params(0).Value = .ID
                    End If
                    Params(1).Value = .Active
                    Params(2).Value = .Deleted
                    Params(3).Value = .Reason_Type
                    Params(4).Value = .Reason_Name
                    Params(5).Value = .Reason_Text
                    If .ID <= 0 Then
                        Params(6).Value = .CreatedBy
                    Else
                        Params(6).Value = .ModifiedBy
                    End If
                End With


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSYSText", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oFinTextInfo.ID Then
                    oFinTextInfo.ID = Params(0).Value
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
