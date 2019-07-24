' -------------------------------------------------------------------------------
' MUSTER.DataAccess.FinancialActivityDB
' Provides the means for marshalling Financial Activity state to/from the repository
' 
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB      06/23/2005   Original class definition
' 
' 
' Function                  Description
' -------------------------------------------------------------------------------    
' 
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess

Public Class FinancialActivityDB
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
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.FinancialActivityInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.FinancialActivityInfo
                End If
                strSQL = "spGetFinActivity"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ActivityID").Value = nVal

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.FinancialActivityInfo(drSet.Item("Activity_ID"), _
                            AltIsDBNull(drSet.Item("ActivityDesc"), ""), _
                            AltIsDBNull(drSet.Item("ActivityDescShort"), ""), _
                            AltIsDBNull(drSet.Item("TimeAndMaterials"), 0), _
                            AltIsDBNull(drSet.Item("CostPlus"), 0), _
                            AltIsDBNull(drSet.Item("FixedPrice"), 0), _
                            AltIsDBNull(drSet.Item("DueDateStatement"), ""), _
                            AltIsDBNull(drSet.Item("ReimbursementCondition"), 0), _
                            AltIsDBNull(drSet.Item("ReimbursementConditionDesc"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("TimeAndMaterialsDesc"), 0), _
                            AltIsDBNull(drSet.Item("CostPlusDesc"), 0), _
                            AltIsDBNull(drSet.Item("FixedPriceDesc"), 0), _
                            AltIsDBNull(drSet.Item("CoverTemplate"), ""), _
                            AltIsDBNull(drSet.Item("NoticeTemplate"), ""))

                Else

                    Return New MUSTER.Info.FinancialActivityInfo
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
        Public Function DBGetAll() As MUSTER.Info.FinancialActivityCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim colText As New MUSTER.Info.FinancialActivityCollection

            Try

                strSQL = "spGetFinActivity"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ActivityID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Dim otmpObject As New MUSTER.Info.FinancialActivityInfo(drSet.Item("Activity_ID"), _
                            AltIsDBNull(drSet.Item("ActivityDesc"), ""), _
                            AltIsDBNull(drSet.Item("ActivityDescShort"), ""), _
                            AltIsDBNull(drSet.Item("TimeAndMaterials"), 0), _
                            AltIsDBNull(drSet.Item("CostPlus"), 0), _
                            AltIsDBNull(drSet.Item("FixedPrice"), 0), _
                            AltIsDBNull(drSet.Item("DueDateStatement"), ""), _
                            AltIsDBNull(drSet.Item("ReimbursementCondition"), 0), _
                            AltIsDBNull(drSet.Item("ReimbursementConditionDesc"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("TimeAndMaterialsDesc"), 0), _
                            AltIsDBNull(drSet.Item("CostPlusDesc"), 0), _
                            AltIsDBNull(drSet.Item("FixedPriceDesc"), 0), _
                            AltIsDBNull(drSet.Item("CoverTemplate"), ""), _
                            AltIsDBNull(drSet.Item("NoticeTemplate"), ""))


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

        Public Function DBGetCostFormatSpecs(ByVal costformat As String) As DataTable

            Dim dt As DataTable
            Dim ds As DataSet
            Dim strSQL As String

            Dim Params As Collection

            Try

                strSQL = "spGetCostFormatSpec"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@CostFormat").Value = costformat

                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 Then
                    Return ds.Tables(0)
                Else
                    Return Nothing
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function DBGetActivityDocs(ByVal activityID As Integer) As DataTable

            Dim dt As DataTable
            Dim ds As DataSet
            Dim strSQL As String

            Dim Params As Collection

            Try

                strSQL = "spTecFin_GetAllTechDocsForActivity"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ActivityID").Value = activityID

                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 Then
                    Return ds.Tables(0)
                Else
                    Return Nothing
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function



        Public Function DBGetAllTechDocsForActivity() As DataTable

            Dim dt As DataTable
            Dim ds As DataSet
            Dim strSQL As String

            Try

                strSQL = "spGetTecDocumentForFinancial"


                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL)

                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 Then
                    Return ds.Tables(0)
                Else
                    Return Nothing
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function



        'Updates the data table for the activity ID to tech Doc ID relationship
        Public Sub DBPutFinActivityTechDocRelationship(ByVal Activity_ID As Integer, ByVal DocID As Integer, ByVal IsSentToFinancialDoc As Boolean, _
                                                       ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)


            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Financial, Integer))) Then
                    returnVal = "You do not have rights to save a Financial Activity."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutTecFin_Activity_Doc")

                Params(0).Value = Activity_ID
                Params(1).Value = DocID
                Params(2).Value = IsSentToFinancialDoc

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutTecFin_Activity_Doc", Params)

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub


        'Clears the data from the activity/tec relationship by activity_id 
        Public Sub DBRemoveDocsFromActivityID(ByVal Activity_ID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)

            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Financial, Integer))) Then
                    returnVal = "You do not have rights to save a Financial Activity."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spClearTecFin_Activity_Doc_by_ID")

                Params(0).Value = Activity_ID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spClearTecFin_Activity_Doc_by_ID", Params)

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub



        ' Operation to send the INFO object to the repository
        Public Sub Put(ByRef oFinActInfo As MUSTER.Info.FinancialActivityInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Financial, Integer))) Then
                    returnVal = "You do not have rights to save a Financial Activity."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFinActivity")

                With oFinActInfo
                    If .ActivityID = 0 Or .ActivityID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                    Else
                        Params(0).Value = .ActivityID
                    End If
                    Params(1).Value = .ActivityDesc
                    Params(2).Value = .ActivityDescShort
                    Params(3).Value = .TimeAndMaterials
                    Params(4).Value = .CostPlus
                    Params(5).Value = .FixedPrice
                    Params(6).Value = IIf(IsNothing(.DueDateStatement), "", .DueDateStatement)
                    Params(7).Value = .ReimbursementCondition
                    Params(8).Value = .Active
                    Params(9).Value = .Deleted
                    If .ActivityID <= 0 Then
                        Params(10).Value = .CreatedBy
                    Else
                        Params(10).Value = .ModifiedBy
                    End If

                    If Params.GetUpperBound(0) >= 12 Then
                        Params(11).Value = .CoverTemplateDoc
                        Params(12).Value = .NoticeTemplateDoc
                    End If


                End With

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFinActivity", Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oFinActInfo.ActivityID Then
                    oFinActInfo.ActivityID = Params(0).Value
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