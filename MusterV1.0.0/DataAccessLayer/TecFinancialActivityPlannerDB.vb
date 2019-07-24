' -------------------------------------------------------------------------------
' MUSTER.DataAccess.RecFinancialActivityPlannerDB
' Provides the means for marshalling Financial/Technical Activity planning state to/from the repository
' 
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0    Thomas Franey  05/06/2009   Original class definition
' 
' 
' Function                  Description
' -------------------------------------------------------------------------------    
' 
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess

    Public Class TecFinancialActivityPlannerDB
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
        Public Function DBGetByID(ByVal nVal As Int64, ByVal nActivity As Int64) As MUSTER.Info.TecFinancialActivityPlannerInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strAct As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.TecFinancialActivityPlannerInfo
                End If
                strSQL = "spGetTecFinActivityPlanner"
                strVal = nVal
                strAct = nActivity

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@EventID").Value = nVal
                Params("@ActivityID").Value = nActivity

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()

                    Return New MUSTER.Info.TecFinancialActivityPlannerInfo(drSet.Item("EventID"), _
                            drSet.Item("ActivityID"), _
                            AltIsDBNull(drSet.Item("Duration"), 0), _
                            AltIsDBNull(drSet.Item("Cost"), 0.0))

                Else

                    Return New MUSTER.Info.TecFinancialActivityPlannerInfo
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
        Public Function DBGetAll() As MUSTER.Info.TecFinancialActivityPlannerCollection
            Dim coll As Info.TecFinancialActivityPlannerCollection
            Dim drSet As DataSet
            Dim strVal As String
            Dim strAct As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetTecFinActivityPlanner"
                strVal = "0"
                strAct = "-1"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@EventID").Value = 0
                Params("@ActivityID").Value = -1

                drSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Not drSet Is Nothing AndAlso drSet.Tables(0).Rows.Count > 0 Then

                    coll = New Info.TecFinancialActivityPlannerCollection

                    For Each rec As DataRow In drSet.Tables(0).Rows


                        coll.Add(New MUSTER.Info.TecFinancialActivityPlannerInfo(rec.Item("EventID"), _
                                rec.Item("ActivityID"), _
                                AltIsDBNull(rec.Item("Duration"), 0), _
                                AltIsDBNull(rec.Item("Cost"), 0.0)))

                    Next

                    Return coll
                Else

                        Return New MUSTER.Info.TecFinancialActivityPlannerCollection
                End If

            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    drSet.Dispose()
                End If
            End Try
        End Function

        ' Operation to send the INFO object to the repository
        Public Sub Put(ByRef oTecFinActPlannerInfo As MUSTER.Info.TecFinancialActivityPlannerInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(IIf(moduleID = 616, SqlHelper.EntityTypes.Financial, SqlHelper.EntityTypes.TechnicalActivity), Integer))) Then
                    returnVal = "You do not have rights to save a Financial Activity or a Technical Activity."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutTecFinActivityPlanner")

                With oTecFinActPlannerInfo
                    If .ActivityTypeID = -1 Or .EventID = 0 Or .EventID = -1 Then
                        Params(0).Value = 0 'System.DBNull.Value
                        Params(1).Value = 0
                    Else
                        Params(0).Value = .EventID
                        Params(1).Value = .ActivityTypeID
                    End If

                    Params(2).Value = .Duration
                    Params(3).Value = .Cost

                End With

                If Params(0).Value <> 0 Then
                    SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutTecFinActivityPlanner", Params)
                Else
                    returnVal = "A Technical Event and Activity Type needs to be defined before saving a record."
                    Exit Sub

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