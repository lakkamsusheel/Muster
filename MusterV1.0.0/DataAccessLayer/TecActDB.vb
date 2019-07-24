'-------------------------------------------------------------------------------
' MUSTER.DataAccess.TecActDB
'   Provides the means for marshalling Technical Activity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials        Date        Description
'  1.0        JVC         05/31/05    Original class definition.
'  1.1        JvC         07/28/05    Added InUse to detect association of activity
'  1.2    Thomas Franey   06/10/09    Added Activity Cost Planning 
'                                       to existing LUST events.
' Function                  Description
' DBGetByID(ID)         Returns a Technical Activity Object indicated by Lust Remediation ID
' DBGetByEventID(ID)    Returns an Technical Activity Collection indicated by Lust Event ID
' DBGetDS(SQL)          Returns a resultant Dataset by running query specified by the string arg SQL
' Put(oTecDoc)        Saves the Technical Activity passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'

Imports System.Data.SqlClient
Imports MUSTER.Info.TecActInfo

Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class TecActDB
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

        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.TecActInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim oTecActInfo As MUSTER.Info.TecActInfo
            Dim oTecDocDB As New MUSTER.DataAccess.TecDocDB

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.TecActInfo
                End If
                strSQL = "spGetTecActivities"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ACTIVITY_ID").Value = nVal
                Params("@ACTIVITY_NAME").Value = ""

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    oTecActInfo = New MUSTER.Info.TecActInfo(drSet.Item("Activity_ID"), _
                            AltIsDBNull(drSet.Item("Activity_Name"), String.Empty), _
                            AltIsDBNull(drSet.Item("Action_days"), 0), _
                            AltIsDBNull(drSet.Item("Warn_Days"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("MODIFIED_BY"), ""), _
                            AltIsDBNull(drSet.Item("MODIFIED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("CostMode"), ActivityCostModeEnum.NotFundable), _
                            AltIsDBNull(drSet.Item("EstimatedCost"), 0.0) _
                    )

                    oTecActInfo.DocumentsCollection = oTecDocDB.GetByActivity(oTecActInfo.ID)
                    Return oTecActInfo
                Else

                    Return New MUSTER.Info.TecActInfo
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
        Public Function DBGetByName(ByVal strVal As String) As MUSTER.Info.TecActInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Dim oTecActInfo As MUSTER.Info.TecActInfo
            Dim oTecDocDB As New MUSTER.DataAccess.TecDocDB
            Try
                If strVal = "" Then
                    Return New MUSTER.Info.TecActInfo
                End If
                strSQL = "spGetTecActivities"
                strVal = strVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ACTIVITY_NAME").Value = strVal
                Params("@ACTIVITY_ID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    oTecActInfo = New MUSTER.Info.TecActInfo(drSet.Item("Activity_ID"), _
                            AltIsDBNull(drSet.Item("Activity_Name"), String.Empty), _
                            AltIsDBNull(drSet.Item("Action_days"), 0), _
                            AltIsDBNull(drSet.Item("Warn_Days"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("MODIFIED_BY"), ""), _
                            AltIsDBNull(drSet.Item("MODIFIED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("CostMode"), ActivityCostModeEnum.NotFundable), _
                            AltIsDBNull(drSet.Item("EstimatedCost"), 0.0) _
                    )

                    oTecActInfo.DocumentsCollection = oTecDocDB.GetByActivity(oTecActInfo.ID)
                    Return oTecActInfo
                Else

                    Return New MUSTER.Info.TecActInfo
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
        Public Function GetAllInfo() As MUSTER.Info.TecActCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim oTecDocDB As New MUSTER.DataAccess.TecDocDB

            Try
                strSQL = "spGetTecActivities"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ACTIVITY_ID").Value = 0
                Params("@ACTIVITY_NAME").Value = ""

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.TecActCollection
                While drSet.Read
                    Dim oTecActInfo As New MUSTER.Info.TecActInfo(drSet.Item("Activity_ID"), _
                            AltIsDBNull(drSet.Item("Activity_Name"), String.Empty), _
                            AltIsDBNull(drSet.Item("Action_days"), 0), _
                            AltIsDBNull(drSet.Item("Warn_Days"), 0), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("MODIFIED_BY"), ""), _
                            AltIsDBNull(drSet.Item("MODIFIED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("Active"), 1), _
                            AltIsDBNull(drSet.Item("Deleted"), 0), _
                            AltIsDBNull(drSet.Item("CostMode"), oTecActInfo.ActivityCostModeEnum.NotFundable), _
                            AltIsDBNull(drSet.Item("EstimatedCost"), 0.0) _
                    )

                    oTecActInfo.DocumentsCollection = oTecDocDB.GetByActivity(oTecActInfo.ID)
                    colEntities.Add(oTecActInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Sub Put(ByRef oTecActInfo As MUSTER.Info.TecActInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strSQL As String
            Dim Params() As SqlParameter
            Dim PurgeParams() As SqlParameter
            Dim DocParams() As SqlParameter
            Dim oTecDocInfo As MUSTER.Info.TecDocInfo

            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.TechnicalActivity, Integer))) AndAlso _
                    Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Financial, Integer))) Then
                    returnVal = "You do not have rights to save a Technical Activity."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutTecActivity"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                With oTecActInfo
                    If .ID <= 0 Then
                        Params(0).Value = System.DBNull.Value
                    Else
                        Params(0).Value = .ID
                    End If
                    Params(1).Value = .Name
                    Params(2).Value = .ActDays
                    Params(3).Value = .WarnDays
                    Params(4).Value = .Active
                    Params(5).Value = .Deleted
                    If .ID <= 0 Then
                        Params(6).Value = .CreatedBy
                    Else
                        Params(6).Value = .ModifiedBy
                    End If

                    Params(7).Value = .CostMode

                    Params(8).SqlDbType = SqlDbType.Money
                    Params(8).Value = .Cost

                End With


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

                'Perform check for New ID and assign, if necessary
                If Params(0).Value <> oTecActInfo.ID Then
                    oTecActInfo.ID = Params(0).Value
                End If

                'Purge all ActivityDocuments for this Activity
                strSQL = "spPurgeTecActivity_Documents"
                PurgeParams = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                PurgeParams(0).Value = oTecActInfo.ID
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, PurgeParams)

                'Add ActivityDocuments for this Activity
                strSQL = "spPutTecActivity_Documents"
                DocParams = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                DocParams(0).Value = oTecActInfo.ID
                For Each oTecDocInfo In oTecActInfo.DocumentsCollection.Values
                    DocParams(1).Value = oTecDocInfo.ID
                    SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, DocParams)
                Next

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function InUse(ByRef oTecActInfo As MUSTER.Info.TecActInfo) As Boolean
            Dim strSQL As String
            Dim Params As Collection
            Dim ActCount As Int32

            Try
                strSQL = "SELECT COUNT(*) FROM vLUSTEVENTACTIVITIES_STATES WHERE DELETED = 0 AND ACTIVITY_TYPE_ID = " & oTecActInfo.ID
                ActCount = SqlHelper.ExecuteScalar(_strConn, CommandType.Text, strSQL)
                Return IIf(ActCount > 0, True, False)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
                '
                '  Couldn't determine, so say it's in use
                '
                Return True
            End Try

        End Function
#End Region
    End Class
End Namespace
