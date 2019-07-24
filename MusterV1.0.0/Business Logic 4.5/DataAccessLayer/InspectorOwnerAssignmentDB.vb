'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectorOwnerAssignmentDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      RFF/KKM     06/21/2005  Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as InspectorOwnerAssignment to build other objects.
'       Replace keyword "InspectorOwnerAssignment" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes            ' Reqd for Inserting Null Values on Dates

Namespace MUSTER.DataAccess
    Public Class InspectorOwnerAssignmentDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo(ByVal InspID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectorOwnerAssignmentsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAEInspectorOwnerAssignment"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@STAFF_ID").Value = InspID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colInspectorOwnerAssignment As New MUSTER.Info.InspectorOwnerAssignmentsCollection
                While drSet.Read
                    Dim oInspectorOwnerAssignmentInfo As New MUSTER.Info.InspectorOwnerAssignmentInfo(drSet.Item("ID"), _
                                                                            InspID, _
                                                                            drSet.Item("OWNER_ID"), _
                                                                           String.Empty, _
                                                                            CDate("01/01/0001"), _
                                                                            String.Empty, _
                                                                            CDate("01/01/0001"), _
                                                                            0)
                    oInspectorOwnerAssignmentInfo.OWNER = drSet.Item("OWNER")
                    oInspectorOwnerAssignmentInfo.Facilities = drSet.Item("FACILITIES")
                    colInspectorOwnerAssignment.Add(oInspectorOwnerAssignmentInfo)
                End While
                If Not drSet.IsClosed Then drSet.Close()
                Return colInspectorOwnerAssignment
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DBGetByID(ByVal ID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectorOwnerAssignmentInfo
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                If ID <= 0 Then
                    Return New MUSTER.Info.InspectorOwnerAssignmentInfo
                End If

                strSQL = "spGetCAEInspectorOwnerAssignment"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ID").Value = IIf(ID = 0, DBNull.Value, ID)
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                Return New MUSTER.Info.InspectorOwnerAssignmentInfo(drSet.Item("INS_OWNER_ID"), _
                                                                            drSet.Item("STAFF_ID"), _
                                                                            drSet.Item("OWNER_ID"), _
                                                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                            drSet.Item("DATE_CREATED"), _
                                                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                            drSet.Item("DELETED"))

                If Not drSet.IsClosed Then drSet.Close()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBGetOwnerFacilities() As DataSet
            Dim dtData As DataSet
            Dim strSql As String
            Try
                strSql = "select * from vCAEOwnerFacility"
                dtData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSql)
                Return dtData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Put(ByRef obj As MUSTER.Info.InspectorOwnerAssignmentInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.CAE, Integer))) Then
                    returnVal = "You do not have rights to save Inspector Owner Association."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCAEInspectorOwnerAssignment"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If obj.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = obj.ID
                End If
                Params(1).Value = IIf(obj.STAFF_ID = 0, DBNull.Value, obj.STAFF_ID)
                Params(2).Value = IIf(obj.OWNER_ID = 0, DBNull.Value, obj.OWNER_ID)
                Params(3).Value = DBNull.Value
                Params(4).Value = DBNull.Value
                Params(5).Value = DBNull.Value
                Params(6).Value = DBNull.Value
                Params(7).Value = obj.DELETED

                If obj.ID <= 0 Then
                    Params(8).Value = obj.CREATED_BY
                Else
                    Params(8).Value = obj.LAST_EDITED_BY
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If obj.ID <= 0 Then
                    obj.ID = Params(0).Value
                End If
                obj.CREATED_BY = IsNull(Params(3).Value, String.Empty)
                obj.DATE_CREATED = IsNull(Params(4).Value, CDate("01/01/0001"))
                obj.LAST_EDITED_BY = AltIsDBNull(Params(5).Value, String.Empty)
                obj.DATE_LAST_EDITED = AltIsDBNull(Params(6).Value, CDate("01/01/0001"))
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
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetAvailableOwnerFacilities(ByVal inspID As String) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCAEInspectorAvailableOwnerAssignment"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = DBNull.Value
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBgetOwnersInConflictOfManagerTerritory(Optional ByVal managerID As Integer = 0) As DataTable
            Dim dsData As DataSet
            Dim strSQL As String
            Try
                If managerID = 0 Then
                    strSQL = "select * from vCNE_Conflicted_Facilties_Owner_Territories"
                Else
                    strSQL = String.Format("exec sp_CNE_ListConflicted_Facilities_Owner_Territories_For_Manager {0}", managerID)
                End If

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                If Not dsData Is Nothing AndAlso Not dsData.Tables(0) Is Nothing Then
                    Return dsData.Tables(0)
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function DBgetAvailableInspectorsForManagement() As DataTable
            Dim dsData As DataSet
            Dim strSQL As String
            Try
                strSQL = "select * from vUnassignedInspectors"

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                If Not dsData Is Nothing AndAlso Not dsData.Tables(0) Is Nothing Then
                    Return dsData.Tables(0)
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Sub DBSetAssignmentToCNEManager(ByVal managerID As Integer, ByVal inspectorID As Integer, ByVal mode As Integer)

            Dim strSQL As String
            Dim Params() As SqlParameter

            Try
                strSQL = "sp_CNE_Update_Inspector_Manager_Assignment"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = managerID
                Params(1).Value = inspectorID
                Params(2).Value = mode

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Sub

        Public Function DBgetInspectorsUnderManager(Optional ByVal managerID As Integer = 0) As DataTable

            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter

            Try
                strSQL = "sp_CNE_GetCNEManagerInspectorList"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = managerID
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Not dsData Is Nothing AndAlso Not dsData.Tables(0) Is Nothing Then
                    Return dsData.Tables(0)
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function


        Public Function DBGetCNEmanagers() As DataTable
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "sp_CNE_Manager_Fac_Count"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = 0
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Not dsData Is Nothing AndAlso Not dsData.Tables(0) Is Nothing Then
                    Return dsData.Tables(0)
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace
