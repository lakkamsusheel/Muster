'-------------------------------------------------------------------------------
' MUSTER.DataAccess.UserGroupDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        AN      11/29/04    Original class definition.
'  1.1        JC      12/28/04    Changed references to Deleted so they are not inverted
'                                 Changed Put to update GroupID if newly added
'  1.2        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        EN      02/10/05     Modified 01/01/1901  to 01/01/0001 
'  1.4        AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByName, DBGetByID
'  1.5        AB      02/17/05    Added Finally to the Try/Catch to close all datareaders
'  1.6        AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.7        AN      06/14/05    Added Active Flag
'
'
' Function                  Description
' GetAllInfo()        Returns an UserGroupCollection containing all UserGroup objects in the repository.
' DBGetByName(NAME)   Returns an UserGroupInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an UserGroupInfo object indicated by arg ID.
'-------------------------------------------------------------------------------
'
' TODO - Add to application 1/03 JVC2
' TODO - check properties and operations against list.
'
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class UserGroupDB
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
        Public Function GetAllInfo() As MUSTER.Info.UserGroupCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetUserGroup"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Group_Name").Value = DBNull.Value
                Params("@Group_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colUserGroups As New MUSTER.Info.UserGroupCollection
                While drSet.Read
                    Dim oUserGroupInfo As New MUSTER.Info.UserGroupInfo(False, _
                                                                        drSet.Item("GROUP_ID"), _
                                                                        drSet.Item("GROUP_NAME"), _
                                                                        drSet.Item("GROUP_DESCRIPTION"), _
                                                                        False, _
                                                                        drSet.Item("Deleted"), _
                                                                        drSet.Item("Active"), _
                                                                        drSet.Item("CreatedBy"), _
                                                                        drSet.Item("CreatedOn"), _
                                                                        AltIsDBNull(drSet.Item("ModifiedBy"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("ModifiedOn"), CDate("01/01/0001")))



                    colUserGroups.Add(oUserGroupInfo)
                End While

                Return colUserGroups
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByName(ByVal strVal As String) As MUSTER.Info.UserGroupInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetUserGroup"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Group_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Group_Name").Value = strVal

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS WHERE GROUP_NAME = '" & strVal & "'")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserGroupInfo(False, _
                                                            drSet.Item("GROUP_ID"), _
                                                            drSet.Item("GROUP_NAME"), _
                                                            drSet.Item("GROUP_DESCRIPTION"), _
                                                            False, _
                                                            drSet.Item("Deleted"), _
                                                            drSet.Item("Active"), _
                                                            drSet.Item("CreatedBy"), _
                                                            drSet.Item("CreatedOn"), _
                                                            AltIsDBNull(drSet.Item("ModifiedBy"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("ModifiedOn"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.UserGroupInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.UserGroupInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetUserGroup"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Group_Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Group_ID").Value = nVal

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS WHERE GROUP_ID = '" & nVal & "'")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserGroupInfo(False, _
                                                            drSet.Item("GROUP_ID"), _
                                                            drSet.Item("GROUP_NAME"), _
                                                            drSet.Item("GROUP_DESCRIPTION"), _
                                                            False, _
                                                            drSet.Item("Deleted"), _
                                                            drSet.Item("Active"), _
                                                            drSet.Item("CreatedBy"), _
                                                            drSet.Item("CreatedOn"), _
                                                            AltIsDBNull(drSet.Item("ModifiedBy"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("ModifiedOn"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.UserGroupInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
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

        Public Sub Put(ByRef oUserGroupInfo As MUSTER.Info.UserGroupInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.User, Integer))) Then
                    returnVal = "You do not have rights to save User Group."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(6) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutUserGroups")
                Params(0).Value = oUserGroupInfo.ID
                Params(1).Value = oUserGroupInfo.Name
                Params(2).Value = oUserGroupInfo.Description
                Params(3).Value = oUserGroupInfo.Deleted
                Params(4).Value = oUserGroupInfo.Active
                Params(5).Value = 0

                If oUserGroupInfo.ID <= 0 Then
                    Params(6).Value = oUserGroupInfo.CreatedBy
                Else
                    Params(6).Value = oUserGroupInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutUserGroups", Params)
                If Params(5).Value <> 0 Then
                    oUserGroupInfo.ID = Params(5).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function DBGetGroupModuleRel(ByVal groupID As Integer, Optional ByVal moduleID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.GroupModuleRelationsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetGroupModuleRel"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@GROUP_ID").Value = groupID
                Params("@MODULE_ID").Value = IIf(moduleID <= 0, DBNull.Value, moduleID)
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colGroupModuleRel As New MUSTER.Info.GroupModuleRelationsCollection
                While drSet.Read
                    Dim oGroupModuleRelInfo As New MUSTER.Info.GroupModuleRelationInfo(drSet.Item("GROUP_ID"), _
                                                                        drSet.Item("MODULE_ID"), _
                                                                        drSet.Item("WRITE_ACCESS"), _
                                                                        drSet.Item("READ_ACCESS"), _
                                                                        drSet.Item("Deleted"), _
                                                                        drSet.Item("MODULENAME"))
                    colGroupModuleRel.Add(oGroupModuleRelInfo)
                End While
                Return colGroupModuleRel
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Sub PutGroupModuleRel(ByRef groupModuleRelInfo As MUSTER.Info.GroupModuleRelationInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutGroupModuleRel"
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.User, Integer))) Then
                    returnVal = "You do not have rights to save User Group."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = groupModuleRelInfo.GroupID
                Params(1).Value = groupModuleRelInfo.ModuleID
                Params(2).Value = groupModuleRelInfo.WriteAccess
                Params(3).Value = groupModuleRelInfo.READACCESS
                Params(4).Value = groupModuleRelInfo.Deleted

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
    End Class
End Namespace

