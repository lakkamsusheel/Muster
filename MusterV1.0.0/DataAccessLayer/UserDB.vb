'-------------------------------------------------------------------------------
' MUSTER.DataAccess.UserDB
'   Provides the means for marshalling User state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        PN      12/4/04    Original class definition.
'  1.1        JC      12/31/04   Removed archiving of UserInfo object - moved
'                                  the call to the pUser.Save method.
'  1.2        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        EN      02/10/05    Modified 01/01/1901  to 01/01/0001 
'  1.4        AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByName, DBGetByID
'  1.5        AB      02/17/05    Added Finally to the Try/Catch to close all datareaders
'  1.6        AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.7        AB      02/28/05    Modified Get functions based upon changes made to 
'                                     make several nullable fields non-nullable
'  1.8        MR      03/04/05    Modified GetAllInfo() to set Deleted Parameter value as DBNULL.
'  1.9        AB      03/06/05    Added DBGetPMHead
'
' Function                  Description
' GetAllInfo()        Returns an UserCollection containing all User objects in the repository.
' DBGetByName(NAME)   Returns an UserInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an UserInfo object indicated by arg ID.
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/2005 - JVC 2
' TODO - check properties and operations lists.
'

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class UserDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function GetAllInfo() As MUSTER.Info.UserCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetUser"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@User_ID").Value = DBNull.Value
                Params("@Staff_ID").Value = DBNull.Value
                Params("@Deleted").Value = DBNull.Value
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_UST_STAFF_MASTER ORDER BY USER_ID")
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colUsers As New MUSTER.Info.UserCollection
                While drSet.Read
                    Dim oUserInfo As New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                    colUsers.Add(oUserInfo)
                End While

                Return colUsers
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByName(ByVal UserId As String) As MUSTER.Info.UserInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetUser"
                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_UST_STAFF_MASTER WHERE USER_ID = '" & UserId & "'")

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Staff_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@User_ID").Value = UserId

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                Else
                    Return New MUSTER.Info.UserInfo(-1, UserId, String.Empty, String.Empty, String.Empty, 0, 0, String.Empty, False, False, False, False, False, False, False, False, False, False, String.Empty, Now, String.Empty, CDate("01/01/0001"), False)
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal ID As Int64) As MUSTER.Info.UserInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetUser"
                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_UST_STAFF_MASTER WHERE Staff_ID =" & ID.ToString)

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@User_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Staff_ID").Value = ID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                Else
                    Return New MUSTER.Info.UserInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetPMHead() As MUSTER.Info.UserInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetUserPMHead"

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                Else
                    Return New MUSTER.Info.UserInfo(-1, String.Empty, String.Empty, String.Empty, String.Empty, 0, 0, String.Empty, False, False, False, False, False, False, False, False, False, False, String.Empty, Now, String.Empty, CDate("01/01/0001"), False)
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetClosureHead() As MUSTER.Info.UserInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetUserClosureHead"

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                Else
                    Return New MUSTER.Info.UserInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetCAEHead() As MUSTER.Info.UserInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetUserCAEHead"

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                Else
                    Return New MUSTER.Info.UserInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetFEEHead() As MUSTER.Info.UserInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetUserFEEHead"

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                Else
                    Return New MUSTER.Info.UserInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetExecutiveDirector() As MUSTER.Info.UserInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetUserExecutiveDirector"

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                Else
                    Return New MUSTER.Info.UserInfo
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
        Public Function Put(ByRef oUserInf As MUSTER.Info.UserInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal OverrideRights As Boolean = False) As Integer
            Try

                If Not OverrideRights Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.User, Integer))) Then
                        returnVal = "You do not have rights to save an User."
                        Exit Function
                    Else
                        returnVal = String.Empty
                    End If
                Else
                    returnVal = String.Empty
                End If

                Dim nTempStaffId As Integer
                Dim Params(19) As SqlParameter

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSYSUSTSTAFFMASTER")
                Params(0).Value = IIFIsIntegerNull(oUserInf.UserKey, System.DBNull.Value)
                Params(1).Value = String.Empty                  ' Designation
                Params(2).Value = oUserInf.ID                   ' User_ID
                Params(3).Value = oUserInf.EncryptedPassword    ' Password
                Params(4).Value = oUserInf.Name                 ' User_Name
                Params(5).Value = oUserInf.ManagerID            ' Manager_ID
                Params(6).Value = AltIsDBNull(oUserInf.EmailAddress, String.Empty)
                Params(7).Value = AltIsDBNull(oUserInf.PhoneNumber, String.Empty)
                Params(8).Value = oUserInf.DefaultModule        ' Default Module
                Params(9).Value = oUserInf.Deleted              ' Deleted Flag
                Params(10).Value = oUserInf.HEAD_PM
                Params(11).Value = oUserInf.HEAD_CLOSURE
                Params(12).Value = oUserInf.HEAD_REGISTRATION
                Params(13).Value = oUserInf.HEAD_INSPECTION
                Params(14).Value = oUserInf.HEAD_CANDE
                Params(15).Value = oUserInf.HEAD_FEES
                Params(16).Value = oUserInf.HEAD_FINANCIAL
                Params(17).Value = oUserInf.HEAD_ADMIN
                Params(18).Value = oUserInf.Active
                Params(0).Direction = ParameterDirection.InputOutput

                If oUserInf.UserKey <= 0 Then
                    Params(19).Value = oUserInf.CreatedBy
                Else
                    Params(19).Value = oUserInf.ModifiedBy
                End If
                Params(20).Value = oUserInf.EXECUTIVE_DIRECTOR

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSYSUSTSTAFFMASTER", Params)
                If oUserInf.UserKey <= 0 Then
                    nTempStaffId = Params(0).Value
                    oUserInf.UserKey = Params(0).Value
                Else
                    nTempStaffId = 0
                End If
                Return nTempStaffId
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function ModuleHeadCheck(ByVal HeadField As String, ByVal UserID As String) As Boolean
            Dim strSQL As String
            Dim dsset As New DataSet
            strSQL = "SELECT * FROM tblSYS_UST_STAFF_MASTER Where User_ID <> '" & UserID & "' And " & HeadField & "=1 and deleted = 0"
            Try
                dsset = DBGetDS(strSQL)
                If dsset.Tables(0).Rows.Count > 0 Then
                    Return True
                Else
                    Return False
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function DBGetUserGroupRel(ByVal staffID As Int64, Optional ByVal groupID As Int64 = 0, Optional ByVal showActiveInactiveGroups As Boolean = True, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.UserGroupRelationsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetUserGroupRel"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@STAFF_ID").Value = staffID
                Params("@GROUP_ID").Value = IIf(groupID <= 0, DBNull.Value, groupID)
                Params("@ACTIVE").Value = showActiveInactiveGroups
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colUserGroupRel As New MUSTER.Info.UserGroupRelationsCollection
                While drSet.Read
                    Dim oUserGroupRelInfo As New MUSTER.Info.UserGroupRelationInfo(drSet.Item("STAFF_ID"), _
                                                                        drSet.Item("GROUP_ID"), _
                                                                        drSet.Item("INACTIVE"), _
                                                                        drSet.Item("DELETED"), _
                                                                        drSet.Item("GROUP"))
                    colUserGroupRel.Add(oUserGroupRelInfo)
                End While
                Return colUserGroupRel
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Sub PutUserGroupRel(ByRef userGroupRelInfo As MUSTER.Info.UserGroupRelationInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutUserGroupRel"
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.User, Integer))) Then
                    returnVal = "You do not have rights to save User Group."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = userGroupRelInfo.StaffID
                Params(1).Value = userGroupRelInfo.GroupID
                Params(2).Value = userGroupRelInfo.Deleted

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function DBGetManagedUsers(ByVal staffID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.UserCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetManagedUsers"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@STAFF_ID").Value = staffID
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_USER_GROUPS")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colManagedUsers As New MUSTER.Info.UserCollection
                While drSet.Read
                    Dim oManagedUserInfo As New MUSTER.Info.UserInfo(drSet.Item("STAFF_ID"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("USER_NAME"), _
                                                          drSet.Item("EMAIL_ADDRESS"), _
                                                          drSet.Item("PHONE_NUMBER"), _
                                                          drSet.Item("DEFAULT_MODULE"), _
                                                          drSet.Item("MANAGER_ID"), _
                                                          drSet.Item("PASSWORD"), _
                                                          drSet.Item("DELETED"), _
                                                          AltIsDBNull(drSet.Item("HEAD_PM"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CLOSURE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_REGISTRATION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_INSPECTION"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_CANDE"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FEES"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_FINANCIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("HEAD_ADMIN"), 0), _
                                                          AltIsDBNull(drSet.Item("ACTIVE"), 0), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          drSet.Item("LAST_EDITED_BY"), _
                                                          drSet.Item("DATE_LAST_EDITED"), _
                                                          AltIsDBNull(drSet.Item("EXECUTIVE_DIRECTOR"), 0))
                    colManagedUsers.Add(oManagedUserInfo)
                End While
                Return colManagedUsers
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally

                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Sub PutModuleEntityRel(ByVal moduleID As Integer, ByVal entityType As Integer, ByVal deleted As Boolean)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutModuleEntityRel"
            Try
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = moduleID
                Params(1).Value = entityType
                Params(2).Value = deleted

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function HasAccess(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal EntityTypeID As Integer) As Boolean
            Try
                Return SqlHelper.HasWriteAccess(moduleID, staffID, EntityTypeID)

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Sub SyncDB(ByVal strSQL As String)
            Try
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.Text, strSQL)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
    End Class
End Namespace
