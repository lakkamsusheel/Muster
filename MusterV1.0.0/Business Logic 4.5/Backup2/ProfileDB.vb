'-------------------------------------------------------------------------------
' MUSTER.DataAccess.ReportDB
'   Provides the means for marshalling ProfileInfo state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/23/04    Original class definition.
'  1.1        JVC2      11/29/04    Altered DBGetByKey to support retrieval of groups of 
'                                    ProfileInfo objects in the event that a partial key
'                                    is supplied.
'  1.2        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        EN        02/10/05    Modified 01/01/1901  to 01/01/0001 
'  1.4        AB        02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByKey
'  1.5        AB        02/16/05    Added Finally to the Try/Catch to close all datareaders
'  1.6        AB        02/16/05    Removed any IsNull calls for fields the DB requires
'  1.7        MR        02/17/05    Altered to pass Profile Key parameter value and DBNULL for Optional Parameters in DBGetByKey
'  
' Operations
' Function                  Description
' GetAllInfo([showDeleted]) Returns a ProfileCollection containing all ProfileInfo objects in the repository.
' DBGetByKey(User, Key, [Mod1], [Mod2], [Deleted])
'                           Returns a ProfileCollection containing all keys which match the partial (or full)
'                            key supplied.
' DBGetByKey(strArray(), [showDeleted])
'                           Returns a ProfileCollection containing all keys which match the partial (or full)
'                            key supplied in the array of string.
' DBGetDS(strSQL)           Returns a dataset containing the results of the select query supplied in strSQL.
'
' Put(ProfileInfo)          Updates the repository with the information supplied in ProfileInfo.  Inserts the
'                            data if no matching ProfileInfo is in the repository.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class ProfileDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo(Optional ByVal showDeleted As Boolean = False) As Muster.Info.ProfileCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                'strSQL = "SELECT * FROM tblSYS_PROFILE_INFO "
                'strSQL += IIf(Not showDeleted, " WHERE DELETED <> 1 ", "")
                'strSQL += "ORDER BY USER_ID, PROFILE_KEY"

                strSQL = "spGetSYSProfile"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colProfiles As New Muster.Info.ProfileCollection
                While drSet.Read
                    Dim oProfileInfo As New MUSTER.Info.ProfileInfo(drSet.Item("USER_ID"), _
                                                          drSet.Item("PROFILE_KEY"), _
                                                          drSet.Item("PROFILE_MODIFIER_1"), _
                                                          drSet.Item("PROFILE_MODIFIER_2"), _
                                                          drSet.Item("PROFILE_VALUE"), _
                                                          drSet.Item("DELETED"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                    colProfiles.Add(oProfileInfo)
                End While
                Return colProfiles
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByKey(ByVal strUser As String, ByVal strKey As String, Optional ByVal strMod1 As String = "", Optional ByVal strMod2 As String = "", Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ProfileCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim oProfInf As MUSTER.Info.ProfileInfo
            Dim oProfCol As New MUSTER.Info.ProfileCollection
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblSYS_PROFILE_INFO WHERE USER_ID = '" & strUser & "' AND PROFILE_KEY = '" & strKey & "' "
            'strSQL += IIf(strMod1 <> String.Empty, " AND PROFILE_MODIFIER_1 = '" & strMod1 & "' ", "")
            'strSQL += IIf(strMod2 <> String.Empty, " AND PROFILE_MODIFIER_2 = '" & strMod2 & "' ", "")
            'strSQL += IIf(Not showDeleted, " AND DELETED <> 1", "")

            Try
                strSQL = "spGetSYSProfile"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@User_ID").Value = strUser
                Params("@Profile_Key").value = strKey
                Params("@Profile_Modifier1").Value = IIf(strMod1 <> String.Empty, strMod1, DBNull.Value)
                Params("@Profile_Modifier2").Value = IIf(strMod2 <> String.Empty, strMod2, DBNull.Value)
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        oProfInf = New MUSTER.Info.ProfileInfo(drSet.Item("USER_ID"), _
                                                          drSet.Item("PROFILE_KEY"), _
                                                          drSet.Item("PROFILE_MODIFIER_1"), _
                                                          drSet.Item("PROFILE_MODIFIER_2"), _
                                                          drSet.Item("PROFILE_VALUE"), _
                                                          drSet.Item("DELETED"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                        oProfCol.Add(oProfInf)
                    End While
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
            Return oProfCol
        End Function
        Public Function DBGetByKey(ByVal strArray() As String, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ProfileCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim oProfInf As MUSTER.Info.ProfileInfo
            Dim oProfCol As New MUSTER.Info.ProfileCollection
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblSYS_PROFILE_INFO WHERE USER_ID = '" & strArray(0) & "' AND PROFILE_KEY = '" & strArray(1) & "' "
            'If strArray.Length >= 3 Then
            '    If strArray(2) <> String.Empty Then
            '        strSQL += " AND PROFILE_MODIFIER_1 = '" & strArray(2) & "' "
            '    End If
            'End If
            'If strArray.Length >= 4 Then
            '    If strArray(3) <> String.Empty Then
            '        strSQL += " AND PROFILE_MODIFIER_2 = '" & strArray(3) & "' "
            '    End If
            'End If
            'strSQL += IIf(Not showDeleted, " AND DELETED <> 1", "")

            Try

                strSQL = "spGetSYSProfile"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@User_ID").Value = strArray(0)
                Params("@Profile_Key").value = strArray(1)
                Params("@Profile_Modifier1").Value = DBNull.Value
                Params("@Profile_Modifier2").Value = DBNull.Value
                If strArray.Length >= 3 Then
                    If strArray(2) <> String.Empty Then
                        Params("@Profile_Modifier1").Value = strArray(2)
                    End If
                End If
                If strArray.Length >= 4 Then
                    If strArray(3) <> String.Empty Then
                        Params("@Profile_Modifier2").Value = strArray(3)
                    End If
                End If
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read()
                        oProfInf = New MUSTER.Info.ProfileInfo(drSet.Item("USER_ID"), _
                                                          drSet.Item("PROFILE_KEY"), _
                                                          drSet.Item("PROFILE_MODIFIER_1"), _
                                                          drSet.Item("PROFILE_MODIFIER_2"), _
                                                          drSet.Item("PROFILE_VALUE"), _
                                                          drSet.Item("DELETED"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                        oProfCol.Add(oProfInf)
                    End While
                End If
            Catch ex As Exception
                If Not drSet.IsClosed Then drSet.Close()
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
            Return oProfCol
        End Function

        '
        ' Note - there is no DBGetByName since profile datum has no "ID" per se
        '
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

        Public Sub Put(ByRef oProfInf As MUSTER.Info.ProfileInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal OverrideRights As Boolean = False)
            Try

                If Not OverrideRights Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Profile, Integer))) Then
                        returnVal = "You do not have rights to save Profile."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                Else
                    returnVal = String.Empty
                End If

                Dim Params(7) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSYSPROFILEINFO")
                Params(0).Value = oProfInf.User
                Params(1).Value = oProfInf.ProfileKey
                Params(2).Value = oProfInf.ProfileMod1
                Params(3).Value = oProfInf.ProfileMod2
                Params(4).Value = oProfInf.ProfileValue
                Params(5).Value = oProfInf.Deleted
                If oProfInf.User = "SYSTEM" Then
                    Params(6).Value = oProfInf.CreatedBy
                Else
                    Params(6).Value = oProfInf.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSYSPROFILEINFO", Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
    End Class
End Namespace


