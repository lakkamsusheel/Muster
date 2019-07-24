'-------------------------------------------------------------------------------
' MUSTER.DataAccess.FavSearchDB
'   Provides the means for marshalling FavSearchChildInfo & FavSearchParentInfo state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR       12/5/04     Original class definition.
'  1.1        AN       12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        MR        1/7/04     Added DBGetByParentID and DBGetByChildID functions
'  1.3        JVC2     01/21/05    Altered PUTs to update ID in place, rather than passing
'                                    ID back to caller for replacement.
'                                    Changed GetAllChildInfo to order results by CRITERION_ORDER
'                                    Modified DBGetByParentID to include PUBLIC_FLAG in definition.
'                                    Modified GetAllParentInfo to include PUBLIC_FLAG in definition.
'  1.4       AB      02/08/05      Replaced dynamic SQL with stored procedures in the following
'                                     Functions:  GetAllParentInfo, GetAllChildInfo, DBGetByParentID, DBGetByChildID
'                                     Modified DBGetByParentID and DBGetByChildID to use Reader as well
'  1.5       AB      02/15/05      Added Finally to the Try/Catch to close all datareaders
'  1.6       AB      02/16/05      Removed any IsNull calls for fields the DB requires
'  1.7       AB      02/18/05      Set all parameters for SP, that are not required, to NULL
'
' Operations
' Function                  Description
' GetAllParentInfo([strUserId]) Returns a FavSearchParentCollection containing all FavSearchParentInfo objects in the repository.
'
' DBGetByParentID(nSearchID)    Returns FavSearchParentInfo for the given SearchID
'
' DBGetByChildID(nCriteriaID)    Returns FavSearchChildInfo for the given SearchID
'
' GetAllChildInfo([strUserId]) Returns a FavSearchChildCollection containing all FavSearchChildInfo objects in the repository.
'
' Put(FavSearchChildInfo)          Updates the repository with the information supplied in FavSearchChildInfo.  Inserts the
'                                    data if no matching FavSearchChildInfo is in the repository.
' Put(FavSearchParentInfo)         Updates the repository with the information supplied in FavSearchParentInfo.  Inserts the
'                                    data if no matching FavSearchParentInfo is in the repository.
'-------------------------------------------------------------------------------
'
'  TODO - JVC2 - Integrate into application 1/21/05
'  TODO - JVC2 - Copy spPutSearch - added CriterionOrder to the arg list
'
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class FavSearchDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function GetAllParentInfo(Optional ByVal strUserID As String = Nothing) As Muster.Info.FavSearchParentCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblSYS_SEARCH WHERE SEARCH_USER = '" + strUserID + "' Order by SEARCH_ID"
            Try
                strSQL = "spGetSYSSearch"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Search_User").Value = strUserID
                Params("@Search_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colParent As New MUSTER.Info.FavSearchParentCollection
                While drSet.Read
                    Dim oParentInfo As New MUSTER.Info.FavSearchParentInfo(drSet.Item("SEARCH_ID"), _
                                                                           drSet.Item("SEARCH_USER"), _
                                                                           drSet.Item("SEARCH_TYPE"), _
                                                                           drSet.Item("SEARCH_NAME"), _
                                                                           AltIsDBNull(drSet.Item("LUST_STATUS"), "UNKNOWN"), _
                                                                           AltIsDBNull(drSet.Item("TANK_STATUS"), "UNKNOWN"), _
                                                                           drSet.Item("PUBLIC_FLAG"))
                    colParent.Add(oParentInfo)
                End While
                Return colParent
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function GetAllChildInfo(Optional ByVal strUserID As String = Nothing) As Muster.Info.FavSearchChildCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblSYS_SEARCH_CRITERIA WHERE SEARCH_ID IN (SELECT SEARCH_ID FROM tblSYS_SEARCH WHERE SEARCH_USER = '" + strUserID + "') Order by CRITERION_ORDER"
            Try
                strSQL = "spGetSYSSearch_Criteria"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Search_User").Value = strUserID
                Params("@Search_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colChild As New MUSTER.Info.FavSearchChildCollection
                While drSet.Read
                    Dim oChildInfo As New MUSTER.Info.FavSearchChildInfo(drSet.Item("CRITERION_ID"), _
                                                                         AltIsDBNull(drSet.Item("CRITERION_ORDER"), 0), _
                                                                         drSet.Item("CRITERION_NAME"), _
                                                                         drSet.Item("CRITERION_VALUE"), _
                                                                         drSet.Item("CRITERION_DATA_TYPE"), _
                                                                         drSet.Item("SEARCH_ID"))

                    colChild.Add(oChildInfo)
                End While
                Return colChild
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByParentID(ByVal nSearchID As Int64) As Muster.Info.FavSearchParentInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Dim oParentInfo As MUSTER.Info.FavSearchParentInfo

            strSQL = "spGetSYSSearch"

            Try

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Search_ID").Value = nSearchID
                Params("@Search_User").Value = DBNull.Value
                Params("@OrderBy").Value = 1

                If nSearchID = 0 Then
                    Return New MUSTER.Info.FavSearchParentInfo
                End If
                'drSet = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, "SELECT * FROM tblSYS_SEARCH WHERE SEARCH_ID  = " & nSearchID.ToString)
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    oParentInfo = New MUSTER.Info.FavSearchParentInfo(drSet.Item("SEARCH_ID"), _
                                                                           drSet.Item("SEARCH_USER"), _
                                                                           drSet.Item("SEARCH_TYPE"), _
                                                                           drSet.Item("SEARCH_NAME"), _
                                                                           AltIsDBNull(drSet.Item("LUST_STATUS"), "UNKNOWN"), _
                                                                           AltIsDBNull(drSet.Item("TANK_STATUS"), "UNKNOWN"), _
                                                                           drSet.Item("PUBLIC_FLAG"))
                End While
                Return oParentInfo
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByChildID(ByVal nCriteria As Int64) As Muster.Info.FavSearchChildInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            strSQL = "spGetSYSSearch"

            Try

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Search_ID").Value = nCriteria
                Params("@Search_User").Value = DBNull.Value
                Params("@OrderBy").Value = 1


                If nCriteria = 0 Then
                    Return New MUSTER.Info.FavSearchChildInfo
                End If
                'drSet = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, "SELECT * FROM tblSYS_SEARCH WHERE SEARCH_ID  = " & nCriteria.ToString)
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim oChildInfo As MUSTER.Info.FavSearchChildInfo
                If drSet.HasRows Then
                    drSet.Read()
                    oChildInfo = New MUSTER.Info.FavSearchChildInfo(drSet.Item("CRITERION_ID"), _
                                                                         drSet.Item("CRITERION_ID"), _
                                                                         drSet.Item("CRITERION_NAME"), _
                                                                         drSet.Item("CRITERION_VALUE"), _
                                                                         drSet.Item("CRITERION_DATA_TYPE"), _
                                                                         drSet.Item("SEARCH_ID"))
                    Return oChildInfo
                Else
                    Return New MUSTER.Info.FavSearchChildInfo
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Sub Put(ByRef oParentInfo As Muster.Info.FavSearchParentInfo)
            Try
                Dim ParentParams() As SqlParameter
                ParentParams = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSearch")
                ParentParams(0).Direction = ParameterDirection.InputOutput
                If oParentInfo.ID <= 0 Then
                    ParentParams(0).Value = 0
                Else
                    ParentParams(0).Value = CInt(oParentInfo.ID)
                End If

                ParentParams(1).Value = oParentInfo.Name
                ParentParams(2).Value = oParentInfo.SearchType
                ParentParams(3).Value = oParentInfo.User
                ParentParams(4).Value = oParentInfo.IsPublic
                ParentParams(5).Value = oParentInfo.Deleted
                ParentParams(6).Value = oParentInfo.LustStatus
                ParentParams(7).Value = oParentInfo.TankStatus

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSearch", ParentParams)

                If ParentParams(0).Value <> 0 Then
                    oParentInfo.ID = ParentParams(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Put(ByRef oChildInfo As Muster.Info.FavSearchChildInfo)
            Try
                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutSearchCriterion")
                Params(0).Direction = ParameterDirection.InputOutput
                If oChildInfo.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = CInt(oChildInfo.ID)
                End If
                Params(1).Value = oChildInfo.Order
                Params(2).Value = oChildInfo.ParentID
                Params(3).Value = oChildInfo.CriterionName
                Params(4).Value = oChildInfo.CriterionValue
                Params(5).Value = oChildInfo.CriterionDataType
                Params(6).Value = oChildInfo.Deleted

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutSearchCriterion", Params)
                If Params(0).Value <> 0 Then
                    oChildInfo.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
    End Class
End Namespace
