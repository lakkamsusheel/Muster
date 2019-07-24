'-------------------------------------------------------------------------------
' MUSTER.DataAccess.CalendarDB
'   Provides the means for marshalling Calendar to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/10/04    Original class definition.
'  1.1        JVC2    01/17/05    Added DBGetByOtherID() which returns a collection of 
'                                   CalendarInfo objects.
'
'  1.2        MR       1/27/05    Modified GetAllInfo to select the Associated Group Users Records.
'                                 Removed Redundant PutCalendar Function.
'  1.3        EN      02/10/05    Modified 01/01/1901 to 01/01/0001 
'  1.4        AB      02/14/05    Changed dynamic SQL statement to a parameterized stored procedure in GetAllInfo(), 
'                                 DBGetByID() and DBGetByOtherID()
'  1.5        AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.6        AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.7        AB      02/18/05    Set all parameters for SP, that are not required, to NULL 
'  1.8        JC      02/21/05    Added Calendar_Description to DBGetByOtherID
'  1.9        JVC2    03/25/05    Added OwningEntityType and OwningEntityID arguments to all subs.
'  1.10       MR      03/30/05    Removed Parameters Array size in PUT Function.
'
' Function                              Description
' GetAllInfo(strUSer,showDeleted)   Returns an CalendarCollection containing all CalendarInfo objects from the repository.
' DBGetByID(ID)                     Returns a CalendarInfo object corresponding to a CalendarInfoID         
' DBGetDS(strSQL)                   Returns a dataset containing the results of the select query supplied in strSQL.
' PutCalendar(TankInfo)             Updates the repository with the information supplied in CalendarInfo. Inserts the
'                                       data if no matching CalendarInfo is in the repository.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class CalendarDB
        Private _strConn

#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

#Region "Exposed Operations"
        Public Function GetAllInfo(ByVal strUser As String, Optional ByVal showDeleted As Boolean = False) As Muster.Info.CalendarCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            'strSQL = "select * from TBLSYS_CALENDAR_CALENDAR_INFO"
            'strSQL += " WHERE USER_ID in (select distinct USER_ID from tblsys_profile_info where profile_modifier_1 in (select profile_modifier_1 from tblsys_profile_info where user_id = '" & strUser & "' and Profile_key = 'user groups' ) )"
            'strSQL += IIf(Not showDeleted, " AND DELETED = 0 ", "")
            'Dim drSet As SqlDataReader = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
            Try
                strSQL = "spGetCALENDAR_INFO_All"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@User_ID").Value = strUser
                Params("@OrderBy").Value = 2
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                Dim colCalendar As New MUSTER.Info.CalendarCollection
                While drSet.Read
                    Dim oCalendarInfo As New MUSTER.Info.CalendarInfo(drSet.Item("CALENDAR_INFO_ID"), _
                                                          drSet.Item("NOTIFICATION_DATE"), _
                                                          drSet.Item("DATE_DUE"), _
                                                          AltIsDBNull(drSet.Item("CURRENT_COLOR_CODE"), 0), _
                                                          drSet.Item("TASK_DESCRIPTION"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("SOURCE_USER_ID"), _
                                                          AltIsDBNull(drSet.Item("GROUP_ID"), String.Empty), _
                                                          drSet.Item("DUE_TO_ME"), _
                                                          drSet.Item("TO_DO"), _
                                                          drSet.Item("COMPLETED"), _
                                                          drSet.Item("DELETED"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                          drSet.Item("OWNING_ENTITY_TYPE"), _
                                                          drSet.Item("OWNING_ENTITY_ID"))
                    colCalendar.Add(oCalendarInfo)
                End While

                Return colCalendar
            Catch Ex As Exception
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try

        End Function
        Public Function DBGetByID(ByVal nVal As Int64) As Muster.Info.CalendarInfo
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            'strSQL = "SELECT * FROM TBLSYS_CALENDAR_CALENDAR_INFO WHERE CALENDAR_INFO_ID = " & nVal
            'Try
            '    Dim drSet As SqlDataReader = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
            Try
                strSQL = "spGetCALENDAR_INFO_BY_ID"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@User_ID").Value = DBNull.Value
                Params("@Group_ID").Value = DBNull.Value
                Params("@Calendar_Info_ID").Value = nVal
                Params("@Deleted").Value = False
                Params("@OrderBy").Value = 2

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CalendarInfo(drSet.Item("CALENDAR_INFO_ID"), _
                                                          drSet.Item("NOTIFICATION_DATE"), _
                                                          drSet.Item("DATE_DUE"), _
                                                          AltIsDBNull(drSet.Item("CURRENT_COLOR_CODE"), 0), _
                                                          drSet.Item("TASK_DESCRIPTION"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("SOURCE_USER_ID"), _
                                                          AltIsDBNull(drSet.Item("GROUP_ID"), String.Empty), _
                                                          drSet.Item("DUE_TO_ME"), _
                                                          drSet.Item("TO_DO"), _
                                                          drSet.Item("COMPLETED"), _
                                                          drSet.Item("DELETED"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                          drSet.Item("OWNING_ENTITY_TYPE"), _
                                                          drSet.Item("OWNING_ENTITY_ID"))
                End If

                Return New MUSTER.Info.CalendarInfo
            Catch ex As Exception
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetByOtherID(ByVal strVal As String, Optional ByVal strType As String = "USER", Optional ByVal nEntityType As Int16 = 0, Optional ByVal nEntityID As Int64 = 0) As MUSTER.Info.CalendarCollection
            Dim oColCal As New MUSTER.Info.CalendarCollection
            Dim drSet As SqlDataReader
            'If strType = "USER" Then
            '    strSQL = "SELECT * FROM TBLSYS_CALENDAR_CALENDAR_INFO WHERE USER_ID = '" & strVal & "' AND DELETED = 0"
            'Else
            '    strSQL = "SELECT * FROM TBLSYS_CALENDAR_CALENDAR_INFO WHERE GROUP_ID = '" & strVal & "' AND DELETED = 0"
            'End If
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetCALENDAR_INFO_BY_ID"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@User_ID").Value = DBNull.Value
                Params("@Group_ID").Value = DBNull.Value
                Params("@Calendar_Description").Value = DBNull.Value
                Params("@Calendar_Info_ID").Value = DBNull.Value
                Params("@Calendar_Owning_Entity_Type").value = IIf(nEntityType = 0, DBNull.Value, nEntityType)
                Params("@Calendar_Owning_Entity_ID").value = IIf(nEntityID = 0, DBNull.Value, nEntityID)
                Params("@Deleted").Value = False
                Params("@OrderBy").Value = 1

                Select Case strType
                    Case "USER"
                        Params("@User_ID").Value = strVal
                    Case "GROUP"
                        Params("@Group_ID").Value = strVal
                    Case "DESCRIPTION"
                        Params("@Calendar_Description").Value = strVal
                End Select

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read()
                        oColCal.Add(New MUSTER.Info.CalendarInfo(drSet.Item("CALENDAR_INFO_ID"), _
                                                          drSet.Item("NOTIFICATION_DATE"), _
                                                          drSet.Item("DATE_DUE"), _
                                                          AltIsDBNull(drSet.Item("CURRENT_COLOR_CODE"), 0), _
                                                          drSet.Item("TASK_DESCRIPTION"), _
                                                          drSet.Item("USER_ID"), _
                                                          drSet.Item("SOURCE_USER_ID"), _
                                                          AltIsDBNull(drSet.Item("GROUP_ID"), String.Empty), _
                                                          drSet.Item("DUE_TO_ME"), _
                                                          drSet.Item("TO_DO"), _
                                                          drSet.Item("COMPLETED"), _
                                                          drSet.Item("DELETED"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                          drSet.Item("OWNING_ENTITY_TYPE"), _
                                                          drSet.Item("OWNING_ENTITY_ID")))
                    End While
                End If

                Return oColCal
            Catch ex As Exception
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
            Catch ex As Exception
                Throw ex
            End Try
        End Function
        Public Function Put(ByRef obj As MUSTER.Info.CalendarInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                returnVal = String.Empty
                If Not moduleID = 0 Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Calendar, Integer))) Then
                        returnVal = "You do not have rights to save Calendar."
                        Exit Function
                    End If
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCalendar")
                If obj.CalendarInfoId <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = obj.CalendarInfoId
                End If
                Params(1).Value = obj.NotificationDate
                Params(2).Value = obj.DateDue
                Params(3).Value = obj.CurrentColorCode
                Params(4).Value = obj.TaskDescription
                Params(5).Value = obj.UserId
                Params(6).Value = obj.SourceUserId
                Params(7).Value = obj.GroupId
                Params(8).Value = obj.DueToMe
                Params(9).Value = obj.ToDo
                Params(10).Value = obj.Completed
                Params(11).Value = obj.Deleted
                Params(12).Value = 0
                Params(12).Direction = ParameterDirection.InputOutput
                Params(13).Value = obj.OwningEntityType
                Params(14).Value = obj.OwningEntityID

                If obj.CalendarInfoId <= 0 Then
                    Params(15).Value = obj.CreatedBy
                Else
                    Params(15).Value = obj.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCalendar", Params)

                If Params(12).Value <> 0 Then
                    obj.CalendarInfoId = Params(12).Value
                End If
            Catch ex As Exception
                Throw ex
            End Try

        End Function

        Public Function DBGetParentEntityID(ByVal fromEntityID As Int64, ByVal fromEntityType As Int64, ByVal toEntityType As Int64) As Int64
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet
            Dim returnVal As Int64 = 0
            Try
                strSQL = "spGetParentEntityID"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@fromEntityID").Value = fromEntityID
                Params("@fromEntityType").Value = fromEntityType
                Params("@toEntityType").Value = toEntityType

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not dsData Is Nothing Then
                    If dsData.Tables.Count > 0 Then
                        If dsData.Tables(0).Rows.Count > 0 Then
                            If Not dsData.Tables(0).Rows(0)(0) Is DBNull.Value Then
                                returnVal = dsData.Tables(0).Rows(0)(0)
                            End If
                        End If
                    End If
                End If

                Return returnVal

            Catch ex As Exception
                Throw ex
            End Try
        End Function
#End Region

    End Class
End Namespace
