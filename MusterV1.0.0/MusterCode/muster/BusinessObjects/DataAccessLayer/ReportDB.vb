'-------------------------------------------------------------------------------
' MUSTER.DataAccess.ReportDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/23/04    Original class definition.
'  1.1        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        EN        02/10/05     Modified 01/01/1901  to 01/01/0001
'  1.3        AB        02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByName, DBGetByID
'  1.4        AB        02/16/05    Added Finally to the Try/Catch to close all datareaders
'  1.5        AB        02/16/05    Removed any IsNull calls for fields the DB requires
'  1.6        AB        02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.7        AB        02/28/05    Modified Get functions based upon changes made to 
'                                     make several nullable fields non-nullable
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class ReportDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo() As Muster.Info.ReportsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetSYSReportMaster"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Report_ID").Value = DBNull.Value
                Params("@Report_Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "select * from tblSYS_REPORT_MASTER ORDER BY REPORT_ID")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.ReportsCollection
                While drSet.Read
                    Dim oReportInfo As New MUSTER.Info.ReportInfo(drSet.Item("REPORT_ID"), _
                                                                    drSet.Item("REPORT_NAME"), _
                                                                    drSet.Item("REPORT_MODULE"), _
                                                                    drSet.Item("REPORT_DESC"), _
                                                                    drSet.Item("REPORT_LOC"), _
                                                                    drSet.Item("DELETED"), _
                                                                    drSet.Item("CREATED_BY"), _
                                                                    drSet.Item("DATE_CREATED"), _
                                                                    AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                    AltIsDBNull(drSet.Item("ACTIVE"), 0))
                    colEntities.Add(oReportInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByName(ByVal strVal As String) As MUSTER.Info.ReportInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetSYSReportMaster"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Report_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Report_Name").Value = strVal

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_REPORT_MASTER WHERE REPORT_NAME = '" & strVal & "'")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ReportInfo(drSet.Item("REPORT_ID"), _
                                                        drSet.Item("REPORT_NAME"), _
                                                        drSet.Item("REPORT_MODULE"), _
                                                        drSet.Item("REPORT_DESC"), _
                                                        drSet.Item("REPORT_LOC"), _
                                                        drSet.Item("DELETED"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("ACTIVE"), 0))
                Else
                    Dim oReportInfo As New MUSTER.Info.ReportInfo
                    oReportInfo.Name = strVal
                    Return oReportInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.ReportInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetSYSReportMaster"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Report_Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Report_ID").Value = nVal

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblSYS_REPORT_MASTER WHERE REPORT_ID = " & nVal.ToString)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ReportInfo(drSet.Item("REPORT_ID"), _
                                                        drSet.Item("REPORT_NAME"), _
                                                        drSet.Item("REPORT_MODULE"), _
                                                        drSet.Item("REPORT_DESC"), _
                                                        drSet.Item("REPORT_LOC"), _
                                                        drSet.Item("DELETED"), _
                                                        drSet.Item("CREATED_BY"), _
                                                        drSet.Item("DATE_CREATED"), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("ACTIVE"), 0))
                Else
                    Return New MUSTER.Info.ReportInfo
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

        Public Sub Put(ByRef oRptInf As MUSTER.Info.ReportInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Report, Integer))) Then
                    returnVal = "You do not have rights to save a Report."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params(7) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPUTSYSREPORTMASTER")
                Params(0).Value = oRptInf.ID
                Params(1).Value = AltIsDBNull(oRptInf.Module, 0)
                Params(2).Value = AltIsDBNull(oRptInf.Description, String.Empty)
                Params(3).Value = AltIsDBNull(oRptInf.Name, String.Empty)
                Params(4).Value = AltIsDBNull(oRptInf.Path, String.Empty)
                Params(5).Value = oRptInf.Deleted
                Params(6).Value = oRptInf.Active

                If oRptInf.ID <= 0 Then
                    Params(7).Value = oRptInf.CreatedBy
                Else
                    Params(7).Value = oRptInf.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPUTSYSREPORTMASTER", Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        'Public Function ListReportNames(ByVal bolshowdeleted As Boolean) As DataTable
        '    Dim dtReportNames As DataTable

        '    Dim strSQL As String
        '    Dim dsset As New DataSet
        '    strSQL = "SELECT '' as REPORT_NAME,'' as Report_ID UNION SELECT  REPORT_NAME,REPORT_ID FROM tblSYS_REPORT_MASTER "
        '    strSQL += IIf(Not bolshowdeleted, " WHERE DELETED <> 1", "")
        '    strSQL += " Order by REPORT_NAME"
        '    Try
        '        dsset = DBGetDS(strSQL)
        '        If dsset.Tables(0).Rows.Count > 0 Then
        '            dtReportNames = dsset.Tables(0)
        '        Else
        '            dtReportNames = Nothing
        '        End If
        '        Return dtReportNames
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Public Function ListReportNames(ByVal moduleid As String, ByVal bolshowdeleted As Boolean, Optional ByVal UserID As String = "") As DataTable
        '    Dim dtReportNames As DataTable

        '    Dim strSQL As String
        '    Dim dsset As New DataSet
        '    strSQL = "SELECT '' as REPORT_NAME,'' as Report_ID,'' as REPORT_LOC UNION SELECT  REPORT_NAME as REPORT_NAME,report_id,REPORT_LOC as Report_ID FROM tblSYS_REPORT_MASTER "
        '    strSQL += IIf(moduleid > 0, " Where REPORT_MODULE='" & moduleid & "'", "")
        '    strSQL += IIf(Not bolshowdeleted, IIf(moduleid > 0, " and DELETED <> 1", " where DELETED <> 1"), "")
        '    strSQL += IIf(UserID <> "", " AND Report_NAME in (Select USER_ID from tblSys_Profile_Info where USER_ID in (SELECT  REPORT_NAME FROM tblSYS_REPORT_MASTER) and PROFILE_KEY='USER GROUPS' and PROFILE_VALUE in (Select PROFILE_VALUE from tblSys_Profile_Info Where USER_ID='" & UserID & "' and PROFILE_KEY='USER GROUPS' and PROFILE_VALUE <> 'NONE'))", "")
        '    strSQL += " Order by REPORT_NAME"
        '    Try
        '        dsset = DBGetDS(strSQL)
        '        If dsset.Tables(0).Rows.Count > 0 Then
        '            dtReportNames = dsset.Tables(0)
        '        Else
        '            dtReportNames = Nothing
        '        End If
        '        Return dtReportNames
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        Public Function ReportFileExists(ByVal FilePath As String, ByVal ReportID As String) As Boolean
            Dim strSQL As String
            Dim dsset As New DataSet
            strSQL = "SELECT * FROM tblSYS_REPORT_MASTER Where REPORT_ID <> " & ReportID & " And REPORT_LOC='" & FilePath & "'"
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

        Public Function DBGetReportGroupRel(ByVal reportID As Int64, Optional ByVal groupID As Int64 = 0, Optional ByVal showActiveInactiveGroups As Boolean = True, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ReportGroupRelationsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetReportGroupRel"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@REPORT_ID").Value = reportID
                Params("@GROUP_ID").Value = IIf(groupID <= 0, DBNull.Value, groupID)
                Params("@ACTIVE").Value = showActiveInactiveGroups
                Params("@DELETED").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colReportGroupRel As New MUSTER.Info.ReportGroupRelationsCollection
                While drSet.Read
                    Dim oReportGroupRelInfo As New MUSTER.Info.ReportGroupRelationInfo(drSet.Item("REPORT_ID"), _
                                                                        drSet.Item("GROUP_ID"), _
                                                                        drSet.Item("GROUP_NAME"), _
                                                                        drSet.Item("INACTIVE"), _
                                                                        drSet.Item("DELETED"))
                    colReportGroupRel.Add(oReportGroupRelInfo)
                End While
                Return colReportGroupRel
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Sub PutReportGroupRel(ByRef reportGroupRelInfo As MUSTER.Info.ReportGroupRelationInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutReportGroupRel"
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Report, Integer))) Then
                    returnVal = "You do not have rights to save Report."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = reportGroupRelInfo.ReportID
                Params(1).Value = reportGroupRelInfo.GroupID
                Params(2).Value = reportGroupRelInfo.Deleted

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function DBGetReportsForUser(Optional ByVal staffID As Integer = 0, Optional ByVal moduleID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal showFavReport As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Dim dsData As DataSet
            Try
                strSQL = "spGetReportsForUser"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@STAFF_ID").Value = IIf(staffID <= 0, DBNull.Value, staffID)
                Params("@MODULE_ID").Value = IIf(moduleID <= 0, DBNull.Value, moduleID)
                Params("@DELETED").Value = showDeleted
                Params("@FAV_REPORT").Value = showFavReport
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Sub PutFavReport(ByVal staffID As Integer, ByVal reportID As Integer, ByVal deleted As Boolean, ByVal moduleID As Integer, ByVal securityStaffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutFavReport"
            Try
                If Not SqlHelper.HasWriteAccess(moduleID, securityStaffID, SqlHelper.EntityTypes.Report) Then
                    returnVal = "You do not have rights to save Fav Report."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = staffID
                Params(1).Value = reportID
                Params(2).Value = deleted

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function DBGetCapPreMonthly(Optional ByVal showPrev As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Try
                strSQL = "spGetCAPPreMonthly"
                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = showPrev

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub PutCapPreMonthly(ByVal strCapStatusIDs As String)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutCAPPreMonthly"
            Try
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = strCapStatusIDs

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
    End Class
End Namespace

