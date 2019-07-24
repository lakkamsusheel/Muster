'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionMonitorWellsDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as InspectionMonitorWells to build other objects.
'       Replace keyword "InspectionMonitorWells" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionMonitorWellsDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionMonitorWellsInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionMonitorWellsInfo
                End If
                strSQL = "spGetInspectionMonitorWells"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_MON_WELL_ID").Value = id
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionMonitorWellsInfo(drSet.Item("INS_MON_WELL_ID"), _
                                                                    AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                                    AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                    AltIsDBNull(drSet.Item("TANK_LINE"), False), _
                                                                    AltIsDBNull(drSet.Item("WELL_NUMBER"), 0), _
                                                                    AltIsDBNull(drSet.Item("WELL_DEPTH"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DEPTH_TO_WATER"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DEPTH_TO_SLOTS"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("SURFACE_SEALED"), -1), _
                                                                    AltIsDBNull(drSet.Item("WELL_CAPS"), -1), _
                                                                    AltIsDBNull(drSet.Item("INSPECTORS_OBSERVTIONS"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                    AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                    AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                    AltIsDBNull(drSet.Item("LINE_NUMBER"), 0))
                Else
                    Return New MUSTER.Info.InspectionMonitorWellsInfo
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
        Public Sub Put(ByRef oInspectionMonitorWellsInfo As MUSTER.Info.InspectionMonitorWellsInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection Monitor Wells."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionMonitorWells"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionMonitorWellsInfo.ID <= 0, 0, oInspectionMonitorWellsInfo.ID)
                Params(1).Value = IIf(oInspectionMonitorWellsInfo.InspectionID <= 0, DBNull.Value, oInspectionMonitorWellsInfo.InspectionID)
                Params(2).Value = IIf(oInspectionMonitorWellsInfo.QuestionID = 0, DBNull.Value, oInspectionMonitorWellsInfo.QuestionID)
                Params(3).Value = oInspectionMonitorWellsInfo.TankLine
                Params(4).Value = IIf(oInspectionMonitorWellsInfo.WellNumber <= 0, DBNull.Value, oInspectionMonitorWellsInfo.WellNumber)
                Params(5).Value = oInspectionMonitorWellsInfo.WellDepth
                Params(6).Value = oInspectionMonitorWellsInfo.DepthToWater
                Params(7).Value = oInspectionMonitorWellsInfo.DepthToSlots
                Params(8).Value = IIf(oInspectionMonitorWellsInfo.SurfaceSealed = -1, DBNull.Value, oInspectionMonitorWellsInfo.SurfaceSealed)
                Params(9).Value = IIf(oInspectionMonitorWellsInfo.WellCaps = -1, DBNull.Value, oInspectionMonitorWellsInfo.WellCaps)
                Params(10).Value = oInspectionMonitorWellsInfo.InspectorsObservations
                Params(11).Value = oInspectionMonitorWellsInfo.Deleted
                Params(12).Value = DBNull.Value
                Params(13).Value = DBNull.Value
                Params(14).Value = DBNull.Value
                Params(15).Value = DBNull.Value
                If oInspectionMonitorWellsInfo.ID <= 0 Then
                    Params(16).Value = oInspectionMonitorWellsInfo.CreatedBy
                Else
                    Params(16).Value = oInspectionMonitorWellsInfo.ModifiedBy
                End If
                Params(17).Value = oInspectionMonitorWellsInfo.LineNumber

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionMonitorWellsInfo.ID Then
                    oInspectionMonitorWellsInfo.ID = Params(0).Value
                End If
                oInspectionMonitorWellsInfo.CreatedBy = AltIsDBNull(Params(12).Value, String.Empty)
                oInspectionMonitorWellsInfo.CreatedOn = AltIsDBNull(Params(13).Value, CDate("01/01/0001"))
                oInspectionMonitorWellsInfo.ModifiedBy = AltIsDBNull(Params(14).Value, String.Empty)
                oInspectionMonitorWellsInfo.ModifiedOn = AltIsDBNull(Params(15).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
