'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionCPReadingsDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MNR         6/10/05     Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as InspectionCPReadings to build other objects.
'       Replace keyword "InspectionCPReadings" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionCPReadingsDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionCPReadingsInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionCPReadingsInfo
                End If
                strSQL = "spGetInspectionCPReadings"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_CP_READ_ID").Value = id
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionCPReadingsInfo(drSet.Item("INS_CP_READ_ID"), _
                                                                    AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                                    AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                    AltIsDBNull(drSet.Item("TANK_PIPE_ID"), 0), _
                                                                    AltIsDBNull(drSet.Item("TANK_INDEX"), 0), _
                                                                    AltIsDBNull(drSet.Item("TANK_PIPE_ENTITY_ID"), 0), _
                                                                    AltIsDBNull(drSet.Item("CONTACT_POINT"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("LOCAL_REFER_CELL_PLACEMENT"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("LOCAL_ON"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("REMOTE_OFF"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("PASS_FAIL_INCON"), -1), _
                                                                    AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                    AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                    AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                    AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                    AltIsDBNull(drSet.Item("LINE_NUMBER"), 0), _
                                                                    AltIsDBNull(drSet.Item("REMOTE_REFER_CELL_PLACEMENT"), False), _
                                                                    AltIsDBNull(drSet.Item("GALVANIC_IC"), False), _
                                                                    AltIsDBNull(drSet.Item("GALVANIC_IC_RESPONSE"), -1), _
                                                                    AltIsDBNull(drSet.Item("TESTED_BY_INSPECTOR"), False), _
                                                                    AltIsDBNull(drSet.Item("TESTED_BY_INSPECTOR_RESPONSE"), False))
                Else
                    Return New MUSTER.Info.InspectionCPReadingsInfo
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
        Public Sub Put(ByRef oInspectionCPReadingsInfo As MUSTER.Info.InspectionCPReadingsInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection CP Readings."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionCPReadings"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionCPReadingsInfo.ID <= 0, 0, oInspectionCPReadingsInfo.ID)
                Params(1).Value = IIf(oInspectionCPReadingsInfo.InspectionID <= 0, DBNull.Value, oInspectionCPReadingsInfo.InspectionID)
                Params(2).Value = IIf(oInspectionCPReadingsInfo.QuestionID = 0, DBNull.Value, oInspectionCPReadingsInfo.QuestionID)
                Params(3).Value = IIf(oInspectionCPReadingsInfo.TankPipeID <= 0, DBNull.Value, oInspectionCPReadingsInfo.TankPipeID)
                Params(4).Value = IIf(oInspectionCPReadingsInfo.TankPipeIndex <= 0, DBNull.Value, oInspectionCPReadingsInfo.TankPipeIndex)
                Params(5).Value = IIf(oInspectionCPReadingsInfo.TankPipeEntityID <= 0, DBNull.Value, oInspectionCPReadingsInfo.TankPipeEntityID)
                Params(6).Value = IIf(oInspectionCPReadingsInfo.ContactPoint = String.Empty, DBNull.Value, oInspectionCPReadingsInfo.ContactPoint)
                Params(7).Value = IIf(oInspectionCPReadingsInfo.LocalReferCellPlacement = String.Empty, DBNull.Value, oInspectionCPReadingsInfo.LocalReferCellPlacement)
                Params(8).Value = IIf(oInspectionCPReadingsInfo.LocalOn = String.Empty, DBNull.Value, oInspectionCPReadingsInfo.LocalOn)
                Params(9).Value = IIf(oInspectionCPReadingsInfo.RemoteOff = String.Empty, DBNull.Value, oInspectionCPReadingsInfo.RemoteOff)
                Params(10).Value = IIf(oInspectionCPReadingsInfo.PassFailIncon < 0, DBNull.Value, oInspectionCPReadingsInfo.PassFailIncon)
                Params(11).Value = oInspectionCPReadingsInfo.Deleted
                Params(12).Value = DBNull.Value
                Params(13).Value = DBNull.Value
                Params(14).Value = DBNull.Value
                Params(15).Value = DBNull.Value
                If oInspectionCPReadingsInfo.ID <= 0 Then
                    Params(16).Value = oInspectionCPReadingsInfo.CreatedBy
                Else
                    Params(16).Value = oInspectionCPReadingsInfo.ModifiedBy
                End If
                Params(17).Value = oInspectionCPReadingsInfo.LineNumber
                Params(18).Value = oInspectionCPReadingsInfo.RemoteReferCellPlacement
                Params(19).Value = oInspectionCPReadingsInfo.GalvanicIC
                Params(20).Value = oInspectionCPReadingsInfo.GalvanicICResponse
                Params(21).Value = oInspectionCPReadingsInfo.TestedByInspector
                Params(22).Value = oInspectionCPReadingsInfo.TestedByInspectorResponse

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionCPReadingsInfo.ID Then
                    oInspectionCPReadingsInfo.ID = Params(0).Value
                End If
                oInspectionCPReadingsInfo.CreatedBy = AltIsDBNull(Params(12).Value, String.Empty)
                oInspectionCPReadingsInfo.CreatedOn = AltIsDBNull(Params(13).Value, CDate("01/01/0001"))
                oInspectionCPReadingsInfo.ModifiedBy = AltIsDBNull(Params(14).Value, String.Empty)
                oInspectionCPReadingsInfo.ModifiedOn = AltIsDBNull(Params(15).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
