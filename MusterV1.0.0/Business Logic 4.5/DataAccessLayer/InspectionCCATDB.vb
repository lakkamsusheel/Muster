'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionCCATDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as InspectionCCAT to build other objects.
'       Replace keyword "InspectionCCAT" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionCCATDB
        Private _strConn
        Private _FirstCompartment As Boolean = False

        Private MusterException As New MUSTER.Exceptions.MusterExceptions


#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionCCATInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionCCATInfo
                End If
                strSQL = "spGetInspectionCCAT"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_CCAT_ID").Value = id
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionCCATInfo(drSet.Item("INS_CCAT_ID"), _
                                                                AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("TANK_PIPE_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("TANK_PIPE_ENTITY_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("TANK_PIPE_RESPONSE"), False), _
                                                                AltIsDBNull(drSet.Item("TERMINATION"), False), _
                                                                AltIsDBNull(drSet.Item("TANK_PIPE_RESPONSE_DETAILS"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), AltIsDBNull(drSet.Item("CompartmentID"), 0))
                Else
                    Return New MUSTER.Info.InspectionCCATInfo
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

        Public Function DBGetCCATListForInspection(ByVal inspectID As Integer, Optional ByVal facID As Integer = 0) As DataTable

            Dim Params As Collection
            Dim dsData As DataSet

            Try
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, "spGetInspectionCCATlistByFacility")
                Params("@INSPECTIONID").Value = inspectID
                If inspectID = 0 Then
                    Params("@FACILITY_ID").Value = facID
                End If


                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spGetInspectionCCATlistByFacility", Params)

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex

            End Try

            If Not dsData Is Nothing AndAlso dsData.Tables.Count > 0 Then
                Return dsData.Tables(0)
            Else
                Return Nothing
            End If

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
        Public Sub Put(ByRef oInspectionCCATInfo As MUSTER.Info.InspectionCCATInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal firstCompartment As Boolean)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionCCAT"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionCCATInfo.ID <= 0, 0, oInspectionCCATInfo.ID)
                Params(1).Value = IIf(oInspectionCCATInfo.InspectionID <= 0, DBNull.Value, oInspectionCCATInfo.InspectionID)
                Params(2).Value = IIf(oInspectionCCATInfo.QuestionID = 0, DBNull.Value, oInspectionCCATInfo.QuestionID)
                Params(3).Value = IIf(oInspectionCCATInfo.TankPipeID = 0, DBNull.Value, oInspectionCCATInfo.TankPipeID)
                Params(4).Value = IIf(oInspectionCCATInfo.TankPipeEntityID <= 0, DBNull.Value, oInspectionCCATInfo.TankPipeEntityID)
                Params(5).Value = oInspectionCCATInfo.TankPipeResponse
                Params(6).Value = oInspectionCCATInfo.Termination
                Params(7).Value = oInspectionCCATInfo.TankPipeResponseDetail
                Params(8).Value = oInspectionCCATInfo.Deleted
                Params(9).Value = DBNull.Value
                Params(10).Value = DBNull.Value
                Params(11).Value = DBNull.Value
                Params(12).Value = DBNull.Value

                If oInspectionCCATInfo.ID <= 0 Then
                    Params(13).Value = oInspectionCCATInfo.CreatedBy
                Else
                    Params(13).Value = oInspectionCCATInfo.ModifiedBy
                End If

                Params(14).Value = oInspectionCCATInfo.CompartmentID
                Params(15).Value = IIf(firstCompartment, 1, 0)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionCCATInfo.ID Then
                    oInspectionCCATInfo.ID = Params(0).Value
                End If
                oInspectionCCATInfo.CreatedBy = AltIsDBNull(Params(9).Value, String.Empty)
                oInspectionCCATInfo.CreatedOn = AltIsDBNull(Params(10).Value, CDate("01/01/0001"))
                oInspectionCCATInfo.ModifiedBy = AltIsDBNull(Params(11).Value, String.Empty)
                oInspectionCCATInfo.ModifiedOn = AltIsDBNull(Params(12).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
