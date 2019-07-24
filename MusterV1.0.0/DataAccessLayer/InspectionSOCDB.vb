'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionSOCDB
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
' NOTE: This file to be used as InspectionSOC to build other objects.
'       Replace keyword "InspectionSOC" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionSOCDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionSOCInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionSOCInfo
                End If
                strSQL = "spGetInspectionSOC"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_TOS_ID").Value = id
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionSOCInfo(drSet.Item("INS_TOS_ID"), _
                                                            AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                            AltIsDBNull(drSet.Item("FSOC_LK_PREVENT"), -1), _
                                                            AltIsDBNull(drSet.Item("FSOC_LK_PRE_CITATION"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("FSOC_LK_PRE_LINE_NUMBERS"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("FAC_SOC_LK_DETECTION"), -1), _
                                                            AltIsDBNull(drSet.Item("FSOC_LK_DET_CITATION"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("FSOC_LK_DET_LINENUMBERS"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("FSOC_LK_PRE_LK_DET"), -1), _
                                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("CAE_OVERRIDE"), False))
                Else
                    Return New MUSTER.Info.InspectionSOCInfo
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
        Public Sub Put(ByRef oInspectionSOCInfo As MUSTER.Info.InspectionSOCInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection SOC."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionSOC"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionSOCInfo.ID <= 0, 0, oInspectionSOCInfo.ID)
                Params(1).Value = IIf(oInspectionSOCInfo.InspectionID <= 0, DBNull.Value, oInspectionSOCInfo.InspectionID)
                Params(2).Value = IIf(oInspectionSOCInfo.LeakPrevention <= 0, DBNull.Value, oInspectionSOCInfo.LeakPrevention)
                Params(3).Value = IIf(oInspectionSOCInfo.LeakPreventionCitation = String.Empty, DBNull.Value, oInspectionSOCInfo.LeakPreventionCitation)
                Params(4).Value = IIf(oInspectionSOCInfo.LeakPreventionLineNumbers = String.Empty, DBNull.Value, oInspectionSOCInfo.LeakPreventionLineNumbers)
                Params(5).Value = IIf(oInspectionSOCInfo.LeakDetection <= 0, DBNull.Value, oInspectionSOCInfo.LeakDetection)
                Params(6).Value = IIf(oInspectionSOCInfo.LeakDetectionCitation = String.Empty, DBNull.Value, oInspectionSOCInfo.LeakDetectionCitation)
                Params(7).Value = IIf(oInspectionSOCInfo.LeakDetectionLineNumbers = String.Empty, DBNull.Value, oInspectionSOCInfo.LeakDetectionLineNumbers)
                Params(8).Value = IIf(oInspectionSOCInfo.LeakPreventionDetection <= 0, DBNull.Value, oInspectionSOCInfo.LeakPreventionDetection)
                Params(9).Value = oInspectionSOCInfo.Deleted
                Params(10).Value = DBNull.Value
                Params(11).Value = DBNull.Value
                Params(12).Value = DBNull.Value
                Params(13).Value = DBNull.Value

                If oInspectionSOCInfo.ID <= 0 Then
                    Params(14).Value = oInspectionSOCInfo.CreatedBy
                Else
                    Params(14).Value = oInspectionSOCInfo.ModifiedBy
                End If
                Params(15).Value = oInspectionSOCInfo.CAEOverride

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionSOCInfo.ID Then
                    oInspectionSOCInfo.ID = Params(0).Value
                End If
                oInspectionSOCInfo.CreatedBy = AltIsDBNull(Params(10).Value, String.Empty)
                oInspectionSOCInfo.CreatedOn = AltIsDBNull(Params(11).Value, CDate("01/01/0001"))
                oInspectionSOCInfo.ModifiedBy = AltIsDBNull(Params(12).Value, String.Empty)
                oInspectionSOCInfo.ModifiedOn = AltIsDBNull(Params(13).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
