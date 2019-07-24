'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionResponsesDB
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
' NOTE: This file to be used as InspectionResponses to build other objects.
'       Replace keyword "InspectionResponses" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionResponsesDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionResponsesInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionResponsesInfo
                End If
                strSQL = "spGetInspectionResponses"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_RESP_ID").Value = id
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionResponsesInfo(drSet.Item("INS_RESP_ID"), _
                                                                AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("SOC"), False), _
                                                                AltIsDBNull(drSet.Item("RESPONSE"), -1), _
                                                                AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.InspectionResponsesInfo
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
        Public Sub Put(ByRef oInspectionResponsesInfo As MUSTER.Info.InspectionResponsesInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection Responses."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionResponses"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionResponsesInfo.ID <= 0, 0, oInspectionResponsesInfo.ID)
                Params(1).Value = oInspectionResponsesInfo.InspectionID
                Params(2).Value = oInspectionResponsesInfo.QuestionID
                Params(3).Value = oInspectionResponsesInfo.SOC
                Params(4).Value = oInspectionResponsesInfo.Response
                Params(5).Value = oInspectionResponsesInfo.Deleted
                Params(6).Value = DBNull.Value
                Params(7).Value = DBNull.Value
                Params(8).Value = DBNull.Value
                Params(9).Value = DBNull.Value

                If oInspectionResponsesInfo.ID <= 0 Then
                    Params(10).Value = oInspectionResponsesInfo.CreatedBy
                Else
                    Params(10).Value = oInspectionResponsesInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionResponsesInfo.ID Then
                    oInspectionResponsesInfo.ID = Params(0).Value
                End If

                oInspectionResponsesInfo.CreatedBy = AltIsDBNull(Params(6).Value, String.Empty)
                oInspectionResponsesInfo.CreatedOn = AltIsDBNull(Params(7).Value, CDate("01/01/0001"))
                oInspectionResponsesInfo.ModifiedBy = AltIsDBNull(Params(8).Value, String.Empty)
                oInspectionResponsesInfo.ModifiedOn = AltIsDBNull(Params(9).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
