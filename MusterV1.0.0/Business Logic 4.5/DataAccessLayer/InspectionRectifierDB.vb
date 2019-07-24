'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionRectifierDB
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
' NOTE: This file to be used as InspectionRectifier to build other objects.
'       Replace keyword "InspectionRectifier" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionRectifierDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionRectifierInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionRectifierInfo
                End If
                strSQL = "spGetInspectionRectifier"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_RECT_ID").Value = id
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionRectifierInfo(drSet.Item("INS_RECT_ID"), _
                                                                AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("RECITIFIER_ON"), False), _
                                                                AltIsDBNull(drSet.Item("INOP_HOW_LONG"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("VOLTS"), 0.0), _
                                                                AltIsDBNull(drSet.Item("AMPS"), 0.0), _
                                                                AltIsDBNull(drSet.Item("HOURS"), 0.0), _
                                                                AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.InspectionRectifierInfo
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
        Public Sub Put(ByRef oInspectionRectifierInfo As MUSTER.Info.InspectionRectifierInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection Rectifier."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionRectifier"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionRectifierInfo.ID <= 0, 0, oInspectionRectifierInfo.ID)
                Params(1).Value = IIf(oInspectionRectifierInfo.InspectionID <= 0, DBNull.Value, oInspectionRectifierInfo.InspectionID)
                Params(2).Value = IIf(oInspectionRectifierInfo.QuestionID = 0, DBNull.Value, oInspectionRectifierInfo.QuestionID)
                Params(3).Value = oInspectionRectifierInfo.RectifierOn
                Params(4).Value = IIf(oInspectionRectifierInfo.InopHowLong = String.Empty, DBNull.Value, oInspectionRectifierInfo.InopHowLong)
                Params(5).Value = IIf(oInspectionRectifierInfo.Volts <= 0.0, DBNull.Value, oInspectionRectifierInfo.Volts)
                Params(6).Value = IIf(oInspectionRectifierInfo.Amps <= 0.0, DBNull.Value, oInspectionRectifierInfo.Amps)
                Params(7).Value = IIf(oInspectionRectifierInfo.Hours <= 0.0, DBNull.Value, oInspectionRectifierInfo.Hours)
                Params(8).Value = oInspectionRectifierInfo.Deleted
                Params(9).Value = DBNull.Value
                Params(10).Value = DBNull.Value
                Params(11).Value = DBNull.Value
                Params(12).Value = DBNull.Value

                If oInspectionRectifierInfo.ID <= 0 Then
                    Params(13).Value = oInspectionRectifierInfo.CreatedBy
                Else
                    Params(13).Value = oInspectionRectifierInfo.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionRectifierInfo.ID Then
                    oInspectionRectifierInfo.ID = Params(0).Value
                End If
                oInspectionRectifierInfo.CreatedBy = AltIsDBNull(Params(9).Value, String.Empty)
                oInspectionRectifierInfo.CreatedOn = AltIsDBNull(Params(10).Value, CDate("01/01/0001"))
                oInspectionRectifierInfo.ModifiedBy = AltIsDBNull(Params(11).Value, String.Empty)
                oInspectionRectifierInfo.ModifiedOn = AltIsDBNull(Params(12).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
