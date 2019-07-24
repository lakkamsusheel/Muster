'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionDiscrepDB
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
' NOTE: This file to be used as InspectionDiscrep to build other objects.
'       Replace keyword "InspectionDiscrep" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionDiscrepDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionDiscrepInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If id <= 0 Then
                    Return New MUSTER.Info.InspectionDiscrepInfo
                End If
                strSQL = "spGetInspectionDiscrep"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_DESCREP_ID").Value = id
                Params("@INSPECTION_ID").Value = DBNull.Value
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionDiscrepInfo(drSet.Item("INS_DESCREP_ID"), _
                                                                AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("DESCRIPTION"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                                                AltIsDBNull(drSet.Item("DISCREP_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("INS_CIT_ID"), 0))
                Else
                    Return New MUSTER.Info.InspectionDiscrepInfo
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
        Public Function DBGetByOtherID(Optional ByVal inspID As Int64 = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionDiscrepsCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Dim discrepCollection As New MUSTER.Info.InspectionDiscrepsCollection

            If inspID <= 0 Then
                Return discrepCollection
            End If

            Try
                strSQL = "spGetInspectionDiscrep"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@INS_DESCREP_ID").Value = DBNull.Value
                Params("@INSPECTION_ID").Value = inspID
                Params("@DELETED").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                While drSet.Read
                    Dim discrepInfo As New MUSTER.Info.InspectionDiscrepInfo(drSet.Item("INS_DESCREP_ID"), _
                                                                AltIsDBNull(drSet.Item("INSPECTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("QUESTION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("DESCRIPTION"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("RESCINDED"), False), _
                                                                AltIsDBNull(drSet.Item("DISCREP_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("INS_CIT_ID"), 0))
                    discrepCollection.Add(discrepInfo)
                End While
                Return discrepCollection
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
        Public Sub Put(ByRef oInspectionDiscrepInfo As MUSTER.Info.InspectionDiscrepInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection Discrepency."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spPutInspectionDiscrep"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oInspectionDiscrepInfo.ID <= 0, 0, oInspectionDiscrepInfo.ID)
                Params(1).Value = IIf(oInspectionDiscrepInfo.InspectionID <= 0, DBNull.Value, oInspectionDiscrepInfo.InspectionID)
                Params(2).Value = IIf(oInspectionDiscrepInfo.QuestionID = 0, DBNull.Value, oInspectionDiscrepInfo.QuestionID)
                Params(3).Value = IIf(oInspectionDiscrepInfo.Description = String.Empty, DBNull.Value, oInspectionDiscrepInfo.Description)
                Params(4).Value = oInspectionDiscrepInfo.Deleted
                Params(5).Value = DBNull.Value
                Params(6).Value = DBNull.Value
                Params(7).Value = DBNull.Value
                Params(8).Value = DBNull.Value

                If oInspectionDiscrepInfo.ID <= 0 Then
                    Params(9).Value = oInspectionDiscrepInfo.CreatedBy
                Else
                    Params(9).Value = oInspectionDiscrepInfo.ModifiedBy
                End If
                Params(10).Value = oInspectionDiscrepInfo.Rescinded
                If Date.Compare(oInspectionDiscrepInfo.DiscrepReceived, CDate("01/01/0001")) = 0 Then
                    Params(11).Value = DBNull.Value
                Else
                    Params(11).Value = oInspectionDiscrepInfo.DiscrepReceived
                End If
                Params(12).Value = oInspectionDiscrepInfo.InspCitID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oInspectionDiscrepInfo.ID Then
                    oInspectionDiscrepInfo.ID = Params(0).Value
                End If
                oInspectionDiscrepInfo.CreatedBy = AltIsDBNull(Params(5).Value, String.Empty)
                oInspectionDiscrepInfo.CreatedOn = AltIsDBNull(Params(6).Value, CDate("01/01/0001"))
                oInspectionDiscrepInfo.ModifiedBy = AltIsDBNull(Params(7).Value, String.Empty)
                oInspectionDiscrepInfo.ModifiedOn = AltIsDBNull(Params(8).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
