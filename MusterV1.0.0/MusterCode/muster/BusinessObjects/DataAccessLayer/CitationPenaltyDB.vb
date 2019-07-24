'-------------------------------------------------------------------------------
' MUSTER.DataAccess.CitationPenaltyDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      KKM/RAF     06/25/2005  Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as CitationPenalty to build other objects.
'       Replace keyword "CitationPenalty" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class CitationPenaltyDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo(Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CitationPenaltysCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAECitationPenalty"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@CITATION_ID").Value = DBNull.Value
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)

                Dim colEntities As New MUSTER.Info.CitationPenaltysCollection
                While drSet.Read
                    Dim oCitationPenaltyInfo As New MUSTER.Info.CitationPenaltyInfo(drSet.Item("CITATION_ID"), _
                                                                                     AltIsDBNull(drSet.Item("StateCitation"), String.Empty), _
                                                                                     AltIsDBNull(drSet.Item("FederalCitation"), String.Empty), _
                                                                                     AltIsDBNull(drSet.Item("Section"), String.Empty), _
                                                                                     AltIsDBNull(drSet.Item("Description"), String.Empty), _
                                                                                     AltIsDBNull(drSet.Item("Category"), String.Empty), _
                                                                                    drSet.Item("Small"), _
                                                                                    drSet.Item("Medium"), _
                                                                                    drSet.Item("Large"), _
                                                                                    AltIsDBNull(drSet.Item("CorrectiveAction"), String.Empty), _
                                                                                    AltIsDBNull(drSet.Item("EPA"), CDate("01/01/0001")), _
                                                                                    AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                                    AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                                    AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                                    AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                                    drSet.Item("DELETED"))
                    colEntities.Add(oCitationPenaltyInfo)
                End While
                If Not drSet.IsClosed Then drSet.Close()
                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function DBGetByID(ByVal ID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CitationPenaltyInfo
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAECitationPenalty"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@CITATION_ID").Value = IIf(ID = 0, DBNull.Value, ID)
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CitationPenaltyInfo(drSet.Item("CITATION_ID"), _
                                                                                         AltIsDBNull(drSet.Item("StateCitation"), String.Empty), _
                                                                                         AltIsDBNull(drSet.Item("FederalCitation"), String.Empty), _
                                                                                         AltIsDBNull(drSet.Item("Section"), String.Empty), _
                                                                                         AltIsDBNull(drSet.Item("Description"), String.Empty), _
                                                                                         AltIsDBNull(drSet.Item("Category"), String.Empty), _
                                                                                        drSet.Item("Small"), _
                                                                                        drSet.Item("Medium"), _
                                                                                        drSet.Item("Large"), _
                                                                                        AltIsDBNull(drSet.Item("CorrectiveAction"), String.Empty), _
                                                                                        AltIsDBNull(drSet.Item("EPA"), CDate("01/01/0001")), _
                                                                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                                        drSet.Item("DELETED"))
                Else
                    Return New MUSTER.Info.CitationPenaltyInfo
                End If
                If Not drSet.IsClosed Then drSet.Close()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
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
        Public Sub Put(ByRef obj As MUSTER.Info.CitationPenaltyInfo)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spPutCAECitationPenalty"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                If obj.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = obj.ID
                End If
                Params(1).Value = IIf(obj.StateCitation = 0, DBNull.Value, obj.StateCitation)
                Params(2).Value = IsNull(obj.StateCitation, String.Empty)
                Params(3).Value = IsNull(obj.FederalCitation, String.Empty)
                Params(4).Value = IsNull(obj.Section, String.Empty)
                Params(5).Value = IsNull(obj.Description, String.Empty)
                Params(6).Value = IsNull(obj.Category, String.Empty)
                Params(7).Value = IIf(obj.Small = 0, DBNull.Value, obj.Small)
                Params(8).Value = IIf(obj.Medium = 0, DBNull.Value, obj.Medium)
                Params(9).Value = IIf(obj.Large = 0, DBNull.Value, obj.Large)
                Params(10).Value = IsNull(obj.CorrectiveAction, String.Empty)
                Params(11).Value = IsNull(obj.EPA, String.Empty)
                Params(12).Value = DBNull.Value
                Params(13).Value = DBNull.Value
                Params(14).Value = DBNull.Value
                Params(15).Value = DBNull.Value
                Params(16).Value = obj.Deleted
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If obj.ID <= 0 Then
                    obj.ID = Params(0).Value
                End If
                obj.CREATED_BY = IsNull(Params(12).Value, String.Empty)
                obj.DATE_CREATED = IsNull(Params(13).Value, CDate("01/01/0001"))
                obj.LAST_EDITED_BY = AltIsDBNull(Params(14).Value, String.Empty)
                obj.DATE_LAST_EDITED = AltIsDBNull(Params(15).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
    End Class
End Namespace
