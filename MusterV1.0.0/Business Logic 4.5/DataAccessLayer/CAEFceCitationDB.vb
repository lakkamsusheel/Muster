'-------------------------------------------------------------------------------
' MUSTER.DataAccess.CAEFceCitationDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MKK       08/15/2005   Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class CAEFceCitationDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo(Optional ByVal FCECitationID As Integer = Nothing, Optional ByVal showDeleted As Integer = 0) As MUSTER.Info.CAEFceCitationCollection
            Dim drSet As SqlDataReader
            Try
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblCAE_FACILITY_CITATIONS where deleted = " & showDeleted.ToString)
                Dim colEntities As New MUSTER.Info.CAEFceCitationCollection
                While drSet.Read
                    Dim oCAEFceCitationInfo As New MUSTER.Info.CAEFceCitationInfo(drSet.Item("FACILITY_CITATION_ID"), _
                                                          drSet.Item("FACILITY_ID"), _
                                                          drSet.Item("CITATION_ID"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                          drSet.Item("DELETED"))
                    colEntities.Add(oCAEFceCitationInfo)
                End While
                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.CAEFceCitationInfo
            Dim drSet As SqlDataReader
            Try
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, "SELECT * FROM tblCAE_FACILITY_CITATIONS WHERE FACILITY_CITATION_ID = " & nVal.ToString)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CAEFceCitationInfo(drSet.Item("FACILITY_CITATION_ID"), _
                                                          drSet.Item("FACILITY_ID"), _
                                                          drSet.Item("CITATION_ID"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                          drSet.Item("DELETED"))
                Else
                    Return New MUSTER.Info.CAEFceCitationInfo
                End If
            Catch ex As Exception
                If Not drSet.IsClosed Then drSet.Close()
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

        Public Sub Put(ByRef oCAEFceCitationInfo As MUSTER.Info.CAEFceCitationInfo)
            Dim strSQL As String
            Try
                Dim Params() As SqlParameter
                strSQL = "SPPutCAEFCECitation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IsNull(oCAEFceCitationInfo.ID, System.DBNull.Value)
                Params(1).Value = IsNull(oCAEFceCitationInfo.FacilityID, System.DBNull.Value)
                Params(2).Value = IsNull(oCAEFceCitationInfo.CitationID, System.DBNull.Value)
                Params(3).Value = DBNull.Value
                Params(4).Value = DBNull.Value
                Params(5).Value = DBNull.Value
                Params(6).Value = DBNull.Value
                Params(7).Value = IsNull(oCAEFceCitationInfo.Deleted, False)
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oCAEFceCitationInfo.ID Then
                    oCAEFceCitationInfo.ID = Params(0).Value
                End If
                oCAEFceCitationInfo.CREATED_BY = IsNull(Params(4).Value, String.Empty)
                oCAEFceCitationInfo.DATE_CREATED = IsNull(Params(5).Value, CDate("01/01/0001"))
                oCAEFceCitationInfo.LAST_EDITED_BY = AltIsDBNull(Params(6).Value, String.Empty)
                oCAEFceCitationInfo.DATE_LAST_EDITED = AltIsDBNull(Params(7).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
