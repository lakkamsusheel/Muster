'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectorCountyAssociationDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      RFF/KKM     06/21/2005  Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as InspectorCountyAssociation to build other objects.
'       Replace keyword "InspectorCountyAssociation" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes            ' Reqd for Inserting Null Values on Dates

Namespace MUSTER.DataAccess
    Public Class InspectorCountyAssociationDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetbyInspectorID(ByVal InspID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectorCountyAssociationsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAEInspectorCountyAssignment"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@STAFF_ID").Value = InspID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colInspectorCountyAssociation As New MUSTER.Info.InspectorCountyAssociationsCollection
                While drSet.Read
                    Dim oInspectorCountyAssociationInfo As New MUSTER.Info.InspectorCountyAssociationInfo(drSet.Item("ID"), _
                                                                            InspID, _
                                                                            drSet.Item("FIPS"), _
                                                                            String.Empty, _
                                                                             CDate("01/01/0001"), _
                                                                           String.Empty, _
                                                                             CDate("01/01/0001"), _
                                                                            0)
                    oInspectorCountyAssociationInfo.County = drSet.Item("COUNTY")
                    oInspectorCountyAssociationInfo.Facilities = drSet.Item("FACILITIES")
                    colInspectorCountyAssociation.Add(oInspectorCountyAssociationInfo)
                End While
                Return colInspectorCountyAssociation
                If Not drSet.IsClosed Then drSet.Close()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function DBGetByID(ByVal ID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectorCountyAssociationInfo
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                If ID <= 0 Then
                    Return New MUSTER.Info.InspectorCountyAssociationInfo
                End If

                strSQL = "spGetCAEInspectorOwnerAssignment"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ID").Value = IIf(ID = 0, DBNull.Value, ID)
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                Return New MUSTER.Info.InspectorCountyAssociationInfo(drSet.Item("INS_OWNER_ID"), _
                                                                            drSet.Item("STAFF_ID"), _
                                                                            drSet.Item("FIPS_CODE"), _
                                                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                            drSet.Item("DELETED"))
            Catch ex As Exception
                If Not drSet.IsClosed Then drSet.Close()
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        'Public Function DBGetCountyFacilities() As DataSet
        '    Dim dtData As DataSet
        '    Dim strSQL As String
        '    Try
        '        strSQL = "select * from VCAECountyFacilities"
        '        dtData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
        '        Return dtData
        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function DBGetInspectors() As DataSet
        '    Dim dtData As DataSet
        '    Dim strSQL As String
        '    Try
        '        strSQL = "select * from vCAEInspectors"
        '        dtData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
        '        Return dtData
        '    Catch Ex As Exception
        '        MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        Public Sub Put(ByRef obj As MUSTER.Info.InspectorCountyAssociationInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.CAE, Integer))) Then
                    returnVal = "You do not have rights to save Inspector County Association."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCAEInspectorCountyAssignment"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If obj.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = obj.ID
                End If
                Params(1).Value = IIf(obj.STAFF_ID = 0, DBNull.Value, obj.STAFF_ID)
                Params(2).Value = IIf(obj.FIPS_CODE = 0, DBNull.Value, obj.FIPS_CODE)
                Params(3).Value = DBNull.Value
                Params(4).Value = DBNull.Value
                Params(5).Value = DBNull.Value
                Params(6).Value = DBNull.Value
                Params(7).Value = obj.DELETED

                If obj.ID <= 0 Then
                    Params(8).Value = obj.CREATED_BY
                Else
                    Params(8).Value = obj.LAST_EDITED_BY
                End If


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If obj.ID <= 0 Then
                    obj.ID = Params(0).Value
                End If
                obj.CREATED_BY = IsNull(Params(3).Value, String.Empty)
                obj.DATE_CREATED = IsNull(Params(4).Value, CDate("01/01/0001"))
                obj.LAST_EDITED_BY = AltIsDBNull(Params(5).Value, String.Empty)
                obj.DATE_LAST_EDITED = AltIsDBNull(Params(6).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function DBGetAvailableCountyFacilities(ByVal inspID As String) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCAEInspectorAvailableCountyAssignment"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = DBNull.Value
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Public Function DBGetTotalCountyOwnerFacilities(ByVal inspID As String) As DataSet
        '    Dim dsData As DataSet
        '    Dim strSQL As String
        '    Dim Params() As SqlParameter
        '    Try
        '        strSQL = "spGetCAECountyOwnerFacilities"
        '        Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
        '        Params(0).Value = inspID
        '        dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
        '        Return dsData
        '    Catch ex As Exception
        '        MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetFacilityCount(ByVal strSQL As String) As Integer
            Dim Count As Integer
            Try
                Count = SqlHelper.ExecuteNonQuery(_strConn, CommandType.Text, strSQL)
                Return Count
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace
