'-------------------------------------------------------------------------------
' MUSTER.DataAccess.ProviderDB
'   Provides the means for marshalling Provider state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR     5/21/04    Original class definition.
'                                  
'
' Function                  Description
' GetAllInfo()      Returns an ProviderCollection containing all Course objects in the repository
' DBGetByID(ID)     Returns an ProviderInfo object indicated by int arg ID
' DBGetByName(ProviderName)     Returns an ProviderInfo object indicated by int arg Name
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(Entity)       Saves the Provider passed as an argument, to the DB
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class ProviderDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAllInfo() As MUSTER.Info.ProviderCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCOMProviders"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Provider_ID").Value = DBNull.Value
                Params("@Provider_Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colProvider As New MUSTER.Info.ProviderCollection
                While drSet.Read

                    Dim oProviderInfo As New MUSTER.Info.ProviderInfo(drSet.Item("PROVIDER_ID"), _
                                            drSet.Item("ACTIVE"), _
                                            drSet.Item("PROVIDER_NAME"), _
                                            drSet.Item("ABBREV"), _
                                            drSet.Item("DEPARTMENT"), _
                                            drSet.Item("WEBSITE"), _
                                            drSet.Item("DELETED"), _
                                            drSet.Item("CREATED_BY"), _
                                            drSet.Item("DATE_CREATED"), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                    colProvider.Add(oProviderInfo)
                End While

                Return colProvider
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByName(ByVal nVal As Integer) As MUSTER.Info.ProviderInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetCOMProviders"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Provider_ID").Value = DBNull.Value
                Params("@Provider_Name").Value = strVal
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ProviderInfo(drSet.Item("PROVIDER_ID"), _
                                      drSet.Item("ACTIVE"), _
                                      drSet.Item("PROVIDER_NAME"), _
                                      drSet.Item("ABBREV"), _
                                      drSet.Item("DEPARTMENT"), _
                                      drSet.Item("WEBSITE"), _
                                      drSet.Item("DELETED"), _
                                      drSet.Item("CREATED_BY"), _
                                      drSet.Item("DATE_CREATED"), _
                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                Else
                    Return New MUSTER.Info.ProviderInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try

        End Function
        Public Function DBGetByID(ByVal nVal As Integer) As MUSTER.Info.ProviderInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                strSQL = "spGetCOMProviders"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Provider_ID").Value = strVal
                Params("@Provider_Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ProviderInfo(drSet.Item("PROVIDER_ID"), _
                                      AltIsDBNull(drSet.Item("ACTIVE"), False), _
                                      AltIsDBNull(drSet.Item("PROVIDER_NAME"), String.Empty), _
                                      AltIsDBNull(drSet.Item("ABBREV"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DEPARTMENT"), String.Empty), _
                                      AltIsDBNull(drSet.Item("WEBSITE"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DELETED"), False), _
                                      AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))

                Else
                    Return New MUSTER.Info.ProviderInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
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
        Public Sub Put(ByRef oProviderInf As MUSTER.Info.ProviderInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Provider, Integer))) Then
                    returnVal = "You do not have rights to save Provider."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCOMProviders")

                If oProviderInf.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oProviderInf.ID
                End If
                Params(1).Value = oProviderInf.Active
                Params(2).Value = oProviderInf.ProviderName
                Params(3).Value = oProviderInf.Abbrev
                Params(4).Value = oProviderInf.Department
                Params(5).Value = oProviderInf.Website
                Params(6).Value = DBNull.Value
                Params(7).Value = DBNull.Value
                Params(8).Value = DBNull.Value
                Params(9).Value = DBNull.Value
                Params(10).Value = oProviderInf.Deleted

                If oProviderInf.ID <= 0 Then
                    Params(11).Value = oProviderInf.CreatedBy
                Else
                    Params(11).Value = oProviderInf.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCOMProviders", Params)
                If oProviderInf.ID <= 0 Then
                    oProviderInf.ID = Params(0).Value
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
    End Class
End Namespace
