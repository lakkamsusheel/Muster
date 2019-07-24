'-------------------------------------------------------------------------------
' MUSTER.DataAccess.AddressDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/06/04    Original class definition.
'  1.1        KJ      12/23/04    Added more descriptions in the header. Changed the showDeleted to make it consistent
'  1.2        EN      12/27/04    Changed from   ExecuteDataset to ExecuteReader in following methods
'                                 DBGetByID(),DBGetByAddressTypeID()
'  1.3        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.4        MNR     01/13/05    Changed from ExecuteDataset to ExecuteReader in DBGetByID()
'  1.5        EN      02/10/05    Modified 01/01/1901 to 01/01/0001 
'  1.6        AB      02/14/05    Changed dynamic SQL statement to a parameterized stored procedure in GetAllInfo(), 
'                                 DBGetByID() and DBGetByAddressTypeID()
'  1.7        AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.8        AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.9        AB      02/18/05    Set all parameters that are not required to NULL 
'                                   - they appear to be retaining values once cached
'  1.10       AB      02/23/05    Modified Get and Put functions based upon changes made to 
'                                   make several nullable fields non-nullable

'  1.11 Thomas Franey 12/2/09    Added the DB Logic to implement Physicaltown Into the Put/Get Statement
'
' Function                              Description
' GetAllInfo(showDeleted)   Returns an AddressCollection containing all AddressInfo objects from the repository.
' DBGetByID(ID)             Returns a AddressInfo object corresponding to a Address_ID
' DBGetByAddressTypeID(ID)  Returns a AddressInfo object corresponding to a Address_Type_ID
' DBGetDS(strSQL)           Returns a dataset containing the results of the select query supplied in strSQL.
' PutAddress(AddressInfo)   Updates the repository with the information supplied in AddressInfo. Inserts the
'                               data if no matching AddressInfo is in the repository.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class AddressDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
#Region "Exposed Operations"
        Public Function GetAllInfo(Optional ByVal showDeleted As Boolean = False) As Muster.Info.AddressCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader

            'Dim strSQL As String = "SELECT * FROM tblREG_Address_Master"
            'strSQL += IIf(Not showDeleted, " WHERE DELETED <> 1 ", "")
            'strSQL += "ORDER BY ADDRESS_ID"
            Try
                strSQL = "spGetAddress_Master"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Address_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colAddresses As New Muster.Info.AddressCollection
                While drSet.Read
                    Dim oAddressInfo As New MUSTER.Info.AddressInfo(drSet.Item("ADDRESS_ID"), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_TYPE_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("ENTITY_TYPE"), 0), _
                                                          drSet.Item("ADDRESS_LINE_ONE"), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_TWO"), String.Empty), _
                                                          drSet.Item("CITY"), _
                                                          drSet.Item("STATE"), _
                                                          drSet.Item("ZIP"), _
                                                          drSet.Item("FIPS_CODE"), _
                                                          AltIsDBNull(drSet.Item("START_DATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("END_DATE"), CDate("01/01/0001")), _
                                                          drSet.Item("DELETED"), _
                                                          drSet.Item("CREATED_BY"), _
                                                          drSet.Item("DATE_CREATED"), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("COUNTY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE_FOR_ENSITE"), drSet.Item("ADDRESS_LINE_ONE")), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_TWO_FOR_ENSITE"), String.Empty), AltIsDBNull(drSet("PHYSICALTOWN"), String.Empty))
                    colAddresses.Add(oAddressInfo)
                End While

                Return colAddresses
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Int64, Optional ByVal showDeleted As Boolean = False) As Muster.Info.AddressInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                If nVal <= 0 Then
                    Return New MUSTER.Info.AddressInfo
                End If
                strVal = nVal

                strSQL = "spGetAddress_Master"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@Address_ID").Value = strVal
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                'strSQL = "SELECT * FROM tblREG_Address_Master WHERE ADDRESS_ID = '" + strVal + "'"
                'If Not showDeleted Then
                '    strSQL += " AND DELETED <> 1"
                'End If
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.AddressInfo(drSet.Item("ADDRESS_ID"), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_TYPE_ID"), 0), _
                                                      AltIsDBNull(drSet.Item("ENTITY_TYPE"), 0), _
                                                      drSet.Item("ADDRESS_LINE_ONE"), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_TWO"), String.Empty), _
                                                      drSet.Item("CITY"), _
                                                      drSet.Item("STATE"), _
                                                      drSet.Item("ZIP"), _
                                                      drSet.Item("FIPS_CODE"), _
                                                      AltIsDBNull(drSet.Item("START_DATE"), CDate("01/01/0001")), _
                                                      AltIsDBNull(drSet.Item("END_DATE"), CDate("01/01/0001")), _
                                                      drSet.Item("DELETED"), _
                                                      drSet.Item("CREATED_BY"), _
                                                      drSet.Item("DATE_CREATED"), _
                                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                      AltIsDBNull(drSet.Item("COUNTY"), String.Empty), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE_FOR_ENSITE"), drSet.Item("ADDRESS_LINE_ONE")), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_TWO_FOR_ENSITE"), String.Empty), AltIsDBNull(drSet("PHYSICALTOWN"), String.Empty))
                Else
                    Return New MUSTER.Info.AddressInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not (drSet Is Nothing) Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByAddressTypeID(ByVal nVal As Int64) As Muster.Info.AddressInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                If nVal <= 0 Then
                    Return New MUSTER.Info.AddressInfo
                End If
                strSQL = "spGetAddress_Master_By_AddressTypeID"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@Address_Type_ID").Value = nVal
                Params("@OrderBy").Value = 2
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.AddressInfo(drSet.Item("ADDRESS_ID"), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_TYPE_ID"), 0), _
                                                      AltIsDBNull(drSet.Item("ENTITY_TYPE"), 0), _
                                                      drSet.Item("ADDRESS_LINE_ONE"), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_TWO"), String.Empty), _
                                                      drSet.Item("CITY"), _
                                                      drSet.Item("STATE"), _
                                                      drSet.Item("ZIP"), _
                                                      drSet.Item("FIPS_CODE"), _
                                                      AltIsDBNull(drSet.Item("START_DATE"), CDate("01/01/0001")), _
                                                      AltIsDBNull(drSet.Item("END_DATE"), CDate("01/01/0001")), _
                                                      drSet.Item("DELETED"), _
                                                      drSet.Item("CREATED_BY"), _
                                                      drSet.Item("DATE_CREATED"), _
                                                      AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                      AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                      AltIsDBNull(drSet.Item("COUNTY"), String.Empty), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE_FOR_ENSITE"), drSet.Item("ADDRESS_LINE_ONE")), _
                                                      AltIsDBNull(drSet.Item("ADDRESS_TWO_FOR_ENSITE"), String.Empty), AltIsDBNull(drSet("PHYSICALTOWN"), String.Empty))
                Else
                    Return New MUSTER.Info.AddressInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not (drSet Is Nothing) Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
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
        'Public Sub PutAddress(ByRef obj As Muster.Info.AddressInfo)
        '    Try
        '        Dim Params() As SqlParameter
        '        Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutAddress")
        '        Params(0).Value = IIf(obj.AddressId <= 0, 0, obj.AddressId)
        '        Params(1).Value = obj.AddressTypeId
        '        Params(2).Value = obj.EntityType
        '        Params(3).Value = obj.AddressLine1
        '        Params(4).Value = IIf(obj.AddressLine2 = String.Empty, DBNull.Value, obj.AddressLine2)
        '        Params(5).Value = obj.City
        '        Params(6).Value = obj.State
        '        Params(7).Value = obj.Zip
        '        Params(8).Value = IIf(obj.FIPSCode = String.Empty, DBNull.Value, obj.FIPSCode)
        '        SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutAddress", Params)
        '        If Params(0).Value <> obj.AddressId Then
        '            obj.AddressId = Params(0).Value
        '        End If
        '    Catch ex As Exception
        '        MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        Public Sub Put(ByRef obj As MUSTER.Info.AddressInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal entityTypeID As Integer, ByVal entity As Integer)
            Try

                If obj.EntityType = CType(SqlHelper.EntityTypes.Owner, Integer) Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Owner, Integer))) Then
                        returnVal = "You do not have rights to save a Owner."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                ElseIf obj.EntityType = CType(SqlHelper.EntityTypes.Facility, Integer) Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Facility, Integer))) Then
                        returnVal = "You do not have rights to save a Facility."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                Else
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Address, Integer))) Then
                        returnVal = "You do not have rights to save the Address."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                End If

                Dim strSQL As String = "spPutAddress"
                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(obj.AddressId <= 0, 0, obj.AddressId)
                Params(1).Value = IIf(obj.AddressTypeId <= 0, DBNull.Value, obj.AddressTypeId)
                Params(2).Value = obj.EntityType
                Params(3).Value = obj.AddressLine1.Trim
                Params(4).Value = IIf(obj.AddressLine2.Trim = String.Empty, DBNull.Value, obj.AddressLine2.Trim)
                Params(5).Value = obj.City.Trim
                Params(6).Value = obj.State.Trim
                Params(7).Value = obj.Zip.Trim
                Params(8).Value = IIf(obj.FIPSCode.Trim = String.Empty, DBNull.Value, obj.FIPSCode.Trim)
                Params(9).Value = obj.Deleted

                If obj.AddressId <= 0 Then
                    Params(10).Value = obj.CreatedBy
                Else
                    Params(10).Value = obj.ModifiedBy
                End If
                Params(11).Value = IIf(obj.AddressLine1ForEnsite = String.Empty, obj.AddressLine1, obj.AddressLine1ForEnsite)
                Params(12).Value = IIf(obj.AddressLine2ForEnsite = String.Empty, obj.AddressLine2, obj.AddressLine2ForEnsite)
                Params(13).Value = IIf(obj.PhsycalTown = String.Empty, DBNull.Value, obj.PhsycalTown)
                Params(14).Value = entityTypeID
                Params(15).Value = entity


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> obj.AddressId Then
                    obj.AddressId = Params(0).Value
                End If
                obj.AddressLine1ForEnsite = Params(11).Value
                If Params(12).Value Is DBNull.Value Then
                    obj.AddressLine2ForEnsite = String.Empty
                Else
                    obj.AddressLine2ForEnsite = Params(12).Value
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
    End Class
End Namespace
