'-------------------------------------------------------------------------------
' MUSTER.DataAccess.ComAddressDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR      5/24/05    Original class definition.
'
'
' Function                              Description
' GetAllInfo(showDeleted)   Returns an ComAddressCollection containing all ComAddressInfo objects from the repository.
' DBGetByID(ID)             Returns a ComAddressInfo object corresponding to a Address_ID
' DBGetDS(strSQL)           Returns a dataset containing the results of the select query supplied in strSQL.
' PutAddress(ComAddressInfo)   Updates the repository with the information supplied in ComAddressInfo. Inserts the
'                               data if no matching ComAddressInfo is in the repository.
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class ComAddressDB
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
        Public Function GetAllInfo(Optional ByVal nAddressID As Integer = 0, Optional ByVal nCompanyID As Integer = 0, Optional ByVal nLicenseeID As Integer = 0, Optional ByVal nProviderID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ComAddressCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader

            Try
                strSQL = "spGetCOMAddress"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Com_Address_ID").Value = IIf(nAddressID = 0, DBNull.Value, nAddressID.ToString)
                Params("@Company_ID").Value = IIf(nCompanyID = 0, DBNull.Value, nCompanyID.ToString)
                Params("@Licensee_ID").Value = IIf(nLicenseeID = 0, DBNull.Value, nLicenseeID.ToString)
                Params("@Provider_ID").Value = IIf(nProviderID = 0, DBNull.Value, nProviderID.ToString)
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colAddresses As New MUSTER.Info.ComAddressCollection
                While drSet.Read

                    Dim oAddressInfo As New MUSTER.Info.ComAddressInfo(drSet.Item("COM_ADDRESS_ID"), _
                                                          AltIsDBNull(drSet.Item("COMPANY_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_LINE_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("CITY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("STATE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("ZIP"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("FIPS_CODE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_NUMBER_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("EXT_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_ONE_COMMENT"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_NUMBER_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("EXT_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_TWO_COMMENT"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("CELL_NUMBER"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("FAX_NUMBER"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DELETED"), False), _
                                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
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
        Public Function DBGetByID(ByVal nVal As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ComAddressInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                strVal = nVal

                strSQL = "spGetCOMAddress"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@Com_Address_ID").Value = strVal
                Params("@Company_ID").Value = DBNull.Value
                Params("@Licensee_ID").Value = DBNull.Value
                Params("@Provider_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ComAddressInfo(drSet.Item("COM_ADDRESS_ID"), _
                                                       AltIsDBNull(drSet.Item("COMPANY_ID"), 0), _
                                                       AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                                       AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                                       AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("ADDRESS_LINE_TWO"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("CITY"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("STATE"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("ZIP"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("FIPS_CODE"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("PHONE_NUMBER_ONE"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("EXT_ONE"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("PHONE_ONE_COMMENT"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("PHONE_NUMBER_TWO"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("EXT_TWO"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("PHONE_TWO_COMMENT"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("CELL_NUMBER"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("FAX_NUMBER"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("DELETED"), False), _
                                                       AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                       AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                       AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.ComAddressInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetCompanyAddress(ByVal nCompanyID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsSet As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Try

                strSQL = "spGetCOMAddress"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Com_Address_ID").Value = DBNull.Value
                Params("@Company_ID").Value = IIf(nCompanyID = 0, DBNull.Value, nCompanyID.ToString)
                Params("@Licensee_ID").Value = DBNull.Value
                Params("@Provider_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 2
                Params("@Deleted").Value = False
                dsSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsSet
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetByTypeID(ByVal nAddressID As Integer, Optional ByVal nCompanyID As Integer = 0, Optional ByVal nLicenseeID As Integer = 0, Optional ByVal nProviderID As Integer = 0) As MUSTER.Info.ComAddressInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try

                strSQL = "spGetCOMAddress"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@Com_Address_ID").Value = IIf(nAddressID = 0, DBNull.Value, nAddressID.ToString)
                Params("@Company_ID").Value = IIf(nCompanyID = 0, DBNull.Value, nCompanyID.ToString)
                Params("@Licensee_ID").Value = IIf(nLicenseeID = 0, DBNull.Value, nLicenseeID.ToString)
                Params("@Provider_ID").Value = IIf(nProviderID = 0, DBNull.Value, nProviderID.ToString)
                Params("@OrderBy").Value = 2
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ComAddressInfo(drSet.Item("COM_ADDRESS_ID"), _
                                                          AltIsDBNull(drSet.Item("COMPANY_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_LINE_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("CITY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("STATE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("ZIP"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("FIPS_CODE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_NUMBER_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("EXT_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_ONE_COMMENT"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_NUMBER_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("EXT_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_TWO_COMMENT"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("CELL_NUMBER"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("FAX_NUMBER"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DELETED"), False), _
                                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.ComAddressInfo
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
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetByProviderID(Optional ByVal nProviderID As Integer = 0) As MUSTER.Info.ComAddressInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try

                strSQL = "SELECT * FROM tblCOM_COMPANY_ADD_PHONE_DETAILS WHERE PROVIDER_ID = " & nProviderID.ToString
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ComAddressInfo(drSet.Item("COM_ADDRESS_ID"), _
                                                          AltIsDBNull(drSet.Item("COMPANY_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("PROVIDER_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("ADDRESS_LINE_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("CITY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("STATE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("ZIP"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("FIPS_CODE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_NUMBER_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("EXT_ONE"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_ONE_COMMENT"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_NUMBER_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("EXT_TWO"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("PHONE_TWO_COMMENT"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("CELL_NUMBER"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("FAX_NUMBER"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DELETED"), False), _
                                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.ComAddressInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Sub PutAddress(ByRef obj As MUSTER.Info.ComAddressInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Company, Integer))) Then
                    returnVal = "You do not have rights to save Company Address."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutCOMAddressPhoneDetails")
                If obj.AddressId <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = obj.AddressId
                End If

                Params(1).Value = IIf(obj.CompanyId = 0, DBNull.Value, obj.CompanyId)
                Params(2).Value = IIf(obj.LicenseeID = 0, DBNull.Value, obj.LicenseeID)
                Params(3).Value = IIf(obj.ProviderID = 0, DBNull.Value, obj.ProviderID)
                Params(4).Value = obj.AddressLine1
                Params(5).Value = obj.AddressLine2
                Params(6).Value = obj.City
                Params(7).Value = obj.State
                Params(8).Value = obj.Zip
                Params(9).Value = AltIsDBNull(obj.FIPSCode, String.Empty)

                If obj.Phone1 = "(___)___-____" Then
                    Params(10).Value = String.Empty
                Else
                    Params(10).Value = IsNull(obj.Phone1, String.Empty)
                End If

                Params(11).Value = obj.Ext1
                Params(12).Value = obj.Phone1Comment

                If obj.Phone2 = "(___)___-____" Then
                    Params(13).Value = String.Empty
                Else
                    Params(13).Value = IsNull(obj.Phone2, String.Empty)
                End If
                Params(14).Value = obj.Ext2
                Params(15).Value = obj.Phone2Comment
                If obj.Cell = "(___)___-____" Then
                    Params(16).Value = String.Empty
                Else
                    Params(16).Value = IsNull(obj.Cell, String.Empty)
                End If
                If obj.Fax = "(___)___-____" Then
                    Params(17).Value = String.Empty
                Else
                    Params(17).Value = IsNull(obj.Fax, String.Empty)
                End If
                Params(18).Value = DBNull.Value
                Params(19).Value = DBNull.Value
                Params(20).Value = DBNull.Value
                Params(21).Value = DBNull.Value
                Params(22).Value = obj.Deleted

                If obj.AddressId <= 0 Then
                    Params(23).Value = obj.CreatedBy
                Else
                    Params(23).Value = obj.ModifiedBy
                End If



                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutCOMAddressPhoneDetails", Params)
                If obj.AddressId <= 0 Then
                    obj.AddressId = Params(0).Value
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
    End Class
End Namespace
