'-------------------------------------------------------------------------------
' MUSTER.DataAccess.ContactDatumDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      KKM         03/28/2005  Class definition
'  1.1      MR          04/29/05    Added Ext1 and Ext2 Parameters in Get and Put Functions.
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes

Namespace MUSTER.DataAccess
    Public Class ContactDatumDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByID(Optional ByVal ContactID As Integer = 0, Optional ByVal entityID As Integer = 0, Optional ByVal moduleID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetContactByID"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ContactID").Value = IIf(ContactID = 0, System.DBNull.Value, ContactID)
                Params("@EntityID").Value = IIf(entityID = 0, System.DBNull.Value, entityID)
                Params("@ModuleID").Value = IIf(moduleID = 0, System.DBNull.Value, moduleID)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetAll(Optional ByVal entityID As Integer = 0, Optional ByVal moduleID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ContactDatumCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Dim colContactDatum As New MUSTER.Info.ContactDatumCollection
            Try
                strSQL = "spGetCONDetails"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@EntityID").Value = IIf(entityID = 0, System.DBNull.Value, entityID)
                Params("@ModuleID").Value = IIf(moduleID = 0, System.DBNull.Value, moduleID)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        Dim contactDatumInfo As New MUSTER.Info.ContactDatumInfo(drSet.Item("ContactID"), _
                                                            (AltIsDBNull(drSet.Item("IsPerson"), 0)), _
                                                            (AltIsDBNull(drSet.Item("Org_Entity_Code"), 0)), _
                                                            (AltIsDBNull(drSet.Item("Company_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Title"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Prefix"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("First_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Middle_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Last_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Suffix"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Address_One"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Address_Two"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("City"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("State"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("ZipCode"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("FIPS_Code"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Phone_Number_One"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Phone_Number_Two"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Ext_One"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Ext_Two"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Fax_Number"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Cell_Number"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Email_Address"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Email_Address_Personal"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Vendor_Number"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001"))), _
                                                            (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001"))), _
                                                            (AltIsDBNull(drSet.Item("Deleted"), False)))
                        colContactDatum.Add(contactDatumInfo)
                    End While
                End If
                Return colContactDatum
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetInfoByID(ByVal contactID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ContactDatumInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetCONDetails_ByContactID"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ContactID").Value = contactID

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ContactDatumInfo(drSet.Item("ContactID"), _
                                                            (AltIsDBNull(drSet.Item("IsPerson"), 0)), _
                                                            (AltIsDBNull(drSet.Item("Org_Entity_Code"), 0)), _
                                                            (AltIsDBNull(drSet.Item("Company_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Title"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Prefix"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("First_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Middle_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Last_Name"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Suffix"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Address_One"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Address_Two"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("City"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("State"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("ZipCode"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("FIPS_Code"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Phone_Number_One"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Phone_Number_Two"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Ext_One"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Ext_Two"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Fax_Number"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Cell_Number"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Email_Address"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Email_Address_Personal"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("Vendor_Number"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001"))), _
                                                            (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                                                            (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001"))), _
                                                            (AltIsDBNull(drSet.Item("Deleted"), False)))
                Else
                    Return New MUSTER.Info.ContactDatumInfo
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function Put(ByRef oContactDatumInfo As MUSTER.Info.ContactDatumInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal NewAddress As Boolean)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Dim dtTempDate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Contact, Integer))) Then
                    returnVal = "You do not have rights to save Contact."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCONContact"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If oContactDatumInfo.ID > 0 Then
                    Params(0).Value = oContactDatumInfo.ID
                Else
                    Params(0).Value = 0
                End If
                Params(1).Value = IsNull(oContactDatumInfo.IsPerson, System.DBNull.Value)
                Params(2).Value = IsNull(oContactDatumInfo.orgCode, System.DBNull.Value)
                Params(3).Value = IsNull(oContactDatumInfo.companyName, String.Empty)
                Params(4).Value = IsNull(oContactDatumInfo.Title, String.Empty)
                Params(5).Value = IsNull(oContactDatumInfo.Prefix, String.Empty)
                Params(6).Value = IsNull(oContactDatumInfo.FirstName, String.Empty)
                Params(7).Value = IsNull(oContactDatumInfo.MiddleName, String.Empty)
                Params(8).Value = IsNull(oContactDatumInfo.LastName, String.Empty)
                Params(9).Value = IsNull(oContactDatumInfo.suffix, String.Empty)
                Params(10).Value = IsNull(oContactDatumInfo.AddressLine1, String.Empty)
                Params(11).Value = IsNull(oContactDatumInfo.AddressLine2, String.Empty)
                Params(12).Value = IsNull(oContactDatumInfo.City, String.Empty)
                Params(13).Value = IsNull(oContactDatumInfo.State, String.Empty)
                Params(14).Value = IsNull(oContactDatumInfo.ZipCode, String.Empty)
                Params(15).Value = IsNull(oContactDatumInfo.FipsCode, String.Empty)

                If oContactDatumInfo.Phone1 = "(___)___-____" Then
                    Params(16).Value = String.Empty
                Else
                    Params(16).Value = IsNull(oContactDatumInfo.Phone1, String.Empty)
                End If

                If oContactDatumInfo.Phone2 = "(___)___-____" Then
                    Params(17).Value = String.Empty
                Else
                    Params(17).Value = IsNull(oContactDatumInfo.Phone2, String.Empty)
                End If


                Params(18).Value = IsNull(oContactDatumInfo.Ext1, String.Empty)
                Params(19).Value = IsNull(oContactDatumInfo.Ext2, String.Empty)
                If oContactDatumInfo.Fax = "(___)___-____" Then
                    Params(20).Value = String.Empty
                Else
                    Params(20).Value = IsNull(oContactDatumInfo.Fax, String.Empty)
                End If

                If oContactDatumInfo.Cell = "(___)___-____" Then
                    Params(21).Value = String.Empty
                Else
                    Params(21).Value = IsNull(oContactDatumInfo.Cell, String.Empty)
                End If
                Params(22).Value = IsNull(oContactDatumInfo.publicEmail, String.Empty)
                Params(23).Value = IsNull(oContactDatumInfo.privateEmail, String.Empty)
                If oContactDatumInfo.VendorNumber = "0" Or oContactDatumInfo.VendorNumber = String.Empty Then
                    Params(24).Value = String.Empty
                Else
                    Params(24).Value = oContactDatumInfo.VendorNumber
                End If
                Params(25).Value = IsNull(oContactDatumInfo.deleted, System.DBNull.Value)
                Params(26).Value = DBNull.Value
                Params(27).Value = SqlDateTime.Null
                Params(28).Value = DBNull.Value
                Params(29).Value = SqlDateTime.Null

                If oContactDatumInfo.ID <= 0 Then
                    Params(30).Value = oContactDatumInfo.CreatedBy
                Else
                    Params(30).Value = oContactDatumInfo.modifiedBy
                End If

                Params(31).Value = DBNull.Value
                Params(32).Value = IIf(NewAddress, 1, 0)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oContactDatumInfo.ID Then
                    oContactDatumInfo.ID = Params(0).Value
                End If
                'oContactDatumInfo.CreatedBy = AltIsDBNull(Params(25).Value, String.Empty)
                'oContactDatumInfo.CreatedOn = AltIsDBNull(Params(26).Value, CDate("01/01/0001"))
                'oContactDatumInfo.modifiedBy = AltIsDBNull(Params(27).Value, String.Empty)
                'If Not Date.Compare(oContactDatumInfo.modifiedOn, dtTempDate) = 0 Then
                'oContactDatumInfo.modifiedOn = AltIsDBNull(Params(28).Value, CDate("01/01/0001"))
                'End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetDS(ByVal strSQL As String, ByVal ParamArray Args As Object()) As DataSet
            Dim dsData As DataSet
            Try

                dsData = SqlHelper.ExecuteDataset(_strConn, strSQL, Args)

                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
    End Class
End Namespace