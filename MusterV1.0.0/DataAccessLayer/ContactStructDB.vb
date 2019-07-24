'-------------------------------------------------------------------------------
' MUSTER.DataAccess.ContactDatumDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      KKM         03/30/2005  Class definition
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
    Public Class ContactStructDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public ReadOnly Property SqlHelperProperty() As SqlHelper
            Get
                Dim sqlHelp As SqlHelper
                Return sqlHelp
            End Get
        End Property

        Public Function DBGetInfoByID(ByVal contactID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ContactStructInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetCONContactRelationship"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ContactAssocID").Value = contactID.ToString

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                Else
                    Return New MUSTER.Info.ContactStructInfo
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function

        Public Function DBGetMainDS(Optional ByVal nEntityID As Integer = 0, Optional ByVal ModuleID As Integer = 0, Optional ByVal SortOrder As Integer = 1, Optional ByVal ShowDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Try
                strSQL = "spGetCONStruct"
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DBGetContactsForAllModules(Optional ByVal nEntityID As Integer = 0, Optional ByVal nEntities As String = "") As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetCONTACTSFORALLMODULES"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@EntityID").Value = nEntityID
                Params("@Entities").Value = nEntities

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
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
        Public Function DBGetContactStruct(Optional ByVal nEntityAssocID As Integer = 0, Optional ByVal ModuleID As Integer = 0, Optional ByVal SortOrder As Integer = 1, Optional ByVal ShowDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCONEntityAssoc"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If nEntityAssocID = 0 Then
                    Params(0).Value = System.DBNull.Value
                Else
                    Params(0).Value = nEntityAssocID
                End If
                If ModuleID = 0 Then
                    Params(1).Value = System.DBNull.Value
                Else
                    Params(1).Value = ModuleID
                End If
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
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

        Public Function DBGetFunction(ByVal strSQL As String, ByVal ParamArray Args As Object()) As Object
            Dim dsData As Object
            Try

                dsData = SqlHelper.ExecuteScalar(_strConn, strSQL, Args)

                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Sub spPutReconciliation(ByRef contactStructInfo As MUSTER.Info.ContactStructInfo)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spPutCONReconciliation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(contactStructInfo.entityAssocID > 0, contactStructInfo.entityAssocID, 0)
                Params(1).Value = contactStructInfo.EntityID
                Params(2).Value = contactStructInfo.entityType
                Params(3).Value = contactStructInfo.ContactAssocID
                Params(4).Value = contactStructInfo.ContactTypeID
                Params(5).Value = contactStructInfo.moduleID
                Params(6).Value = contactStructInfo.EntityAssocActive
                Params(7).Value = IIf(contactStructInfo.ccInfo = String.Empty, DBNull.Value, contactStructInfo.ccInfo)
                Params(8).Value = IIf(contactStructInfo.displayAs = String.Empty, DBNull.Value, contactStructInfo.displayAs)
                Params(9).Value = contactStructInfo.EntityAssocdeleted
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> contactStructInfo.entityAssocID Then
                    contactStructInfo.entityAssocID = Params(0).Value
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function DBGetChildContacts(ByVal parentContactID As Integer)
            Dim params() As SqlParameter
            Dim ds As DataSet
            Try
                Dim strSQL As String = "spGetCONChildContacts"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                params(0).Value = parentContactID
                params(1).Value = 1
                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return ds
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBGetContactAliases(ByVal ContactID As Integer, Optional ByVal ComboBox As Boolean = False)
            Dim params() As SqlParameter
            Dim ds As DataSet
            Try
                Dim strSQL As String = "spGetCONAliases"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                params(0).Value = ContactID
                params(1).Value = IIf(ComboBox, 1, 0)
                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return ds
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBGetContactAddresses(ByVal contactID As Integer, Optional ByVal ComboBox As Boolean = False)
            Dim params() As SqlParameter
            Dim ds As DataSet
            Try
                Dim strSQL As String = "spGetCONAddresses"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                params(0).Value = contactID
                params(1).Value = IIf(ComboBox, 1, 0)
                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return ds
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Sub DBPutContactAlias(ByVal contactID As Integer, ByVal aliasName As String, ByVal deleted As Integer, ByVal user As String, Optional ByVal contactaliasID As Object = Nothing)
            Dim params() As SqlParameter
            Try
                Dim strSQL As String = "spPutCONContactAlias"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                params(0).Value = IIf(contactaliasID Is Nothing, DBNull.Value, CInt(contactaliasID))
                params(0).Direction = ParameterDirection.InputOutput
                params(1).Value = IIf(contactID = 0, DBNull.Value, contactID)
                params(2).Value = IIf(aliasName = String.Empty, DBNull.Value, aliasName)
                params(3).Value = deleted
                params(4).Value = user

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, params)

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function DBRemoveContactAddresses(ByVal ContactID As Integer, ByVal addressID As Integer)
            Dim params() As SqlParameter
            Dim ds As DataSet
            Try
                Dim strSQL As String = "spConDelAddresses"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                params(0).Value = ContactID
                params(1).Value = addressID
                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return ds
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBSetContactAddressesAsMainAddress(ByVal ContactID As Integer, ByVal addressID As Integer)
            Dim params() As SqlParameter
            Dim ds As DataSet
            Try
                Dim strSQL As String = "spConSetToMainAddresses"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                params(0).Value = ContactID
                params(1).Value = addressID
                ds = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return ds
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function DBGetSearchContact(Optional ByVal contactName As String = Nothing, Optional ByVal address As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal phone1 As String = Nothing, Optional ByVal phone2 As String = Nothing, Optional ByVal cell As String = Nothing, Optional ByVal fax As String = Nothing, Optional ByVal email As String = Nothing, Optional ByVal spName As String = "") As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim params() As SqlParameter
            Try
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, spName)
                params(0).Value = IIf(contactName = String.Empty, String.Empty, contactName)
                params(1).Value = IIf(address = String.Empty, DBNull.Value, address)
                params(2).Value = IIf(city = String.Empty, DBNull.Value, city)
                params(3).Value = IIf(state = String.Empty, DBNull.Value, state)
                params(4).Value = IIf(phone1 = String.Empty, DBNull.Value, phone1)
                params(5).Value = IIf(phone2 = String.Empty, DBNull.Value, phone2)
                params(6).Value = IIf(cell = String.Empty, DBNull.Value, cell)
                params(7).Value = IIf(fax = String.Empty, DBNull.Value, fax)
                params(8).Value = IIf(email = String.Empty, DBNull.Value, email)
                strSQL = spName
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function getReconciliation(Optional ByVal nOldContactAssocId As Integer = 0) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim params() As SqlParameter
            Try
                strSQL = "spGetCONReconciliation"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                params(0).Value = IIf(nOldContactAssocId = 0, System.DBNull.Value, nOldContactAssocId)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Sub UpdateReconciliation(ByVal strAccept As String, ByVal strReject As String, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Contact, Integer))) Then
                    returnVal = "You do not have rights to Update Reconciliation."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spModifyCONReconciliation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(strAccept = String.Empty, DBNull.Value, strAccept)
                Params(1).Value = IIf(strReject = String.Empty, DBNull.Value, strReject)
                Params(2).Value = UserID
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Sub PutContactRelationship(ByRef oContactStructInfo As MUSTER.Info.ContactStructInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal user As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Contact, Integer))) Then
                    returnVal = "You do not have rights to save save Contact Relationship."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCONContactRelation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = IIf(oContactStructInfo.ContactAssocID > 0, oContactStructInfo.ContactAssocID, 0)
                Params(1).Value = oContactStructInfo.ParentContactID
                Params(2).Value = IIf(oContactStructInfo.ChildContactID = 0, 0, oContactStructInfo.ChildContactID)
                Params(3).Value = oContactStructInfo.ContactAssocActive
                If (oContactStructInfo.dateAssociated = CDate("01/01/0001")) Then
                    Params(4).Value = System.DateTime.Now
                Else
                    Params(4).Value = oContactStructInfo.dateAssociated
                End If
                Params(5).Value = oContactStructInfo.ContactAssocdeleted

                Params(6).Value = user

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oContactStructInfo.ContactAssocID Then
                    oContactStructInfo.ContactAssocID = Params(0).Value
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub PutContactTypesAdmin(ByVal contactTypeID As Integer, ByVal contactType As String, ByVal moduleID As Integer, ByVal deleted As Integer, ByVal letterContactType As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Contact, Integer))) Then
                    returnVal = "You do not have rights to save Contact Types."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCONContactTypes"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)

                Params(0).Value = contactTypeID
                Params(1).Value = contactType
                Params(2).Value = moduleID
                Params(3).Value = letterContactType
                Params(4).Value = deleted
                Params(5).Value = UserID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub PutEntityContactRelationship(ByRef oContactStructInfo As MUSTER.Info.ContactStructInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal userID As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Contact, Integer))) Then
                    returnVal = "You do not have rights to save Entity Contact Relationship."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCONEntityContactRelation"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(oContactStructInfo.entityAssocID > 0, oContactStructInfo.entityAssocID, 0)
                Params(1).Value = oContactStructInfo.EntityID
                Params(2).Value = oContactStructInfo.entityType
                Params(3).Value = oContactStructInfo.ContactAssocID
                Params(4).Value = oContactStructInfo.ContactTypeID
                Params(5).Value = oContactStructInfo.moduleID
                Params(6).Value = oContactStructInfo.EntityAssocActive
                Params(7).Value = IIf(oContactStructInfo.ccInfo = String.Empty, DBNull.Value, oContactStructInfo.ccInfo)
                Params(8).Value = IIf(oContactStructInfo.displayAs = String.Empty, DBNull.Value, oContactStructInfo.displayAs)
                Params(9).Value = oContactStructInfo.EntityAssocdeleted
                Params(10).Value = userID
                Params(11).Value = IIf(oContactStructInfo.PreferredAlias = 0, DBNull.Value, oContactStructInfo.PreferredAlias)
                Params(12).Value = IIf(oContactStructInfo.PreferredAddress = 0, DBNull.Value, oContactStructInfo.PreferredAddress)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oContactStructInfo.entityAssocID Then
                    oContactStructInfo.entityAssocID = Params(0).Value
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function DBModifyContactAddress(ByVal oContactStructInfo As MUSTER.Info.ContactStructInfo, ByVal oContactDatumInfo As MUSTER.Info.ContactDatumInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal nUserResponse As Integer = -1, Optional ByRef ConEntityUpdateFlag As Integer = 0) As Integer
            Dim Params() As SqlParameter
            Dim strSQL As String
            Dim result As Integer
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Contact, Integer))) Then
                    returnVal = "You do not have rights to Modify Contact Address."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spModifyCONAddress"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IsNull(oContactStructInfo.EntityID, System.DBNull.Value)
                Params(1).Value = IsNull(oContactStructInfo.entityType, System.DBNull.Value)
                Params(2).Value = IsNull(oContactStructInfo.entityAssocID, System.DBNull.Value)
                Params(3).Value = IsNull(oContactStructInfo.ContactAssocID, System.DBNull.Value)
                Params(4).Value = IsNull(oContactDatumInfo.ID, System.DBNull.Value)
                Params(5).Value = IsNull(oContactStructInfo.moduleID, System.DBNull.Value)
                Params(6).Value = IsNull(oContactDatumInfo.IsPerson, System.DBNull.Value)
                Params(7).Value = IsNull(oContactDatumInfo.orgCode, System.DBNull.Value)
                Params(8).Value = IsNull(oContactDatumInfo.companyName, String.Empty)
                Params(9).Value = IsNull(oContactDatumInfo.Title, String.Empty)
                Params(10).Value = IsNull(oContactDatumInfo.Prefix, String.Empty)
                Params(11).Value = IsNull(oContactDatumInfo.FirstName, String.Empty)
                Params(12).Value = IsNull(oContactDatumInfo.MiddleName, String.Empty)
                Params(13).Value = IsNull(oContactDatumInfo.LastName, String.Empty)
                Params(14).Value = IsNull(oContactDatumInfo.suffix, String.Empty)
                Params(15).Value = IsNull(oContactDatumInfo.AddressLine1, String.Empty)
                Params(16).Value = IsNull(oContactDatumInfo.AddressLine2, String.Empty)
                Params(17).Value = IsNull(oContactDatumInfo.City, String.Empty)
                Params(18).Value = IsNull(oContactDatumInfo.State, String.Empty)
                Params(19).Value = IsNull(oContactDatumInfo.ZipCode, String.Empty)
                Params(20).Value = IsNull(oContactDatumInfo.FipsCode, String.Empty)

                If oContactDatumInfo.Phone1 = "(___)___-____" Then
                    Params(21).Value = String.Empty
                Else
                    Params(21).Value = IsNull(oContactDatumInfo.Phone1, String.Empty)
                End If

                If oContactDatumInfo.Phone2 = "(___)___-____" Then
                    Params(22).Value = String.Empty
                Else
                    Params(22).Value = IsNull(oContactDatumInfo.Phone2, String.Empty)
                End If

                Params(23).Value = IsNull(oContactDatumInfo.Ext1, String.Empty)
                Params(24).Value = IsNull(oContactDatumInfo.Ext2, String.Empty)

                If oContactDatumInfo.Fax = "(___)___-____" Then
                    Params(25).Value = String.Empty
                Else
                    Params(25).Value = IsNull(oContactDatumInfo.Fax, String.Empty)
                End If

                If oContactDatumInfo.Cell = "(___)___-____" Then
                    Params(26).Value = String.Empty
                Else
                    Params(26).Value = IsNull(oContactDatumInfo.Cell, String.Empty)
                End If

                Params(27).Value = IsNull(oContactDatumInfo.publicEmail, String.Empty)
                Params(28).Value = IsNull(oContactDatumInfo.privateEmail, String.Empty)
                If oContactDatumInfo.VendorNumber = "0" Or oContactDatumInfo.VendorNumber = String.Empty Then
                    Params(29).Value = String.Empty
                Else
                    Params(29).Value = oContactDatumInfo.VendorNumber
                End If
                Params(30).Value = oContactDatumInfo.deleted
                Params(31).Value = nUserResponse
                Params(32).Value = IsNull(oContactDatumInfo.CreatedBy, String.Empty)
                Params(33).Value = IsNull(oContactDatumInfo.CreatedOn, CDate("01/01/0001"))
                Params(34).Value = IsNull(oContactDatumInfo.modifiedBy, String.Empty)
                Params(35).Value = IsNull(oContactDatumInfo.modifiedOn, CDate("01/01/0001"))
                Params(36).Value = System.DBNull.Value
                Params(37).Value = UserID
                Params(38).Value = ConEntityUpdateFlag

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                result = Params(36).Value
                ConEntityUpdateFlag = Params(38).Value
                Return result
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub DBModifyContactAddressModule(ByVal oContactStructInfo As MUSTER.Info.ContactStructInfo, ByVal oContactDatumInfo As MUSTER.Info.ContactDatumInfo, ByVal result As String, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Contact, Integer))) Then
                    returnVal = "You do not have rights to Modify Contact Address."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spModifyCONAddressModule"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IsNull(oContactStructInfo.EntityID, System.DBNull.Value)
                Params(1).Value = IsNull(oContactStructInfo.entityType, System.DBNull.Value)
                Params(2).Value = IsNull(oContactStructInfo.entityAssocID, System.DBNull.Value)
                Params(3).Value = IsNull(oContactStructInfo.ContactAssocID, System.DBNull.Value)
                Params(4).Value = IsNull(oContactDatumInfo.ID, System.DBNull.Value)
                Params(5).Value = IsNull(oContactStructInfo.moduleID, System.DBNull.Value)
                Params(6).Value = IsNull(oContactDatumInfo.IsPerson, System.DBNull.Value)
                Params(7).Value = IsNull(oContactDatumInfo.orgCode, System.DBNull.Value)
                Params(8).Value = IsNull(oContactDatumInfo.companyName, String.Empty)
                Params(9).Value = IsNull(oContactDatumInfo.Title, String.Empty)
                Params(10).Value = IsNull(oContactDatumInfo.Prefix, String.Empty)
                Params(11).Value = IsNull(oContactDatumInfo.FirstName, String.Empty)
                Params(12).Value = IsNull(oContactDatumInfo.MiddleName, String.Empty)
                Params(13).Value = IsNull(oContactDatumInfo.LastName, String.Empty)
                Params(14).Value = IsNull(oContactDatumInfo.suffix, String.Empty)
                Params(15).Value = IsNull(oContactDatumInfo.AddressLine1, String.Empty)
                Params(16).Value = IsNull(oContactDatumInfo.AddressLine2, String.Empty)
                Params(17).Value = IsNull(oContactDatumInfo.City, String.Empty)
                Params(18).Value = IsNull(oContactDatumInfo.State, String.Empty)
                Params(19).Value = IsNull(oContactDatumInfo.ZipCode, String.Empty)
                Params(20).Value = IsNull(oContactDatumInfo.FipsCode, String.Empty)

                If oContactDatumInfo.Phone1 = "(___)___-____" Then
                    Params(21).Value = String.Empty
                Else
                    Params(21).Value = IsNull(oContactDatumInfo.Phone1, String.Empty)
                End If

                If oContactDatumInfo.Phone2 = "(___)___-____" Then
                    Params(22).Value = String.Empty
                Else
                    Params(22).Value = IsNull(oContactDatumInfo.Phone2, String.Empty)
                End If

                Params(23).Value = IsNull(oContactDatumInfo.Ext1, String.Empty)
                Params(24).Value = IsNull(oContactDatumInfo.Ext2, String.Empty)

                If oContactDatumInfo.Fax = "(___)___-____" Then
                    Params(25).Value = String.Empty
                Else
                    Params(25).Value = IsNull(oContactDatumInfo.Fax, String.Empty)
                End If

                If oContactDatumInfo.Cell = "(___)___-____" Then
                    Params(26).Value = String.Empty
                Else
                    Params(26).Value = IsNull(oContactDatumInfo.Cell, String.Empty)
                End If


                Params(27).Value = IsNull(oContactDatumInfo.publicEmail, String.Empty)
                Params(28).Value = IsNull(oContactDatumInfo.privateEmail, String.Empty)
                If oContactDatumInfo.VendorNumber = "0" Or oContactDatumInfo.VendorNumber = String.Empty Then
                    Params(29).Value = String.Empty
                Else
                    Params(29).Value = oContactDatumInfo.VendorNumber
                End If

                Params(30).Value = oContactDatumInfo.deleted
                Params(31).Value = IsNull(oContactDatumInfo.CreatedBy, String.Empty)
                Params(32).Value = IsNull(oContactDatumInfo.CreatedOn, CDate("01/01/0001"))
                Params(33).Value = IsNull(oContactDatumInfo.modifiedBy, String.Empty)
                Params(34).Value = IsNull(oContactDatumInfo.modifiedOn, CDate("01/01/0001"))
                Params(35).Value = IsNull(result, String.Empty)
                Params(36).Value = UserID
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function DBGetContactsByEntityAndModule(Optional ByVal nEntityID As Integer = 0, Optional ByVal ModuleID As Integer = 0) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCONStruct"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IsNull(nEntityID, System.DBNull.Value)
                Params(1).Value = IsNull(ModuleID, System.DBNull.Value)
                Params(2).Value = True
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DBGetFilteredContacts(Optional ByVal nEntityID As Integer = 0, Optional ByVal ModuleID As Integer = 0, Optional ByVal strEntityIDs As String = "", Optional ByVal bolActive As Boolean = False, Optional ByVal strEntityAssocIDs As String = "", Optional ByVal nEntityType As Integer = 0, Optional ByVal nRelatedEntityType As Integer = 0) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCONStruct2"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                'If nEntityID = 0 Then
                '    Params(0).Value = System.DBNull.Value
                'Else
                '    Params(0).Value = nEntityID
                'End If
                Params(0).Value = nEntityID
                If ModuleID = 0 Then
                    Params(1).Value = System.DBNull.Value
                Else
                    Params(1).Value = ModuleID
                End If
                If bolActive = False Then
                    Params(2).Value = System.DBNull.Value
                Else
                    Params(2).Value = bolActive
                End If
                If strEntityIDs = String.Empty Then
                    Params(3).Value = System.DBNull.Value
                Else
                    Params(3).Value = strEntityIDs
                End If
                If strEntityAssocIDs = String.Empty Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = strEntityAssocIDs
                End If
                If nEntityType = 0 Then
                    Params(5).Value = System.DBNull.Value
                Else
                    Params(5).Value = nEntityType
                End If
                If nRelatedEntityType = 0 Then
                    Params(6).Value = System.DBNull.Value
                Else
                    Params(6).Value = nRelatedEntityType
                End If
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)



                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
    End Class
End Namespace