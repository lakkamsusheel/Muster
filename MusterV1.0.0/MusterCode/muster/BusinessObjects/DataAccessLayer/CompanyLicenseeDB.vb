'-------------------------------------------------------------------------------
' MUSTER.DataAccess.CompanyLicenseeDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      RAF/MKK     05/17/2005  Original class definition
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
    Public Class CompanyLicenseeDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetAssociations(ByVal nCompanyID As Integer, Optional ByVal showDeleted As Boolean = False, Optional ByVal nLicenseeID As Integer = 0) As MUSTER.Info.CompanyLicenseeCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Dim colAssociations As New MUSTER.Info.CompanyLicenseeCollection
            Try
                strSQL = "spGetCOMCompanyLicenseeAssociation"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@COMPANY_ID").Value = nCompanyID
                Params("@COM_LICENSEE_ID").value = IIf(nLicenseeID = 0, DBNull.Value, nLicenseeID)
                Params("@DELETED").Value = showDeleted
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        Dim AssociationInfo As New MUSTER.Info.CompanyLicenseeInfo(drSet.Item("COM_LICENSEE_ID"), _
                                            AltIsDBNull(drSet.Item("COMPANY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COM_ADDRESS_ID"), 0), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                        colAssociations.Add(AssociationInfo)
                    End While
                End If
                Return colAssociations
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByAssociationID(ByVal AssociationID As String, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CompanyLicenseeInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetCOMCompanyLicenseeAssociation"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@COMPANY_ID").Value = DBNull.Value
                Params("@COM_LICENSEE_ID").Value = AssociationID
                Params("@DELETED").Value = showDeleted
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CompanyLicenseeInfo(drSet.Item("COM_LICENSEE_ID"), _
                                            AltIsDBNull(drSet.Item("COMPANY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                            AltIsDBNull(drSet.Item("COM_ADDRESS_ID"), 0), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.CompanyLicenseeInfo
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
        Public Sub Put(ByRef obj As MUSTER.Info.CompanyLicenseeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Licensee, Integer))) Then
                    returnVal = "You do not have rights to save Licensee."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCOMCompanyLicensee"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If obj.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = obj.ID
                End If
                Params(1).Value = IIf(obj.CompanyID = 0, DBNull.Value, obj.CompanyID)
                Params(2).Value = IIf(obj.LicenseeID = 0, DBNull.Value, obj.LicenseeID)
                Params(3).Value = IIf(obj.ComLicAddressID = 0, DBNull.Value, obj.ComLicAddressID)
                Params(4).Value = obj.Deleted
                Params(5).Value = DBNull.Value
                Params(6).Value = DBNull.Value
                Params(7).Value = DBNull.Value
                Params(8).Value = DBNull.Value
                If obj.ID <= 0 Then
                    Params(9).Value = obj.CreatedBy
                Else
                    Params(9).Value = obj.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If obj.ID <= 0 Then
                    obj.ID = Params(0).Value
                End If
                obj.CreatedBy = IsNull(Params(5).Value, String.Empty)
                obj.CreatedOn = IsNull(Params(6).Value, CDate("01/01/0001"))
                obj.ModifiedBy = AltIsDBNull(Params(7).Value, String.Empty)
                obj.ModifiedOn = AltIsDBNull(Params(8).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
    End Class
End Namespace
