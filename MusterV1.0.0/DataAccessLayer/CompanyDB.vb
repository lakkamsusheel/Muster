'-------------------------------------------------------------------------------
' MUSTER.DataAccess.CompanyDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        RAF/MKK      05/15/2005   Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
'
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes

Namespace MUSTER.DataAccess
    Public Class CompanyDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByInfoID(ByVal ComID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.CompanyInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetCOMCompany"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@COMPANY_ID").Value = ComID
                Params("@DELETED").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.CompanyInfo(drSet.Item("COMPANY_ID"), _
                                                        (AltIsDBNull(drSet.Item("CERT_RESPON"), String.Empty)), _
                                                        (AltIsDBNull(drSet.Item("COMPANY_NAME"), String.Empty)), _
                                                        (AltIsDBNull(drSet.Item("FIN_RESP_END_DATE"), CDate("01/01/0001"))), _
                                                        (AltIsDBNull(drSet.Item("EMAIL_ADDRESS"), String.Empty)), _
                                                        (AltIsDBNull(drSet.Item("PRO_ENGIN"), String.Empty)), _
                                                        (AltIsDBNull(drSet.Item("PRO_ENGIN_ADD_ID"), 0)), _
                                                        (AltIsDBNull(drSet.Item("PRO_ENGIN_NUMBER"), String.Empty)), _
                                                        (AltIsDBNull(drSet.Item("PRO_ENGIN_APP_APRV_DATE"), CDate("01/01/0001"))), _
                                                        (AltIsDBNull(drSet.Item("PRO_ENGIN_LIABIL_DATE"), CDate("01/01/0001"))), _
                                                        (AltIsDBNull(drSet.Item("PRO_GEOLO"), String.Empty)), _
                                                        (AltIsDBNull(drSet.Item("PRO_GEOLO_ADD_ID"), 0)), _
                                                        (AltIsDBNull(drSet.Item("PRO_GEOLO_NUMBER"), String.Empty)), _
                                                        (AltIsDBNull(drSet.Item("CTIAC"), False)), _
                                                        (AltIsDBNull(drSet.Item("CTC"), False)), _
                                                        (AltIsDBNull(drSet.Item("PTTT"), False)), _
                                                        (AltIsDBNull(drSet.Item("LDC"), False)), _
                                                        (AltIsDBNull(drSet.Item("USTSE"), False)), _
                                                        (AltIsDBNull(drSet.Item("WL"), False)), _
                                                        (AltIsDBNull(drSet.Item("TST"), False)), _
                                                        (AltIsDBNull(drSet.Item("ERAC"), False)), _
                                                        (AltIsDBNull(drSet.Item("IRAC"), False)), _
                                                        (AltIsDBNull(drSet.Item("EC"), False)), _
                                                        (AltIsDBNull(drSet.Item("ED"), False)), _
                                                        (AltIsDBNull(drSet.Item("TL"), False)), _
                                                        (AltIsDBNull(drSet.Item("CE"), False)), _
                                                        (AltIsDBNull(drSet.Item("CM"), False)), _
                                                        (AltIsDBNull(drSet.Item("ACTIVE"), False)), _
                                                        AltIsDBNull(drSet.Item("DELETED"), False), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("PRO_ENGIN_EMAIL"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("PRO_GEOLO_EMAIL"), String.Empty))
                Else
                    Return New MUSTER.Info.CompanyInfo
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
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
        Public Sub put(ByRef oCompanyInfo As MUSTER.Info.CompanyInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Dim dtTempDate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Company, Integer))) Then
                    returnVal = "You do not have rights to save Company."
                End If

                strSQL = "spPutCOMCompany"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IsNull(oCompanyInfo.ID, System.DBNull.Value)
                Params(1).Value = IsNull(oCompanyInfo.CERT_RESPON, String.Empty)
                Params(2).Value = oCompanyInfo.COMPANY_NAME
                If Date.Compare(oCompanyInfo.FIN_RESP_END_DATE, dtTempDate) = 0 Then
                    Params(3).Value = SqlDateTime.Null
                Else
                    Params(3).Value = oCompanyInfo.FIN_RESP_END_DATE.Date
                End If
                Params(4).Value = IsNull(oCompanyInfo.EMAIL_ADDRESS, String.Empty)
                Params(5).Value = IsNull(oCompanyInfo.PRO_ENGIN, String.Empty)
                Params(6).Value = IsNull(oCompanyInfo.PRO_ENGIN_ADD_ID, System.DBNull.Value)
                Params(7).Value = IsNull(oCompanyInfo.PRO_ENGIN_NUMBER, String.Empty)
                If Date.Compare(oCompanyInfo.PRO_ENGIN_APP_APRV_DATE, dtTempDate) = 0 Then
                    Params(8).Value = SqlDateTime.Null
                Else
                    Params(8).Value = oCompanyInfo.PRO_ENGIN_APP_APRV_DATE.Date
                End If
                If Date.Compare(oCompanyInfo.PRO_ENGIN_LIABIL_DATE, dtTempDate) = 0 Then
                    Params(9).Value = SqlDateTime.Null
                Else
                    Params(9).Value = oCompanyInfo.PRO_ENGIN_LIABIL_DATE.Date
                End If
                Params(10).Value = IsNull(oCompanyInfo.PRO_GEOLO, String.Empty)
                Params(11).Value = IsNull(oCompanyInfo.PRO_GEOLO_ADD_ID, System.DBNull.Value)
                Params(12).Value = IsNull(oCompanyInfo.PRO_GEOLO_NUMBER, String.Empty)
                Params(13).Value = IsNull(oCompanyInfo.CTIAC, System.DBNull.Value)
                Params(14).Value = IsNull(oCompanyInfo.CTC, System.DBNull.Value)
                Params(15).Value = IsNull(oCompanyInfo.PTTT, System.DBNull.Value)
                Params(16).Value = IsNull(oCompanyInfo.LDC, System.DBNull.Value)
                Params(17).Value = IsNull(oCompanyInfo.USTSE, System.DBNull.Value)
                Params(18).Value = IsNull(oCompanyInfo.WL, System.DBNull.Value)
                Params(19).Value = IsNull(oCompanyInfo.TST, System.DBNull.Value)
                Params(20).Value = IsNull(oCompanyInfo.ERAC, System.DBNull.Value)
                Params(21).Value = IsNull(oCompanyInfo.IRAC, System.DBNull.Value)
                Params(22).Value = IsNull(oCompanyInfo.EC, System.DBNull.Value)
                Params(23).Value = IsNull(oCompanyInfo.ED, System.DBNull.Value)
                Params(24).Value = IsNull(oCompanyInfo.TL, System.DBNull.Value)
                Params(25).Value = IsNull(oCompanyInfo.CE, System.DBNull.Value)
                Params(26).Value = IsNull(oCompanyInfo.ACTIVE, System.DBNull.Value)
                Params(27).Value = String.Empty 'IsNull(oContactDatumInfo.CreatedBy, String.Empty)
                Params(28).Value = DBNull.Value 'IsNull(oContactDatumInfo.CreatedOn, CDate("01/01/0001"))
                Params(29).Value = String.Empty  'IsNull(oContactDatumInfo.modifiedBy, String.Empty)
                Params(30).Value = DBNull.Value 'IsNull(oContactDatumInfo.modifiedOn, CDate("01/01/0001"))
                Params(31).Value = IsNull(oCompanyInfo.DELETED, False)

                If oCompanyInfo.ID <= 0 Then
                    Params(32).Value = oCompanyInfo.CREATED_BY
                Else
                    Params(32).Value = oCompanyInfo.LAST_EDITED_BY
                End If

                Params(33).Value = IsNull(oCompanyInfo.PRO_ENGIN_EMAIL, String.Empty)
                Params(34).Value = IsNull(oCompanyInfo.PRO_GEOLO_EMAIL, String.Empty)
                Params(35).Value = IsNull(oCompanyInfo.CM, System.DBNull.Value)
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oCompanyInfo.ID Then
                    oCompanyInfo.ID = Params(0).Value
                End If
                oCompanyInfo.CREATED_BY = IsNull(Params(27).Value, String.Empty)
                oCompanyInfo.DATE_CREATED = IsNull(Params(28).Value, CDate("01/01/0001"))
                oCompanyInfo.LAST_EDITED_BY = AltIsDBNull(Params(29).Value, String.Empty)
                oCompanyInfo.DATE_LAST_EDITED = AltIsDBNull(Params(30).Value, CDate("01/01/0001"))
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetAssociatedLicensees(ByVal compID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Dim dsData As DataSet
            Try
                strSQL = "spGetCOMCompanyLicensees"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = compID
                Params(1).Value = showDeleted
                Params(2).Value = 1
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function searchLicensee(Optional ByVal LicenseeName As String = Nothing, Optional ByVal companyName As String = Nothing, Optional ByVal LicenseeAddress As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal erac As Boolean = False, Optional ByVal irac As Boolean = False, Optional ByVal spName As String = Nothing) As DataSet
            Dim dsData As DataSet
            Dim params() As SqlParameter
            Try
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, spName)
                params(0).Value = IIf(LicenseeName = String.Empty, DBNull.Value, LicenseeName)
                params(1).Value = IIf(companyName = String.Empty, DBNull.Value, companyName)
                params(2).Value = IIf(LicenseeAddress = String.Empty, DBNull.Value, LicenseeAddress)
                params(3).Value = IIf(city = String.Empty, DBNull.Value, city)
                params(4).Value = IIf(state = String.Empty, DBNull.Value, state)
                params(5).Value = IIf(erac = False, DBNull.Value, erac)
                params(6).Value = IIf(irac = False, DBNull.Value, irac)
                params(7).Value = 1
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, spName, params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBsearchTecCompany(Optional ByVal companyName As String = Nothing, Optional ByVal LicenseeAddress As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal erac As Boolean = False, Optional ByVal irac As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spTECCOMSearch"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                params(0).Value = IIf(companyName = String.Empty, DBNull.Value, companyName)
                params(1).Value = IIf(LicenseeAddress = String.Empty, DBNull.Value, LicenseeAddress)
                params(2).Value = IIf(city = String.Empty, DBNull.Value, city)
                params(3).Value = IIf(state = String.Empty, DBNull.Value, state)
                params(4).Value = IIf(erac = False, DBNull.Value, erac)
                params(5).Value = IIf(irac = False, DBNull.Value, irac)
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

    End Class
End Namespace
