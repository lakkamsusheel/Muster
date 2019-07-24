'-------------------------------------------------------------------------------
' MUSTER.DataAccess.LicenseeDB
'   Provides the means for marshalling Address to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       Raf      05/07/05    Original class definition.
'
' Function                  Description
' 
''-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes            ' Reqd for Inserting Null Values on Dates

Namespace MUSTER.DataAccess
    Public Class LicenseeDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByID(ByVal LicenseeID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetCOMLicensees"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Licensee_ID").Value = LicenseeID
                Params("@DELETED").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LicenseeInfo(drSet.Item("LICENSEE_ID"), _
                                       AltIsDBNull(drSet.Item("TITLE"), String.Empty), _
                                       AltIsDBNull(drSet.Item("FIRST_NAME"), String.Empty), _
                                       AltIsDBNull(drSet.Item("MIDDLE_NAME"), String.Empty), _
                                       AltIsDBNull(drSet.Item("LAST_NAME"), String.Empty), _
                                       AltIsDBNull(drSet.Item("SUFFIX"), String.Empty), _
                                       AltIsDBNull(drSet.Item("LICENSE_NUMBER_PREFIX"), String.Empty), _
                                       AltIsDBNull(drSet.Item("LICENSE_NUMBER"), 0), _
                                       AltIsDBNull(drSet.Item("EMAIL_ADDRESS"), String.Empty), _
                                       AltIsDBNull(drSet.Item("ASSOCATED_COMPANY_ID"), 0), _
                                        AltIsDBNull(drSet.Item("HIRE_STATUS"), String.Empty), _
                                       AltIsDBNull(drSet.Item("EMPLOYEE_LETTER"), False), _
                                       AltIsDBNull(drSet.Item("STATUS"), 0), _
                                        AltIsDBNull(drSet.Item("CMSTATUS"), 0), _
                                       AltIsDBNull(drSet.Item("OVERRIDE_EXPIRE"), False), _
                                       AltIsDBNull(drSet.Item("CERT_TYPE"), 0), _
                                       AltIsDBNull(drSet.Item("CMCERT_TYPE"), 0), _
                                       AltIsDBNull(drSet.Item("APP_RECVD_DATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("ORGIN_ISSUED_DATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("ISSUED_DATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("LICENSE_EXPIRE_DATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("EXCEPT_GRANT_DATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                       AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                       AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("DELETED"), False), _
                                       AltIsDBNull(drSet.Item("STATUS_DESC"), String.Empty), _
                                       AltIsDBNull(drSet.Item("CMSTATUS_DESC"), String.Empty), _
                                       AltIsDBNull(drSet.Item("CERT_TYPE_DESC"), String.Empty), _
                                       AltIsDBNull(drSet.Item("CMCERT_TYPE_DESC"), String.Empty), _
                                       AltIsDBNull(drSet.Item("EXTENSION_DEADLINE_DATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("COMPLIANCEMANAGER"), False), _
                                       AltIsDBNull(drSet.Item("ISLICENSEE"), False), _
                                       AltIsDBNull(drSet.Item("INITCERTDATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("INITCERTBY"), 0), _
                                       AltIsDBNull(drSet.Item("INITCERTBYDESC"), String.Empty), _
                                       AltIsDBNull(drSet.Item("RETRAINDATE1"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("RETRAINDATE2"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("RETRAINDATE3"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("REVOKEDATE"), CDate("01/01/0001")), _
                                       AltIsDBNull(drSet.Item("RETRAINREQDATE"), CDate("01/01/0001")))
                Else
                    Return New MUSTER.Info.LicenseeInfo
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByCompanyID(Optional ByVal CompanyID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim strVal As String = ""
            Dim Params As Collection
            Dim colLicensee As New MUSTER.Info.LicenseeCollection
            Try
                strSQL = "spGetCOMCompanyLicensees"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@COMPANYID").Value = CompanyID
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        Dim LicenseeInfo As New MUSTER.Info.LicenseeInfo(drSet.Item("LICENSEE_ID"), _
                                            AltIsDBNull(drSet.Item("TITLE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("FIRST_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("MIDDLE_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LAST_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("SUFFIX"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LICENSE_NUMBER_PREFIX"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LICENSE_NUMBER"), 0), _
                                            AltIsDBNull(drSet.Item("EMAIL_ADDRESS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("ASSOCATED_COMPANY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("HIRE_STATUS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("EMPLOYEE_LETTER"), False), _
                                            AltIsDBNull(drSet.Item("STATUS"), 0), _
                                            AltIsDBNull(drSet.Item("CMSTATUS"), 0), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_EXPIRE"), False), _
                                            AltIsDBNull(drSet.Item("CERT_TYPE"), 0), _
                                            AltIsDBNull(drSet.Item("CMCERT_TYPE"), 0), _
                                            AltIsDBNull(drSet.Item("APP_RECVD_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("ORGIN_ISSUED_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("ISSUED_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LICENSE_EXPIRE_DATE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("EXCEPT_GRANT_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("STATUS_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("CMSTATUS_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("CERT_TYPE_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("CMCERT_TYPE_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("EXTENSION_DEADLINE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("COMPLIANCEMANAGER"), False), _
                                            AltIsDBNull(drSet.Item("ISLICENSEE"), False), _
                                            AltIsDBNull(drSet.Item("INITCERTDATE"), String.Empty), _
                                        AltIsDBNull(drSet.Item("INITCERTBY"), 0), _
                                        AltIsDBNull(drSet.Item("INITCERTBYDESC"), String.Empty), _
                                        AltIsDBNull(drSet.Item("RETRAINDATE1"), String.Empty), _
                                        AltIsDBNull(drSet.Item("RETRAINDATE2"), String.Empty), _
                                        AltIsDBNull(drSet.Item("RETRAINDATE3"), String.Empty), _
                                        AltIsDBNull(drSet.Item("REVOKEDATE"), String.Empty), _
                                        AltIsDBNull(drSet.Item("RETRAINREQDATE"), String.Empty))
                        colLicensee.Add(LicenseeInfo)
                    End While
                End If
                Return colLicensee
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBManagerGetByCompanyID(Optional ByVal CompanyID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LicenseeCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim strVal As String = ""
            Dim Params As Collection
            Dim colLicensee As New MUSTER.Info.LicenseeCollection
            Try
                strSQL = "spGetCOMCompanyManagers"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@COMPANYID").Value = CompanyID
                Params("@Deleted").Value = showDeleted
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        Dim LicenseeInfo As New MUSTER.Info.LicenseeInfo(drSet.Item("LICENSEE_ID"), _
                                            AltIsDBNull(drSet.Item("TITLE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("FIRST_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("MIDDLE_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LAST_NAME"), String.Empty), _
                                            AltIsDBNull(drSet.Item("SUFFIX"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LICENSE_NUMBER_PREFIX"), String.Empty), _
                                            AltIsDBNull(drSet.Item("LICENSE_NUMBER"), 0), _
                                            AltIsDBNull(drSet.Item("EMAIL_ADDRESS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("ASSOCATED_COMPANY_ID"), 0), _
                                            AltIsDBNull(drSet.Item("HIRE_STATUS"), String.Empty), _
                                            AltIsDBNull(drSet.Item("EMPLOYEE_LETTER"), False), _
                                            AltIsDBNull(drSet.Item("STATUS"), 0), _
                                            AltIsDBNull(drSet.Item("CMSTATUS"), 0), _
                                            AltIsDBNull(drSet.Item("OVERRIDE_EXPIRE"), False), _
                                            AltIsDBNull(drSet.Item("CERT_TYPE"), 0), _
                                            AltIsDBNull(drSet.Item("CMCERT_TYPE"), 0), _
                                            AltIsDBNull(drSet.Item("APP_RECVD_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("ORGIN_ISSUED_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("ISSUED_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LICENSE_EXPIRE_DATE"), String.Empty), _
                                            AltIsDBNull(drSet.Item("EXCEPT_GRANT_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("DELETED"), False), _
                                            AltIsDBNull(drSet.Item("STATUS_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("CMSTATUS_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("CERT_TYPE_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("CMCERT_TYPE_DESC"), String.Empty), _
                                            AltIsDBNull(drSet.Item("EXTENSION_DEADLINE_DATE"), CDate("01/01/0001")), _
                                            AltIsDBNull(drSet.Item("COMPLIANCEMANAGER"), False), _
                                            AltIsDBNull(drSet.Item("ISLICENSEE"), False), _
                                            AltIsDBNull(drSet.Item("INITCERTDATE"), String.Empty), _
                                       AltIsDBNull(drSet.Item("INITCERTBY"), 0), _
                                       AltIsDBNull(drSet.Item("INITCERTBYDESC"), String.Empty), _
                                       AltIsDBNull(drSet.Item("RETRAINDATE1"), String.Empty), _
                                       AltIsDBNull(drSet.Item("RETRAINDATE2"), String.Empty), _
                                       AltIsDBNull(drSet.Item("RETRAINDATE3"), String.Empty), _
                                       AltIsDBNull(drSet.Item("REVOKEDATE"), String.Empty), _
                                       AltIsDBNull(drSet.Item("RETRAINREQDATE"), String.Empty))
                        colLicensee.Add(LicenseeInfo)
                    End While
                End If
                Return colLicensee
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetCMList(ByVal facID As Int64) As DataSet
            Dim strSQL As String
            Dim Params(1) As SqlParameter
            Dim dsData As DataSet
            Try
                If facID = 0 Then
                    dsData = New DataSet
                    Return dsData
                End If
                strSQL = "spGetComplianceManagerList"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                '      Params(1).Value = showDeleted
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetCheckList(ByVal LicenseeID As Integer, ByVal LicenseeType As Integer) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub Put(ByRef oLicenseeInfo As MUSTER.Info.LicenseeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Dim dtTempDate As Date
            Dim tempDateStr As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Licensee, Integer))) Then
                    returnVal = "You do not have rights to save Licensee."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCOMLicensees"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IsNull(oLicenseeInfo.ID, System.DBNull.Value)
                Params(1).Value = IsNull(oLicenseeInfo.TITLE, String.Empty)
                Params(2).Value = IsNull(oLicenseeInfo.FIRST_NAME, String.Empty)
                Params(3).Value = IsNull(oLicenseeInfo.MIDDLE_NAME, String.Empty)
                Params(4).Value = IsNull(oLicenseeInfo.LAST_NAME, String.Empty)
                Params(5).Value = IsNull(oLicenseeInfo.SUFFIX, String.Empty)
                oLicenseeInfo.LICENSEE_NUMBER_PREFIX = ""
                If oLicenseeInfo.CertTypeDesc.ToUpper = "CLOSURE" Then
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "C"
                ElseIf oLicenseeInfo.CertTypeDesc.ToUpper = "INSTALL" Then
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "N"
                Else
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "_"
                End If
                If oLicenseeInfo.HIRE_STATUS.StartsWith("HX") Or _
                    oLicenseeInfo.HIRE_STATUS.StartsWith("HB") Or _
                    oLicenseeInfo.HIRE_STATUS.StartsWith("RX") Or _
                    oLicenseeInfo.HIRE_STATUS.StartsWith("RB") Then
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += oLicenseeInfo.HIRE_STATUS.Substring(0, 2)
                End If
                'If oLicenseeInfo.HIRE_STATUS = "HX - For Hire - owner" Then
                '    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "HX"
                'ElseIf oLicenseeInfo.HIRE_STATUS = "HB- For Hire - Employee" Then
                '    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "HB"
                'ElseIf oLicenseeInfo.HIRE_STATUS = "RX - Not for Hire" Then
                '    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "RX"
                'End If
                Params(6).Value = IsNull(oLicenseeInfo.LICENSEE_NUMBER_PREFIX, String.Empty)
                If oLicenseeInfo.LICENSEE_NUMBER = String.Empty Then
                    Params(7).Value = System.DBNull.Value
                Else
                    Params(7).Value = Integer.Parse(oLicenseeInfo.LICENSEE_NUMBER)
                End If

                Params(8).Value = IsNull(oLicenseeInfo.EMAIL_ADDRESS, String.Empty)
                Params(9).Value = IsNull(oLicenseeInfo.ASSOCATED_COMPANY_ID, System.DBNull.Value)
                Params(10).Value = IsNull(oLicenseeInfo.HIRE_STATUS, String.Empty)
                Params(11).Value = IsNull(oLicenseeInfo.EMPLOYEE_LETTER, System.DBNull.Value)
                Params(12).Value = IsNull(oLicenseeInfo.STATUS, 0)
                Params(13).Value = IsNull(oLicenseeInfo.OVERRIDE_EXPIRE, System.DBNull.Value)
                Params(14).Value = IsNull(oLicenseeInfo.CertType, 0)
                If Date.Compare(oLicenseeInfo.APP_RECVD_DATE, dtTempDate) = 0 Then
                    Params(15).Value = SqlDateTime.Null
                Else
                    Params(15).Value = oLicenseeInfo.APP_RECVD_DATE
                End If
                If Date.Compare(oLicenseeInfo.ORIGIN_ISSUED_DATE, dtTempDate) = 0 Then
                    Params(16).Value = SqlDateTime.Null
                Else
                    Params(16).Value = oLicenseeInfo.ORIGIN_ISSUED_DATE
                End If
                If Date.Compare(oLicenseeInfo.ISSUED_DATE, dtTempDate) = 0 Then
                    Params(17).Value = SqlDateTime.Null
                Else
                    Params(17).Value = oLicenseeInfo.ISSUED_DATE
                End If
                If (oLicenseeInfo.LICENSE_EXPIRE_DATE.Length() = 0) OrElse (oLicenseeInfo.LICENSE_EXPIRE_DATE = "#12:00:00AM#") Then
                    Params(18).Value = SqlDateTime.Null
                Else
                    Params(18).Value = oLicenseeInfo.LICENSE_EXPIRE_DATE
                End If
                ' If Date.Compare(oLicenseeInfo.LICENSE_EXPIRE_DATE, dtTempDate) = 0 Then
                'Params(18).Value = SqlDateTime.Null
                'Else
                '   Params(18).Value = oLicenseeInfo.LICENSE_EXPIRE_DATE
                'End If
                If Date.Compare(oLicenseeInfo.EXCEPT_GRANT_DATE, dtTempDate) = 0 Then
                    Params(19).Value = SqlDateTime.Null
                Else
                    Params(19).Value = oLicenseeInfo.EXCEPT_GRANT_DATE
                End If
                Params(20).Value = DBNull.Value
                Params(21).Value = DBNull.Value
                Params(22).Value = DBNull.Value
                Params(23).Value = DBNull.Value
                Params(24).Value = IsNull(oLicenseeInfo.DELETED, False) 'IsNull(oContactDatumInfo.modifiedOn, CDate("01/01/0001"))

                If oLicenseeInfo.ID <= 0 Then
                    Params(25).Value = oLicenseeInfo.CREATED_BY
                Else
                    Params(25).Value = oLicenseeInfo.LAST_EDITED_BY
                End If
                If Date.Compare(oLicenseeInfo.EXTENSION_DEADLINE_DATE, dtTempDate) = 0 Then
                    Params(26).Value = SqlDateTime.Null
                Else
                    Params(26).Value = oLicenseeInfo.EXTENSION_DEADLINE_DATE
                End If
                Params(27).Value = IsNull(oLicenseeInfo.COMPLIANCEMANAGER, 0) 'ComplianceManager
                Params(28).Value = 1 'IsLicensee
                If (oLicenseeInfo.INITCERTDATE.Length() = 0) OrElse (oLicenseeInfo.INITCERTDATE = "#12:00:00AM#") Then
                    Params(29).Value = SqlDateTime.Null
                Else
                    Params(29).Value = oLicenseeInfo.INITCERTDATE
                End If

                Params(30).Value = IsNull(oLicenseeInfo.INITCERTBY, DBNull.Value)
                If (oLicenseeInfo.RETRAINDATE1.Length() = 0) OrElse (oLicenseeInfo.RETRAINDATE1 = "#12:00:00AM#") Then
                    Params(31).Value = SqlDateTime.Null
                Else
                    Params(31).Value = oLicenseeInfo.RETRAINDATE1
                End If
                If (oLicenseeInfo.RETRAINDATE2.Length() = 0) OrElse (oLicenseeInfo.RETRAINDATE2 = "#12:00:00AM#") Then
                    Params(32).Value = SqlDateTime.Null
                Else
                    Params(32).Value = oLicenseeInfo.RETRAINDATE2
                End If
                If (oLicenseeInfo.RETRAINDATE3.Length() = 0) OrElse (oLicenseeInfo.RETRAINDATE3 = "#12:00:00AM#") Then
                    Params(33).Value = SqlDateTime.Null
                Else
                    Params(33).Value = oLicenseeInfo.RETRAINDATE3
                End If
                If (oLicenseeInfo.REVOKEDATE.Length() = 0) OrElse (oLicenseeInfo.REVOKEDATE = "#12:00:00AM#") Then
                    Params(34).Value = SqlDateTime.Null
                Else
                    Params(34).Value = oLicenseeInfo.REVOKEDATE
                End If
                Params(35).Value = IsNull(oLicenseeInfo.CMSTATUS, 0)
                Params(36).Value = IsNull(oLicenseeInfo.CMCertType, 0)
                If (oLicenseeInfo.RETRAINREQDATE.Length() = 0) OrElse (oLicenseeInfo.RETRAINREQDATE = "#12:00:00AM#") Then
                    Params(37).Value = SqlDateTime.Null
                Else
                    Params(37).Value = oLicenseeInfo.RETRAINREQDATE
                End If
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oLicenseeInfo.ID Then
                    oLicenseeInfo.ID = Params(0).Value
                End If
                oLicenseeInfo.LICENSEE_NUMBER = AltIsDBNull(Params(7).Value, 0)
                oLicenseeInfo.CREATED_BY = IsNull(Params(20).Value, String.Empty)
                oLicenseeInfo.DATE_CREATED = IsNull(Params(21).Value, CDate("01/01/0001"))
                oLicenseeInfo.LAST_EDITED_BY = AltIsDBNull(Params(22).Value, String.Empty)
                oLicenseeInfo.DATE_LAST_EDITED = AltIsDBNull(Params(23).Value, CDate("01/01/0001"))
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub CMPut(ByRef oLicenseeInfo As MUSTER.Info.LicenseeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal CompanyID As Integer = 0, Optional ByVal CompanyAddrID As Integer = 0)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Dim dtTempDate As Date
            Dim tempDateStr As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Licensee, Integer))) Then
                    returnVal = "You do not have rights to save Licensee."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCOMLicensees"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IsNull(oLicenseeInfo.ID, System.DBNull.Value)
                Params(1).Value = IsNull(oLicenseeInfo.TITLE, String.Empty)
                Params(2).Value = IsNull(oLicenseeInfo.FIRST_NAME, String.Empty)
                Params(3).Value = IsNull(oLicenseeInfo.MIDDLE_NAME, String.Empty)
                Params(4).Value = IsNull(oLicenseeInfo.LAST_NAME, String.Empty)
                Params(5).Value = IsNull(oLicenseeInfo.SUFFIX, String.Empty)
                oLicenseeInfo.LICENSEE_NUMBER_PREFIX = ""
                If oLicenseeInfo.CertTypeDesc.ToUpper = "CLOSURE" Then
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "C"
                ElseIf oLicenseeInfo.CertTypeDesc.ToUpper = "INSTALL" Then
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "N"
                Else
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "_"
                End If
                If oLicenseeInfo.HIRE_STATUS.StartsWith("HX") Or _
                    oLicenseeInfo.HIRE_STATUS.StartsWith("HB") Or _
                    oLicenseeInfo.HIRE_STATUS.StartsWith("RX") Or _
                    oLicenseeInfo.HIRE_STATUS.StartsWith("RB") Then
                    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += oLicenseeInfo.HIRE_STATUS.Substring(0, 2)
                End If
                'If oLicenseeInfo.HIRE_STATUS = "HX - For Hire - owner" Then
                '    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "HX"
                'ElseIf oLicenseeInfo.HIRE_STATUS = "HB- For Hire - Employee" Then
                '    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "HB"
                'ElseIf oLicenseeInfo.HIRE_STATUS = "RX - Not for Hire" Then
                '    oLicenseeInfo.LICENSEE_NUMBER_PREFIX += "RX"
                'End If
                Params(6).Value = IsNull(oLicenseeInfo.LICENSEE_NUMBER_PREFIX, String.Empty)
                If oLicenseeInfo.LICENSEE_NUMBER = String.Empty Then
                    Params(7).Value = System.DBNull.Value
                Else
                    Params(7).Value = Integer.Parse(oLicenseeInfo.LICENSEE_NUMBER)
                End If

                Params(8).Value = IsNull(oLicenseeInfo.EMAIL_ADDRESS, String.Empty)
                Params(9).Value = IsNull(oLicenseeInfo.ASSOCATED_COMPANY_ID, System.DBNull.Value)
                Params(10).Value = IsNull(oLicenseeInfo.HIRE_STATUS, String.Empty)
                Params(11).Value = IsNull(oLicenseeInfo.EMPLOYEE_LETTER, System.DBNull.Value)
                Params(12).Value = IsNull(oLicenseeInfo.STATUS, 0)
                Params(13).Value = IsNull(oLicenseeInfo.OVERRIDE_EXPIRE, System.DBNull.Value)
                Params(14).Value = IsNull(oLicenseeInfo.CertType, 0)
                If Date.Compare(oLicenseeInfo.APP_RECVD_DATE, dtTempDate) = 0 Then
                    Params(15).Value = SqlDateTime.Null
                Else
                    Params(15).Value = oLicenseeInfo.APP_RECVD_DATE
                End If
                If Date.Compare(oLicenseeInfo.ORIGIN_ISSUED_DATE, dtTempDate) = 0 Then
                    Params(16).Value = SqlDateTime.Null
                Else
                    Params(16).Value = oLicenseeInfo.ORIGIN_ISSUED_DATE
                End If
                If Date.Compare(oLicenseeInfo.ISSUED_DATE, dtTempDate) = 0 Then
                    Params(17).Value = SqlDateTime.Null
                Else
                    Params(17).Value = oLicenseeInfo.ISSUED_DATE
                End If
                If (oLicenseeInfo.LICENSE_EXPIRE_DATE.Length() = 0) OrElse (oLicenseeInfo.LICENSE_EXPIRE_DATE = "#12:00:00AM#") Then
                    Params(18).Value = SqlDateTime.Null
                Else
                    Params(18).Value = oLicenseeInfo.LICENSE_EXPIRE_DATE
                End If
                ' If Date.Compare(oLicenseeInfo.LICENSE_EXPIRE_DATE, dtTempDate) = 0 Then
                'Params(18).Value = SqlDateTime.Null
                'Else
                '   Params(18).Value = oLicenseeInfo.LICENSE_EXPIRE_DATE
                'End If
                If Date.Compare(oLicenseeInfo.EXCEPT_GRANT_DATE, dtTempDate) = 0 Then
                    Params(19).Value = SqlDateTime.Null
                Else
                    Params(19).Value = oLicenseeInfo.EXCEPT_GRANT_DATE
                End If
                Params(20).Value = DBNull.Value
                Params(21).Value = DBNull.Value
                Params(22).Value = DBNull.Value
                Params(23).Value = DBNull.Value
                Params(24).Value = IsNull(oLicenseeInfo.DELETED, False) 'IsNull(oContactDatumInfo.modifiedOn, CDate("01/01/0001"))

                If oLicenseeInfo.ID <= 0 Then
                    Params(25).Value = oLicenseeInfo.CREATED_BY
                Else
                    Params(25).Value = oLicenseeInfo.LAST_EDITED_BY
                End If
                If Date.Compare(oLicenseeInfo.EXTENSION_DEADLINE_DATE, dtTempDate) = 0 Then
                    Params(26).Value = SqlDateTime.Null
                Else
                    Params(26).Value = oLicenseeInfo.EXTENSION_DEADLINE_DATE
                End If
                Params(27).Value = oLicenseeInfo.COMPLIANCEMANAGER 'ComplianceManager
                Params(28).Value = oLicenseeInfo.ISLICENSEE 'IsLicensee
                If (oLicenseeInfo.INITCERTDATE.Length() = 0) OrElse (oLicenseeInfo.INITCERTDATE = "#12:00:00AM#") OrElse (oLicenseeInfo.INITCERTDATE = "#12:00:00 AM#") Then
                    Params(29).Value = SqlDateTime.Null
                Else
                    Params(29).Value = oLicenseeInfo.INITCERTDATE
                End If

                Params(30).Value = IsNull(oLicenseeInfo.INITCERTBY, DBNull.Value)
                If (oLicenseeInfo.RETRAINDATE1.Length() = 0) OrElse (oLicenseeInfo.RETRAINDATE1 = "#12:00:00AM#") OrElse (oLicenseeInfo.RETRAINDATE1 = "#12:00:00 AM#") Then
                    Params(31).Value = SqlDateTime.Null
                Else
                    Params(31).Value = oLicenseeInfo.RETRAINDATE1
                End If
                If (oLicenseeInfo.RETRAINDATE2.Length() = 0) OrElse (oLicenseeInfo.RETRAINDATE2 = "#12:00:00AM#") OrElse (oLicenseeInfo.RETRAINDATE2 = "#12:00:00 AM#") Then
                    Params(32).Value = SqlDateTime.Null
                Else
                    Params(32).Value = oLicenseeInfo.RETRAINDATE2
                End If
                If (oLicenseeInfo.RETRAINDATE3.Length() = 0) OrElse (oLicenseeInfo.RETRAINDATE3 = "#12:00:00AM#") OrElse (oLicenseeInfo.RETRAINDATE3 = "#12:00:00 AM#") Then
                    Params(33).Value = SqlDateTime.Null
                Else
                    Params(33).Value = oLicenseeInfo.RETRAINDATE3
                End If
                If (oLicenseeInfo.REVOKEDATE.Length() = 0) OrElse (oLicenseeInfo.REVOKEDATE = "#12:00:00AM#") OrElse (oLicenseeInfo.REVOKEDATE = "#12:00:00 AM#") Then
                    Params(34).Value = SqlDateTime.Null
                Else
                    Params(34).Value = oLicenseeInfo.REVOKEDATE
                End If
                Params(35).Value = IsNull(oLicenseeInfo.CMSTATUS, 0)
                Params(36).Value = IsNull(oLicenseeInfo.CMCertType, 0)
                If (oLicenseeInfo.RETRAINREQDATE.Length() = 0) OrElse (oLicenseeInfo.RETRAINREQDATE = "#12:00:00AM#") OrElse (oLicenseeInfo.RETRAINREQDATE = "#12:00:00 AM#") Then
                    Params(37).Value = SqlDateTime.Null
                Else
                    Params(37).Value = oLicenseeInfo.RETRAINREQDATE
                End If
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(0).Value <> oLicenseeInfo.ID Then
                    oLicenseeInfo.ID = Params(0).Value
                End If
                oLicenseeInfo.LICENSEE_NUMBER = AltIsDBNull(Params(7).Value, 0)
                oLicenseeInfo.CREATED_BY = IsNull(Params(20).Value, String.Empty)
                oLicenseeInfo.DATE_CREATED = IsNull(Params(21).Value, CDate("01/01/0001"))
                oLicenseeInfo.LAST_EDITED_BY = AltIsDBNull(Params(22).Value, String.Empty)
                oLicenseeInfo.DATE_LAST_EDITED = AltIsDBNull(Params(23).Value, CDate("01/01/0001"))
                'Add spPutCOMCompanyLicensee for add/edit UST compliance manager
                'Added by Hua Cao on 09/10/2012
                If CompanyID <> 0 Then
                    Dim ParamsComLic() As SqlParameter
                    Dim strSQLComLic As String
                    strSQLComLic = "spPutCOMCompanyLicensee"
                    ParamsComLic = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQLComLic)

                    ParamsComLic(0).Value = 0
                    ParamsComLic(1).Value = CompanyID
                    ParamsComLic(2).Value = oLicenseeInfo.ID
                    ParamsComLic(3).Value = CompanyAddrID
                    ParamsComLic(4).Value = 0
                    ParamsComLic(5).Value = DBNull.Value
                    ParamsComLic(6).Value = DBNull.Value
                    ParamsComLic(7).Value = DBNull.Value
                    ParamsComLic(8).Value = DBNull.Value

                    If oLicenseeInfo.ID <= 0 Then
                        ParamsComLic(9).Value = oLicenseeInfo.CREATED_BY
                    Else
                        ParamsComLic(9).Value = oLicenseeInfo.LAST_EDITED_BY
                    End If

                    SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQLComLic, ParamsComLic)
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

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
        Public Function GetLicenseeList(Optional ByVal showDeleted As Boolean = False, Optional ByVal companyID As Integer = -1) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Try
                strSQl = "SELECT * FROM dbo.vCOM_LICENSEELIST"


                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQl)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetManagerList(Optional ByVal showDeleted As Boolean = False, Optional ByVal companyID As Integer = -1) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Try
                strSQl = "SELECT * FROM dbo.vCOM_MANAGERLIST"


                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQl)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetLicenseesByType(ByVal _dt As DateTime, Optional ByVal InputType As String = "RENEWAL", Optional ByVal showDeleted As Boolean = False, Optional ByVal showLetterGeneratedOnly As Int16 = -1) As DataSet
            Dim dsData As DataSet
            Dim strSQl As String
            Dim params() As SqlParameter
            Try
                strSQl = "spGetCOMRenewalReminderExpiration"
                params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQl)
                params(0).Value = InputType
                params(1).Value = showDeleted
                If Date.Compare(_dt, CDate("01/01/0001")) = 0 Then
                    params(2).Value = DBNull.Value
                Else
                    params(2).Value = _dt.Date
                End If
                If showLetterGeneratedOnly = -1 Then
                    params(3).Value = DBNull.Value
                Else
                    params(3).Value = showLetterGeneratedOnly
                End If
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQl, params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetPriorCompanies(ByVal licID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Dim dsData As DataSet
            Try
                strSQL = "spGetCOMLicenseePriorCompanies"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = licID
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function ProcessRRE(Optional ByVal LicenseeID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As Boolean
            Dim strSQL As String
            Dim Params() As SqlParameter
            Dim dsData As DataSet
            Dim strResult As String = String.Empty
            Try
                strSQL = "spCOMProcessRRE"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = LicenseeID
                Params(1).Value = showDeleted

                strResult = SqlHelper.ExecuteScalar(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If strResult.ToUpper = "SUCCESS" Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub UpdateRenewals(ByVal LicenseeIDs As String, ByVal InputType As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spCOMModifyRenewalLicensees"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(LicenseeIDs = String.Empty, DBNull.Value, LicenseeIDs)
                Params(1).Value = InputType
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function DBGetLicenseesToBeProcessed() As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Try
                strSQL = "spPROCESSLICENSEES"
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetLicenseeStatus(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String = String.Empty
            Try
                If showBlankPropertyName Then strSQL = SqlHelper.AddBlankPropertyName(164) + " UNION "
                strSQL += "SELECT * FROM vCOM_LICENSEESTATUS WHERE PROPERTY_ACTIVE = " + IIf(showInActive, "PROPERTY_ACTIVE", "'YES'")
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetManagerStatus(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String = String.Empty
            Try
                If showBlankPropertyName Then strSQL = SqlHelper.AddBlankPropertyName(164) + " UNION "
                strSQL += "SELECT * FROM vCOM_MANAGERSTATUS WHERE PROPERTY_ACTIVE = " + IIf(showInActive, "PROPERTY_ACTIVE", "'YES'")
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetLicenseeCertificationType(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String = String.Empty
            Try
                If showBlankPropertyName Then strSQL = SqlHelper.AddBlankPropertyName(165) + " UNION "
                strSQL += "SELECT * FROM vCOM_LICENSEECERTTYPE WHERE PROPERTY_ACTIVE = " + IIf(showInActive, "PROPERTY_ACTIVE", "'YES'")
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetManagerInitCertBy(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String = String.Empty
            Try
                If showBlankPropertyName Then strSQL = SqlHelper.AddBlankPropertyName(165) + " UNION "
                strSQL += "SELECT * FROM vCOM_MANAGERINITCERTBY WHERE PROPERTY_ACTIVE = " + IIf(showInActive, "PROPERTY_ACTIVE", "'YES'")
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetLicenseeHireStatus(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String = String.Empty
            Try
                If showBlankPropertyName Then strSQL = SqlHelper.AddBlankPropertyName(166) + " UNION "
                strSQL += "SELECT * FROM vCOM_LICENSEEHIRESTATUS WHERE PROPERTY_ACTIVE = " + IIf(showInActive, "PROPERTY_ACTIVE", "'YES'")
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace
