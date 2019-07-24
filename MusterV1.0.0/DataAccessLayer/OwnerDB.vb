'-------------------------------------------------------------------------------
' MUSTER.DataAccess.OwnerDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR     12/03/04      Original class definition.
'  1.1        AN      12/30/04      Added Try catch and Exception Handling/Logging
'  1.2        MNR      1/07/05      Added param for Deleted field in Put(..) function
'  1.3        MNR      1/11/05      Added GetFaciliies(ByVal OwnrID As Integer) function
'  1.4        MNR      1/17/05      Renamed GetFaciliies(..) function to GetFacilitiesTankStatus(..)
'  1.5        EN      02/10/05      Modified 01/01/1901  to 01/01/0001 
'  1.6        AB      02/15/05      Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByID
'  1.7        AB      02/15/05      Added Finally to the Try/Catch to close all datareaders
'  1.8        AB      02/16/05      Removed any IsNull calls for fields the DB requires
'  1.9        AB      02/18/05      Set all parameters for SP, that are not required, to NULL
'  1.10       AB      02/28/05      Modified Get functions based upon changes made to 
'                                     make several nullable fields non-nullable
'  1.11       AB      03/04/05      Added GetCAPParticipationLevel
'  1.12       JVC2    03/07/05      Changed Modified By and Modified On to reflect state after PUT
'  1.13       MR      03/07/05      Modified PUT function to pass EmptyString or NULL Dates Values for DBNULL
'  1.14       MR      03/14/05      Changed Created By and Created On to reflect state after PUT
'  1.15       AB      03/21/05      Added GetFacilitiesLUSTSummary
'
' Function                  Description
' GetAllInfo()      Returns an EntityCollection containing all Entity objects in the repository
' DBGetByID(ID)     Returns an EntityInfo object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(Entity)       Saves the Entity passed as an argument, to the DB
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class OwnerDB
        Private _strConn
        Private MusterException As MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            Try
                If MusterXCEP Is Nothing Then
                    MusterException = New MUSTER.Exceptions.MusterExceptions
                Else
                    MusterException = MusterXCEP
                End If
                If strDBConn = String.Empty Then
                    Dim oCnn As New ConnectionSettings
                    _strConn = oCnn.cnString
                    oCnn = Nothing
                Else
                    _strConn = strDBConn
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
#End Region
        Public Function GetAllInfo() As MUSTER.Info.OwnersCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetOwner"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Owner_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                'SqlHelper.ExecuteReader(_strConn, CommandType.Text, "select * from tblREG_OWNER ORDER BY OWNER_ID")

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.OwnersCollection
                While drSet.Read
                    Dim oOwnerInfo As New MUSTER.Info.OwnerInfo(drSet.Item("OWNER_ID"), _
                                                                AltIsDBNull(drSet.Item("ORGANIZATION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("PERSON_ID"), 0), _
                                                                drSet.Item("PHONE_NUMBER_ONE"), _
                                                                drSet.Item("PHONE_NUMBER_TWO"), _
                                                                drSet.Item("FAX_NUMBER"), _
                                                                drSet.Item("EMAIL_ADDRESS"), _
                                                                drSet.Item("EMAIL_ADDRESS_PERSONAL"), _
                                                                drSet.Item("ADDRESS_ID"), _
                                                                AltIsDBNull(drSet.Item("DATE_CAP_SIGNUP"), CDate("01/01/0001")), _
                                                                AltIsDBNull(drSet.Item("CAP_CURRENT_STATUS"), False), _
                                                                drSet.Item("OWNER_TYPE"), _
                                                                AltIsDBNull(drSet.Item("BP2K_OWNER_TYPE"), 0), _
                                                                AltIsDBNull(drSet.Item("FEES_PROFILE_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("FEES_STATUS"), False), _
                                                                AltIsDBNull(drSet.Item("COMPLIANCE_PROFILE_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("COMPLIACE_STATUS"), False), _
                                                               drSet.Item("ACTIVE"), _
                                                               AltIsDBNull(drSet.Item("FEE_ACTIVE"), False), _
                                                                AltIsDBNull(drSet.Item("ENSITE_ORGANIZATION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("ENSITE_PERSON_ID"), 0), _
                                                               AltIsDBNull(drSet.Item("ENSITE_AGENCY_INTEREST_ID"), False), _
                                                               drSet.Item("CUST_ENTITY_CODE"), _
                                                                drSet.Item("CUST_TYPE_CODE"), _
                                                               drSet.Item("DELETED"), _
                                                                drSet.Item("CREATED_BY"), _
                                                                drSet.Item("DATE_CREATED"), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                drSet.Item("OWNER_L2C_SNIPPET"), _
                                                                AltIsDBNull(drSet.Item("BP2K_OWNER_ID"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("CAP_PARTICIPATION_LEVEL"), String.Empty))

                    'AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE"), String.Empty), _
                    'AltIsDBNull(drSet.Item("ADDRESS_TWO"), String.Empty), _
                    'AltIsDBNull(drSet.Item("CITY"), String.Empty), _
                    'AltIsDBNull(drSet.Item("STATE"), String.Empty), _
                    'AltIsDBNull(drSet.Item("ZIP"), String.Empty), _
                    'AltIsDBNull(drSet.Item("FIPS_CODE"), String.Empty), _

                    colEntities.Add(oOwnerInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Integer, Optional ByVal showDeleted As Boolean = False, _
                                Optional ByVal intInspectionID As Integer = Integer.MinValue) As MUSTER.Info.OwnerInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblREG_OWNER WHERE OWNER_ID = '" + strVal + "'"
            'If Not showDeleted Then
            '    strSQL += " AND DELETED <> 1"
            'End If

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.OwnerInfo
                End If
                If Not (intInspectionID = Integer.MinValue) Then
                    strSQL = "spGetOwnerArchived"
                    Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                    Params("@Owner_ID").Value = nVal
                    Params("@Inspection_ID").Value = intInspectionID
                Else
                    strSQL = "spGetOwner"
                    Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                    Params("@Owner_ID").Value = nVal
                    Params("@OrderBy").Value = 1
                    Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)
                End If
                strVal = nVal


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.OwnerInfo(drSet.Item("OWNER_ID"), _
                                                                AltIsDBNull(drSet.Item("ORGANIZATION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("PERSON_ID"), 0), _
                                                                drSet.Item("PHONE_NUMBER_ONE"), _
                                                                drSet.Item("PHONE_NUMBER_TWO"), _
                                                                drSet.Item("FAX_NUMBER"), _
                                                                drSet.Item("EMAIL_ADDRESS"), _
                                                                drSet.Item("EMAIL_ADDRESS_PERSONAL"), _
                                                                drSet.Item("ADDRESS_ID"), _
                                                                AltIsDBNull(drSet.Item("DATE_CAP_SIGNUP"), CDate("01/01/0001")), _
                                                               (AltIsDBNull(drSet.Item("CAP_CURRENT_STATUS"), False)), _
                                                                drSet.Item("OWNER_TYPE"), _
                                                                AltIsDBNull(drSet.Item("BP2K_OWNER_TYPE"), 0), _
                                                                AltIsDBNull(drSet.Item("FEES_PROFILE_ID"), 0), _
                                                               (AltIsDBNull(drSet.Item("FEES_STATUS"), False)), _
                                                                AltIsDBNull(drSet.Item("COMPLIANCE_PROFILE_ID"), 0), _
                                                               (AltIsDBNull(drSet.Item("COMPLIACE_STATUS"), False)), _
                                                               drSet.Item("ACTIVE"), _
                                                               (AltIsDBNull(drSet.Item("FEE_ACTIVE"), False)), _
                                                                AltIsDBNull(drSet.Item("ENSITE_ORGANIZATION_ID"), 0), _
                                                                AltIsDBNull(drSet.Item("ENSITE_PERSON_ID"), 0), _
                                                               (AltIsDBNull(drSet.Item("ENSITE_AGENCY_INTEREST_ID"), False)), _
                                                               drSet.Item("CUST_ENTITY_CODE"), _
                                                                drSet.Item("CUST_TYPE_CODE"), _
                                                               (drSet.Item("DELETED")), _
                                                                drSet.Item("CREATED_BY"), _
                                                                drSet.Item("DATE_CREATED"), _
                                                                AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                drSet.Item("OWNER_L2C_SNIPPET"), _
                                                                AltIsDBNull(drSet.Item("BP2K_OWNER_ID"), String.Empty), _
                                                                AltIsDBNull(drSet.Item("CAP_PARTICIPATION_LEVEL"), String.Empty))

                Else

                    Return New MUSTER.Info.OwnerInfo
                End If
            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetPreviousFacs(ByVal nVal As Integer) As DataSet
            Dim dsSet As DataSet
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal <> 0 Then
                    strSQL = "spGetPreviousFacilities"
                    Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                    Params("@OwnerID").Value = nVal

                    dsSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                End If

            Catch ex As Exception

                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
            End Try
            Return dsSet

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
        Public Function DBGetDS(ByVal ownerID As Integer, Optional ByVal [Module] As String = "", Optional ByVal showDeleted As Boolean = False, Optional ByVal facID As Int64 = 0, Optional ByVal tankID As Int64 = 0) As DataSet
            Dim dsData As DataSet
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Try
                strSQL = "spGetAllByOwnerID"
                Dim Params(1) As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = ownerID.ToString
                Params(1).Value = showDeleted
                Params(2).Value = IIf(facID = 0, DBNull.Value, facID)
                Params(3).Value = IIf(tankID = 0, DBNull.Value, tankID)
                Params(4).Value = IIf([Module] = "", DBNull.Value, [Module])
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Put(ByRef oOwnrInf As MUSTER.Info.OwnerInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String)
            Dim Params() As SqlParameter
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Owner, Integer))) Then
                    returnVal = "You do not have rights to save an Owner."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutRegOwner")
                If oOwnrInf.ID < 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oOwnrInf.ID
                End If
                If oOwnrInf.OrganizationID = 0 Then
                    Params(1).Value = System.DBNull.Value
                Else
                    Params(1).Value = oOwnrInf.OrganizationID
                End If
                If oOwnrInf.PersonID = 0 And Not oOwnrInf.OrganizationID = 0 Then
                    Params(2).Value = System.DBNull.Value
                Else
                    Params(2).Value = oOwnrInf.PersonID
                End If
                'If oOwnrInf.PhoneNumberOne = String.Empty Then
                '    Params(3).Value = System.DBNull.Value
                'Else
                If oOwnrInf.PhoneNumberOne = "(___)___-____" Then
                    Params(3).Value = String.Empty
                Else
                    Params(3).Value = oOwnrInf.PhoneNumberOne
                End If

                'End If
                'If oOwnrInf.PhoneNumberTwo = String.Empty Then
                '    Params(4).Value = System.DBNull.Value
                'Else
                If oOwnrInf.PhoneNumberOne = "(___)___-____" Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = oOwnrInf.PhoneNumberTwo
                End If

                'End If
                'If oOwnrInf.Fax = String.Empty Then
                'Params(5).Value = System.DBNull.Value
                'Else
                Params(5).Value = oOwnrInf.Fax
                'End If
                'If oOwnrInf.EmailAddress = String.Empty Then
                '    Params(6).Value = System.DBNull.Value
                'Else
                Params(6).Value = oOwnrInf.EmailAddress
                'End If
                'If oOwnrInf.EmailAddressPersonal = String.Empty Then
                '    Params(7).Value = System.DBNull.Value
                'Else
                Params(7).Value = oOwnrInf.EmailAddressPersonal
                'End If
                'If oOwnrInf.AddressId = 0 Then
                'Params(8).Value = System.DBNull.Value
                'Else
                Params(8).Value = oOwnrInf.AddressId
                'End If
                If Date.Compare(oOwnrInf.DateCapSignUp, CDate("01/01/0001")) = 0 Then
                    Params(9).Value = System.DBNull.Value
                Else
                    Params(9).Value = oOwnrInf.DateCapSignUp.Date
                End If
                Params(10).Value = oOwnrInf.CapCurrentStatus
                'If oOwnrInf.OwnerType = 0 Then
                '    Params(11).Value = System.DBNull.Value
                'Else
                Params(11).Value = oOwnrInf.OwnerType
                'End If
                If oOwnrInf.BP2KType = 0 Then
                    Params(12).Value = System.DBNull.Value
                Else
                    Params(12).Value = oOwnrInf.BP2KType
                End If
                If oOwnrInf.FeesProfileID = 0 Then
                    Params(13).Value = System.DBNull.Value
                Else
                    Params(13).Value = oOwnrInf.FeesProfileID
                End If
                Params(14).Value = oOwnrInf.FeesStatus
                If oOwnrInf.ComplianceProfileID = 0 Then
                    Params(15).Value = System.DBNull.Value
                Else
                    Params(15).Value = oOwnrInf.ComplianceProfileID
                End If
                Params(16).Value = oOwnrInf.ComplianceStatus
                Params(17).Value = oOwnrInf.Active
                Params(18).Value = oOwnrInf.FeeActive
                If oOwnrInf.EnsiteOrganizationID = 0 Then
                    Params(19).Value = System.DBNull.Value
                Else
                    Params(19).Value = oOwnrInf.EnsiteOrganizationID
                End If
                If oOwnrInf.EnsitePersonID = 0 Then
                    Params(20).Value = System.DBNull.Value
                Else
                    Params(20).Value = oOwnrInf.EnsitePersonID
                End If
                Params(21).Value = oOwnrInf.EnsiteAgencyInterestID
                Params(22).Value = 0
                'If oOwnrInf.Description = 0 Then
                '    Params(22).Value = System.DBNull.Value
                'Else
                '   Params(22).Value = oOwnrInf.Description
                'End If
                'If oOwnrInf.CustEntityCode = 0 Then
                '    Params(23).Value = System.DBNull.Value
                'Else
                Params(23).Value = oOwnrInf.CustEntityCode
                'End If
                'If oOwnrInf.CustTypeCode = 0 Then
                '    Params(24).Value = System.DBNull.Value
                'Else
                Params(24).Value = oOwnrInf.CustTypeCode
                'End If
                Params(25).Value = oOwnrInf.Deleted
                Params(26).Value = oOwnrInf.OwnerL2CSnippet
                Params(27).Value = System.DBNull.Value
                Params(28).Value = System.DBNull.Value
                Params(29).Value = System.DBNull.Value
                Params(30).Value = System.DBNull.Value
                Params(31).Value = strUser
                'Params(32).Value = oOwnrInf.AddressLine2
                'Params(33).Value = oOwnrInf.City
                'Params(34).Value = oOwnrInf.State
                'Params(35).Value = oOwnrInf.Zip
                'Params(36).Value = oOwnrInf.FIPSCode

                'If oOwnrInf.ID <= 0 Then
                '    Params(31).Value = oOwnrInf.CreatedBy
                'Else
                '    Params(31).Value = oOwnrInf.ModifiedBy
                'End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutRegOwner", Params)
                If Params(0).Value <> oOwnrInf.ID Then
                    oOwnrInf.ID = Params(0).Value
                    oOwnrInf.CapParticipationLevel = "NONE (0/0)"
                End If
                oOwnrInf.ModifiedBy = AltIsDBNull(Params(27).Value, String.Empty)
                oOwnrInf.ModifiedOn = AltIsDBNull(Params(28).Value, CDate("01/01/0001"))
                oOwnrInf.CreatedBy = AltIsDBNull(Params(29).Value, String.Empty)
                oOwnrInf.CreatedOn = AltIsDBNull(Params(30).Value, CDate("01/01/0001"))
                oOwnrInf.Active = Params(17).Value
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetFacilitiesTankStatus(ByVal OwnrID As Integer) As DataSet
            Dim Params(1) As SqlParameter
            Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spSELFacility_Tank_Status")
            Dim dsData As DataSet
            Params(0).Value = OwnrID
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spSELFacility_Tank_Status", Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetFacilitiesLUSTSummary(ByVal OwnrID As Integer) As DataSet
            Dim Params(1) As SqlParameter
            Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spSELFacility_LUST_Summary")
            Dim dsData As DataSet
            Params(0).Value = OwnrID
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spSELFacility_LUST_Summary", Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetFacilitiesFinancialSummaryTable(ByVal OwnrID As Integer) As DataSet
            Dim Params(1) As SqlParameter
            Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spSELFacility_Financial_Summary")
            Dim dsData As DataSet
            Params(0).Value = OwnrID
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spSELFacility_Financial_Summary", Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Public Function DBGetPreviousFacilities(ByVal OwnrID As Integer) As DataSet
        '    Dim dsData As DataSet
        '    Try
        '        Dim strSQL As String = "select distinct a.facility_id,a.name,B.[ADDRESS_LINE_ONE],B.[CITY],A.[Date_Transferred] from tblreg_facility A INNER JOIN tblREG_ADDRESS_MASTER B ON A.ADDRESS_ID = B.ADDRESS_ID where a.facility_id in(select facility_id from tblreg_facility_history where owner_id = " + OwnrID.ToString + ") and a.facility_id not in (select facility_id from tblreg_facility where owner_id = " + OwnrID.ToString + ")"
        '        dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
        '        Return dsData
        '    Catch ex As Exception
        '        MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Public Function GetCAPParticipationLevelOld(ByVal OwnrID As Integer) As String
            Dim Params As New SqlParameter
            Dim dsData As DataSet
            Dim strRet As String

            Params.SqlDbType = SqlDbType.BigInt
            Params.ParameterName = "@Owner_ID"
            Params.Value = OwnrID

            Try
                strRet = SqlHelper.ExecuteScalar(_strConn, CommandType.StoredProcedure, "spGetOwnerCAPParticipation", Params)
                Return strRet
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub PutOwnerActive(ByVal bolStatus As Boolean, ByVal ownerID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal userID As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Owner, Integer))) Then
                    returnVal = "You do not have rights to update an owner."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutREGOWNERACTIVE"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = bolStatus
                Params(1).Value = ownerID
                Params(2).Value = userID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub CheckRegistrationActivityRights(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Owner, Integer))) Then
                    returnVal = "You do not have rights to Process Registration Activity."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub DBRollOverTankCAPDates(ByVal facID As Integer, ByVal tnkID As Integer, ByVal dtCPDate As Date, ByVal dtLIInspected As Date, ByVal dtTTDate As Date, ByVal dtSpillTested As Date, ByVal dtOverfillInspected As Date, ByVal dtTankSecondary As Date, ByVal dtTankElectronic As Date, ByVal dtATG As Date, ByVal userID As String)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spRollOverCAPTankDates"
            Try
                ' No need to check for security here as security is already checked when user clicked menu item in muster container
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                Params(1).Value = userID
                Params(2).Value = tnkID
                If Date.Compare(dtCPDate, CDate("01/01/0001")) = 0 Then
                    Params(3).Value = System.DBNull.Value
                Else
                    Params(3).Value = dtCPDate
                End If
                If Date.Compare(dtLIInspected, CDate("01/01/0001")) = 0 Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = dtLIInspected
                End If
                If Date.Compare(dtTTDate, CDate("01/01/0001")) = 0 Then
                    Params(5).Value = System.DBNull.Value
                Else
                    Params(5).Value = dtTTDate
                End If
                If Date.Compare(dtSpillTested, CDate("01/01/0001")) = 0 Then
                    Params(6).Value = System.DBNull.Value
                Else
                    Params(6).Value = dtSpillTested
                End If
                If Date.Compare(dtOverfillInspected, CDate("01/01/0001")) = 0 Then
                    Params(7).Value = System.DBNull.Value
                Else
                    Params(7).Value = dtOverfillInspected
                End If
                If Date.Compare(dtTankSecondary, CDate("01/01/0001")) = 0 Then
                    Params(8).Value = System.DBNull.Value
                Else
                    Params(8).Value = dtTankSecondary
                End If
                If Date.Compare(dtTankElectronic, CDate("01/01/0001")) = 0 Then
                    Params(9).Value = System.DBNull.Value
                Else
                    Params(9).Value = dtTankElectronic
                End If
                If Date.Compare(dtATG, CDate("01/01/0001")) = 0 Then
                    Params(10).Value = System.DBNull.Value
                Else
                    Params(10).Value = dtATG
                End If
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub DBRollOverPipeCAPDates(ByVal facID As Integer, ByVal pipeID As Integer, ByVal dtCPDate As Date, ByVal dtTermCPTestDate As Date, ByVal dtALLDTestDate As Date, ByVal dtTTDate As Date, ByVal dtShear As Date, ByVal dtPipeSecondary As Date, ByVal dtPipeElectronic As Date, ByVal userID As String)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spRollOverCAPPipeDates"
            Try
                ' No need to check for security here as security is already checked when user clicked menu item in muster container
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                Params(1).Value = userID
                Params(2).Value = pipeID
                If Date.Compare(dtCPDate, CDate("01/01/0001")) = 0 Then
                    Params(3).Value = System.DBNull.Value
                Else
                    Params(3).Value = dtCPDate
                End If
                If Date.Compare(dtTermCPTestDate, CDate("01/01/0001")) = 0 Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = dtTermCPTestDate
                End If
                If Date.Compare(dtALLDTestDate, CDate("01/01/0001")) = 0 Then
                    Params(5).Value = System.DBNull.Value
                Else
                    Params(5).Value = dtALLDTestDate
                End If
                If Date.Compare(dtTTDate, CDate("01/01/0001")) = 0 Then
                    Params(6).Value = System.DBNull.Value
                Else
                    Params(6).Value = dtTTDate
                End If
                If Date.Compare(dtShear, CDate("01/01/0001")) = 0 Then
                    Params(7).Value = System.DBNull.Value
                Else
                    Params(7).Value = dtShear
                End If
                If Date.Compare(dtPipeSecondary, CDate("01/01/0001")) = 0 Then
                    Params(8).Value = System.DBNull.Value
                Else
                    Params(8).Value = dtPipeSecondary
                End If
                If Date.Compare(dtPipeElectronic, CDate("01/01/0001")) = 0 Then
                    Params(9).Value = System.DBNull.Value
                Else
                    Params(9).Value = dtPipeElectronic
                End If
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub DBTransferOwnerBillingByFacilities(ByVal facilities As String, ByVal oldOwner As Integer, ByVal newOwner As Integer)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spFees_TransferOwnershipBilling"
            Try
                ' No need to check for security here as security is already checked when user clicked menu item in muster container
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facilities
                Params(1).Value = oldOwner
                Params(2).Value = newOwner

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function DBGetOwnerSummary(ByVal nOWnerID As Integer) As DataSet
            Dim dsSet As DataSet
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetOwnerSummary"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@OWNER_ID").Value = nOWnerID
                dsSet = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsSet
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
            End Try
        End Function
        Public Function DBCheckWriteAccess(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal entityType As Integer) As Boolean
            Try
                Return SqlHelper.HasWriteAccess(moduleID, staffID, entityType)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub DBClearCAPAnnualSummary(ByVal processingYear As Integer, ByVal mode As Integer, ByVal ownerID As Integer, Optional ByVal fac As String = "")
            Dim Params() As SqlParameter
            Dim strSQL As String = "spClearCAPAnnualSummary"
            Try
                ' No need to check for security here as security is already checked when user clicked menu item in muster container
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = processingYear
                Params(1).Value = mode
                Params(2).Value = ownerID
                Params(3).Value = fac

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub DBClearCAPAnnualCalendar(ByVal processingYear As Integer, ByVal ownerID As Integer, Optional ByVal mode As Integer = 0, Optional ByVal month As Integer = -1, Optional ByVal fac As String = "")
            Dim Params() As SqlParameter
            Dim strSQL As String = "spClearCAPAnnualCalendar"
            Try
                ' No need to check for security here as security is already checked when user clicked menu item in muster container
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = processingYear
                Params(1).Value = ownerID
                Params(2).Value = mode
                Params(3).Value = month
                Params(4).Value = fac

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub DBSaveCAPAnnualSummary(ByVal processingYear As Integer, ByVal linePosition As Integer, ByVal ownerID As Integer, _
                                            ByVal ownerName As String, ByVal facilityID As Integer, ByVal facName As String, _
                                            ByVal facAddr1 As String, ByVal facAddrCity As String, ByVal facAddrState As String, _
                                            ByVal facAddrZip As String, ByVal desc As String, ByVal isDescPeriodicTestReq As Boolean, _
                                            ByVal isDescHeading As Boolean, ByVal isDescSubHeading As Boolean, ByVal createdBy As String, ByVal mode As Integer)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutCAPAnnualSummary"
            Try
                ' No need to check for security here as security is already checked when user clicked menu item in muster container
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = processingYear
                Params(1).Value = linePosition
                Params(2).Value = ownerID
                Params(3).Value = ownerName
                Params(4).Value = facilityID
                Params(5).Value = facName
                Params(6).Value = facAddr1
                Params(7).Value = facAddrCity
                Params(8).Value = facAddrState
                Params(9).Value = facAddrZip
                Params(10).Value = desc
                Params(11).Value = isDescPeriodicTestReq
                Params(12).Value = isDescHeading
                Params(13).Value = isDescSubHeading
                Params(14).Value = createdBy
                Params(15).Value = mode
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub


        Public Sub DBSaveCAPAnnualCalendar(ByVal processingYear As Integer, ByVal ownerID As Integer, ByVal ownerName As String, _
                                                    ByVal month As Integer, ByVal facilityID As Integer, ByVal facName As String, _
                                                    ByVal city As String, ByVal requirements As String, ByVal createdBy As String, _
                                                    Optional ByVal mode As Integer = 0, Optional ByVal procMonth As Integer = -1)
            Dim Params() As SqlParameter
            Dim strSQL As String = "spPutCAPAnnualCalendar"
            Try
                ' No need to check for security here as security is already checked when user clicked menu item in muster container
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = processingYear
                Params(1).Value = ownerID
                Params(2).Value = ownerName
                Params(3).Value = month
                Params(4).Value = facilityID
                Params(5).Value = facName
                Params(6).Value = city
                Params(7).Value = requirements
                Params(8).Value = createdBy
                Params(9).Value = mode
                Params(10).Value = procMonth

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub


    End Class
End Namespace
