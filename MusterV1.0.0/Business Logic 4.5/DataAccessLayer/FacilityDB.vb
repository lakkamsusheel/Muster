'-------------------------------------------------------------------------------
' MUSTER.DataAccess.FacilityDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       EN      12/06/04    Original class definition.
'  1.1       EN      12/29/04    Modified all methods to pass 3 new parameters(datum,Method,LocationType)
'  1.2       AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3       MNR     01/12/04    Added DBGetByOwnerID(..) function
'  1.4       JVC2    01/19/05    Added DBGetPreviousOwners function
'  1.5       MNR     01/28/05    Added DBGetFacStatus function
'                                Modified Put Sub to get new Facility Status using DBGetFacStatus
'  1.6       AB      02/08/05    Replaced dynamic SQL with stored procedures in the following
'                                Functions:  DBGetAllInfo, DBGetByID, DBGetByName, DBGetByOwnerID
'  1.7       AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.8       AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.9       AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.10      AB      02/23/05    Modified Get and Put functions based upon changes made to 
'                                   make several nullable fields non-nullable
'  1.11      MR      03/07/05    Changed Modified By and Modified On to reflect state after PUT
'                                   Changed StoredProc 'spPutFacility_test' to 'spPutFacility' in PUT function.
'                                   Modified PUT function to pass EmptyString or NULL Dates Values for DBNULL
'  1.12      MR      03/10/05    Changed AltISDBNULL validation in all the Get Functions.
'  1.13      MR      03/14/05    Changed Created By and Created On to reflect state after PUT
' 1.14  Thomas Franey    02/16/09     Added MGPTF Status to show latest tech Status of any closures of tanks
' 1.20  Thomas Franey    06/16/09     Added Deignated Operator to database GEt and Put view/procedures 
'
' Function                  Description
' DBGetAllInfo()    Returns an Facility Collection containing all Facility objects in the repository
' DBGetByID(ID)     Returns an FacilityInfo object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(Facility)     Saves the Facility passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class FacilityDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions


#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function GetDataRow(ByVal DBViewName As String) As DataRow
            Dim dsReturn As New DataSet

            Dim dtReturn As DataRow
            Dim strSQL As String
            Try
                strSQL = "Exec " & DBViewName

                dsReturn = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)

                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0).Rows(0)
                Else
                    dtReturn = Nothing
                End If

                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.Info")) Then MusterException.Publish(Ex, Nothing, Nothing)
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
        Public Function DBGetPreviousOwners(ByVal FacID As Int64) As DataSet
            Dim dsData As DataSet
            Dim Params() As SqlParameter
            Try
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spGetPreviousOwners")
                Params(0).Value = FacID
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, "spGetPreviousOwners", Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetAllInfo(Optional ByVal OwnerID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FacilityCollection
            Dim colFacility As New MUSTER.Info.FacilityCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader

            Try
                'strSQL = "SELECT * FROM tblREG_FACILITY WHERE 1=1"
                'strSQL += IIf(Not showDeleted, " AND DELETED <> 1 ", "")
                'strSQL += IIf(OwnerID <> 0, " AND OWNER_ID =" & OwnerID, "")
                'strSQL += " ORDER BY FACILITY_ID"

                strSQL = "spGetFacility"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Facility_ID").Value = DBNull.Value
                Params("@Owner_ID").Value = OwnerID
                Params("@Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    Dim oFacinfo As New MUSTER.Info.FacilityInfo(drSet.Item("FACILITY_ID"), _
                     (AltIsDBNull(drSet.Item("FACILITY_AIID"), 0)), _
                     drSet.Item("NAME"), _
                     drSet.Item("OWNER_ID"), _
                    drSet.Item("ADDRESS_ID"), _
                     (AltIsDBNull(drSet.Item("BILLING_ADDRESS_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_SECONDS"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_SECONDS"), -1)), _
                     drSet.Item("PHONE"), _
                     (AltIsDBNull(drSet.Item("DATUM"), 0)), _
                     (AltIsDBNull(drSet.Item("METHOD"), 0)), _
                     drSet.Item("FAX"), _
                     (AltIsDBNull(drSet.Item("FEES_PROFILE_ID"), 0)), _
                     drSet.Item("FACILITY_TYPE"), _
                     (AltIsDBNull(drSet.Item("FEES_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_CIU_NUMBER"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_CANDIDATE"), False)), _
                     (AltIsDBNull(drSet.Item("CITATION_PROFILE_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_LUST_STATUS"), 0)), _
                     drSet.Item("FUEL_BRAND"), _
                     drSet.Item("FACILITY_DESCRIPTION"), _
                     (AltIsDBNull(drSet.Item("SIGNATURE_NEEDED"), False)), _
                     (AltIsDBNull(drSet.Item("DATE_RECD"), CDate("01/01/0001"))), _
                     (AltIsDBNull(drSet.Item("DATE_TRANSFERRED"), CDate("01/01/0001"))), _
                     (AltIsDBNull(drSet.Item("FACILITY_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("DELETED"), False)), _
                    (AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty)), _
                    (AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001"))), _
                    (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                    (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001"))), _
                    (AltIsDBNull(drSet.Item("DATE_POWEROFF"), CDate("01/01/0001"))), _
                    (AltIsDBNull(drSet.Item("LOCATION_TYPE"), 0)), _
                    (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), False)), _
                    (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION_DATE"), Nothing)), _
                    AltIsDBNull(drSet.Item("LICENSEEID"), 0), _
                    AltIsDBNull(drSet.Item("CONTRACTORID"), 0), _
                    AltIsDBNull(drSet.Item("NAME_FOR_ENSITE"), drSet.Item("NAME")), _
                    AltIsDBNull(drSet.Item("MGPTFSTATUS"), String.Empty), , , AltIsDBNull(drSet("DesignatedOperator"), String.Empty), _
                    AltIsDBNull(drSet.Item("DesignatedManager"), String.Empty))
                 '   AltIsDBNull(drSet.Item("MGPTFSTATUS"), String.Empty), , , AltIsDBNull(drSet("DesignatedManager"), String.Empty))

                    'AltIsDBNull(drSet.Item("ADDRESS_LINE_ONE"), String.Empty), _
                    'AltIsDBNull(drSet.Item("ADDRESS_TWO"), String.Empty), _
                    'AltIsDBNull(drSet.Item("CITY"), String.Empty), _
                    'AltIsDBNull(drSet.Item("STATE"), String.Empty), _
                    'AltIsDBNull(drSet.Item("ZIP"), String.Empty), _
                    'AltIsDBNull(drSet.Item("FIPS_CODE"), String.Empty), _

                    colFacility.Add(oFacinfo)
                End While

                Return colFacility
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal nFacilityID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FacilityInfo

            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nFacilityID = 0 Then
                    Return New MUSTER.Info.FacilityInfo
                End If
                strVal = nFacilityID
                'strSQL = "select * from tblREG_FACILITY where FACILITY_ID= '" + strVal + "'"
                'If Not showDeleted Then
                '    strSQL += " AND DELETED <> 1"
                'End If
                strSQL = "spGetFacility"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Facility_ID").Value = strVal
                Params("@Owner_ID").Value = DBNull.Value
                Params("@Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()


                    Return New MUSTER.Info.FacilityInfo(drSet.Item("FACILITY_ID"), _
                     (AltIsDBNull(drSet.Item("FACILITY_AIID"), 0)), _
                     (AltIsDBNull(drSet.Item("NAME"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("OWNER_ID"), 0)), _
                     drSet.Item("ADDRESS_ID"), _
                     (AltIsDBNull(drSet.Item("BILLING_ADDRESS_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_SECONDS"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_SECONDS"), -1)), _
                     (AltIsDBNull(drSet.Item("PHONE"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("DATUM"), Nothing)), _
                     (AltIsDBNull(drSet.Item("METHOD"), Nothing)), _
                     (AltIsDBNull(drSet.Item("FAX"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("FEES_PROFILE_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("FACILITY_TYPE"), 0)), _
                     (AltIsDBNull(drSet.Item("FEES_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_CIU_NUMBER"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_CANDIDATE"), False)), _
                     (AltIsDBNull(drSet.Item("CITATION_PROFILE_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_LUST_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("FUEL_BRAND"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("FACILITY_DESCRIPTION"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("SIGNATURE_NEEDED"), False)), _
                     (AltIsDBNull(drSet.Item("DATE_RECD"), Nothing)), _
                     (AltIsDBNull(drSet.Item("DATE_TRANSFERRED"), Nothing)), _
                     (AltIsDBNull(drSet.Item("FACILITY_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("DELETED"), False)), _
                     (AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001"))), _
                     (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001"))), _
                     (AltIsDBNull(drSet.Item("DATE_POWEROFF"), Nothing)), _
                     (AltIsDBNull(drSet.Item("LOCATION_TYPE"), Nothing)), _
                     (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), False)), _
                    (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION_DATE"), Nothing)), _
                    AltIsDBNull(drSet.Item("LICENSEEID"), 0), _
                    AltIsDBNull(drSet.Item("CONTRACTORID"), 0), _
                    AltIsDBNull(drSet.Item("NAME_FOR_ENSITE"), drSet.Item("NAME")), _
                    AltIsDBNull(drSet.Item("MGPTFSTATUS"), String.Empty), , , AltIsDBNull(drSet.Item("DesignatedOperator"), String.Empty), _
                    AltIsDBNull(drSet.Item("DesignatedManager"), String.Empty))
                '    AltIsDBNull(drSet.Item("MGPTFSTATUS"), String.Empty), , , AltIsDBNull(drSet.Item("DesignatedManager"), String.Empty))
                    'IIf(AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), "YES").ToString.ToUpper = "YES", True, False), _
                Else
                    Return New MUSTER.Info.FacilityInfo
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
        Public Function DBGetByName(ByVal nVal As String) As MUSTER.Info.FacilityInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = String.Empty Then
                    Return New MUSTER.Info.FacilityInfo
                End If
                strSQL = "spGetFacility"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Facility_ID").Value = DBNull.Value
                Params("@Owner_ID").Value = DBNull.Value
                Params("@Name").Value = nVal
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()

                    Return New MUSTER.Info.FacilityInfo(drSet.Item("FACILITY_ID"), _
                     (AltIsDBNull(drSet.Item("FACILITY_AIID"), 0)), _
                     drSet.Item("NAME"), _
                     drSet.Item("OWNER_ID"), _
                     drSet.Item("ADDRESS_ID"), _
                     (AltIsDBNull(drSet.Item("BILLING_ADDRESS_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_SECONDS"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_SECONDS"), -1)), _
                     drSet.Item("PHONE"), _
                     (AltIsDBNull(drSet.Item("DATUM"), Nothing)), _
                     (AltIsDBNull(drSet.Item("METHOD"), Nothing)), _
                     drSet.Item("FAX"), _
                     (AltIsDBNull(drSet.Item("FEES_PROFILE_ID"), 0)), _
                     drSet.Item("FACILITY_TYPE"), _
                     (AltIsDBNull(drSet.Item("FEES_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_CIU_NUMBER"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_CANDIDATE"), False)), _
                     (AltIsDBNull(drSet.Item("CITATION_PROFILE_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_LUST_STATUS"), 0)), _
                     drSet.Item("FUEL_BRAND"), _
                     drSet.Item("FACILITY_DESCRIPTION"), _
                     (AltIsDBNull(drSet.Item("SIGNATURE_NEEDED"), False)), _
                     (AltIsDBNull(drSet.Item("DATE_RECD"), Nothing)), _
                     (AltIsDBNull(drSet.Item("DATE_TRANSFERRED"), Nothing)), _
                     (AltIsDBNull(drSet.Item("FACILITY_STATUS"), 0)), _
                     IIf(drSet.Item("DELETED").ToString.ToUpper = "YES", True, False), _
                    (AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty)), _
                    (AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001"))), _
                    (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                    (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001"))), _
                    (AltIsDBNull(drSet.Item("DATE_POWEROFF"), Nothing)), _
                    (AltIsDBNull(drSet.Item("LOCATION_TYPE"), Nothing)), _
                    (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), False)), _
                    (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION_DATE"), Nothing)), _
                    AltIsDBNull(drSet.Item("LICENSEEID"), 0), _
                    AltIsDBNull(drSet.Item("CONTRACTORID"), 0), _
                    AltIsDBNull(drSet.Item("NAME_FOR_ENSITE"), drSet.Item("NAME")), _
                    AltIsDBNull(drSet.Item("MGPTFSTATUS"), String.Empty), , , AltIsDBNull(drSet.Item("DesignatedOperator"), String.Empty), _
                    AltIsDBNull(drSet.Item("DesignatedManager"), String.Empty))

                Else
                    Return New MUSTER.Info.FacilityInfo
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

        Public Function Put(ByRef oFacility As MUSTER.Info.FacilityInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal AdviceIDForTransfers As Integer = 0) As Integer


            Dim Params() As SqlParameter
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Facility, Integer))) Then
                    returnVal = "You do not have rights to save a Facility."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutFacility")

                If oFacility.ID < 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oFacility.ID
                End If
                If oFacility.AIID = 0 Then
                    Params(1).Value = System.DBNull.Value
                Else
                    Params(1).Value = oFacility.AIID
                End If
                'If oFacility.Name = String.Empty Then
                '    Params(2).Value = System.DBNull.Value
                'Else
                '    Params(2).Value = oFacility.Name
                'End If
                Params(2).Value = oFacility.Name

                Params(3).Value = IsNull(oFacility.OwnerID, 0)
                Params(4).Value = IsNull(oFacility.AddressID, 0)
                'P1 12/31/04 start
                Params(5).Value = IIf(oFacility.LatitudeDegree < 0, System.DBNull.Value, oFacility.LatitudeDegree)
                Params(6).Value = IIf(oFacility.LatitudeMinutes < 0, System.DBNull.Value, oFacility.LatitudeMinutes)
                Params(7).Value = IIf(oFacility.LatitudeSeconds < 0, System.DBNull.Value, oFacility.LatitudeSeconds)
                Params(8).Value = IIf(oFacility.LongitudeDegree < 0, System.DBNull.Value, oFacility.LongitudeDegree)
                Params(9).Value = IIf(oFacility.LongitudeMinutes < 0, System.DBNull.Value, oFacility.LongitudeMinutes)
                Params(10).Value = IIf(oFacility.LongitudeSeconds < 0, System.DBNull.Value, oFacility.LongitudeSeconds)
                'P1 12/31/04 end
                'P1 04/01/05 start
                'Params(11).Value = IsNull(oFacility.Phone, System.DBNull.Value)
                'Params(12).Value = IsNull(oFacility.Fax, System.DBNull.Value)
                'Params(13).Value = IIFIsIntegerNull(oFacility.FacilityType, System.DBNull.Value)
                'Params(14).Value = IsNull(oFacility.FuelBrand, System.DBNull.Value)
                If oFacility.Phone = "(___)___-____" Then
                    Params(11).Value = String.Empty
                Else
                    Params(11).Value = IsNull(oFacility.Phone, String.Empty)
                End If
                'Params(11).Value = IsNull(oFacility.Phone, String.Empty)
                Params(12).Value = IsNull(oFacility.Fax, String.Empty)
                Params(13).Value = IIFIsIntegerNull(oFacility.FacilityType, 0)
                Params(14).Value = IsNull(oFacility.FuelBrand, String.Empty)
                If Date.Compare(oFacility.DateReceived, CDate("01/01/0001")) = 0 Then
                    Params(15).Value = System.DBNull.Value
                Else
                    Params(15).Value = oFacility.DateReceived.Date
                End If
                If Date.Compare(oFacility.DateTransferred, CDate("01/01/0001")) = 0 Then
                    Params(16).Value = System.DBNull.Value
                Else
                    Params(16).Value = oFacility.DateTransferred.Date
                End If
                Params(17).Value = IIFIsIntegerNull(oFacility.FacilityStatus, System.DBNull.Value)
                Params(18).Value = oFacility.Deleted
                If Date.Compare(oFacility.DatePowerOff, CDate("01/01/0001")) = 0 Then
                    Params(19).Value = System.DBNull.Value
                Else
                    Params(19).Value = oFacility.DatePowerOff.Date
                End If
                Params(20).Value = IsNull(oFacility.SignatureOnNF, System.DBNull.Value)
                Params(21).Value = 0
                Params(21).Direction = ParameterDirection.InputOutput
                Params(22).Value = IIFIsIntegerNull(oFacility.Datum, System.DBNull.Value)
                Params(23).Value = IIFIsIntegerNull(oFacility.Method, System.DBNull.Value)
                Params(24).Value = IIFIsIntegerNull(oFacility.LocationType, System.DBNull.Value)
                Params(25).Value = IsNull(oFacility.CAPCandidate, System.DBNull.Value)
                Params(26).Value = IsNull(oFacility.UpcomingInstallation, System.DBNull.Value)
                If Date.Compare(oFacility.UpcomingInstallationDate, CDate("01/01/0001")) = 0 Then
                    Params(27).Value = System.DBNull.Value
                Else
                    Params(27).Value = oFacility.UpcomingInstallationDate.Date
                End If
                Params(28).Value = IIFIsIntegerNull(oFacility.CapStatus, System.DBNull.Value)
                'new parameters added to stored procedure - Elango 

                Params(29).Value = IIFIsIntegerNull(oFacility.BillingAddressID, System.DBNull.Value)
                Params(30).Value = IIFIsIntegerNull(oFacility.FeesProfileId, System.DBNull.Value)
                Params(31).Value = IIFIsIntegerNull(oFacility.FeesStatus, System.DBNull.Value)
                Params(32).Value = IIFIsIntegerNull(oFacility.CurrentCIUNumber, System.DBNull.Value)
                Params(33).Value = IIFIsIntegerNull(oFacility.CitationProfileID, System.DBNull.Value)
                Params(34).Value = IIFIsIntegerNull(oFacility.CurrentLUSTStatus, System.DBNull.Value)
                'P1 04/01/05 start

                'If oFacility.FacilityDescription = String.Empty Then
                '    Params(35).Value = System.DBNull.Value
                'Else
                '    Params(35).Value = oFacility.FacilityDescription
                'End If
                Params(35).Value = oFacility.FacilityDescription
                Params(36).Value = DBNull.Value
                Params(37).Value = DBNull.Value
                Params(38).Value = DBNull.Value
                Params(39).Value = DBNull.Value
                Params(40).Value = strUser
                'Params(41).Value = oFacility.AddressLine2
                'Params(42).Value = oFacility.City
                'Params(43).Value = oFacility.State
                'Params(44).Value = oFacility.Zip
                'Params(45).Value = oFacility.FIPSCode

                'If oFacility.ID <= 0 Then
                '    Params(40).Value = oFacility.CreatedBy
                'Else
                '    Params(40).Value = oFacility.ModifiedBy
                'End If
                Params(41).Value = oFacility.LicenseeID
                Params(42).Value = oFacility.ContractorID
                Params(43).Value = IIf(oFacility.NameForEnsite = String.Empty, oFacility.Name, oFacility.NameForEnsite)
                Params(44).Value = oFacility.DesignatedOperator
                Params(45).Value = oFacility.DesignatedManager
                If AdviceIDForTransfers = 0 Then
                    Params(46).Value = DBNull.Value
                    Params(46).Direction = ParameterDirection.InputOutput
                Else
                    Params(46).Value = AdviceIDForTransfers
                    Params(46).Direction = ParameterDirection.InputOutput


                End If

                Dim nResult As Integer = SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutFacility", Params)

                If CInt(Params(21).Value) > 0 Then
                    oFacility.ID = Params(21).Value
                End If
                oFacility.ModifiedBy = AltIsDBNull(Params(36).Value, String.Empty)
                oFacility.ModifiedOn = AltIsDBNull(Params(37).Value, CDate("01/01/0001"))
                oFacility.CreatedBy = AltIsDBNull(Params(38).Value, String.Empty)
                oFacility.CreatedOn = AltIsDBNull(Params(39).Value, CDate("01/01/0001"))
                oFacility.CapStatus = AltIsDBNull(Params(28).Value, 0)
                oFacility.NameForEnsite = AltIsDBNull(Params(43).Value, oFacility.Name)
                AdviceIDForTransfers = AltIsDBNull(Params(46).Value, 0)

                Return AdviceIDForTransfers

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DBGetByOwnerID(ByVal nOwnerID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FacilityCollection
            Dim drSet As SqlDataReader
            Dim colFacility As New MUSTER.Info.FacilityCollection
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nOwnerID = 0 Then
                    Return colFacility
                End If

                strVal = nOwnerID
                strSQL = "spGetFacility"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Owner_ID").Value = strVal
                Params("@Facility_ID").Value = DBNull.Value
                Params("@Name").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    Dim oFacinfo
                    While drSet.Read()
                        Dim oFacilityInfo As New MUSTER.Info.FacilityInfo(drSet.Item("FACILITY_ID"), _
                                             (AltIsDBNull(drSet.Item("FACILITY_AIID"), 0)), _
                     (AltIsDBNull(drSet.Item("NAME"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("OWNER_ID"), 0)), _
                     drSet.Item("ADDRESS_ID"), _
                     (AltIsDBNull(drSet.Item("BILLING_ADDRESS_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LATITUDE_SECONDS"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_DEGREE"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_MINUTES"), -1)), _
                     (AltIsDBNull(drSet.Item("LONGITUDE_SECONDS"), -1)), _
                     (AltIsDBNull(drSet.Item("PHONE"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("DATUM"), Nothing)), _
                     (AltIsDBNull(drSet.Item("METHOD"), Nothing)), _
                     (AltIsDBNull(drSet.Item("FAX"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("FEES_PROFILE_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("FACILITY_TYPE"), 0)), _
                     (AltIsDBNull(drSet.Item("FEES_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_CIU_NUMBER"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("CAP_CANDIDATE"), False)), _
                     (AltIsDBNull(drSet.Item("CITATION_PROFILE_ID"), 0)), _
                     (AltIsDBNull(drSet.Item("CURRENT_LUST_STATUS"), 0)), _
                     (AltIsDBNull(drSet.Item("FUEL_BRAND"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("FACILITY_DESCRIPTION"), String.Empty)), _
                     (AltIsDBNull(drSet.Item("SIGNATURE_NEEDED"), False)), _
                     (AltIsDBNull(drSet.Item("DATE_RECD"), Nothing)), _
                     (AltIsDBNull(drSet.Item("DATE_TRANSFERRED"), Nothing)), _
                     (AltIsDBNull(drSet.Item("FACILITY_STATUS"), 0)), _
                     IIf(drSet.Item("DELETED").ToString.ToUpper = "YES", True, False), _
                    (AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty)), _
                    (AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001"))), _
                    (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                    (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001"))), _
                    (AltIsDBNull(drSet.Item("DATE_POWEROFF"), Nothing)), _
                    (AltIsDBNull(drSet.Item("LOCATION_TYPE"), Nothing)), _
                    (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), False)), _
                    (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION_DATE"), Nothing)), _
                    AltIsDBNull(drSet.Item("LICENSEEID"), 0), _
                    AltIsDBNull(drSet.Item("CONTRACTORID"), 0), _
                    AltIsDBNull(drSet.Item("NAME_FOR_ENSITE"), drSet.Item("NAME")), _
                    AltIsDBNull(drSet.Item("MGPTFSTATUS"), String.Empty), , , AltIsDBNull(drSet.Item("DesignatedOperator"), String.Empty), _
                    AltIsDBNull(drSet.Item("DesignatedManager"), String.Empty))

                        '(AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), False)), _

                        'Dim oFacilityInfo As New MUSTER.Info.FacilityInfo(drSet.Item("FACILITY_ID"), _
                        ' (AltIsDBNull(drSet.Item("FACILITY_AIID"), 0)), _
                        ' drSet.Item("NAME"), _
                        ' drSet.Item("OWNER_ID"), _
                        ' drSet.Item("ADDRESS_ID"), _
                        ' (AltIsDBNull(drSet.Item("BILLING_ADDRESS_ID"), 0)), _
                        ' (AltIsDBNull(drSet.Item("LATITUDE_DEGREE"), -1)), _
                        ' (AltIsDBNull(drSet.Item("LATITUDE_MINUTES"), -1)), _
                        ' (AltIsDBNull(drSet.Item("LATITUDE_SECONDS"), -1)), _
                        ' (AltIsDBNull(drSet.Item("LONGITUDE_DEGREE"), -1)), _
                        ' (AltIsDBNull(drSet.Item("LONGITUDE_MINUTES"), -1)), _
                        ' (AltIsDBNull(drSet.Item("LONGITUDE_SECONDS"), -1)), _
                        ' drSet.Item("PHONE"), _
                        ' (AltIsDBNull(drSet.Item("DATUM"), Nothing)), _
                        ' (AltIsDBNull(drSet.Item("METHOD"), Nothing)), _
                        ' drSet.Item("FAX"), _
                        ' (AltIsDBNull(drSet.Item("FEES_PROFILE_ID"), 0)), _
                        ' drSet.Item("FACILITY_TYPE"), _
                        ' (AltIsDBNull(drSet.Item("FEES_STATUS"), 0)), _
                        ' (AltIsDBNull(drSet.Item("CURRENT_CIU_NUMBER"), 0)), _
                        ' (AltIsDBNull(drSet.Item("CAP_STATUS"), 0)), _
                        ' (AltIsDBNull(drSet.Item("CAP_CANDIDATE"), False)), _
                        ' (AltIsDBNull(drSet.Item("CITATION_PROFILE_ID"), 0)), _
                        ' (AltIsDBNull(drSet.Item("CURRENT_LUST_STATUS"), 0)), _
                        ' drSet.Item("FUEL_BRAND"), _
                        ' drSet.Item("FACILITY_DESCRIPTION"), _
                        ' (AltIsDBNull(drSet.Item("SIGNATURE_NEEDED"), False)), _
                        ' (AltIsDBNull(drSet.Item("DATE_RECD"), Nothing)), _
                        ' (AltIsDBNull(drSet.Item("DATE_TRANSFERRED"), Nothing)), _
                        ' (AltIsDBNull(drSet.Item("FACILITY_STATUS"), 0)), _
                        ' (drSet.Item("DELETED")), _
                        ' drSet.Item("CREATED_BY"), _
                        ' drSet.Item("DATE_CREATED"), _
                        ' (AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty)), _
                        ' (AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), Nothing)), _
                        ' (AltIsDBNull(drSet.Item("DATE_POWEROFF"), Nothing)), _
                        ' (AltIsDBNull(drSet.Item("LOCATION_TYPE"), Nothing)), _
                        ' (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), False)), _
                        ' (AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION_DATE"), Nothing)))

                        'IIf(AltIsDBNull(drSet.Item("CAP_CANDIDATE"), "YES").ToString.ToUpper = "YES", True, False), _
                        'IIf(AltIsDBNull(drSet.Item("SIGNATURE_NEEDED"), "YES").ToString.ToUpper = "YES", True, False), _
                        'IIf(AltIsDBNull(drSet.Item("DELETED"), "YES").ToString.ToUpper = "YES", True, False), _
                        '(AltIsDBNull(drSet.Item("UPCOMING_INSTALLATION"), False)), _

                        colFacility.Add(oFacilityInfo)
                    End While

                    Return colFacility
                Else
                    Return New MUSTER.Info.FacilityCollection
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
        Public Function DBGetFacStatus(ByVal FacID As Integer, Optional ByVal showDeleted As Boolean = False) As Integer
            Dim Params() As SqlParameter
            Dim newStatus As Integer
            Dim strSQL As String
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetFacilityStatus"
                'Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spSELGetFacilityStatus")
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = FacID
                Params(1).Value = showDeleted
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    newStatus = drSet.Item("PROPERTY_ID")
                End If
                If Not drSet.IsClosed Then drSet.Close()
                Return newStatus
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetLustStatus(ByVal FacID As Integer) As Integer
            Dim Params() As SqlParameter
            Dim newStatus As Integer
            Dim strSQL As String
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetFacilityLustStatus"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = FacID
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    newStatus = drSet.Item("SiteCount")
                End If
                If Not drSet.IsClosed Then drSet.Close()
                Return newStatus
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        'This needs to be called irrespective of whether the user has rights or not.
        Public Sub PutFacStatus(ByVal facID As Integer, ByVal status As Integer)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spPutRegFacilityStatus"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                Params(1).Value = status
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetCAPStatus(ByVal facID As Integer) As Integer
            Dim drSet As SqlDataReader
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spUpdateFacilityCAPStatus"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

                strSQL = "select CAP_STATUS from tblreg_facility where deleted = 0 and facility_id = " + facID.ToString
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return drSet.Item("CAP_STATUS")
                Else
                    Return 0
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try

        End Function

        Public Function DBHasCIUTOSITanks(ByVal facID As Integer) As Boolean

            Dim drSet As SqlDataReader
            Dim Params() As SqlParameter
            Dim strSQL As String
            Dim DS As DataSet = Nothing
            Dim answer As Boolean = False

            Try
                strSQL = "spGetRegFacilityWithAvailableTanks"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                DS = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not DS Is Nothing AndAlso DS.Tables.Count > 0 AndAlso DS.Tables(0).Rows.Count > 0 AndAlso DS.Tables(0).Rows(0).Item("HasTanks") = "Yes" Then
                    answer = True
                Else
                    answer = False
                End If


                Return answer

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not DS Is Nothing Then
                    DS.Dispose()
                End If
            End Try

        End Function


        Public Function DBSaveTANKCAPData(ByVal isInspection As Boolean, ByVal facID As Integer, ByVal tankID As Integer, ByVal dateSpillTested As Object, _
                                     ByVal dateOverfillTested As Object, ByVal dateTankElecInsp As Object, ByVal dateLastTCP As Object, _
                                     ByVal dateLineInteriorInspect As Object, ByVal dateATGInsp As Object, ByVal dateTTT As Object, ByVal userID As String, _
        ByVal dateLineInteriorInstalled As Object, ByVal dateSpillInstalled As Object, ByVal dateOverfillInstalled As Object) As Boolean


            Dim Params() As SqlParameter
            Dim strSQL As String

            Try
                strSQL = "spPutCAPTANK"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = isInspection
                Params(1).Value = facID
                Params(2).Value = tankID
                Params(3).Value = IIf(dateSpillTested Is DBNull.Value, Nothing, dateSpillTested)
                Params(4).Value = IIf(dateOverfillTested Is DBNull.Value, Nothing, dateOverfillTested)
                Params(5).Value = IIf(dateTankElecInsp Is DBNull.Value, Nothing, dateTankElecInsp)
                Params(6).Value = IIf(dateLastTCP Is DBNull.Value, Nothing, dateLastTCP)
                Params(7).Value = IIf(dateLineInteriorInspect Is DBNull.Value, Nothing, dateLineInteriorInspect)
                Params(8).Value = IIf(dateATGInsp Is DBNull.Value, Nothing, dateATGInsp)
                Params(9).Value = IIf(dateTTT Is DBNull.Value, Nothing, dateTTT)
                Params(10).Value = userID
                Params(11).Value = IIf(dateLineInteriorInstalled Is DBNull.Value, Nothing, dateLineInteriorInstalled)
                Params(12).Value = IIf(dateSpillInstalled Is DBNull.Value, Nothing, dateSpillInstalled)
                Params(13).Value = IIf(dateOverfillInstalled Is DBNull.Value, Nothing, dateOverfillInstalled)


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try


            Return True


        End Function



        Public Function DBSavePIPECAPData(ByVal isInspection As Boolean, ByVal facID As Integer, ByVal pipeID As Integer, ByVal tankID As Integer, ByVal dateALLD_Test As Object, _
                                        ByVal dateLTT As Object, ByVal datePipeCPTest As Object, ByVal dateTermCPTest As Object, _
                                        ByVal datePipeShearTest As Object, ByVal datePipeSecInsp As Object, ByVal datePipeElecInsp As Object, ByVal userID As String) As Boolean

            Dim Params() As SqlParameter
            Dim strSQL As String

            Try
                strSQL = "spPutCAPPIPE"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = isInspection
                Params(1).Value = facID
                Params(2).Value = pipeID
                Params(3).Value = tankID
                Params(4).Value = IIf(dateALLD_Test Is DBNull.Value, Nothing, dateALLD_Test)
                Params(5).Value = IIf(dateLTT Is DBNull.Value, Nothing, dateLTT)
                Params(6).Value = IIf(datePipeCPTest Is DBNull.Value, Nothing, datePipeCPTest)
                Params(7).Value = IIf(dateTermCPTest Is DBNull.Value, Nothing, dateTermCPTest)
                Params(8).Value = IIf(datePipeShearTest Is DBNull.Value, Nothing, datePipeShearTest)
                Params(9).Value = IIf(datePipeSecInsp Is DBNull.Value, Nothing, datePipeSecInsp)
                Params(10).Value = IIf(datePipeElecInsp Is DBNull.Value, Nothing, datePipeElecInsp)
                Params(11).Value = userID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

            Return True

        End Function


        ''Public Sub PutCAPStatus(ByVal facID As Integer, ByVal status As Integer)
        ''    Dim Params() As SqlParameter
        ''    Dim strSQL As String
        ''    Try
        ''        strSQL = "spPutRegFacilityCAPStatus"
        ''        Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
        ''        Params(0).Value = facID
        ''        Params(1).Value = status
        ''        SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
        ''    Catch ex As Exception
        ''        MusterException.Publish(ex, Nothing, Nothing)
        ''        Throw ex
        ''    End Try
        ''End Sub
        Public Sub PutLustStatus(ByVal facID As Integer, ByVal status As Integer)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try
                strSQL = "spPutRegFacilityLustStatus"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                Params(1).Value = status
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function DBGetDS(ByVal ownerID As Integer, Optional ByVal [Module] As String = "", _
                                Optional ByVal showDeleted As Boolean = False, Optional ByVal facID As Int64 = 0, _
                                Optional ByVal tankID As Int64 = 0, _
                                Optional ByVal intInspectionID As Integer = Integer.MinValue) As DataSet
            Dim dsData As DataSet
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Try
                Dim Params(1) As SqlParameter
                If Not (intInspectionID = Integer.MinValue) Then
                    strSQL = "spGetAllByOwnerIDArchived"
                    Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                    Params(0).Value = ownerID.ToString
                    Params(1).Value = showDeleted
                    Params(2).Value = IIf(facID = 0, DBNull.Value, facID)
                    Params(3).Value = IIf(tankID = 0, DBNull.Value, tankID)
                    Params(4).Value = IIf([Module] = "", DBNull.Value, [Module])
                    Params(5).Value = IIf(intInspectionID = 0, DBNull.Value, intInspectionID)
                Else
                    strSQL = "spGetAllByOwnerID"
                    Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                    Params(0).Value = ownerID.ToString
                    Params(1).Value = showDeleted
                    Params(2).Value = IIf(facID = 0, DBNull.Value, facID)
                    Params(3).Value = IIf(tankID = 0, DBNull.Value, tankID)
                    Params(4).Value = IIf([Module] = "", DBNull.Value, [Module])
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

