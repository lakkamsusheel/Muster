'-------------------------------------------------------------------------------
' MUSTER.DataAccess.TankDB
'   Provides the means for marshalling Tank to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/15/04    Original class definition.
'  1.1        KJ      12/23/04    Changed all my Functions to have Null dates to be shown in DateTimePicker.
'  1.2        EN      02/10/05    Modified 01/01/1901  to 01/01/0001 
'  1.3        AB      02/15/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByFacilityID, DBGetByID
'  1.4        AB      02/16/05    Added Finally to the Try/Catch to close all datareaders
'  1.5        AB      02/16/05    Removed any IsNull calls for fields the DB requires
'  1.6        AB      02/17/05    Changed Put() to return the Tank_Index from the SProc as well as the ID
'  1.7        AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.8        AB      02/28/05    Modified Get functions based upon changes made to 
'                                     make several nullable fields non-nullable
'  1.9        MR      03/07/05    Changed Modified By and Modified On to reflect state after PUT
'  1.10       MR      03/14/05    Changed Created By and Created On to reflect state after PUT
'                                       Added NULL Validation for Created By and Created On in all the GET Methods.
'   1.11       TMF      02/18/2009     Added replacement tank msgbox in Copy Profile to determine if it is replacement

'
'
' Function                              Description
' GetAllInfo(showDeleted)   Returns an TankCollection containing all Tank objects in the repository.
' DBGetByFacilityID(nFacilityID,showDeleted)    Returns a TankCollection containing all Tank objects corresponding to a Facility ID
' DBGetByID()               Returns a TankInfo object corresponding to a TankID         
' DBGetDS(strSQL)           Returns a dataset containing the results of the select query supplied in strSQL.
' DBGetArrayList()          Returns an ArrayList. This might not be needed. - Check with J if I need it.
' PutTank(TankInfo)         Updates the repository with the information supplied in TankInfo. Inserts the
'                               data if no matching TankInfo is in the repository.
'
' IMP:  For showing blank/NULL dates in DateTimePicker I am setting the dates to CDate("01/01/0001")
'       When storing in the database I have to again check if the date is NULL. 
'       And Set Dates to SqlDateTime.Null
'       This is because by default the CDate("01/01/0001") corresponds to "# 12:00:00 AM#" which throws overflow exception in SQL Server
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/05 - JVC2
' TODO - check properties and operations against list.
'

Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes            ' Reqd for Inserting Null Vales on Dates

Namespace MUSTER.DataAccess
    Public Class TankDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function GetAllInfo(Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.TankCollection
            Dim strSQL As String
            Dim Params As Collection

            Dim drSet As SqlDataReader
            'strSQL = "select * from tblReg_Tank"
            'strSQL += IIf(Not showDeleted, " WHERE DELETED <> 1 ", "")
            Try
                strSQL = "spGetTank"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Tank_ID").value = DBNull.Value
                Params("@Facility_ID").value = DBNull.Value
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colTank As New MUSTER.Info.TankCollection
                While drSet.Read

                    'AltIsDBNull(drSet.Item("DateSecondaryContainmentLastInspected"), CDate("01/01/0001")), _

                    Dim oTankInfo As New MUSTER.Info.TankInfo(drSet.Item("TANK_ID"), _
                                                          drSet.Item("TANK_INDEX"), _
                                                          drSet.Item("FACILITY_ID"), _
                                                          AltIsDBNull(drSet.Item("TANKSTATUS"), 0), _
                                                          AltIsDBNull(drSet.Item("DATERECEIVED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("MANIFOLD"), False), _
                                                          AltIsDBNull(drSet.Item("COMPARTMENT"), False), _
                                                          AltIsDBNull(drSet.Item("TANKCAPACITY"), 0), _
                                                          AltIsDBNull(drSet.Item("SUBSTANCE"), 0), _
                                                          AltIsDBNull(drSet.Item("CASNUMBER"), 0), _
                                                          AltIsDBNull(drSet.Item("SUBSTANCECOMMENTS_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("DATELASTUSED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DATECLOSURERECEIVED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DATECLOSED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("CLOSURESTATUSDESC"), 0), _
                                                          AltIsDBNull(drSet.Item("CLOSURETYPE"), 0), _
                                                          AltIsDBNull(drSet.Item("INERTMATERIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKMATDESC"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKMODDESC"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKOTHERMATERIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("OVERFILLINSTALLED"), False), _
                                                          AltIsDBNull(drSet.Item("SPILLINSTALLED"), False), _
                                                          AltIsDBNull(drSet.Item("LICENSEEID"), 0), _
                                                          AltIsDBNull(drSet.Item("CONTRACTORID"), 0), _
                                                          AltIsDBNull(drSet.Item("DATESIGNED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DATEINSTALLEDTANK"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateSpillPreventionInstalled"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateSpillPreventionLastTested"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateOverfillPreventionInstalled"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateOverfillPreventionLastInspected"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateElectronicDeviceInspected"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateATGLastInspected"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("SMALLDELIVERY"), False), _
                                                          AltIsDBNull(drSet.Item("TANKEMERGEN"), False), _
                                                          AltIsDBNull(drSet.Item("PLANNEDINSTDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LASTTCPDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LINEDINTERIORINSTALLDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LINEDINTERIORINSPECTDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TCPINSTALLDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TTTDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TANKLD"), 0), _
                                                          AltIsDBNull(drSet.Item("OVERFILLTYPE"), 0), _
                                                          AltIsDBNull(drSet.Item("REVOKEREASON"), 0), _
                                                          AltIsDBNull(drSet.Item("REVOKEDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DatePhysicallyTagged"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("PROHIBITION"), False), _
                                                          AltIsDBNull(drSet.Item("TIGHTFILLADAPTERS"), False), _
                                                          AltIsDBNull(drSet.Item("DROPTUBE"), False), _
                                                          AltIsDBNull(drSet.Item("TANKCPTYPE"), 0), _
                                                          AltIsDBNull(drSet.Item("PLACEDINSERVICEDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TANKTYPES"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKLOCATION_DESCRIPTION"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("TANKMANUFACTURER"), 0), _
                                                          AltIsDBNull(drSet.Item("DELETED"), False), _
                                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                           AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                    colTank.Add(oTankInfo)
                End While
                Return colTank
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByFacilityID(ByVal nFacilityID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.TankCollection
            Dim colTank As New MUSTER.Info.TankCollection
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            'strSQL = "SELECT * FROM tblREG_TANK WHERE FACILITY_ID = '" + strVal + "'"
            'If Not showDeleted Then
            '    strSQL += " AND DELETED <> 1"
            'End If
            Try
                If nFacilityID = 0 Then
                    Return colTank
                End If
                strVal = nFacilityID
                strSQL = "spGetTank"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Tank_ID").value = DBNull.Value
                Params("@Facility_ID").Value = nFacilityID
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)
                Params("@OrderBy").Value = 1

                'AltIsDBNull(drSet.Item("DateSecondaryContainmentLastInspected"), CDate("01/01/0001")), _

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read
                        Dim oTankInfo As New MUSTER.Info.TankInfo(drSet.Item("TANK_ID"), _
                                                           drSet.Item("TANK_INDEX"), _
                                                           drSet.Item("FACILITY_ID"), _
                                                           AltIsDBNull(drSet.Item("TANKSTATUS"), 0), _
                                                           AltIsDBNull(drSet.Item("DATERECEIVED"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("MANIFOLD"), False), _
                                                           AltIsDBNull(drSet.Item("COMPARTMENT"), False), _
                                                           AltIsDBNull(drSet.Item("TANKCAPACITY"), 0), _
                                                           AltIsDBNull(drSet.Item("SUBSTANCE"), 0), _
                                                           AltIsDBNull(drSet.Item("CASNUMBER"), 0), _
                                                           AltIsDBNull(drSet.Item("SUBSTANCECOMMENTS_ID"), 0), _
                                                           AltIsDBNull(drSet.Item("DATELASTUSED"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DATECLOSURERECEIVED"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DATECLOSED"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("CLOSURESTATUSDESC"), 0), _
                                                           AltIsDBNull(drSet.Item("CLOSURETYPE"), 0), _
                                                           AltIsDBNull(drSet.Item("INERTMATERIAL"), 0), _
                                                           AltIsDBNull(drSet.Item("TANKMATDESC"), 0), _
                                                           AltIsDBNull(drSet.Item("TANKMODDESC"), 0), _
                                                           AltIsDBNull(drSet.Item("TANKOTHERMATERIAL"), 0), _
                                                           AltIsDBNull(drSet.Item("OVERFILLINSTALLED"), False), _
                                                           AltIsDBNull(drSet.Item("SPILLINSTALLED"), False), _
                                                           AltIsDBNull(drSet.Item("LICENSEEID"), 0), _
                                                           AltIsDBNull(drSet.Item("CONTRACTORID"), 0), _
                                                           AltIsDBNull(drSet.Item("DATESIGNED"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DATEINSTALLEDTANK"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DateSpillPreventionInstalled"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DateSpillPreventionLastTested"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DateOverfillPreventionInstalled"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DateOverfillPreventionLastInspected"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DateElectronicDeviceInspected"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DateATGLastInspected"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("SMALLDELIVERY"), False), _
                                                           AltIsDBNull(drSet.Item("TANKEMERGEN"), False), _
                                                           AltIsDBNull(drSet.Item("PLANNEDINSTDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("LASTTCPDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("LINEDINTERIORINSTALLDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("LINEDINTERIORINSPECTDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("TCPINSTALLDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("TTTDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("TANKLD"), 0), _
                                                           AltIsDBNull(drSet.Item("OVERFILLTYPE"), 0), _
                                                           AltIsDBNull(drSet.Item("REVOKEREASON"), 0), _
                                                           AltIsDBNull(drSet.Item("REVOKEDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("DatePhysicallyTagged"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("TIGHTFILLADAPTERS"), False), _
                                                           AltIsDBNull(drSet.Item("PROHIBITION"), False), _
                                                           AltIsDBNull(drSet.Item("DROPTUBE"), False), _
                                                           AltIsDBNull(drSet.Item("TANKCPTYPE"), 0), _
                                                           AltIsDBNull(drSet.Item("PLACEDINSERVICEDATE"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("TANKTYPES"), 0), _
                                                           AltIsDBNull(drSet.Item("TANKLOCATION_DESCRIPTION"), String.Empty), _
                                                           AltIsDBNull(drSet.Item("TANKMANUFACTURER"), 0), _
                                                           AltIsDBNull(drSet.Item("DELETED"), False), _
                                                           AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                           AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                           AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                        colTank.Add(oTankInfo)
                    End While
                    'Else
                    'Dim oTankInfo As New Muster.Info.TankInfo
                    'colTank.Add(otankinfo)
                End If
                Return colTank
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.TankInfo
            Dim strSQL As String
            'Dim Params As Collection
            Dim Params() As SqlParameter
            Dim drSet As SqlDataReader
            'strSQL = "SELECT * FROM tblREG_TANK WHERE TANK_ID = " & nVal
            'strSQL += IIf(Not showDeleted, " AND DELETED <> 1 ", "")
            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.TankInfo
                End If

                strSQL = "spGetTank"
                'Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                'Params("@Tank_ID").Value = nVal
                'Params("@Facility_ID").Value = DBNull.Value
                'Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)
                'Params("@OrderBy").Value = 1
                Params(0).Value = nVal
                Params(1).Value = DBNull.Value
                Params(2).Value = IIf(showDeleted, DBNull.Value, False)
                Params(3).Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.TankInfo(drSet.Item("TANK_ID"), _
                                                          drSet.Item("TANK_INDEX"), _
                                                          drSet.Item("FACILITY_ID"), _
                                                          AltIsDBNull(drSet.Item("TANKSTATUS"), 0), _
                                                          AltIsDBNull(drSet.Item("DATERECEIVED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("MANIFOLD"), False), _
                                                          AltIsDBNull(drSet.Item("COMPARTMENT"), False), _
                                                          AltIsDBNull(drSet.Item("TANKCAPACITY"), 0), _
                                                          AltIsDBNull(drSet.Item("SUBSTANCE"), 0), _
                                                          AltIsDBNull(drSet.Item("CASNUMBER"), 0), _
                                                          AltIsDBNull(drSet.Item("SUBSTANCECOMMENTS_ID"), 0), _
                                                          AltIsDBNull(drSet.Item("DATELASTUSED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DATECLOSURERECEIVED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DATECLOSED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("CLOSURESTATUSDESC"), 0), _
                                                          AltIsDBNull(drSet.Item("CLOSURETYPE"), 0), _
                                                          AltIsDBNull(drSet.Item("INERTMATERIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKMATDESC"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKMODDESC"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKOTHERMATERIAL"), 0), _
                                                          AltIsDBNull(drSet.Item("OVERFILLINSTALLED"), False), _
                                                          AltIsDBNull(drSet.Item("SPILLINSTALLED"), False), _
                                                          AltIsDBNull(drSet.Item("LICENSEEID"), 0), _
                                                          AltIsDBNull(drSet.Item("CONTRACTORID"), 0), _
                                                          AltIsDBNull(drSet.Item("DATESIGNED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DATEINSTALLEDTANK"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateSpillPreventionInstalled"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateSpillPreventionLastTested"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateOverfillPreventionInstalled"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateOverfillPreventionLastInspected"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateElectronicDeviceInspected"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DateATGLastInspected"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("SMALLDELIVERY"), False), _
                                                          AltIsDBNull(drSet.Item("TANKEMERGEN"), False), _
                                                          AltIsDBNull(drSet.Item("PLANNEDINSTDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LASTTCPDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LINEDINTERIORINSTALLDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LINEDINTERIORINSPECTDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TCPINSTALLDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TTTDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TANKLD"), 0), _
                                                          AltIsDBNull(drSet.Item("OVERFILLTYPE"), 0), _
                                                          AltIsDBNull(drSet.Item("REVOKEREASON"), 0), _
                                                          AltIsDBNull(drSet.Item("REVOKEDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("DatePhysicallyTagged"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("PROHIBITION"), False), _
                                                          AltIsDBNull(drSet.Item("TIGHTFILLADAPTERS"), False), _
                                                          AltIsDBNull(drSet.Item("DROPTUBE"), False), _
                                                          AltIsDBNull(drSet.Item("TANKCPTYPE"), 0), _
                                                          AltIsDBNull(drSet.Item("PLACEDINSERVICEDATE"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("TANKTYPES"), 0), _
                                                          AltIsDBNull(drSet.Item("TANKLOCATION_DESCRIPTION"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("TANKMANUFACTURER"), 0), _
                                                          AltIsDBNull(drSet.Item("DELETED"), False), _
                                                          AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                          AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                          AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")))
                End If
                Return New MUSTER.Info.TankInfo
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
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
        Private Function DBGetArrayList(ByVal strSQL As String) As ArrayList
            Dim LookupArrList As New ArrayList
            Dim dtRdrLookup As SqlDataReader
            Try
                dtRdrLookup = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
                DBGetArrayList = Nothing
                'Read each row from the dataReader
                While dtRdrLookup.Read()
                    'Adding the data read from datareader to ArrayList
                    LookupArrList.Add(New LookupProperty(IIf(dtRdrLookup("PROPERTY_NAME") Is System.DBNull.Value, "", dtRdrLookup("PROPERTY_NAME")), IIf(dtRdrLookup("PROPERTY_ID") Is System.DBNull.Value, "", dtRdrLookup("PROPERTY_ID"))))
                End While
                DBGetArrayList = LookupArrList

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not dtRdrLookup.IsClosed Then dtRdrLookup.Close()
            End Try
        End Function
        Public Sub Put(ByRef obj As MUSTER.Info.TankInfo, ByRef facCapStatus As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolReplacementTank As Boolean = False, Optional ByVal bolSaveToInspectionMirror As Boolean = False)
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Tank, Integer))) Then
                    returnVal = "You do not have rights to save a Tank."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim dataStr As String = String.Empty
                Dim dir As ParameterDirection = ParameterDirection.Input

                If bolSaveToInspectionMirror Then
                    dataStr = "spPutInsTank"
                Else
                    dataStr = "spPutRegTank"
                    dir = ParameterDirection.InputOutput

                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, dataStr)
                Dim dtTempDate As Date      'For comparing the Null date
                If obj.TankId < 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = obj.TankId
                End If
                Params(0).Direction = dir
                Params(1).Value = obj.TankIndex
                Params(1).Direction = dir
                Params(2).Value = obj.FacilityId
                Params(3).Value = obj.TankStatus
                If Date.Compare(obj.DateReceived, dtTempDate) = 0 Then
                    Params(4).Value = SqlDateTime.Null
                Else
                    Params(4).Value = obj.DateReceived.Date
                End If

                Params(5).Value = AltIsDBNull(obj.Manifold, 0)
                Params(6).Value = AltIsDBNull(obj.Compartment, 0)
                'If obj.TankCapacity = 0 Then
                '    Params(7).Value = System.DBNull.Value
                'Else
                Params(7).Value = AltIsDBNull(obj.TankCapacity, 0)
                'End If
                'If obj.Substance = 0 Then
                '    Params(8).Value = System.DBNull.Value
                'Else
                Params(8).Value = AltIsDBNull(obj.Substance, 0)
                'End If
                'If obj.CASNumber = 0 Then
                '    Params(9).Value = System.DBNull.Value
                'Else
                Params(9).Value = AltIsDBNull(obj.CASNumber, 0)
                'End If
                If obj.SubstanceCommentsID = 0 Then
                    Params(10).Value = System.DBNull.Value
                Else
                    Params(10).Value = obj.SubstanceCommentsID
                End If
                If Date.Compare(obj.DateLastUsed, dtTempDate) = 0 Then
                    Params(11).Value = SqlDateTime.Null
                Else
                    Params(11).Value = obj.DateLastUsed.Date
                End If

                If Date.Compare(obj.DateClosureReceived, dtTempDate) = 0 Then
                    Params(12).Value = SqlDateTime.Null
                Else
                    Params(12).Value = obj.DateClosureReceived.Date
                End If

                If Date.Compare(obj.DateClosed, dtTempDate) = 0 Then
                    Params(13).Value = SqlDateTime.Null
                Else
                    Params(13).Value = obj.DateClosed.Date
                End If

                'If obj.ClosureStatusDesc = 0 Then
                '    Params(14).Value = System.DBNull.Value
                'Else
                Params(14).Value = AltIsDBNull(obj.ClosureStatusDesc, 0)
                'End If
                'If obj.InertMaterial = 0 Then
                '    Params(15).Value = System.DBNull.Value
                'Else
                Params(15).Value = AltIsDBNull(obj.InertMaterial, 0)
                'End If
                Params(16).Value = AltIsDBNull(obj.TankMatDesc, 0)
                Params(17).Value = AltIsDBNull(obj.TankModDesc, 0)
                If obj.TankOtherMaterial = 0 Then
                    Params(18).Value = 351
                Else
                    Params(18).Value = obj.TankOtherMaterial
                End If
                Params(19).Value = AltIsDBNull(obj.OverFillInstalled, 0)
                Params(20).Value = AltIsDBNull(obj.SpillInstalled, 0)
                If obj.LicenseeID = 0 Then
                    Params(21).Value = DBNull.Value
                Else
                    Params(21).Value = obj.LicenseeID
                End If
                If obj.ContractorID = 0 Then
                    Params(22).Value = DBNull.Value
                Else
                    Params(22).Value = obj.ContractorID
                End If
                If Date.Compare(obj.DateSigned, dtTempDate) = 0 Then
                    Params(23).Value = SqlDateTime.Null
                Else
                    Params(23).Value = obj.DateSigned.Date
                End If

                If Date.Compare(obj.DateInstalledTank, dtTempDate) = 0 Then
                    Params(24).Value = SqlDateTime.Null
                Else
                    Params(24).Value = obj.DateInstalledTank.Date
                End If
                If Date.Compare(obj.DateSpillInstalled, dtTempDate) = 0 Then
                    Params(25).Value = SqlDateTime.Null
                Else
                    Params(25).Value = obj.DateSpillInstalled.Date
                End If
                If Date.Compare(obj.DateSpillTested, dtTempDate) = 0 Then
                    Params(26).Value = SqlDateTime.Null
                Else
                    Params(26).Value = obj.DateSpillTested.Date
                End If
                If Date.Compare(obj.DateOverfillInstalled, dtTempDate) = 0 Then
                    Params(27).Value = SqlDateTime.Null
                Else
                    Params(27).Value = obj.DateOverfillInstalled.Date
                End If
                If Date.Compare(obj.DateOverfillTested, dtTempDate) = 0 Then
                    Params(28).Value = SqlDateTime.Null
                Else
                    Params(28).Value = obj.DateOverfillTested.Date
                End If
                'If Date.Compare(obj.DateTankSecInsp, dtTempDate) = 0 Then
                Params(29).Value = SqlDateTime.Null
                ' Else
                    ' Params(29).Value = obj.DateTankSecInsp.Date
                'End If
                If Date.Compare(obj.DateTankElecInsp, dtTempDate) = 0 Then
                    Params(30).Value = SqlDateTime.Null
                Else
                    Params(30).Value = obj.DateTankElecInsp.Date
                End If
                If Date.Compare(obj.DateATGInsp, dtTempDate) = 0 Then
                    Params(31).Value = SqlDateTime.Null
                Else
                    Params(31).Value = obj.DateATGInsp.Date
                End If

                Params(32).Value = AltIsDBNull(obj.SmallDelivery, 0)
                Params(33).Value = AltIsDBNull(obj.TankEmergen, 0)
                If Date.Compare(obj.PlannedInstDate, dtTempDate) = 0 Then
                    Params(34).Value = SqlDateTime.Null
                Else
                    Params(34).Value = obj.PlannedInstDate.Date
                End If

                If Date.Compare(obj.LastTCPDate, dtTempDate) = 0 Then
                    Params(35).Value = SqlDateTime.Null
                Else
                    Params(35).Value = obj.LastTCPDate.Date
                End If

                If Date.Compare(obj.LinedInteriorInstallDate, dtTempDate) = 0 Then
                    Params(36).Value = SqlDateTime.Null
                Else
                    Params(36).Value = obj.LinedInteriorInstallDate.Date
                End If

                If Date.Compare(obj.LinedInteriorInspectDate, dtTempDate) = 0 Then
                    Params(37).Value = SqlDateTime.Null
                Else
                    Params(37).Value = obj.LinedInteriorInspectDate.Date
                End If

                If Date.Compare(obj.TCPInstallDate, dtTempDate) = 0 Then
                    Params(38).Value = SqlDateTime.Null
                Else
                    Params(38).Value = obj.TCPInstallDate.Date
                End If

                If Date.Compare(obj.TTTDate, dtTempDate) = 0 Then
                    Params(39).Value = SqlDateTime.Null
                Else
                    Params(39).Value = obj.TTTDate.Date
                End If

                Params(40).Value = AltIsDBNull(obj.TankLD, 0)
                'If obj.OverFillType = 0 Then
                '    Params(34).Value = System.DBNull.Value
                'Else
                Params(41).Value = AltIsDBNull(obj.OverFillType, 0)
                'End If
                Params(42).Value = AltIsDBNull(obj.TightFillAdapters, 0)
                Params(43).Value = AltIsDBNull(obj.DropTube, 0)
                'If obj.TankCPType = 0 Then
                '    Params(37).Value = System.DBNull.Value
                'Else
                Params(44).Value = AltIsDBNull(obj.TankCPType, 0)
                'End If
                If Date.Compare(obj.PlacedInServiceDate, dtTempDate) = 0 Then
                    Params(45).Value = SqlDateTime.Null
                Else
                    Params(45).Value = obj.PlacedInServiceDate.Date
                End If

                'If obj.TankTypes = 0 Then
                '    Params(39).Value = System.DBNull.Value
                'Else
                Params(46).Value = AltIsDBNull(obj.TankTypes, 0)
                'End If
                'If obj.TankLocationDescription = String.Empty Then
                '    Params(40).Value = System.DBNull.Value
                'Else
                Params(47).Value = AltIsDBNull(obj.TankLocationDescription, 0)
                'End If
                Params(48).Value = AltIsDBNull(obj.TankManufacturer, 0)
                Params(49).Value = obj.Deleted
                Params(50).Value = DBNull.Value
                Params(51).Value = DBNull.Value
                Params(52).Value = DBNull.Value
                Params(53).Value = DBNull.Value
                Params(54).Value = facCapStatus
                Params(55).Value = obj.ClosureType
                Params(56).Value = strUser
                Params(57).Value = bolReplacementTank
                Params(58).Value = obj.RevokeReason
                Params(59).Value = obj.Prohibition
                '     Params(60).Value = obj.RevokeDate
                If Date.Compare(obj.RevokeDate, dtTempDate) = 0 Then
                    Params(60).Value = SqlDateTime.Null
                Else
                    Params(60).Value = obj.RevokeDate.Date
                End If
                If Date.Compare(obj.DatePhysicallyTagged, dtTempDate) = 0 Then
                    Params(61).Value = SqlDateTime.Null
                Else
                    Params(61).Value = obj.DatePhysicallyTagged.Date
                End If
                'If obj.TankId <= 0 Then
                '    Params(49).Value = obj.CreatedBy
                'Else
                '    Params(49).Value = obj.ModifiedBy
                'End If



                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, dataStr, Params)
                If Params(0).Value <> obj.TankId Then
                    obj.TankId = Params(0).Value
                End If
                If Params(1).Value <> obj.TankIndex Then
                    obj.TankIndex = Params(1).Value
                End If
                obj.ModifiedBy = AltIsDBNull(Params(50).Value, String.Empty)
                obj.ModifiedOn = AltIsDBNull(Params(51).Value, CDate("01/01/0001"))
                obj.CreatedBy = AltIsDBNull(Params(52).Value, String.Empty)
                obj.CreatedOn = AltIsDBNull(Params(53).Value, CDate("01/01/0001"))
                facCapStatus = Params(54).Value
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        'Public Sub DeleteTank(ByVal nTankID As Integer)
        '    Dim Params() As SqlParameter
        '    Try
        '        Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spDeleteTank")
        '        Params(0).Value = nTankID
        '        SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spDeleteTank", Params)
        '    Catch ex As Exception

        '    End Try
        'End Sub
        Public Function CopyTankProfile(ByVal tnkID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String) As Integer
            Dim strSQL As String
            Dim Params As Collection
            Dim replacementTank As Object = "N"
            Dim newID As Integer = tnkID
            Dim drSet As SqlClient.SqlDataReader

            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Tank, Integer))) Then
                    returnVal = "You do not have rights to copy Tank Profile."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spCopyTankProfile"

                If MsgBox("Will this be a replacement tank? ", MsgBoxStyle.YesNo, "Copy Tank Profile Function") = MsgBoxResult.Yes Then
                    replacementTank = "Y"
                End If

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@TankID").value = tnkID
                Params("@UserID").value = UserID
                Params("@Replacement").value = replacementTank
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not drSet Is Nothing AndAlso drSet.HasRows Then

                    While (drSet.Read)
                        newID = drSet.Item("TANK ID")
                    End While

                    drSet.Close()
                End If

                Return newID

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally

            End Try
        End Function
        Public Function DBGetCompanyDetails(Optional ByVal LicenseeID As Integer = 0) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCOMCompanyNAME"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = LicenseeID
                Params(1).Value = DBNull.Value

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetDS(ByVal ownerID As Integer, Optional ByVal [Module] As String = "", Optional ByVal showDeleted As Boolean = False, Optional ByVal facID As Int64 = 0, Optional ByVal tankID As Int64 = 0) As DataSet
            Dim dsData As DataSet
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
        Public Function DBGetAttachedPipeIDs(ByVal tankID As Integer) As String
            Try
                Dim strSQL As String
                strSQL = "select distinct pipe_id from tblREG_COMPARTMENTS_PIPES where deleted = 0 and pipe_id in (" + _
                            "select pipe_id from tblREG_COMPARTMENTS_PIPES where deleted = 0 and tank_id = " + tankID.ToString + ") " + _
                            "group by pipe_id having count(tank_id) > 1"
                Dim dsPipeIDs As DataSet = DBGetDS(strSQL)
                strSQL = ""
                For Each dr As DataRow In dsPipeIDs.Tables(0).Rows
                    strSQL += dr("pipe_id").ToString + ", "
                Next
                If strSQL.Length > 0 Then
                    strSQL = strSQL.Trim.TrimEnd(",")
                End If
                Return strSQL
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub DetachPipe(ByVal pipeID As Integer, ByVal tankID As Integer, ByVal compNum As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String)
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Pipe, Integer))) Then
                    returnVal = "You do not have rights to Detach Pipe(s)."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim strSQL As String = "spDetachPipe"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = pipeID
                Params(1).Value = tankID
                Params(2).Value = compNum
                Params(3).Value = strUser

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public ReadOnly Property SqlHelperProperty() As SqlHelper
            Get
                Dim sqlHelp As SqlHelper
                Return sqlHelp
            End Get
        End Property
    End Class
End Namespace
