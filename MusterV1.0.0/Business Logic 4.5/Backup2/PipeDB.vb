'-------------------------------------------------------------------------------
' MUSTER.DataAccess.PipeDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR     12/14/04    Original class definition.
'  1.2        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        EN      02/09/05    Modified the strQuery in DBGetByID method. 
'  1.4        EN      02/10/05    Modified 01/01/1901 to 01/01/0001
'  1.5        AB      02/08/05    Replaced dynamic SQL with stored procedures in the following
'                                    Functions:  GetAllInfo, DBGetByID, DBGetByTankID
'  1.6        AB      02/15/05    Added Finally to the Try/Catch to close all datareaders
'  1.7        AB      02/17/05    Removed any IsNull calls for fields the DB requires
'  1.8        AB      02/17/05    Changed Put() to return the Pipe_Index from the SProc as well as the Pipe_ID
'  1.9        AB      02/18/05    Set all parameters for SP, that are not required, to NULL
'  1.10       AB      02/28/05    Modified Get functions based upon changes made to 
'                                     make several nullable fields non-nullable
'  1.11       MR      03/07/05    Changed Modified By and Modified On to reflect state after PUT
'  1.12       MR      03/14/05    Changed Created By and Created On to reflect state after PUT
'
'  1.13       TF      02/19/2009  Added pipe Status and Tank Status To DataTable for AvailablePipes

'
' Function                  Description
' GetAllInfo()      Returns an EntityCollection containing all Entity objects in the repository
' DBGetByID(ID)     Returns an EntityInfo object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(Entity)       Saves the Entity passed as an argument, to the DB
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Imports System.Data.SqlTypes

Namespace MUSTER.DataAccess
    Public Class PipeDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function GetAllInfo() As MUSTER.Info.PipesCollection
            Dim drSet As SqlDataReader
            Dim Params As Collection
            Dim strSQL As String

            Try
                strSQL = "spGetPipe"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Pipe_ID").Value = DBNull.Value
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = False

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                Dim colEntities As New MUSTER.Info.PipesCollection
                While drSet.Read
                    Dim opipeInfo As New MUSTER.Info.PipeInfo(drSet.Item("PIPE_ID"), _
                                        drSet.Item("PIPE_INDEX"), _
                                        drSet.Item("FACILITY_ID"), _
                                        AltIsDBNull(drSet.Item("COMPARTMENTS_PIPES_TANKID"), 0), _
                                        AltIsDBNull(drSet.Item("ALLD_TEST"), String.Empty), _
                                        AltIsDBNull(drSet.Item("ALLD_TEST_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("CAS_NUMBER"), 0), _
                                        AltIsDBNull(drSet.Item("CLOSURE_STATUS_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("CLOSURETYPE"), 0), _
                                        AltIsDBNull(drSet.Item("COMPOSITE_PRIMARY"), 0), _
                                        AltIsDBNull(drSet.Item("COMPOSITE_SECONDARY"), 0), _
                                        AltIsDBNull(drSet.Item("CONTAIN_SUMPDISP"), False), _
                                        AltIsDBNull(drSet.Item("CONTAIN_SUMPTANK"), False), _
                                        AltIsDBNull(drSet.Item("DATE_CLOSED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_LAST_USED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_CLOSURE_RECD"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_RECD"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_SIGNED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("ALLD_TYPE"), 0), _
                                        AltIsDBNull(drSet.Item("INERT_MATERIAL"), 0), _
                                        AltIsDBNull(drSet.Item("LCP_INSTALL_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                        AltIsDBNull(drSet.Item("CONTRACTOR_ID"), 0), _
                                        AltIsDBNull(drSet.Item("LTT_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_TEST"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateSheerValueTest"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateSecondaryContainmentInspect"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateElectronicDeviceInspect"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_TYPE"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_INSTALL_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_LD"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MANUFACTURER"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MAT_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MOD_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_OTHER_MATERIAL"), String.Empty), _
                                        AltIsDBNull(drSet.Item("PIPE_STATUS_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_TYPE_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPING_COMMENTS"), String.Empty), _
                                        AltIsDBNull(drSet.Item("PIPE_INSTALLATION_PLANNED_FOR"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PLACED_IN_SERVICE_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("SUBSTANCE_COMMENTS"), 0), _
                                        AltIsDBNull(drSet.Item("SUBSTANCE_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("TERM_CP_LAST_TESTED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERM_CP_TYPE_TANK"), 0), _
                                        AltIsDBNull(drSet.Item("TERM_CP_TYPE_DISP"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_INSTALLED_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERMINATION_CP_INSTALLED_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERMINATION_TYPE_DISP"), 0), _
                                        AltIsDBNull(drSet.Item("TERMINATION_TYPE_TANK"), 0), _
                                        AltIsDBNull(drSet.Item("DELETED"), False), _
                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), , AltIsDBNull(drSet.Item("PARENT_PIPE_ID"), 0), AltIsDBNull(drSet.Item("HAS_EXTENSIONS"), False))
                    colEntities.Add(opipeInfo)
                End While

                Return colEntities

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByID(ByVal nVal As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.PipeInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.PipeInfo
                End If
                strVal = nVal
                strSQL = "spGetPipe"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)

                Params("@Pipe_ID").Value = strVal
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.PipeInfo(drSet.Item("PIPE_ID"), _
                                        drSet.Item("PIPE_INDEX"), _
                                        drSet.Item("FACILITY_ID"), _
                                        AltIsDBNull(drSet.Item("COMPARTMENTS_PIPES_TANKID"), 0), _
                                        AltIsDBNull(drSet.Item("ALLD_TEST"), String.Empty), _
                                        AltIsDBNull(drSet.Item("ALLD_TEST_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("CAS_NUMBER"), 0), _
                                        AltIsDBNull(drSet.Item("CLOSURE_STATUS_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("CLOSURETYPE"), 0), _
                                        AltIsDBNull(drSet.Item("COMPOSITE_PRIMARY"), 0), _
                                        AltIsDBNull(drSet.Item("COMPOSITE_SECONDARY"), 0), _
                                        AltIsDBNull(drSet.Item("CONTAIN_SUMPDISP"), False), _
                                        AltIsDBNull(drSet.Item("CONTAIN_SUMPTANK"), False), _
                                        AltIsDBNull(drSet.Item("DATE_CLOSED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_LAST_USED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_CLOSURE_RECD"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_RECD"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_SIGNED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("ALLD_TYPE"), 0), _
                                        AltIsDBNull(drSet.Item("INERT_MATERIAL"), 0), _
                                        AltIsDBNull(drSet.Item("LCP_INSTALL_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                        AltIsDBNull(drSet.Item("CONTRACTOR_ID"), 0), _
                                        AltIsDBNull(drSet.Item("LTT_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_TEST"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateSheerValueTest"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateSecondaryContainmentInspect"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateElectronicDeviceInspect"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_TYPE"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_INSTALL_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_LD"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MANUFACTURER"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MAT_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MOD_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_OTHER_MATERIAL"), String.Empty), _
                                        AltIsDBNull(drSet.Item("PIPE_STATUS_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_TYPE_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPING_COMMENTS"), String.Empty), _
                                        AltIsDBNull(drSet.Item("PIPE_INSTALLATION_PLANNED_FOR"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PLACED_IN_SERVICE_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("SUBSTANCE_COMMENTS"), 0), _
                                        AltIsDBNull(drSet.Item("SUBSTANCE_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("TERM_CP_LAST_TESTED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERM_CP_TYPE_TANK"), 0), _
                                        AltIsDBNull(drSet.Item("TERM_CP_TYPE_DISP"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_INSTALLED_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERMINATION_CP_INSTALLED_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERMINATION_TYPE_DISP"), 0), _
                                        AltIsDBNull(drSet.Item("TERMINATION_TYPE_TANK"), 0), _
                                        AltIsDBNull(drSet.Item("DELETED"), False), _
                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), , AltIsDBNull(drSet.Item("PARENT_PIPE_ID"), 0), _
                                        AltIsDBNull(drSet.Item("HAS_EXTENSIONS"), False))

                Else
                    Return New MUSTER.Info.PipeInfo
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
        Public Sub Put(ByRef oPipeInf As MUSTER.Info.PipeInfo, ByRef facCapStatus As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolSaveToInspectionMirror As Boolean = False)
            Dim dtTempDate As Date
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Pipe, Integer))) Then
                    returnVal = "You do not have rights to save a Pipe."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Dim dir As ParameterDirection = ParameterDirection.Input
                Dim dataStr As String = String.Empty

                If bolSaveToInspectionMirror Then
                    dataStr = "spPutInsPipe"
                Else
                    dataStr = "spPutRegPipe"
                    dir = ParameterDirection.InputOutput

                End If
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, dataStr)
                If oPipeInf.PipeID < 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oPipeInf.PipeID
                End If
                Params(0).Direction = dir
                If oPipeInf.Index = 0 Then
                    Params(1).Value = 1
                Else
                    Params(1).Value = oPipeInf.Index
                End If
                Params(1).Direction = dir
                If oPipeInf.FacilityID <= 0 Then
                    Throw New Exception("Facility ID cannot be <= 0 for Pipe: " + oPipeInf.Index.ToString + " of Tank: " + oPipeInf.TankID.ToString + " CompNum: " + oPipeInf.CompartmentNumber.ToString)
                    Exit Try
                Else
                    Params(2).Value = oPipeInf.FacilityID
                End If
                If oPipeInf.TankID <= 0 Then
                    Throw New Exception("Tank ID cannot be <= 0 for Pipe: " + oPipeInf.Index.ToString + " of Tank: " + oPipeInf.TankID.ToString + " CompNum: " + oPipeInf.CompartmentNumber.ToString)
                    Exit Try
                Else
                    Params(3).Value = oPipeInf.TankID
                End If
                If oPipeInf.ALLDTest = String.Empty Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = oPipeInf.ALLDTest
                End If
                If Date.Compare(oPipeInf.ALLDTestDate, dtTempDate) = 0 Then
                    Params(5).Value = SqlDateTime.Null
                Else
                    Params(5).Value = oPipeInf.ALLDTestDate.Date
                End If
                If oPipeInf.CASNumber = 0 Then
                    Params(6).Value = System.DBNull.Value
                Else
                    Params(6).Value = oPipeInf.CASNumber
                End If
                'If oPipeInf.ClosureStatusDesc = 0 Then
                '    Params(7).Value = System.DBNull.Value
                'Else
                Params(7).Value = oPipeInf.ClosureStatusDesc
                'End If
                'If oPipeInf.CompPrimary = 0 Then
                '    Params(8).Value = System.DBNull.Value
                'Else
                Params(8).Value = oPipeInf.CompPrimary
                'End If
                'If oPipeInf.CompSecondary = 0 Then
                '    Params(9).Value = System.DBNull.Value
                'Else
                Params(9).Value = oPipeInf.CompSecondary
                'End If
                Params(10).Value = oPipeInf.ContainSumpDisp
                Params(11).Value = oPipeInf.ContainSumpTank
                If Date.Compare(oPipeInf.DateClosed, dtTempDate) = 0 Then
                    Params(12).Value = SqlDateTime.Null
                Else
                    Params(12).Value = oPipeInf.DateClosed.Date
                End If
                If Date.Compare(oPipeInf.DateLastUsed, dtTempDate) = 0 Then
                    Params(13).Value = SqlDateTime.Null
                Else
                    Params(13).Value = oPipeInf.DateLastUsed.Date
                End If
                If Date.Compare(oPipeInf.DateClosureRecd, dtTempDate) = 0 Then
                    Params(14).Value = SqlDateTime.Null
                Else
                    Params(14).Value = oPipeInf.DateClosureRecd.Date
                End If
                If Date.Compare(oPipeInf.DateRecd, dtTempDate) = 0 Then
                    Params(15).Value = SqlDateTime.Null
                Else
                    Params(15).Value = oPipeInf.DateRecd.Date
                End If
                If Date.Compare(oPipeInf.DateSigned, dtTempDate) = 0 Then
                    Params(16).Value = SqlDateTime.Null
                Else
                    Params(16).Value = oPipeInf.DateSigned.Date
                End If
                'If oPipeInf.ALLDType = 0 Then
                '    Params(17).Value = System.DBNull.Value
                'Else
                Params(17).Value = oPipeInf.ALLDType
                'End If
                'If oPipeInf.InertMaterial = 0 Then
                '    Params(18).Value = System.DBNull.Value
                'Else
                Params(18).Value = oPipeInf.InertMaterial
                'End If
                If Date.Compare(oPipeInf.LCPInstallDate, dtTempDate) = 0 Then
                    Params(19).Value = SqlDateTime.Null
                Else
                    Params(19).Value = oPipeInf.LCPInstallDate.Date
                End If
                If oPipeInf.LicenseeID = 0 Then
                    Params(20).Value = System.DBNull.Value
                Else
                    Params(20).Value = oPipeInf.LicenseeID
                End If
                If oPipeInf.ContractorID = 0 Then
                    Params(21).Value = System.DBNull.Value
                Else
                    Params(21).Value = oPipeInf.ContractorID
                End If
                If Date.Compare(oPipeInf.LTTDate, dtTempDate) = 0 Then
                    Params(22).Value = SqlDateTime.Null
                Else
                    Params(22).Value = oPipeInf.LTTDate.Date
                End If
                If Date.Compare(oPipeInf.PipeCPTest, dtTempDate) = 0 Then
                    Params(23).Value = SqlDateTime.Null
                Else
                    Params(23).Value = oPipeInf.PipeCPTest.Date
                End If
                'If oPipeInf.PipeCPType = 0 Then
                '    Params(24).Value = System.DBNull.Value
                'Else
                Params(24).Value = oPipeInf.PipeCPType
                'End If
                If Date.Compare(oPipeInf.PipeInstallDate, dtTempDate) = 0 Then
                    Params(25).Value = SqlDateTime.Null
                Else
                    Params(25).Value = oPipeInf.PipeInstallDate.Date
                End If
                'P!

                'If oPipeInf.PipeLD = 0 Then
                '    Params(26).Value = System.DBNull.Value
                'Else
                Params(26).Value = oPipeInf.PipeLD
                'End If
                'If oPipeInf.PipeManufacturer = 0 Then
                '    Params(27).Value = System.DBNull.Value
                'Else
                Params(27).Value = oPipeInf.PipeManufacturer
                'End If
                'If oPipeInf.PipeMatDesc = 0 Then
                '    Params(28).Value = System.DBNull.Value
                'Else
                Params(28).Value = oPipeInf.PipeMatDesc
                'End If
                'If oPipeInf.PipeModDesc = 0 Then
                '    Params(29).Value = System.DBNull.Value
                'Else
                Params(29).Value = oPipeInf.PipeModDesc
                'End If
                'If oPipeInf.PipeOtherMaterial = String.Empty Then
                '    Params(30).Value = System.DBNull.Value
                'Else
                Params(30).Value = oPipeInf.PipeOtherMaterial
                'End If
                'If oPipeInf.PipeStatusDesc = 0 Then
                '    Params(31).Value = System.DBNull.Value
                'Else
                Params(31).Value = oPipeInf.PipeStatusDesc
                'End If
                'If oPipeInf.PipeTypeDesc = 0 Then
                '    Params(32).Value = System.DBNull.Value
                'Else
                Params(32).Value = oPipeInf.PipeTypeDesc
                'End If
                'If oPipeInf.PipingComments = String.Empty Then
                '    Params(33).Value = System.DBNull.Value
                'Else
                Params(33).Value = oPipeInf.PipingComments
                'End If
                If Date.Compare(oPipeInf.PipeInstallationPlannedFor, dtTempDate) = 0 Then
                    Params(34).Value = SqlDateTime.Null
                Else
                    Params(34).Value = oPipeInf.PipeInstallationPlannedFor.Date
                End If
                If Date.Compare(oPipeInf.PlacedInServiceDate, dtTempDate) = 0 Then
                    Params(35).Value = SqlDateTime.Null
                Else
                    Params(35).Value = oPipeInf.PlacedInServiceDate.Date
                End If
                'P!

                'If oPipeInf.SubstanceComments = 0 Then
                '    Params(36).Value = System.DBNull.Value
                'Else
                Params(36).Value = oPipeInf.SubstanceComments
                'End If
                'If oPipeInf.SubstanceDesc = 0 Then
                '    Params(37).Value = System.DBNull.Value
                'Else
                Params(37).Value = oPipeInf.SubstanceDesc
                'End If
                'P!
                If Date.Compare(oPipeInf.TermCPLastTested, dtTempDate) = 0 Then
                    Params(38).Value = SqlDateTime.Null
                Else
                    Params(38).Value = oPipeInf.TermCPLastTested.Date
                End If
                'If oPipeInf.TermCPTypeTank = 0 Then
                '    Params(39).Value = System.DBNull.Value
                'Else
                Params(39).Value = oPipeInf.TermCPTypeTank
                'End If
                'If oPipeInf.TermCPTypeDisp = 0 Then
                '    Params(40).Value = System.DBNull.Value
                'Else
                Params(40).Value = oPipeInf.TermCPTypeDisp
                'End If
                'P!
                If Date.Compare(oPipeInf.PipeCPInstalledDate, dtTempDate) = 0 Then
                    Params(41).Value = SqlDateTime.Null
                Else
                    Params(41).Value = oPipeInf.PipeCPInstalledDate.Date
                End If
                If Date.Compare(oPipeInf.TermCPInstalledDate, dtTempDate) = 0 Then
                    Params(42).Value = SqlDateTime.Null
                Else
                    Params(42).Value = oPipeInf.TermCPInstalledDate.Date
                End If
                'If oPipeInf.TermTypeDisp = 0 Then
                '    Params(43).Value = System.DBNull.Value
                'Else
                Params(43).Value = oPipeInf.TermTypeDisp
                'End If
                'If oPipeInf.TermTypeTank = 0 Then
                '    Params(44).Value = System.DBNull.Value
                'Else
                Params(44).Value = oPipeInf.TermTypeTank
                'End If
                Params(45).Value = DBNull.Value
                Params(46).Value = DBNull.Value
                Params(47).Value = DBNull.Value
                Params(48).Value = DBNull.Value
                Params(49).Value = facCapStatus
                Params(50).Value = oPipeInf.ClosureType
                Params(51).Value = strUser

                If oPipeInf.HasParent Then
                    Params(52).Value = oPipeInf.ParentPipeID
                End If
                If Date.Compare(oPipeInf.DateShearTest, dtTempDate) = 0 Then
                    Params(53).Value = SqlDateTime.Null
                Else
                    Params(53).Value = oPipeInf.DateShearTest.Date
                End If
                If Date.Compare(oPipeInf.DatePipeSecInsp, dtTempDate) = 0 Then
                    Params(54).Value = SqlDateTime.Null
                Else
                    Params(54).Value = oPipeInf.DatePipeSecInsp.Date
                End If
                If Date.Compare(oPipeInf.DatePipeElecInsp, dtTempDate) = 0 Then
                    Params(55).Value = SqlDateTime.Null
                Else
                    Params(55).Value = oPipeInf.DatePipeElecInsp.Date
                End If
                'If oPipeInf.PipeID < 0 Then
                '    Params(51).Value = oPipeInf.CreatedBy
                'Else
                '    Params(51).Value = oPipeInf.ModifiedBy
                'End If



                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, dataStr, Params)
                If Integer.Parse(Params(0).Value) <> oPipeInf.PipeID Then
                    oPipeInf.PipeID = Params(0).Value
                    ' save info in tblreg_compartments_pipes table
                    Me.PutCompartmentsPipe(oPipeInf, moduleID, staffID, returnVal)
                End If
                If Params(1).Value <> oPipeInf.Index Then
                    oPipeInf.Index = Params(1).Value
                End If

                oPipeInf.ModifiedBy = AltIsDBNull(Params(45).Value, String.Empty)
                oPipeInf.ModifiedOn = AltIsDBNull(Params(46).Value, CDate("01/01/0001"))
                oPipeInf.CreatedBy = AltIsDBNull(Params(47).Value, String.Empty)
                oPipeInf.CreatedOn = AltIsDBNull(Params(48).Value, CDate("01/01/0001"))
                facCapStatus = Params(49).Value
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function DBGetByTankID(ByVal nVal As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.PipesCollection
            Dim colPipesLocal As New MUSTER.Info.PipesCollection
            Dim PipesInfoLocal As New MUSTER.Info.PipeInfo
            Dim strVal As String
            Dim strSQL As String
            Dim index As Integer = 0
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                If nVal = 0 Then
                    Return colPipesLocal
                End If
                strVal = nVal
                strSQL = "spGetPipe_andTankID"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Pipe_ID").Value = DBNull.Value
                Params("@Tank_ID").Value = strVal
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = IIf(showDeleted, DBNull.Value, False)

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        PipesInfoLocal = New MUSTER.Info.PipeInfo(drSet.Item("PIPE_ID"), _
                                        drSet.Item("PIPE_INDEX"), _
                                        drSet.Item("FACILITY_ID"), _
                                        AltIsDBNull(drSet.Item("COMPARTMENTS_PIPES_TANKID"), 0), _
                                        AltIsDBNull(drSet.Item("ALLD_TEST"), String.Empty), _
                                        AltIsDBNull(drSet.Item("ALLD_TEST_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("CAS_NUMBER"), 0), _
                                        AltIsDBNull(drSet.Item("CLOSURE_STATUS_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("CLOSURETYPE"), 0), _
                                        AltIsDBNull(drSet.Item("COMPOSITE_PRIMARY"), 0), _
                                        AltIsDBNull(drSet.Item("COMPOSITE_SECONDARY"), 0), _
                                        AltIsDBNull(drSet.Item("CONTAIN_SUMPDISP"), False), _
                                        AltIsDBNull(drSet.Item("CONTAIN_SUMPTANK"), False), _
                                        AltIsDBNull(drSet.Item("DATE_CLOSED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_LAST_USED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_CLOSURE_RECD"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_RECD"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DATE_SIGNED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("ALLD_TYPE"), 0), _
                                        AltIsDBNull(drSet.Item("INERT_MATERIAL"), 0), _
                                        AltIsDBNull(drSet.Item("LCP_INSTALL_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("LICENSEE_ID"), 0), _
                                        AltIsDBNull(drSet.Item("CONTRACTOR_ID"), 0), _
                                        AltIsDBNull(drSet.Item("LTT_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_TEST"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateSheerValueTest"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateSecondaryContainmentInspect"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("DateElectronicDeviceInspect"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_TYPE"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_INSTALL_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PIPE_LD"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MANUFACTURER"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MAT_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_MOD_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_OTHER_MATERIAL"), String.Empty), _
                                        AltIsDBNull(drSet.Item("PIPE_STATUS_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_TYPE_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("PIPING_COMMENTS"), String.Empty), _
                                        AltIsDBNull(drSet.Item("PIPE_INSTALLATION_PLANNED_FOR"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("PLACED_IN_SERVICE_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("SUBSTANCE_COMMENTS"), 0), _
                                        AltIsDBNull(drSet.Item("SUBSTANCE_DESC"), 0), _
                                        AltIsDBNull(drSet.Item("TERM_CP_LAST_TESTED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERM_CP_TYPE_TANK"), 0), _
                                        AltIsDBNull(drSet.Item("TERM_CP_TYPE_DISP"), 0), _
                                        AltIsDBNull(drSet.Item("PIPE_CP_INSTALLED_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERMINATION_CP_INSTALLED_DATE"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("TERMINATION_TYPE_DISP"), 0), _
                                        AltIsDBNull(drSet.Item("TERMINATION_TYPE_TANK"), 0), _
                                        AltIsDBNull(drSet.Item("DELETED"), False), _
                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                        AltIsDBNull(drSet.Item("COMPARTMENT_NUMBER"), 0))
                        colPipesLocal.Add(PipesInfoLocal)
                    End While
                End If
                Return colPipesLocal
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Sub PutCompartmentsPipe(ByVal oPipeInf As MUSTER.Info.PipeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strSQL As String
            Dim Params As Collection
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Pipe, Integer))) Then
                    returnVal = "You do not have rights to save a Compartment Pipe."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutREGCOMPARTMENTSPIPES"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                If oPipeInf.CompartmentNumber <= 0 Then
                    Throw New Exception("CompNum cannot be <= 0 for Pipe : " + oPipeInf.Index.ToString + " of Tank: " + oPipeInf.TankID.ToString + " CompNum: " + oPipeInf.CompartmentNumber.ToString)
                    Exit Sub
                Else
                    Params("@COMPARTMENT_NUMBER").Value = oPipeInf.CompartmentNumber
                End If
                If oPipeInf.TankID <= 0 Then
                    Throw New Exception("TankID cannot be <= 0 for Pipe: " + oPipeInf.Index.ToString + " of Tank: " + oPipeInf.TankID.ToString + " CompNum: " + oPipeInf.CompartmentNumber.ToString)
                    Exit Sub
                Else
                    Params("@TANK_ID").Value = oPipeInf.TankID
                End If
                If oPipeInf.AttachedPipeID <> 0 Then
                    Params("@PIPE_ID").Value = oPipeInf.AttachedPipeID
                Else
                    If oPipeInf.PipeID <= 0 Then
                        Throw New Exception("PipeID cannot be <= 0 for Pipe : " + oPipeInf.Index.ToString + " of Tank: " + oPipeInf.TankID.ToString + " CompNum: " + oPipeInf.CompartmentNumber.ToString)
                        Exit Sub
                    Else
                        Params("@PIPE_ID").Value = oPipeInf.PipeID
                    End If
                End If

                If oPipeInf.PipeID <= 0 Then
                    Params("@UserID").value = oPipeInf.CreatedBy
                Else
                    Params("@UserID").value = oPipeInf.ModifiedBy
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub DeletePipe(ByVal nPipeID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim strSQL As String
            Dim Params As Collection
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Pipe, Integer))) Then
                    returnVal = "You do not have rights to delete a Pipe."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spDeletePipe"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@PipeID").Value = nPipeID
                Params("@UserID").Value = UserID
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function CopyPipeProfile(ByVal PipeID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal Comp_ID As Integer = -1) As Integer
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Pipe, Integer))) Then
                    returnVal = "You do not have rights to copy Pipe Profile."
                    Exit Function
                Else
                    returnVal = String.Empty
                End If

                Dim newID As Integer = PipeID
                strSQL = "spCopyPipeProfile"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@PipeID").Value = PipeID
                Params("@UserID").Value = UserID

                If Comp_ID > -1 Then
                    Params("@Comp_ID").Value = Comp_ID
                End If

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read
                        newID = AltIsDBNull(drSet.Item("PIPE_ID"), PipeID)
                    End While
                End If
                Return newID
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetExistingPipes(ByVal facID As Integer, ByVal tnkID As Integer, ByVal CompNum As Integer) As DataTable
            Dim dtData As New DataTable
            Dim dr As DataRow
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetAvailablePipes"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FACILITY_ID").Value = facID
                Params("@TANK_ID").Value = tnkID
                Params("@COMPARTMENT_NUMBER").Value = CompNum

                dtData.Columns.Add("PipeID", Type.GetType("System.Int64"))
                dtData.Columns.Add("Pipe Site ID", Type.GetType("System.Int64"))
                dtData.Columns.Add("Pipe Site Status", Type.GetType("System.String"))
                dtData.Columns.Add("TankID", Type.GetType("System.Int64"))
                dtData.Columns.Add("Tank Site ID", Type.GetType("System.Int64"))
                dtData.Columns.Add("Tank Site Status", Type.GetType("System.String"))
                dtData.Columns.Add("Compartment Number", Type.GetType("System.Int64"))
                dtData.Columns.Add("Compartment", Type.GetType("System.Int64"))
                dtData.Columns.Add("Has Extensions", Type.GetType("System.Boolean"))
                dtData.Columns.Add("Has Parent", Type.GetType("System.Boolean"))


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    While drSet.Read
                        dr = dtData.NewRow()
                        dr("PipeID") = AltIsDBNull(drSet.Item("PIPE_ID"), Nothing)
                        dr("Pipe Site ID") = AltIsDBNull(drSet.Item("PIPE_SITE_ID"), Nothing)
                        dr("Pipe Site Status") = AltIsDBNull(drSet.Item("PIPE_STATUS"), Nothing)
                        dr("TankID") = AltIsDBNull(drSet.Item("TANK_ID"), Nothing)
                        dr("Tank Site ID") = AltIsDBNull(drSet.Item("TANK_SITE_ID"), Nothing)
                        dr("Tank Site Status") = AltIsDBNull(drSet.Item("TANK_STATUS"), Nothing)
                        dr("Compartment Number") = AltIsDBNull(drSet.Item("COMPARTMENT_NUMBER"), Nothing)
                        dr("Compartment") = AltIsDBNull(drSet.Item("COMPARTMENT"), Nothing)
                        dr("Has Extensions") = AltIsDBNull(drSet.Item("HAS_EXTENSIONS"), False)
                        dr("Has Parent") = AltIsDBNull(drSet.Item("HAS_PARENT"), False)
                        dtData.Rows.Add(dr)
                    End While
                End If
                Return dtData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
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

        Public ReadOnly Property SqlHelperProperty() As SqlHelper
            Get
                Dim sqlHelp As SqlHelper
                Return sqlHelp
            End Get
        End Property
    End Class
End Namespace
