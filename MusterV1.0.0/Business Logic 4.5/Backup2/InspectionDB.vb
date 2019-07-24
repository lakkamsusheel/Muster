'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1        PN                    added Assigneddate  
'   1.2         Kumar 08/02/2005    Added assigned date in the put method as the last paramater, changed the spPutInspection SP too
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as Template to build other objects.
'       Replace keyword "Template" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
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

        Public Function DBGetByID(Optional ByVal id As Integer = 0, Optional ByVal staffID As Int64 = 0, Optional ByVal facID As Int64 = 0, Optional ByVal ownerID As Int64 = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionsCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                If id = 0 And staffID = 0 And facID = 0 And ownerID = 0 Then
                    Return New MUSTER.Info.InspectionsCollection
                End If

                strSQL = "spGetInspection"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@InspectionID").Value = IIf(id = 0, DBNull.Value, id)
                Params("@OwnerID").Value = IIf(ownerID = 0, DBNull.Value, ownerID)
                Params("@FacilityID").Value = IIf(facID = 0, DBNull.Value, facID)
                Params("@StaffID").Value = IIf(staffID = 0, DBNull.Value, staffID)
                Params("@OrderBy").Value = 1
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                Dim colEntities As New MUSTER.Info.InspectionsCollection

                While drSet.Read
                    Dim oInspectionInfo As New MUSTER.Info.InspectionInfo(drSet.Item("INSPECTION_ID"), _
                                                                            drSet.Item("OWNER_ID"), _
                                                                            drSet.Item("FACILITY_ID"), _
                                                                            drSet.Item("INSPECTION_TYPE"), _
                                                                            AltIsDBNull(drSet.Item("ADMIN_COMMENTS"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("SCHEDULED_DATE"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("SCHEDULED_TIME"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("CHECKLIST_GENERATION_DATE"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("SUBMITTED_DATE"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("STAFF_ID"), 0), _
                                                                            AltIsDBNull(drSet.Item("ASSIGNED_DATE"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("SCHEDULED_BY"), String.Empty), _
                                                                            drSet.Item("LETTER_GENERATED"), _
                                                                            AltIsDBNull(drSet.Item("DATE_GENERATED"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("DATE_PLANNED"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("CONDUCT_INSPECTION_ON"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("RESCHEDULED_DATE"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("RESCHEDULED_TIME"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("INSPECTOR_COMMENTS"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("COMPLETED"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                            drSet.Item("DELETED"), _
                                                                            AltIsDBNull(drSet.Item("OWNERS_REP"), String.Empty), _
                                                                            AltIsDBNull(drSet.Item("CAP_DATES_ENTERED"), False), _
                                                                            AltIsDBNull(drSet.Item("INSPECTION_ACCPETED"), False), _
                                                                            AltIsDBNull(drSet.Item("CAE_VIEWED"), CDate("01/01/0001")), _
                                                                            AltIsDBNull(drSet.Item("CHECKLIST_FIRST_SAVED"), CDate("01/01/0001")))

                    colEntities.Add(oInspectionInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
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
        Public Function DBGetInspectionHistory(ByVal facID As Int64, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params(1) As SqlParameter
            Dim dsData As DataSet
            Try
                If facID = 0 Then
                    dsData = New DataSet
                    Return dsData
                End If
                strSQL = "spGetInspectionHistory"
                'Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                'Params("@FacilityID").Value = facID
                'Params("@Deleted").Value = showDeleted
                Params(0).Value = facID
                Params(1).Value = showDeleted
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub Put(ByRef oInspectionInfo As MUSTER.Info.InspectionInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal OverrideRights As Boolean = False)
            Try
                If Not OverrideRights Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                        returnVal = "You do not have rights to save Inspection."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                Else
                    returnVal = String.Empty
                End If

                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutInspection")
                Params(0).Value = IIf(oInspectionInfo.ID < 0, 0, oInspectionInfo.ID)
                Params(1).Value = oInspectionInfo.OwnerID
                Params(2).Value = oInspectionInfo.FacilityID
                Params(3).Value = oInspectionInfo.InspectionType
                Params(4).Value = IIf(oInspectionInfo.AdminComments = String.Empty, DBNull.Value, oInspectionInfo.AdminComments)
                If Date.Compare(oInspectionInfo.ScheduledDate, CDate("01/01/0001")) = 0 Then
                    Params(5).Value = DBNull.Value
                Else
                    Params(5).Value = oInspectionInfo.ScheduledDate
                End If
                Params(6).Value = IIf(oInspectionInfo.ScheduledTime = String.Empty, DBNull.Value, oInspectionInfo.ScheduledTime)
                If Date.Compare(oInspectionInfo.CheckListGenDate, CDate("01/01/0001")) = 0 Then
                    Params(7).Value = DBNull.Value
                Else
                    Params(7).Value = oInspectionInfo.CheckListGenDate
                End If
                If Date.Compare(oInspectionInfo.SubmittedDate, CDate("01/01/0001")) = 0 Then
                    Params(8).Value = DBNull.Value
                Else
                    Params(8).Value = oInspectionInfo.SubmittedDate
                End If
                Params(9).Value = IIf(oInspectionInfo.StaffID = 0, DBNull.Value, oInspectionInfo.StaffID)
                Params(10).Value = IIf(oInspectionInfo.ScheduledBy = String.Empty, DBNull.Value, oInspectionInfo.ScheduledBy)
                Params(11).Value = oInspectionInfo.LetterGenerated
                If Date.Compare(oInspectionInfo.DateLetterGenerated, CDate("01/01/0001")) = 0 Then
                    Params(12).Value = DBNull.Value
                Else
                    Params(12).Value = oInspectionInfo.DateLetterGenerated
                End If
                If Date.Compare(oInspectionInfo.DatePlanned, CDate("01/01/0001")) = 0 Then
                    Params(13).Value = DBNull.Value
                Else
                    Params(13).Value = oInspectionInfo.DatePlanned
                End If
                If Date.Compare(oInspectionInfo.ConductInspectionOn, CDate("01/01/0001")) = 0 Then
                    Params(14).Value = DBNull.Value
                Else
                    Params(14).Value = oInspectionInfo.ConductInspectionOn
                End If
                If Date.Compare(oInspectionInfo.RescheduledDate, CDate("01/01/0001")) = 0 Then
                    Params(15).Value = DBNull.Value
                Else
                    Params(15).Value = oInspectionInfo.RescheduledDate
                End If
                Params(16).Value = IIf(oInspectionInfo.RescheduledTime = String.Empty, DBNull.Value, oInspectionInfo.RescheduledTime)
                Params(17).Value = IIf(oInspectionInfo.InspectorComments = String.Empty, DBNull.Value, oInspectionInfo.InspectorComments)
                If Date.Compare(oInspectionInfo.Completed, CDate("01/01/0001")) = 0 Then
                    Params(18).Value = DBNull.Value
                Else
                    Params(18).Value = oInspectionInfo.Completed
                End If
                Params(19).Value = IIf(oInspectionInfo.CreatedBy = String.Empty, DBNull.Value, oInspectionInfo.CreatedBy)
                If Date.Compare(oInspectionInfo.CreatedOn, CDate("01/01/0001")) = 0 Then
                    Params(20).Value = DBNull.Value
                Else
                    Params(20).Value = oInspectionInfo.CreatedOn
                End If
                Params(21).Value = IIf(oInspectionInfo.ModifiedBy = String.Empty, DBNull.Value, oInspectionInfo.ModifiedBy)
                If Date.Compare(oInspectionInfo.ModifiedOn, CDate("01/01/0001")) = 0 Then
                    Params(22).Value = DBNull.Value
                Else
                    Params(22).Value = oInspectionInfo.ModifiedOn
                End If
                Params(23).Value = oInspectionInfo.Deleted
                If Date.Compare(oInspectionInfo.AssignedDate, CDate("01/01/0001")) = 0 Then
                    Params(24).Value = DBNull.Value
                Else
                    Params(24).Value = oInspectionInfo.AssignedDate
                End If
                Params(25).Value = IIf(oInspectionInfo.OwnersRep = String.Empty, DBNull.Value, oInspectionInfo.OwnersRep)
                Params(26).Value = oInspectionInfo.CAPDatesEntered
                Params(27).Value = oInspectionInfo.InspectionAccepted
                If Date.Compare(oInspectionInfo.CAEViewed, CDate("01/01/0001")) = 0 Then
                    Params(28).Value = DBNull.Value
                Else
                    Params(28).Value = oInspectionInfo.CAEViewed
                End If

                If oInspectionInfo.ID <= 0 Then
                    Params(29).Value = oInspectionInfo.CreatedBy
                Else
                    Params(29).Value = oInspectionInfo.ModifiedBy
                End If
                If Date.Compare(oInspectionInfo.ChecklistFirstSaved, CDate("01/01/0001")) = 0 Then
                    Params(30).Value = DBNull.Value
                Else
                    Params(30).Value = oInspectionInfo.ChecklistFirstSaved
                End If

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutInspection", Params)
                If Params(0).Value <> oInspectionInfo.ID Then
                    oInspectionInfo.ID = Params(0).Value
                End If
                oInspectionInfo.CreatedBy = AltIsDBNull(Params(19).Value, String.Empty)
                oInspectionInfo.CreatedOn = AltIsDBNull(Params(20).Value, CDate("01/01/0001"))
                oInspectionInfo.ModifiedBy = AltIsDBNull(Params(21).Value, String.Empty)
                oInspectionInfo.ModifiedOn = AltIsDBNull(Params(22).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function DBGetTargetFacilities(ByVal nStaffID As Integer, Optional ByVal ownerID As Integer = 0, Optional ByVal bolAllTargetFacs As Boolean = False, Optional ByVal reLoadInspectionID As Integer = 0, Optional ByVal facilityID As Integer = 0) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetTargetFacilities"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = nStaffID
                Params(1).Value = IIf(ownerID = 0, DBNull.Value, ownerID)
                Params(2).Value = bolAllTargetFacs
                Params(3).Value = reLoadInspectionID
                Params(4).Value = facilityID

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DBGetAssignedFacilities(ByVal nStaffID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetAssignedInspection"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(nStaffID = 0, DBNull.Value, nStaffID)
                Params(1).Value = showDeleted

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DBGetCAEFacilityAssignedTo(ByVal facID As Integer) As Integer
            Dim Params() As SqlParameter
            Dim nStaffID As Integer = 0
            Dim strSQL As String
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetCAEFacilityAssignedTo"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    nStaffID = drSet.Item("STAFF_ID")
                End If
                If Not drSet.IsClosed Then drSet.Close()
                Return nStaffID
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Sub PutCAPTankPipeDataBeforeAfterInspection(ByVal inspectionID As Integer, ByVal isBeforeInspectionData As Boolean, ByVal bolDeleteData As Boolean)
            Try
                Dim strSQL As String = "spPutCAPTankPipeDataBeforeAfterInspection"
                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = inspectionID
                Params(1).Value = isBeforeInspectionData
                Params(2).Value = bolDeleteData

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub PutTankPipeDataFromMirrortable(ByVal facilityID As Integer)
            Try
                Dim strSQL As String = "spUpdateTanksPipesFromInspectionViewToTable"
                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facilityID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function HasfacilitybeenInspectedBeforeDB(ByVal facilityID As Integer) As Date
            Try
                Dim strSQL As String = "spGetFacilityLastInspected"
                Dim Params() As SqlParameter
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facilityID

                Dim beenInspected As Date = New Date(1900, 1, 1)

                beenInspected = SqlHelper.ExecuteScalar(_strConn, CommandType.StoredProcedure, strSQL, Params)

                Return beenInspected

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Sub PutInspectionArchive(ByVal facilityID As Integer, ByVal inspectionID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal OverrideRights As Boolean = False, Optional ByVal bolSaveTankPipeTermOnly As Boolean = False)
            Try
                Dim strSQL As String = "spUpdateInspectionArchive"
                Dim Params() As SqlParameter

                If Not OverrideRights Then
                    If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                        returnVal = "You do not have rights to save Inspection."
                        Exit Sub
                    Else
                        returnVal = String.Empty
                    End If
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facilityID
                Params(1).Value = inspectionID
                Params(2).Value = bolSaveTankPipeTermOnly

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub UpdateInspectionSOC(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal inspectionID As Integer, ByVal strUser As String)
            Dim Params() As SqlParameter
            Try
                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spUpdateInspectionSOC")
                Params(0).Value = inspectionID
                Params(1).Value = strUser

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spUpdateInspectionSOC", Params)

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
    End Class

End Namespace
