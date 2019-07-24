'-------------------------------------------------------------------------------
' MUSTER.DataAccess.InspectionChecklistMasterDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/11/05    Original class definition
'
'
' Function                  Description
' GetAllInfo()        Returns an EntityCollection containing all Entity objects in the repository.
' DBGetByName(NAME)   Returns an EntityInfo object indicated by arg NAME.
' DBGetByID(ID)       Returns an EntityInfo object indicated by arg ID.
'
' NOTE: This file to be used as InspectionChecklistMaster to build other objects.
'       Replace keyword "InspectionChecklistMaster" with respective Object name.
'       In addition, change SQL strings and data stream Items as necessary
'       Don't forget to update the history information above!!!
'-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class InspectionChecklistMasterDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region

        Public Function DBGetByID(ByVal id As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.InspectionChecklistMasterInfo
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetInspectionCheckList"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@QuestionID").Value = id
                Params("@Deleted").Value = showDeleted

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionChecklistMasterInfo(drSet.Item("QUESTION_ID"), _
                                                                        AltIsDBNull(drSet.Item("CHKLST_POSITION"), 0), _
                                                                        AltIsDBNull(drSet.Item("CHKLST_ITEM_NUMBER"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("SOC"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("HEADER"), False), _
                                                                        AltIsDBNull(drSet.Item("HEADER_QUESTION_TEXT"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_TANK"), False), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_PIPE"), False), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_PIPETERM"), False), _
                                                                        AltIsDBNull(drSet.Item("CITATION"), 0), _
                                                                        AltIsDBNull(drSet.Item("DISCREP_TEXT"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("WHEN_VISIBLE"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("CCAT"), False), _
                                                                        AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                        AltIsDBNull(drSet.Item("FORE_COLOR"), "BLACK"), _
                                                                        AltIsDBNull(drSet.Item("BACK_COLOR"), "WHITE"))
                Else
                    Return New MUSTER.Info.InspectionChecklistMasterInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetByCheckListItemNum(ByVal chkListItemNum As String) As MUSTER.Info.InspectionChecklistMasterInfo
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetInspectionCheckListItem"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@CHKLST_ITEM_NUMBER").Value = chkListItemNum
                Params("@Deleted").Value = True ' do not change to false as the item retrieved is deleted in db. its used only to maintain the relation between citaton and discrep for assigned inspection / fce

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.InspectionChecklistMasterInfo(drSet.Item("QUESTION_ID"), _
                                                                        AltIsDBNull(drSet.Item("CHKLST_POSITION"), 0), _
                                                                        AltIsDBNull(drSet.Item("CHKLST_ITEM_NUMBER"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("SOC"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("HEADER"), False), _
                                                                        AltIsDBNull(drSet.Item("HEADER_QUESTION_TEXT"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_TANK"), False), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_PIPE"), False), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_PIPETERM"), False), _
                                                                        AltIsDBNull(drSet.Item("CITATION"), 0), _
                                                                        AltIsDBNull(drSet.Item("DISCREP_TEXT"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("WHEN_VISIBLE"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("CCAT"), False), _
                                                                        AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                        AltIsDBNull(drSet.Item("FORE_COLOR"), "BLACK"), _
                                                                        AltIsDBNull(drSet.Item("BACK_COLOR"), "WHITE"))
                Else
                    Return New MUSTER.Info.InspectionChecklistMasterInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
        End Function
        Public Function DBGetChecklistText(Optional ByVal showDeleted As Boolean = False, Optional ByVal showFac_ID As Integer = 0, Optional ByVal useNewCode As Boolean = False, Optional ByVal inspID As Int32 = 0) As MUSTER.Info.InspectionChecklistMastersCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try

                strSQL = IIf(useNewCode, "spGetInspectionCheckListByFacility", "spGetInspectionCheckList")
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@Deleted").Value = showDeleted
                If useNewCode Then
                    Params("@INSPECTIONID").value = inspID

                    If inspID = 0 Then
                        Params("@FACILITY_ID").Value = showFac_ID
                    End If
                Else
                    Params("@FACILITY_ID").Value = showFac_ID

                End If


                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                Dim colEntities As New MUSTER.Info.InspectionChecklistMastersCollection

                While drSet.Read
                    Dim oInspectionCheckListInfo As New MUSTER.Info.InspectionChecklistMasterInfo(drSet.Item("QUESTION_ID"), _
                                                                        AltIsDBNull(drSet.Item("CHKLST_POSITION"), 0), _
                                                                        AltIsDBNull(drSet.Item("CHKLST_ITEM_NUMBER"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("SOC"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("HEADER"), False), _
                                                                        AltIsDBNull(drSet.Item("HEADER_QUESTION_TEXT"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_TANK"), False), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_PIPE"), False), _
                                                                        AltIsDBNull(drSet.Item("APPLIES_TO_PIPETERM"), False), _
                                                                        AltIsDBNull(drSet.Item("CITATION"), 0), _
                                                                        AltIsDBNull(drSet.Item("DISCREP_TEXT"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("WHEN_VISIBLE"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("CCAT"), False), _
                                                                        AltIsDBNull(drSet.Item("DELETED"), False), _
                                                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                                        AltIsDBNull(drSet.Item("FORE_COLOR"), "BLACK"), _
                                                                        AltIsDBNull(drSet.Item("BACK_COLOR"), "WHITE"))
                    colEntities.Add(oInspectionCheckListInfo)
                End While

                Return colEntities
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetResponses(ByVal inspectionID As Int64, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params(1) As SqlParameter
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Try
                strSQL = "spGetInspectionAllResponses"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = inspectionID
                Params(1).Value = showDeleted
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Sub DBSetDateRangeOnInspection(ByVal inspectionID As Int64, ByVal facility_id As Int64, ByRef prevdate As DateTime, ByRef currentdate As DateTime)
            Dim strSQL As String
            Dim Params(1) As SqlParameter
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Try
                strSQL = "spGetInspectionDateRange"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(1).Value = IIf(inspectionID = 0, DBNull.Value, inspectionID)
                Params(0).Value = facility_id

                currentdate = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)
                prevdate = New Date(1900, 1, 1)

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not dsData Is Nothing AndAlso dsData.Tables.Count > 0 AndAlso dsData.Tables(0).Rows.Count > 0 Then
                    With dsData.Tables(0).Rows(0)
                        prevdate = .Item("PrevDate")
                        currentdate = .Item("InspectedDate")
                    End With
                End If


            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function GetDBInspctionCheckListApproval() As Boolean

            Try
                Return SqlHelper.GetFunctionalityApprovalFromDB(_strConn, "DBfastInspCheckList")
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex

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


        Public Function DBGetCLInspectionHistory(ByVal inspectionID As Int64, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim strSQL As String
            Dim Params(1) As SqlParameter
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Try
                strSQL = "spGetInspectionCLHistory"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = inspectionID
                Params(1).Value = showDeleted
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function DBGetAnnouncementLetterProcesses(Optional ByVal ownerId As Integer = 0, Optional ByVal facId As Integer = 0) As DataTable
            Dim strSQL As String
            Dim Params(3) As SqlParameter
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Try
                strSQL = "spGetGuideLineAnnouncementsByFacility"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If ownerId > 0 Then
                    Params(0).Value = ownerId
                End If

                If facId > 0 Then
                    Params(1).Value = facId
                End If
                Params(2).Value = 0
                Params(3).Value = 0


                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not dsData Is Nothing AndAlso dsData.Tables.Count > 0 Then
                    Return dsData.Tables(0)
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function DBGetAnnouncementLetterComponents(Optional ByVal ownerId As Integer = 0, Optional ByVal facId As Integer = 0) As DataTable
            Dim strSQL As String
            Dim Params(3) As SqlParameter
            Dim drSet As SqlDataReader
            Dim dsData As DataSet
            Try
                strSQL = "spGetGuideLineAnnouncementsByFacility"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If ownerId > 0 Then
                    Params(0).Value = ownerId
                End If

                If facId > 0 Then
                    Params(1).Value = facId
                End If
                Params(2).Value = 1
                Params(3).Value = 0


                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Not dsData Is Nothing AndAlso dsData.Tables.Count > 0 Then
                    Return dsData.Tables(0)
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function



        Public Sub DBPutCLInspectionHistory(ByRef id As Int64, ByVal inspectionID As Int64, ByVal staffID As Int64, ByVal insp_Date As Date, ByVal timeIn As String, ByVal timeOut As String, ByVal bolDeleted As Boolean, ByVal moduleID As Integer, ByVal UserID As String, ByRef returnVal As String, ByVal staffIDForSecurity As Integer)
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffIDForSecurity, CType(SqlHelper.EntityTypes.Inspection, Integer))) Then
                    returnVal = "You do not have rights to save Inspection."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutInspectionCLHistory"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(id <= 0, 0, id)
                Params(1).Value = inspectionID
                Params(2).Value = IIf(staffID = 0, DBNull.Value, staffID)
                If Date.Compare(insp_Date, CDate("01/01/0001")) = 0 Then
                    Params(3).Value = DBNull.Value
                Else
                    Params(3).Value = insp_Date
                End If
                Params(4).Value = IIf(timeIn = String.Empty, DBNull.Value, timeIn)
                Params(5).Value = IIf(timeOut = String.Empty, DBNull.Value, timeOut)
                Params(6).Value = bolDeleted

                Params(7).Value = UserID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                id = Params(0).Value
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
    End Class
End Namespace
