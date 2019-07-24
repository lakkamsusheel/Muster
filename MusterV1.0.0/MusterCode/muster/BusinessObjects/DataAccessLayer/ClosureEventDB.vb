'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.ClosureEventDB
'   Provides the means for marshalling Entity state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0     MNR         03/21/05    Original class definition.
'
' Function          Description
' DBGetByID(id)     Returns an EntityInfo object indicated by int arg id
'-------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class ClosureEventDB
        Private _strConn
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#Region "Constructors"
        Public Sub New()
            Dim oCnn As New ConnectionSettings
            _strConn = oCnn.cnString
            oCnn = Nothing
        End Sub
#End Region
        Public Function DBGetByID(ByVal closureID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ClosureEventInfo
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Try
                If closureID = 0 Then
                    Return New MUSTER.Info.ClosureEventInfo
                End If
                strSQL = "spGetCLOClosure"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ClosureID").Value = closureID

                Params("@FacilityID").Value = System.DBNull.Value
                Params("@Deleted").Value = IIf(showDeleted, System.DBNull.Value, False)
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.ClosureEventInfo(drSet.Item("CLOSURE_ID"), _
                                                        drSet.Item("FACILITY_ID"), _
                                                        AltIsDBNull(drSet.Item("FACILITY_SEQUENCE"), 0), _
                                                        drSet.Item("CLOSURE_TYPE_ID"), _
                                                        drSet.Item("CLOSURE_STATUS"), _
                                                        AltIsDBNull(drSet.Item("NOI_RECIEVED"), -1), _
                                                        AltIsDBNull(drSet.Item("NOI_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("OWNER_SIGN"), False), _
                                                        AltIsDBNull(drSet.Item("SCHEDULE_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("CERTIFIED_CONTRACTOR"), 0), _
                                                        AltIsDBNull(drSet.Item("FILL_MATERIAL"), 0), _
                                                        AltIsDBNull(drSet.Item("COMPANY"), 0), _
                                                        AltIsDBNull(drSet.Item("CONTACT"), 0), _
                                                        AltIsDBNull(drSet.Item("VERBAL_WAIVER"), False), _
                                                        AltIsDBNull(drSet.Item("NOI_PROCESSED"), False), _
                                                        AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                                                        drSet.Item("DELETED"), _
                                                        AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("LOCATION"), String.Empty), _
                                                        AltIsDBNull(drSet.Item("CR_CERTIFIED_CONTRACTOR"), 0), _
                                                        AltIsDBNull(drSet.Item("CR_COMPANY"), 0), _
                                                        AltIsDBNull(drSet.Item("CR_CLOSURE_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("CR_CLOSURE_DATE"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("CR_DATE_LAST_USED"), CDate("01/01/0001")), _
                                                        AltIsDBNull(drSet.Item("PROCESS_CLOSURE"), False), _
                                                        AltIsDBNull(drSet.Item("TEC_ID"), 0), _
                                                        AltIsDBNull(drSet.Item("TEC_TYPE"), 0))
                    'AltIsDBNull(drSet.Item("SAMPLES"), String.Empty), _
                    'AltIsDBNull(drSet.Item("TANK_PIPE_LIST"), String.Empty), _
                Else
                    Return New MUSTER.Info.ClosureEventInfo
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
        Public Function DBGetByFacID(Optional ByVal facID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ClosureEventCollection
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Dim Params As Collection
            Dim colCloEvent As New MUSTER.Info.ClosureEventCollection
            Try
                If facID = 0 Then
                    Return colCloEvent
                End If
                strSQL = "spGetCLOClosure"
                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@ClosureID").Value = System.DBNull.Value
                Params("@FacilityID").Value = facID
                Params("@Deleted").Value = IIf(showDeleted, System.DBNull.Value, False)
                Params("@OrderBy").Value = 1

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If drSet.HasRows Then
                    While drSet.Read
                        Dim cloEventInfo As New MUSTER.Info.ClosureEventInfo(drSet.Item("CLOSURE_ID"), _
                                                            drSet.Item("FACILITY_ID"), _
                                                            AltIsDBNull(drSet.Item("FACILITY_SEQUENCE"), 0), _
                                                            drSet.Item("CLOSURE_TYPE_ID"), _
                                                            drSet.Item("CLOSURE_STATUS"), _
                                                            AltIsDBNull(drSet.Item("NOI_RECIEVED"), -1), _
                                                            AltIsDBNull(drSet.Item("NOI_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("OWNER_SIGN"), False), _
                                                            AltIsDBNull(drSet.Item("SCHEDULE_DATE"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("CERTIFIED_CONTRACTOR"), 0), _
                                                            AltIsDBNull(drSet.Item("FILL_MATERIAL"), 0), _
                                                            AltIsDBNull(drSet.Item("COMPANY"), 0), _
                                                            AltIsDBNull(drSet.Item("CONTACT"), 0), _
                                                            AltIsDBNull(drSet.Item("VERBAL_WAIVER"), False), _
                                                            AltIsDBNull(drSet.Item("NOI_PROCESSED"), False), _
                                                            AltIsDBNull(drSet.Item("DUE_DATE"), CDate("01/01/0001")), _
                                                            drSet.Item("DELETED"), _
                                                            AltIsDBNull(drSet.Item("CREATED_BY"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("DATE_CREATED"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("LOCATION"), String.Empty), _
                                                            AltIsDBNull(drSet.Item("CR_CERTIFIED_CONTRACTOR"), 0), _
                                                            AltIsDBNull(drSet.Item("CR_COMPANY"), 0), _
                                                            AltIsDBNull(drSet.Item("CR_CLOSURE_RECEIVED_DATE"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("CR_CLOSURE_DATE"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("CR_DATE_LAST_USED"), CDate("01/01/0001")), _
                                                            AltIsDBNull(drSet.Item("PROCESS_CLOSURE"), False), _
                                                            AltIsDBNull(drSet.Item("TEC_ID"), 0), _
                                                            AltIsDBNull(drSet.Item("TEC_TYPE"), 0))
                        'AltIsDBNull(drSet.Item("SAMPLES"), String.Empty), _
                        'AltIsDBNull(drSet.Item("TANK_PIPE_LIST"), String.Empty), _
                        colCloEvent.Add(cloEventInfo)
                    End While
                End If
                Return colCloEvent
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function
        Public Function DBGetCheckList(ByVal closureID As Integer, ByVal closureType As Integer) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCLOClosureCHKLIST"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                'Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params(0).Value = closureID
                Params(1).Value = closureType
                Params(2).Value = 1
                Params(3).Value = 1

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub Put(ByRef oCloEvt As MUSTER.Info.ClosureEventInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal sendMessageToCNE As Boolean = False)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.ClosureEvent, Integer))) Then
                    returnVal = "You do not have rights to save Closure Event."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCLOClosure"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If oCloEvt.ID < 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oCloEvt.ID
                End If
                Params(1).Value = oCloEvt.FacilityID
                Params(2).Value = oCloEvt.FacilitySequence
                Params(3).Value = oCloEvt.ClosureType
                Params(4).Value = oCloEvt.ClosureStatus
                If oCloEvt.NOIReceived = -1 Then
                    Params(5).Value = System.DBNull.Value
                Else
                    Params(5).Value = oCloEvt.NOIReceived
                End If
                If oCloEvt.NOI_Rcv_Date.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(6).Value = System.DBNull.Value
                Else
                    Params(6).Value = oCloEvt.NOI_Rcv_Date
                End If
                Params(7).Value = oCloEvt.OwnerSign
                If oCloEvt.ScheduledDate.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(8).Value = System.DBNull.Value
                Else
                    Params(8).Value = oCloEvt.ScheduledDate
                End If
                Params(9).Value = oCloEvt.CertContractor
                Params(10).Value = oCloEvt.FillMaterial
                Params(11).Value = oCloEvt.Company
                Params(12).Value = oCloEvt.Contact
                Params(13).Value = oCloEvt.VerbalWaiver
                Params(14).Value = oCloEvt.NOIProcessed
                If oCloEvt.SentToTech.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(15).Value = System.DBNull.Value
                Else
                    Params(15).Value = oCloEvt.SentToTech
                End If
                If oCloEvt.NFAbyClosure.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(16).Value = System.DBNull.Value
                Else
                    Params(16).Value = oCloEvt.NFAbyClosure
                End If
                If oCloEvt.NFAbyTech.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(17).Value = System.DBNull.Value
                Else
                    Params(17).Value = oCloEvt.NFAbyTech
                End If
                If oCloEvt.DueDate.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(18).Value = System.DBNull.Value
                Else
                    Params(18).Value = oCloEvt.DueDate
                End If
                Params(19).Value = oCloEvt.CRCertContractor
                Params(20).Value = oCloEvt.CRCompany
                If oCloEvt.CRClosureReceived.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(21).Value = System.DBNull.Value
                Else
                    Params(21).Value = oCloEvt.CRClosureReceived
                End If
                If oCloEvt.CRClosureDate.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(22).Value = System.DBNull.Value
                Else
                    Params(22).Value = oCloEvt.CRClosureDate
                End If
                If oCloEvt.CRDateLastUsed.CompareTo(CDate("01/01/0001")) = 0 Then
                    Params(23).Value = System.DBNull.Value
                Else
                    Params(23).Value = oCloEvt.CRDateLastUsed
                End If
                Params(24).Value = oCloEvt.ClosureProcessed
                Params(25).Value = oCloEvt.Deleted
                ' CheckList
                If Not oCloEvt.HashTableBoolCheckList Is Nothing Then
                    If oCloEvt.HashTableBoolCheckList.Count > 0 Then
                        Dim itemID As String = String.Empty
                        Dim listOpen As String = String.Empty
                        Dim listDate As String = String.Empty
                        ' if date if 01/01/0001 pass system.dbnull.value
                        'For i As Integer = 0 To oCloEvt.BoolCheckList.Count - 1
                        '    itemID += oCloEvt.BoolCheckList.GetKey(i).ToString + "|"
                        '    listOpen += oCloEvt.BoolCheckList.GetByIndex(i).ToString + "|"
                        '    If Date.Compare(CType(oCloEvt.DateCheckList.GetByIndex(i), Date), CDate("01/01/0001")) = 0 Then
                        '        listDate += "NULL" + "|"
                        '    Else
                        '        listDate += oCloEvt.DateCheckList.GetByIndex(i).ToString + "|"
                        '    End If
                        'Next
                        For Each htEntry As DictionaryEntry In oCloEvt.HashTableBoolCheckList
                            itemID += htEntry.Key.ToString + "|"
                            listOpen += htEntry.Value.ToString + "|"
                            If Date.Compare(CType(oCloEvt.HashTableDateCheckList.Item(htEntry.Key), Date), CDate("01/01/0001")) = 0 Then
                                listDate += "NULL" + "|"
                            Else
                                listDate += CType(oCloEvt.HashTableDateCheckList.Item(htEntry.Key), Date).ToShortDateString + "|"
                            End If
                        Next
                        itemID = itemID.TrimEnd("|")
                        listOpen = listOpen.TrimEnd("|")
                        listDate = listDate.TrimEnd("|")
                        Params(26).Value = itemID
                        Params(27).Value = listOpen
                        Params(28).Value = listDate
                    Else
                        Params(26).Value = System.DBNull.Value
                        Params(27).Value = System.DBNull.Value
                        Params(28).Value = System.DBNull.Value
                    End If
                Else
                    Params(26).Value = System.DBNull.Value
                    Params(27).Value = System.DBNull.Value
                    Params(28).Value = System.DBNull.Value
                End If
                ' Tank Pipe
                'If oCloEvt.TankPipeID.Length > 0 Then
                'If Not oCloEvt.TankPipeID Is Nothing Then
                '    Params(29).Value = oCloEvt.TankPipeID
                '    Params(30).Value = oCloEvt.TankPipeEntity
                'Else
                '    Params(29).Value = System.DBNull.Value
                '    Params(30).Value = System.DBNull.Value
                'End If
                ' passing dbnull for tank pipe id, tank pipe entity and samplestable
                ' as it is handled in the ui directly
                Params(29).Value = System.DBNull.Value
                Params(30).Value = System.DBNull.Value
                Params(31).Value = DBNull.Value
                'If oCloEvt.SamplesTable.Rows.Count > 0 Then
                '    ' construct concatenated string
                '    Dim j As Integer = 0
                '    oCloEvt.Samples = String.Empty
                '    For Each dr As DataRow In oCloEvt.SamplesTable.Rows
                '        For j = 0 To oCloEvt.SamplesTable.Columns.Count - 1
                '            If dr.Item(j) Is System.DBNull.Value Then
                '                oCloEvt.Samples += "NULL|"
                '            Else
                '                oCloEvt.Samples += CType(dr.Item(j), String).ToUpper + "|"
                '            End If
                '        Next
                '        oCloEvt.Samples = oCloEvt.Samples.TrimEnd("|")
                '        oCloEvt.Samples += "~"
                '    Next
                '    oCloEvt.Samples = oCloEvt.Samples.TrimEnd("~")
                '    Params(31).Value = oCloEvt.Samples
                'Else
                '    Params(31).Value = System.DBNull.Value
                'End If
                Params(32).Value = System.DBNull.Value
                Params(33).Value = System.DBNull.Value
                Params(34).Value = System.DBNull.Value
                Params(35).Value = System.DBNull.Value

                If oCloEvt.ID <= 0 Then
                    Params(36).Value = oCloEvt.CreatedBy
                Else
                    Params(36).Value = oCloEvt.ModifiedBy
                End If
                Params(37).Value = oCloEvt.TecID
                Params(38).Value = oCloEvt.TecType
                Params(39).Value = IIf(sendMessageToCNE, 1, 0)


                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Params(0).Value <> oCloEvt.ID Then
                    oCloEvt.ID = Params(0).Value
                    oCloEvt.FacilitySequence = Params(2).Value
                    oCloEvt.ClosureStatus = Params(4).Value
                End If
                'oCloEvt.Samples = AltIsDBNull(Params(31).Value, String.Empty)
                oCloEvt.CreatedBy = AltIsDBNull(Params(32).Value, String.Empty)
                oCloEvt.CreatedOn = AltIsDBNull(Params(33).Value, CDate("01/01/0001"))
                oCloEvt.ModifiedBy = AltIsDBNull(Params(34).Value, String.Empty)
                oCloEvt.ModifiedOn = AltIsDBNull(Params(35).Value, CDate("01/01/0001"))
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetPreviousStatus(ByVal closureID As Integer, ByVal status As Integer, Optional ByVal showDeleted As Boolean = False) As Integer
            Dim strSQL As String
            Dim Params() As SqlParameter
            Dim returnValue As Integer
            Try
                strSQL = "spGetCLOPreviousStatus"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = closureID
                Params(1).Value = status
                Params(2).Value = showDeleted
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If Params(1).Value Is DBNull.Value Then
                    returnValue = status
                Else
                    returnValue = Params(1).Value
                End If
                Return returnValue
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
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetCheckListItems(ByVal closureType As Integer) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCLOClosureCHKLIST_ITEMS"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                'Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params(0).Value = closureType
                Params(1).Value = 1
                Params(2).Value = 1

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
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
        Public Function DBGetClosureSamples(Optional ByVal closureID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCLOClosureSample"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = IIf(closureID = 0, DBNull.Value, closureID)
                Params(1).Value = showDeleted
                'Params(2).Value = 1

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub DBPutClosureSamples(ByRef sampleID As Integer, _
                                    ByVal closureID As Integer, _
                                    ByVal sampleNumber As String, _
                                    ByVal analysisType As Integer, _
                                    ByVal analysisLevel As Integer, _
                                    ByVal sampleMedia As Integer, _
                                    ByVal sampleLocation As Integer, _
                                    ByVal sampleValue As Double, _
                                    ByVal sampleUnits As Integer, _
                                    ByVal sampleConstituent As String, _
                                    ByVal deleted As Boolean, _
                                    ByVal usedID As String, _
                                    ByVal moduleID As Integer, _
                                    ByVal staffID As Integer, _
                                    ByRef returnVal As String)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.ClosureEvent, Integer))) Then
                    returnVal = "You do not have rights to save Closure Event."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCLOClosureSample"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If sampleID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = sampleID
                End If
                Params(1).Value = closureID
                Params(2).Value = sampleNumber
                Params(3).Value = analysisType
                Params(4).Value = analysisLevel
                Params(5).Value = sampleMedia
                Params(6).Value = sampleLocation
                Params(7).Value = sampleValue
                Params(8).Value = sampleUnits
                Params(9).Value = sampleConstituent
                Params(10).Value = deleted
                Params(11).Value = usedID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Params(0).Value <> sampleID Then
                    sampleID = Params(0).Value
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function DBGetClosureTankPipe(ByVal facID As Integer, ByVal closureID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsData As DataSet
            Dim strSQL As String
            Dim Params() As SqlParameter
            Try
                strSQL = "spGetCLOClosureTankPipe"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                Params(0).Value = facID
                Params(1).Value = closureID
                Params(2).Value = showDeleted

                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DBGetTankPipeList(ByVal closureID As Integer, Optional ByVal showDeleted As Boolean = False) As String
            Dim drSet As SqlDataReader
            Dim strSQL As String
            Try
                strSQL = "select dbo.udfGetCloTankPipeList(" + closureID.ToString + ", " + IIf(showDeleted, "1", "0") + ") as TANK_PIPE_LIST"
                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.Text, strSQL)
                If drSet.HasRows Then
                    drSet.Read()
                    Return AltIsDBNull(drSet.Item("TANK_PIPE_LIST"), String.Empty)
                Else
                    Return String.Empty
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
        Public Sub DBPutTankPipe(ByRef cloTankPipeID As Integer, _
                                    ByVal tankPipeID As Integer, _
                                    ByVal tankPipeEntity As Integer, _
                                    ByVal closureID As Integer, _
                                    ByVal deleted As Boolean, _
                                    ByVal moduleID As Integer, _
                                    ByVal staffID As Integer, _
                                    ByRef returnVal As String, _
                                    ByVal UserID As String, _
                                    Optional ByVal analysisType As Integer = -1, _
                                    Optional ByVal analysisLevel As Integer = -1, _
                                    Optional ByVal sampleMedia As Integer = -1, _
                                    Optional ByVal sampleResultsID As Integer = -1)
            Dim Params() As SqlParameter
            Dim strSQL As String
            Try


                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.ClosureEvent, Integer))) Then
                    returnVal = "You do not have rights to save Closure Event."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                strSQL = "spPutCLOClosureTankPipe"
                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, strSQL)
                If cloTankPipeID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = cloTankPipeID
                End If
                Params(1).Value = tankPipeID
                Params(2).Value = tankPipeEntity
                Params(3).Value = closureID
                Params(4).Value = IIf(analysisType = -1, DBNull.Value, analysisType)
                Params(5).Value = IIf(analysisLevel = -1, DBNull.Value, analysisLevel)
                Params(6).Value = IIf(sampleMedia = -1, DBNull.Value, sampleMedia)
                Params(7).Value = IIf(sampleResultsID = -1, DBNull.Value, sampleResultsID)
                Params(8).Value = deleted
                Params(9).Value = UserID

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, strSQL, Params)

                If Params(0).Value <> cloTankPipeID Then
                    cloTankPipeID = Params(0).Value
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
    End Class
End Namespace
