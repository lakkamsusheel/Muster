'-------------------------------------------------------------------------------
' MUSTER.Info.Inspection
'   Provides the container to persist MUSTER Inspection state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/07/05    Original class definition
'
' Function          Description
'
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pInspection
#Region "Public Events"
        Public Event evtInspectionErr(ByVal MsgStr As String)
        Public Event evtInspectionChanged(ByVal bolValue As Boolean)
        Public Event evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String)
#End Region
#Region "Private member variables"
        Private WithEvents oInspectionInfo As MUSTER.Info.InspectionInfo
        Private colInspections As MUSTER.Info.InspectionsCollection
        Private oInspectionDB As New MUSTER.DataAccess.InspectionDB
        Private WithEvents oCheckListMaster As MUSTER.BusinessLogic.pInspectionChecklistMaster
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64
#End Region
#Region "Constructors"
        Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)

            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            oInspectionInfo = New MUSTER.Info.InspectionInfo
            colInspections = New MUSTER.Info.InspectionsCollection
            oInspectionDB = New MUSTER.DataAccess.InspectionDB(strDBConn, MusterXCEP)
            oCheckListMaster = New MUSTER.BusinessLogic.pInspectionChecklistMaster(strDBConn, MusterXCEP)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Int64
            Get
                Return oInspectionInfo.ID
            End Get
        End Property
        Public Property OwnerID() As Int64
            Get
                Return oInspectionInfo.OwnerID
            End Get
            Set(ByVal Value As Int64)
                oInspectionInfo.OwnerID = Value
            End Set
        End Property
        Public Property FacilityID() As Int64
            Get
                Return oInspectionInfo.FacilityID
            End Get
            Set(ByVal Value As Int64)
                oInspectionInfo.FacilityID = Value
            End Set
        End Property
        Public Property InspectionType() As Int64
            Get
                Return oInspectionInfo.InspectionType
            End Get
            Set(ByVal Value As Int64)
                oInspectionInfo.InspectionType = Value
            End Set
        End Property
        Public Property AdminComments() As String
            Get
                Return oInspectionInfo.AdminComments
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.AdminComments = Value
            End Set
        End Property
        Public Property ScheduledDate() As Date
            Get
                Return oInspectionInfo.ScheduledDate
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.ScheduledDate = Value
            End Set
        End Property
        Public Property ScheduledTime() As String
            Get
                Return oInspectionInfo.ScheduledTime
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.ScheduledTime = Value
            End Set
        End Property
        Public Property CheckListGenDate() As Date
            Get
                Return oInspectionInfo.CheckListGenDate
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.CheckListGenDate = Value
            End Set
        End Property
        Public Property SubmittedDate() As Date
            Get
                Return oInspectionInfo.SubmittedDate
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.SubmittedDate = Value
            End Set
        End Property
        Public Property StaffID() As Int64
            Get
                Return oInspectionInfo.StaffID
            End Get
            Set(ByVal Value As Int64)
                oInspectionInfo.StaffID = Value
            End Set
        End Property
        Public Property AssignedDate() As Date
            Get
                Return oInspectionInfo.AssignedDate
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.AssignedDate = Value
            End Set
        End Property
        Public Property ScheduledBy() As String
            Get
                Return oInspectionInfo.ScheduledBy
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.ScheduledBy = Value
            End Set
        End Property
        Public Property LetterGenerated() As Boolean
            Get
                Return oInspectionInfo.LetterGenerated
            End Get
            Set(ByVal Value As Boolean)
                oInspectionInfo.LetterGenerated = Value
            End Set
        End Property
        Public Property DateLetterGenerated() As Date
            Get
                Return oInspectionInfo.DateLetterGenerated
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.DateLetterGenerated = Value
            End Set
        End Property
        Public Property DatePlanned() As Date
            Get
                Return oInspectionInfo.DatePlanned
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.DatePlanned = Value
            End Set
        End Property
        Public Property ConductInspectionOn() As Date
            Get
                Return oInspectionInfo.ConductInspectionOn
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.ConductInspectionOn = Value
            End Set
        End Property
        Public Property RescheduledDate() As Date
            Get
                Return oInspectionInfo.RescheduledDate
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.RescheduledDate = Value
            End Set
        End Property
        Public Property RescheduledTime() As String
            Get
                Return oInspectionInfo.RescheduledTime
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.RescheduledTime = Value
            End Set
        End Property
        Public Property InspectorComments() As String
            Get
                Return oInspectionInfo.InspectorComments
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.InspectorComments = Value
            End Set
        End Property
        Public Property Completed() As Date
            Get
                Return oInspectionInfo.Completed
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.Completed = Value
            End Set
        End Property
        Public Property OwnersRep() As String
            Get
                Return oInspectionInfo.OwnersRep
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.OwnersRep = Value
            End Set
        End Property
        Public Property CAPDatesEntered() As Boolean
            Get
                Return oInspectionInfo.CAPDatesEntered
            End Get
            Set(ByVal Value As Boolean)
                oInspectionInfo.CAPDatesEntered = Value
            End Set
        End Property
        Public Property InspectionAccepted() As Boolean
            Get
                Return oInspectionInfo.InspectionAccepted
            End Get
            Set(ByVal Value As Boolean)
                oInspectionInfo.InspectionAccepted = Value
            End Set
        End Property
        Public Property CAEViewed() As Date
            Get
                Return oInspectionInfo.CAEViewed
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.CAEViewed = Value
            End Set
        End Property
        Public Property ChecklistFirstSaved() As Date
            Get
                Return oInspectionInfo.ChecklistFirstSaved
            End Get
            Set(ByVal Value As Date)
                oInspectionInfo.ChecklistFirstSaved = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oInspectionInfo.Deleted
            End Get
            Set(ByVal value As Boolean)
                oInspectionInfo.Deleted = value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oInspectionInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oInspectionInfo.IsDirty = value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oInspectionInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oInspectionInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oInspectionInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oInspectionInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oInspectionInfo.ModifiedOn
            End Get
        End Property
        Public Property CheckListMaster() As MUSTER.BusinessLogic.pInspectionChecklistMaster
            Get
                Return oCheckListMaster
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pInspectionChecklistMaster)
                oCheckListMaster = Value
            End Set
        End Property
        Public Property InspectionInfo() As MUSTER.Info.InspectionInfo
            Get
                Return oInspectionInfo
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionInfo)
                oInspectionInfo = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"

        Public Function HasfacilityBeenInspected(ByVal facility_id As Integer) As Date
            Return Me.oInspectionDB.HasfacilitybeenInspectedBeforeDB(facility_id)
        End Function

        Public Function Retrieve(Optional ByVal id As Integer = 0, Optional ByVal staffID As Int64 = 0, Optional ByVal facID As Int64 = 0, Optional ByVal ownerID As Int64 = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal bolLoading As Boolean = False) As MUSTER.Info.InspectionInfo
            Dim bolDataAged As Boolean = False
            Dim bolFound As Boolean = False

            Try
                If Not (oInspectionInfo.Deleted Or oInspectionInfo.ID = 0 Or Not oInspectionInfo.IsDirty Or bolLoading) Then
                    Me.ValidateData()
                End If

                Dim oInspectInfoLocal As MUSTER.info.InspectionInfo

                If id = 0 And staffID = 0 And facID = 0 And ownerID = 0 Then
                    bolFound = False
                ElseIf id <> 0 Then
                    ' check in collection
                    oInspectionInfo = colInspections.Item(id)
                    If Not oInspectionInfo Is Nothing Then
                        bolFound = True
                    End If
                ElseIf staffID <> 0 Then
                    For Each oInspectInfoLocal In colInspections.Values
                        If oInspectInfoLocal.StaffID = staffID Then
                            oInspectionInfo = oInspectInfoLocal
                            bolFound = True
                            Exit For
                        End If
                    Next
                ElseIf facID <> 0 Then
                    For Each oInspectInfoLocal In colInspections.Values
                        If oInspectInfoLocal.FacilityID = facID Then
                            oInspectionInfo = oInspectInfoLocal
                            bolFound = True
                            Exit For
                        End If
                    Next
                ElseIf ownerID <> 0 Then
                    For Each oInspectInfoLocal In colInspections.Values
                        If oInspectInfoLocal.FacilityID = facID And _
                        oInspectInfoLocal.OwnerID = ownerID Then
                            oInspectionInfo = oInspectInfoLocal
                            bolFound = True
                            Exit For
                        End If
                    Next
                End If

                If Not bolFound Then
                    Add(id, staffID, facID, ownerID, showDeleted)
                End If

                ' Check for Aged Data here.
                If Not (oInspectionInfo Is Nothing) Then
                    If oInspectionInfo.IsAgedData = True And oInspectionInfo.IsDirty = False Then
                        bolDataAged = True
                        colInspections.Remove(oInspectionInfo)
                    End If
                End If

                If oInspectionInfo Is Nothing Or bolDataAged Then
                    Add(id, staffID, facID, ownerID, showDeleted)
                End If

                oCheckListMaster.InspectionInfo = oInspectionInfo
                Return oInspectionInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal strUser As String = "", Optional ByVal OverrideRights As Boolean = False) As Boolean
            Dim oldID As Integer
            Dim proceedSaving As Boolean = True
            Try
                'If Me.SubmittedDate >= CDate("1/1/1960") Then
                'proceedSaving = False

                'proceedSaving = (MsgBox("This checklist has been submitted and changes will directly be made to registration. Do you still want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

                'End If

                If proceedSaving Then

                    If Not bolValidated And Not oInspectionInfo.Deleted And Not bolDelete Then
                        If Not Me.ValidateData() Then
                            Return False
                        End If
                    End If
                    If Not (oInspectionInfo.ID < 0 And oInspectionInfo.Deleted) Then
                        oldID = oInspectionInfo.ID
                        oInspectionDB.Put(oInspectionInfo, moduleID, staffID, returnVal, OverrideRights)
                        If Not returnVal = String.Empty Then
                            Exit Function
                        End If

                        If Not bolValidated Then
                            If oldID < 0 Then
                                colInspections.ChangeKey(oldID, oInspectionInfo.ID)
                            End If
                            Dim UserID As String = String.Empty
                            If oldID <= 0 Then
                                UserID = oInspectionInfo.CreatedBy
                            Else
                                UserID = oInspectionInfo.ModifiedBy
                            End If
                            oCheckListMaster.Flush(moduleID, staffID, UserID, returnVal, strUser, Me.SubmittedDate >= CDate("1/1/1960"))
                        End If
                        oInspectionInfo.Archive()
                        oInspectionInfo.IsDirty = False
                    End If
                    If Not bolValidated And bolDelete Then
                        If oInspectionInfo.Deleted Then
                            ' check if other inspections are present else load new instance
                            Dim strNext As String = Me.GetNext()
                            Dim strPrev As String = Me.GetPrevious()
                            If strNext = oInspectionInfo.ID Then
                                If strPrev = oInspectionInfo.ID Then
                                    RaiseEvent evtInspectionErr("Inspection " + oInspectionInfo.ID.ToString + " deleted")
                                    colInspections.Remove(oInspectionInfo)
                                    If bolDelete Then
                                        oInspectionInfo = New MUSTER.Info.InspectionInfo
                                    Else
                                        oInspectionInfo = Me.Retrieve(0)
                                    End If
                                Else
                                    RaiseEvent evtInspectionErr("Inspection " + oInspectionInfo.ID.ToString + " deleted")
                                    colInspections.Remove(oInspectionInfo)
                                    oInspectionInfo = Me.Retrieve(strPrev)
                                End If
                            Else
                                RaiseEvent evtInspectionErr("Inspection " + oInspectionInfo.ID.ToString + " deleted")
                                colInspections.Remove(oInspectionInfo)
                                oInspectionInfo = Me.Retrieve(strNext)
                            End If
                        End If
                    End If
                    RaiseEvent evtInspectionChanged(oInspectionInfo.IsDirty)
                    Return True
                End If

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return False
            End Try
        End Function
        Public Sub RetrieveCheckListInfo(ByVal id As Integer, ByVal facID As Int64, ByVal ownerID As Int64, Optional ByVal [readOnly] As Boolean = False, Optional ByVal isOwnerDesignatedOperator As Boolean = True, Optional ByVal path As String = "")
            Try
                If oInspectionInfo.ID <> id Or oInspectionInfo.FacilityID <> facID Or oInspectionInfo.OwnerID <> ownerID Then
                    Retrieve(id, , facID, ownerID)
                End If

                oCheckListMaster.Retrieve(oInspectionInfo, oInspectionInfo.ID, facID, ownerID, [readOnly], isOwnerDesignatedOperator, path)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub RetrieveOwnerFacTanksPipes()
            oCheckListMaster.RetrieveOwnerFacTanksPipes(oInspectionInfo)
        End Sub
        Public Function RetrieveInspectionHistory(ByVal facID As Int64, Optional ByVal showDeleted As Boolean = False) As DataSet
            Try
                Return oInspectionDB.DBGetInspectionHistory(facID, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return New DataSet
            End Try
        End Function
        Public Function ValidateData() As Boolean
            Try
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return False
            End Try
        End Function
        Public Sub SaveCAPTankPipeDataBeforeAfterInspection(ByVal inspectionID As Integer, ByVal isBeforeInspectionData As Boolean, Optional ByVal bolDeleteData As Boolean = False)
            Try
                oInspectionDB.PutCAPTankPipeDataBeforeAfterInspection(inspectionID, isBeforeInspectionData, bolDeleteData)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub UpdateTanksPipesFromMirrorTable(ByVal facilityID As Integer)
            Try
                oInspectionDB.PutTankPipeDataFromMirrortable(facilityID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub UpdateInspectionSOC(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal inspectionID As Integer, ByVal strUser As String)
            Try
                oInspectionDB.UpdateInspectionSOC(moduleID, staffID, returnVal, inspectionID, strUser)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub PutInspectionArchive(ByVal facilityID As Integer, ByVal inspectionID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolSaveTankPipeTermOnly As Boolean = False)
            Try
                oInspectionDB.PutInspectionArchive(facilityID, inspectionID, moduleID, staffID, returnVal, bolSaveTankPipeTermOnly)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal id As Int64, Optional ByVal staffID As Int64 = 0, Optional ByVal facID As Int64 = 0, Optional ByVal ownerID As Int64 = 0, Optional ByVal showDeleted As Boolean = False)
            Try
                Dim colInspectionsLocal As MUSTER.Info.InspectionsCollection = oInspectionDB.DBGetByID(id, staffID, facID, ownerID, showDeleted)
                If colInspectionsLocal.Count = 0 Then
                    oInspectionInfo = New MUSTER.Info.InspectionInfo
                    oInspectionInfo.ID = nID
                    nID -= 1
                    colInspections.Add(oInspectionInfo)
                Else
                    For Each oInspectionInfoLocal As MUSTER.Info.InspectionInfo In colInspectionsLocal.Values
                        oInspectionInfo = oInspectionInfoLocal
                        If oInspectionInfo.ID = 0 Then
                            oInspectionInfo.ID = nID
                            nID -= 1
                        End If
                        colInspections.Add(oInspectionInfo)
                    Next
                End If
                oCheckListMaster.InspectionInfo = oInspectionInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oInspectInfo As MUSTER.Info.InspectionInfo)
            Try
                oInspectionInfo = oInspectInfo
                If oInspectionInfo.ID = 0 Then
                    oInspectionInfo.ID = nID
                    nID -= 1
                End If
                colInspections.Add(oInspectionInfo)
                oCheckListMaster.InspectionInfo = oInspectionInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal id As Int64)
            Try
                If colInspections.Contains(id) Then
                    colInspections.Remove(id)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oInspectInfo As MUSTER.Info.InspectionInfo)
            Try
                If colInspections.Contains(oInspectInfo) Then
                    colInspections.Remove(oInspectInfo)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                Dim IDs As New Collection
                Dim index As Integer
                Dim xInspectInfo As MUSTER.Info.InspectionInfo
                For Each xInspectInfo In colInspections.Values
                    If xInspectInfo.IsDirty Then
                        oInspectionInfo = xInspectInfo
                        If Me.ValidateData() Then
                            If oInspectionInfo.ID < 0 And _
                                Not oInspectionInfo.Deleted Then
                                IDs.Add(oInspectionInfo.ID)
                            End If
                            Me.Save(moduleID, staffID, returnVal, True)
                        Else : Exit For
                        End If
                    End If
                Next
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xInspectInfo = colInspections.Item(colKey)
                        colInspections.ChangeKey(colKey, xInspectInfo.ID)
                    Next
                End If
                RaiseEvent evtInspectionChanged(oInspectionInfo.IsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colInspections.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colInspections.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf colInspections.Count <> 0 Then
                Return colInspections.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oInspectionInfo = New MUSTER.Info.InspectionInfo
        End Sub
        Public Sub Reset()
            oInspectionInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operation"
        Public Function EntityTable() As DataTable
            Dim oInspectionInfoLocal As New MUSTER.Info.InspectionInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("Inspection ID")
                tbEntityTable.Columns.Add("Owner ID")
                tbEntityTable.Columns.Add("Facility ID")
                tbEntityTable.Columns.Add("Inspection Type")
                tbEntityTable.Columns.Add("Admin Comments")
                tbEntityTable.Columns.Add("Schedule Date")
                tbEntityTable.Columns.Add("Schedule Time")
                tbEntityTable.Columns.Add("CheckList Generation Date")
                tbEntityTable.Columns.Add("Submitted Date")
                tbEntityTable.Columns.Add("Staff ID")
                tbEntityTable.Columns.Add("Scheduled By")
                tbEntityTable.Columns.Add("Letter Generated")
                tbEntityTable.Columns.Add("Date Generated")
                tbEntityTable.Columns.Add("Date Planned")
                tbEntityTable.Columns.Add("Conduct Inspection On")
                tbEntityTable.Columns.Add("ReScheduled")
                tbEntityTable.Columns.Add("Inspector Comments")
                tbEntityTable.Columns.Add("Completed")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Created On")
                tbEntityTable.Columns.Add("Modified By")
                tbEntityTable.Columns.Add("Modified On")
                tbEntityTable.Columns.Add("Deleted")

                For Each oInspectionInfoLocal In colInspections.Values
                    dr = tbEntityTable.NewRow()
                    dr("Inspection ID") = oInspectionInfoLocal.ID
                    dr("Owner ID") = oInspectionInfoLocal.OwnerID
                    dr("Facility ID") = oInspectionInfoLocal.FacilityID
                    dr("Inspection Type") = oInspectionInfoLocal.InspectionType
                    dr("Admin Comments") = oInspectionInfoLocal.AdminComments
                    dr("Schedule Date") = oInspectionInfoLocal.ScheduledDate
                    dr("Schedule Time") = oInspectionInfoLocal.ScheduledTime
                    dr("CheckList Generation Date") = oInspectionInfoLocal.CheckListGenDate
                    dr("Submitted Date") = oInspectionInfoLocal.SubmittedDate
                    dr("Staff ID") = oInspectionInfoLocal.StaffID
                    dr("Scheduled By") = oInspectionInfoLocal.ScheduledBy
                    dr("Letter Generated") = oInspectionInfoLocal.LetterGenerated
                    dr("Date Letter Generated") = oInspectionInfoLocal.DateLetterGenerated
                    dr("Date Planned") = oInspectionInfoLocal.DatePlanned
                    dr("Conduct Inspection On") = oInspectionInfoLocal.ConductInspectionOn
                    dr("ReScheduled Date") = oInspectionInfoLocal.RescheduledDate
                    dr("ReScheduled Time") = oInspectionInfoLocal.RescheduledTime
                    dr("Inspector Comments") = oInspectionInfoLocal.InspectorComments
                    dr("Completed") = oInspectionInfoLocal.Completed
                    dr("Created By") = oInspectionInfoLocal.CreatedBy
                    dr("Created On") = oInspectionInfoLocal.CreatedOn
                    dr("Modified By") = oInspectionInfoLocal.ModifiedBy
                    dr("Modified On") = oInspectionInfoLocal.ModifiedOn
                    dr("Deleted") = oInspectionInfoLocal.Deleted
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetTargetFacilities(ByVal nStaffID As Integer, Optional ByVal ownerID As Integer = 0, Optional ByVal bolAllTargetFacs As Boolean = False, Optional ByVal reLoadInspectionID As Integer = 0, Optional ByVal facilityID As Integer = 0) As DataSet
            Dim dsTargetFacilities As New DataSet
            Try
                dsTargetFacilities = oInspectionDB.DBGetTargetFacilities(nStaffID, ownerID, bolAllTargetFacs, reLoadInspectionID, facilityID)
                Return dsTargetFacilities
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetAssignedFacilities(ByVal nStaffID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsAssignedFacilities As New DataSet
            Try
                dsAssignedFacilities = oInspectionDB.DBGetAssignedFacilities(nStaffID, showDeleted)
                Return dsAssignedFacilities
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetInspectors() As DataSet
            Dim strSQL As String
            Try
                'strSQL = "SELECT DISTINCT STAFF_ID, [USER_NAME] FROM tblSYS_UST_STAFF_MASTER " + _
                '"WHERE DEFAULT_MODULE = '615' AND ACTIVE = 0"
                strSQL = "SELECT * FROM vCAEInspectors ORDER BY [USER_NAME]"
                Return oInspectionDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetInspectionTimes() As DataSet
            Dim strSQL As String
            Try
                strSQL = "select * from tblsys_property_master where property_type_id = 151 order by property_position"
                Return oInspectionDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetInspectionTypes() As DataSet
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM vCAEInspectionTypes ORDER BY PROPERTY_NAME"
                Return oInspectionDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetFacOwner(ByVal facID As Integer) As DataSet
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM v_OWNER_NAME WHERE O_ID = (SELECT OWNER_ID FROM TBLREG_FACILITY WHERE FACILITY_ID = " + facID.ToString + ")"
                Return oInspectionDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetCAEFacilityAssignedTo(ByVal facID As Integer) As Integer
            Try
                Return oInspectionDB.DBGetCAEFacilityAssignedTo(facID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "Event Handlers"
        Private Sub TankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String) Handles oCheckListMaster.evtTankValidationErr
            RaiseEvent evtTankValidationErr(tnkID, strMessage)
        End Sub
        Private Sub InspectionChecklistMasterChanged(ByVal bolValue As Boolean) Handles oCheckListMaster.evtInspectionChecklistMasterChanged
            RaiseEvent evtInspectionChanged(bolValue Or oInspectionInfo.IsDirty)
        End Sub
        Private Sub InspectionChanged(ByVal bolValue As Boolean) Handles oInspectionInfo.evtInspectionInfoChanged
            RaiseEvent evtInspectionChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
