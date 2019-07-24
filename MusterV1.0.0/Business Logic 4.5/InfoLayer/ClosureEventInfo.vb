'-------------------------------------------------------------------------------
' MUSTER.Info.ClosureEventInfo
'   Provides the container to persist MUSTER ClosureEvent state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC       03/16/05    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class ClosureEventInfo
#Region "Public Events"
        'Public Delegate Sub ClosureEventInfoChangedEventHandler()
        'Public Event ClosureInfoChanged As ClosureEventInfoChangedEventHandler
        Public Event evtClosureEventInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private nClosureID As Integer
        Private nFacilityID As Integer
        Private nFacilitySequence As Integer
        Private nClosureType As Integer
        Private nClosureStatus As Integer
        Private nNOIReceived As Integer
        ' NOI
        Private dtNOI_Rcv_Date As Date
        Private bolOwnerSign As Boolean
        Private dtScheduledDate As Date
        Private nCertContractor As Integer
        Private nFillMaterial As Integer
        Private nCompany As Integer
        Private nContact As Integer
        Private bolVerbalWaiver As Boolean
        Private bolNOIProcessed As Boolean
        Private dtDueDate As Date
        Private strLocation As String
        'Private bolArrCheckList As SortedList
        'Private dtArrCheckList As SortedList
        Private htBoolCheckList As Hashtable
        Private htDateCheckList As Hashtable
        'private a as DictionaryBase
        Private strTankPipeID As String
        Private strTankPipeEntity As String
        'Private strSamples As String
        Private dtableSamples As DataTable
        Private dtSentToTech As Date
        Private dtNFAByClosure As Date
        Private dtNFAByTech As Date
        ' Closure Report
        Private nCRCertContractor As Integer
        Private nCRCompany As Integer
        Private dtCRClosureReceived As Date
        Private dtCRClosureDate As Date
        Private dtCRDateLastUsed As Date
        Private bolClosureProcessed As Boolean
        Private nTecID As Integer
        Private nTecType As Integer

        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private onClosureID As Integer
        Private onFacilityID As Integer
        Private onFacilitySequence As Integer
        Private onClosureType As Integer
        Private onClosureStatus As Integer
        Private onNOIReceived As Integer
        ' NOI
        Private odtNOI_Rcv_Date As Date
        Private obolOwnerSign As Boolean
        Private odtScheduledDate As Date
        Private onCertContractor As Integer
        Private onFillMaterial As Integer
        Private onCompany As Integer
        Private onContact As Integer
        Private obolVerbalWaiver As Boolean
        Private obolNOIProcessed As Boolean
        Private odtDueDate As Date
        Private ostrLocation As String
        'Private obolArrCheckList As SortedList
        'Private odtArrCheckList As SortedList
        Private ohtBoolCheckList As Hashtable
        Private ohtDateCheckList As Hashtable
        'Private ostrTankPipeID As String
        'Private ostrTankPipeEntity As String
        'Private ostrSamples As String
        Private odtableSamples As DataTable
        Private odtSentToTech As Date
        Private odtNFAByClosure As Date
        Private odtNFAByTech As Date
        ' Closure Report
        Private onCRCertContractor As Integer
        Private onCRCompany As Integer
        Private odtCRClosureReceived As Date
        Private odtCRClosureDate As Date
        Private odtCRDateLastUsed As Date
        Private obolClosureProcessed As Boolean
        Private onTecID As Integer
        Private onTecType As Integer

        ' Common
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private bolIsDirty As Boolean
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private colComments As MUSTER.Info.CommentsCollection
#End Region
#Region "Constructors"
        Public Sub New()
            MyBase.New()
            Init()
            InitializeArray()
            InitializeSamplesTable()
            dtDataAge = Now()
            colComments = New MUSTER.Info.CommentsCollection
        End Sub
        Public Sub New(ByVal ClosureID As Integer, _
                        ByVal FacilityID As Integer, _
                        ByVal FacilitySequence As Integer, _
                        ByVal ClosureType As Integer, _
                        ByVal ClosureStatus As Integer, _
                        ByVal NOIReceived As Integer, _
                        ByVal NOI_Rcv_Date As Date, _
                        ByVal OwnerSign As Boolean, _
                        ByVal ScheduledDate As Date, _
                        ByVal CertContractor As Integer, _
                        ByVal FillMaterial As Integer, _
                        ByVal Company As Integer, _
                        ByVal Contact As Integer, _
                        ByVal VerbalWaiver As Boolean, _
                        ByVal NOIProcessed As Boolean, _
                        ByVal DueDate As Date, _
                        ByVal Deleted As Boolean, _
                        ByVal CreatedBy As String, _
                        ByVal CreatedOn As Date, _
                        ByVal ModifiedBy As String, _
                        ByVal ModifiedOn As Date, _
                        ByVal Location As String, _
                        ByVal CRCertContractor As Integer, _
                        ByVal CRCompany As Integer, _
                        ByVal CRClosureReceived As Date, _
                        ByVal CRClosureDate As Date, _
                        ByVal CRDateLastUsed As Date, _
                        ByVal ClosureProcessed As Boolean, _
                        ByVal TecID As Integer, _
                        ByVal TecType As Integer)
            'ByVal Samples As String, _
            'ByVal TankPipeList As String, _
            onClosureID = ClosureID
            onFacilityID = FacilityID
            onFacilitySequence = FacilitySequence
            onClosureType = ClosureType
            onClosureStatus = ClosureStatus
            onNOIReceived = NOIReceived
            odtNOI_Rcv_Date = NOI_Rcv_Date
            obolOwnerSign = OwnerSign
            odtScheduledDate = ScheduledDate
            onCertContractor = CertContractor
            onFillMaterial = FillMaterial
            onCompany = Company
            onContact = Contact
            obolVerbalWaiver = VerbalWaiver
            obolNOIProcessed = NOIProcessed
            odtDueDate = DueDate
            ostrLocation = Location
            'If TankPipeList.Length > 0 Then
            '    Dim tplist() As String = TankPipeList.Split("~")
            '    ostrTankPipeID = tplist(0)
            '    ostrTankPipeEntity = tplist(1)
            'End If
            'ostrSamples = Samples
            onCRCertContractor = CRCertContractor
            onCRCompany = CRCompany
            odtCRClosureReceived = CRClosureReceived
            odtCRClosureDate = CRClosureDate
            odtCRDateLastUsed = CRDateLastUsed
            obolClosureProcessed = ClosureProcessed
            onTecID = TecID
            onTecType = TecType
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            dtDataAge = Now()
            colComments = New MUSTER.Info.CommentsCollection
            InitializeArray()
            InitializeSamplesTable()
            Me.Reset()
        End Sub
        Sub New(ByVal drCloEvt As DataRow)
            Try
                onClosureID = drCloEvt.Item("CLOSURE_ID")
                onFacilityID = drCloEvt.Item("FACILITY_ID")
                onFacilitySequence = IIf(drCloEvt.Item("FACILITY_SEQUENCE") Is DBNull.Value, 0, drCloEvt.Item("FACILITY_SEQUENCE"))
                onClosureType = drCloEvt.Item("CLOSURE_TYPE")
                onClosureStatus = drCloEvt.Item("CLOSURE_STATUS")
                onNOIReceived = IIf(drCloEvt.Item("NOI_RECEIVED") Is DBNull.Value, -1, drCloEvt.Item("NOI_RECEIVED"))
                odtNOI_Rcv_Date = IIf(drCloEvt.Item("NOI_RECEIVED_DATE") Is DBNull.Value, CDate("01/01/0001"), drCloEvt.Item("NOI_RECEIVED_DATE"))
                odtNOI_Rcv_Date = odtNOI_Rcv_Date.Date
                obolOwnerSign = IIf(drCloEvt.Item("OWNER_SIGN") Is DBNull.Value, False, drCloEvt.Item("OWNER_SIGN"))
                odtScheduledDate = IIf(drCloEvt.Item("SCHEDULED_DATE") Is DBNull.Value, CDate("01/01/0001"), drCloEvt.Item("SCHEDULED_DATE"))
                odtScheduledDate = odtScheduledDate.Date
                onCertContractor = IIf(drCloEvt.Item("CERTIFIED_CONTRACTOR") Is DBNull.Value, 0, drCloEvt.Item("CERTIFIED_CONTRACTOR"))
                onFillMaterial = IIf(drCloEvt.Item("FILL_MATERIAL") Is DBNull.Value, 0, drCloEvt.Item("FILL_MATERIAL"))
                onCompany = IIf(drCloEvt.Item("COMPANY") Is DBNull.Value, 0, drCloEvt.Item("COMPANY"))
                onContact = IIf(drCloEvt.Item("CONTACT") Is DBNull.Value, 0, drCloEvt.Item("CONTACT"))
                obolVerbalWaiver = IIf(drCloEvt.Item("VERBAL_WAIVER") Is DBNull.Value, False, drCloEvt.Item("VERBAL_WAIVER"))
                obolNOIProcessed = drCloEvt.Item("NOI_PROCESSED")
                odtDueDate = IIf(drCloEvt.Item("DUE_DATE") Is DBNull.Value, CDate("01/01/0001"), drCloEvt.Item("DUE_DATE"))
                odtDueDate = odtDueDate.Date
                ostrLocation = IIf(drCloEvt.Item("LOCATION") Is System.DBNull.Value, String.Empty, drCloEvt.Item("LOCATION"))
                'Dim str As String = IIf(drCloEvt.Item("TANK_PIPE_LIST") Is System.DBNull.Value, String.Empty, drCloEvt.Item("TANK_PIPE_LIST"))
                'If str.Length > 0 Then
                '    Dim tplist() As String = str.Split("~")
                '    ostrTankPipeID = tplist(0)
                '    ostrTankPipeEntity = tplist(1)
                'Else
                '    ostrTankPipeID = String.Empty
                '    ostrTankPipeEntity = String.Empty
                'End If
                'ostrSamples = IIf(drCloEvt.Item("SAMPLES") Is System.DBNull.Value, String.Empty, drCloEvt.Item("SAMPLES"))
                onCRCertContractor = IIf(drCloEvt.Item("CERTIFIED_CONTRACTOR_ID") Is DBNull.Value, 0, drCloEvt.Item("CERTIFIED_CONTRACTOR_ID"))
                onCRCompany = IIf(drCloEvt.Item("COMPANY_ID") Is DBNull.Value, 0, drCloEvt.Item("COMPANY_ID"))
                odtCRClosureReceived = IIf(drCloEvt.Item("CLOSURE_RECEIVED_DATE") Is DBNull.Value, CDate("01/01/0001"), drCloEvt.Item("CLOSURE_RECEIVED_DATE"))
                odtCRClosureReceived = odtCRClosureReceived.Date
                odtCRClosureDate = IIf(drCloEvt.Item("CLOSURE_DATE") Is DBNull.Value, CDate("01/01/0001"), odtCRClosureDate.Date)
                odtCRClosureDate = odtCRClosureDate.Date
                odtCRDateLastUsed = IIf(drCloEvt.Item("DATE_LAST_USED") Is DBNull.Value, CDate("01/01/0001"), drCloEvt.Item("DATE_LAST_USED"))
                odtCRDateLastUsed = odtCRDateLastUsed.Date
                obolClosureProcessed = IIf(drCloEvt.Item("PROCESS_CLOSURE") Is DBNull.Value, False, drCloEvt.Item("PROCESS_CLOSURE"))
                obolDeleted = drCloEvt.Item("DELETED")
                ostrCreatedBy = IIf(drCloEvt.Item("CREATED_BY") Is DBNull.Value, String.Empty, drCloEvt.Item("CREATED_BY"))
                odtCreatedOn = IIf(drCloEvt.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drCloEvt.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drCloEvt.Item("LAST_EDITED_BY") Is DBNull.Value, CDate("01/01/0001"), drCloEvt.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drCloEvt.Item("DATE_LAST_EDITED") Is DBNull.Value, String.Empty, drCloEvt.Item("DATE_LAST_EDITED"))
                onTecID = IIf(drCloEvt.Item("TEC_ID") Is DBNull.Value, 0, drCloEvt.Item("TEC_ID"))
                onTecType = IIf(drCloEvt.Item("TEC_TYPE") Is DBNull.Value, 0, drCloEvt.Item("TEC_TYPE"))
                dtDataAge = Now()
                colComments = New MUSTER.Info.CommentsCollection
                InitializeArray()
                InitializeSamplesTable()
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Methods"
        Public Sub Archive()
            onClosureID = nClosureID
            onFacilityID = nFacilityID
            onFacilitySequence = nFacilitySequence
            onClosureType = nClosureType
            onClosureStatus = nClosureStatus
            onNOIReceived = nNOIReceived
            odtNOI_Rcv_Date = dtNOI_Rcv_Date
            obolOwnerSign = bolOwnerSign
            odtScheduledDate = dtScheduledDate
            onCertContractor = nCertContractor
            onFillMaterial = nFillMaterial
            onCompany = nCompany
            onContact = nContact
            obolVerbalWaiver = bolVerbalWaiver
            obolNOIProcessed = bolNOIProcessed
            odtDueDate = dtDueDate
            ostrLocation = strLocation
            'ostrTankPipeID = strTankPipeID
            'ostrTankPipeEntity = strTankPipeEntity
            'ostrSamples = strSamples
            odtableSamples = dtableSamples
            'obolArrCheckList = bolArrCheckList
            'odtArrCheckList = dtArrCheckList
            onCRCertContractor = nCRCertContractor
            onCRCompany = nCRCompany
            odtCRClosureReceived = dtCRClosureReceived
            odtCRClosureDate = dtCRClosureDate
            odtCRDateLastUsed = dtCRDateLastUsed
            obolClosureProcessed = bolClosureProcessed
            onTecID = nTecID
            onTecType = nTecType

            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            ohtBoolCheckList = New Hashtable
            If Not htBoolCheckList Is Nothing Then
                For Each htEntry As DictionaryEntry In htBoolCheckList
                    ohtBoolCheckList.Add(htEntry.Key, htEntry.Value)
                Next
            End If

            ohtDateCheckList = New Hashtable
            If Not htDateCheckList Is Nothing Then
                For Each htEntry As DictionaryEntry In htDateCheckList
                    ohtDateCheckList.Add(htEntry.Key, htEntry.Value)
                Next
            End If

            bolIsDirty = False
        End Sub
        Public Sub Reset()
            If onClosureID > 0 Then
                nClosureID = onClosureID
            End If
            If onFacilityID > 0 Then
                nFacilityID = onFacilityID
            End If
            nFacilitySequence = onFacilitySequence
            nClosureType = onClosureType
            nClosureStatus = onClosureStatus
            nNOIReceived = onNOIReceived
            dtNOI_Rcv_Date = odtNOI_Rcv_Date
            bolOwnerSign = obolOwnerSign
            dtScheduledDate = odtScheduledDate
            nCertContractor = onCertContractor
            nFillMaterial = onFillMaterial
            nCompany = onCompany
            nContact = onContact
            bolVerbalWaiver = obolVerbalWaiver
            bolNOIProcessed = obolNOIProcessed
            dtDueDate = odtDueDate
            strLocation = ostrLocation
            'bolArrCheckList = obolArrCheckList
            'dtArrCheckList = odtArrCheckList
            'strTankPipeID = ostrTankPipeID
            'strTankPipeEntity = ostrTankPipeEntity
            'strSamples = ostrSamples
            dtableSamples = odtableSamples
            nCRCertContractor = onCRCertContractor
            nCRCompany = onCRCompany
            dtCRClosureReceived = odtCRClosureReceived
            dtCRClosureDate = odtCRClosureDate
            dtCRDateLastUsed = odtCRDateLastUsed
            bolClosureProcessed = obolClosureProcessed
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            nTecID = onTecID
            onTecType = onTecType

            htBoolCheckList = New Hashtable
            If Not ohtBoolCheckList Is Nothing Then
                For Each htEntry As DictionaryEntry In ohtBoolCheckList
                    htBoolCheckList.Add(htEntry.Key, htEntry.Value)
                Next
            End If

            htDateCheckList = New Hashtable
            If Not ohtDateCheckList Is Nothing Then
                For Each htEntry As DictionaryEntry In ohtDateCheckList
                    htDateCheckList.Add(htEntry.Key, htEntry.Value)
                Next
            End If

            bolIsDirty = False
            RaiseEvent evtClosureEventInfoChanged(bolIsDirty)
        End Sub

        Public Function DateNuller(ByVal dt As DateTime) As DateTime
            Return IIf(dt > "1/1/1950", dt, "1/1/1950")
        End Function
        Public Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty


            '(strSamples <> ostrSamples) Or _
            '(strTankPipeID <> ostrTankPipeID) Or _
            '(strTankPipeEntity <> ostrTankPipeEntity) Or _
            bolIsDirty = (nClosureType <> onClosureType) Or _
            (nClosureStatus <> onClosureStatus) Or _
            (nNOIReceived <> onNOIReceived) Or _
            (DateNuller(dtNOI_Rcv_Date) <> DateNuller(odtNOI_Rcv_Date)) Or _
            (bolOwnerSign <> obolOwnerSign) Or _
            (DateNuller(dtScheduledDate) <> DateNuller(odtScheduledDate)) Or _
            (nCertContractor <> onCertContractor) Or _
            (nFillMaterial <> onFillMaterial) Or _
            (nCompany <> onCompany) Or _
            (nContact <> onContact) Or _
            (bolVerbalWaiver <> obolVerbalWaiver) Or _
            (bolNOIProcessed <> obolNOIProcessed) Or _
            (DateNuller(dtDueDate) <> DateNuller(odtDueDate)) Or _
            (strLocation <> ostrLocation) Or _
            (nCRCertContractor <> onCRCertContractor) Or _
            (nCRCompany <> onCRCompany) Or _
            (DateNuller(dtCRClosureReceived) <> DateNuller(odtCRClosureReceived)) Or _
            (DateNuller(dtCRClosureDate) <> DateNuller(odtCRClosureDate)) Or _
            (DateNuller(dtCRDateLastUsed) <> DateNuller(odtCRDateLastUsed)) Or _
            (bolClosureProcessed <> obolClosureProcessed) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (DateNuller(dtCreatedOn) <> DateNuller(odtCreatedOn)) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (DateNuller(dtModifiedOn) <> DateNuller(odtModifiedOn)) Or _
            (nTecID <> onTecID) Or _
            (nTecType <> onTecType)
            ' check bolArrCheckList, dtArrCheckList
            Dim index As Integer = 0
            Dim bolChkIsDirty As Boolean = False
            Dim bolChkDtIsDirty As Boolean = False
            Dim bolSamples As Boolean = False
            If nClosureType = onClosureType Then
                If (htBoolCheckList Is Nothing And Not ohtBoolCheckList Is Nothing) Or _
                    (Not htBoolCheckList Is Nothing And ohtBoolCheckList Is Nothing) Then
                    bolChkIsDirty = True
                ElseIf htBoolCheckList Is Nothing And ohtBoolCheckList Is Nothing Then
                    bolChkIsDirty = False
                ElseIf htBoolCheckList.Count <> ohtBoolCheckList.Count Then
                    bolChkIsDirty = True
                Else
                    For Each htEntry As DictionaryEntry In htBoolCheckList
                        If htEntry.Value <> ohtBoolCheckList.Item(htEntry.Key) Then
                            bolChkIsDirty = True
                            Exit For
                        End If
                    Next
                End If
            Else
                bolChkIsDirty = False
            End If
            If nClosureType = onClosureType Then
                If (htDateCheckList Is Nothing And Not ohtDateCheckList Is Nothing) Or _
                    (Not htDateCheckList Is Nothing And ohtDateCheckList Is Nothing) Then
                    bolChkDtIsDirty = True
                ElseIf htDateCheckList Is Nothing And ohtDateCheckList Is Nothing Then
                    bolChkDtIsDirty = False
                ElseIf htDateCheckList.Count <> ohtDateCheckList.Count Then
                    bolChkDtIsDirty = True
                Else
                    For Each htEntry As DictionaryEntry In htDateCheckList
                        If htentry.Value <> ohtDateCheckList.Item(htentry.Key) Then
                            bolChkDtIsDirty = True
                            Exit For
                        End If
                    Next
                End If
            Else
                bolChkDtIsDirty = False
            End If
            'If nClosureType = onClosureType Then
            '    If bolArrCheckList Is Nothing And Not obolArrCheckList Is Nothing Then
            '        bolChkIsDirty = True
            '    ElseIf Not bolArrCheckList Is Nothing And obolArrCheckList Is Nothing Then
            '        bolChkIsDirty = True
            '    ElseIf bolArrCheckList Is Nothing And obolArrCheckList Is Nothing Then
            '        bolChkIsDirty = bolChkIsDirty
            '    ElseIf bolArrCheckList.Count = obolArrCheckList.Count Then
            '        For i As Integer = 0 To bolArrCheckList.Count - 1
            '            If bolArrCheckList.GetByIndex(i) <> obolArrCheckList.GetByIndex(i) Then
            '                bolChkIsDirty = True
            '                Exit For
            '            End If
            '        Next
            '    Else
            '        bolChkIsDirty = True
            '    End If
            'Else
            '    bolChkIsDirty = False
            'End If
            'If nClosureType = onClosureType Then
            '    If dtArrCheckList Is Nothing And Not odtArrCheckList Is Nothing Then
            '        bolChkDtIsDirty = True
            '    ElseIf Not dtArrCheckList Is Nothing And odtArrCheckList Is Nothing Then
            '        bolChkDtIsDirty = True
            '    ElseIf dtArrCheckList Is Nothing And odtArrCheckList Is Nothing Then
            '        bolChkDtIsDirty = bolChkDtIsDirty
            '    ElseIf dtArrCheckList.Count = odtArrCheckList.Count Then
            '        For i As Integer = 0 To dtArrCheckList.Count - 1
            '            If dtArrCheckList.GetByIndex(i) <> odtArrCheckList.GetByIndex(i) Then
            '                bolChkDtIsDirty = True
            '                Exit For
            '            End If
            '        Next
            '    Else
            '        bolChkDtIsDirty = True
            '    End If
            'Else
            '    bolChkDtIsDirty = False
            'End If
            If dtableSamples Is Nothing And Not odtableSamples Is Nothing Then
                bolSamples = True
            ElseIf Not dtableSamples Is Nothing And odtableSamples Is Nothing Then
                bolSamples = True
            ElseIf dtableSamples.Rows.Count = odtableSamples.Rows.Count Then
                For i As Integer = 0 To dtableSamples.Rows.Count - 1
                    For j As Integer = 0 To dtableSamples.Columns.Count - 1
                        If dtableSamples.Rows(i).Item(j) Is System.DBNull.Value And Not odtableSamples.Rows(i).Item(j) Is System.DBNull.Value Then
                            bolSamples = True
                        ElseIf Not dtableSamples.Rows(i).Item(j) Is System.DBNull.Value And odtableSamples.Rows(i).Item(j) Is System.DBNull.Value Then
                            bolSamples = True
                        ElseIf dtableSamples.Rows(i).Item(j) Is System.DBNull.Value And odtableSamples.Rows(i).Item(j) Is System.DBNull.Value Then
                            bolSamples = bolSamples
                        ElseIf dtableSamples.Rows(i).Item(j) <> odtableSamples.Rows(i).Item(j) Then
                            bolSamples = True
                            Exit For
                        End If
                    Next
                    If bolSamples Then Exit For
                Next
            Else
                bolSamples = True
            End If
            bolIsDirty = bolIsDirty Or bolChkIsDirty Or bolChkDtIsDirty Or bolSamples

            RaiseEvent evtClosureEventInfoChanged(bolIsDirty)

        End Sub
#End Region
#Region "Private Methods"
        Private Sub Init()
            onClosureID = 0
            onFacilityID = 0
            onFacilitySequence = 0
            onClosureType = 0
            onClosureStatus = 0
            onNOIReceived = -1
            'odtNOI_Rcv_Date = Now.Date
            odtNOI_Rcv_Date = CDate("01/01/0001")
            obolOwnerSign = False
            'odtScheduledDate = Now.AddDays(30).Date
            odtScheduledDate = CDate("01/01/0001")
            onCertContractor = 0
            onFillMaterial = 0
            onCompany = 0
            onContact = 0
            obolVerbalWaiver = False
            obolNOIProcessed = False
            odtDueDate = CDate("01/01/0001")
            ostrLocation = String.Empty
            strTankPipeID = String.Empty
            strTankPipeEntity = String.Empty
            'ostrSamples = String.Empty
            onCRCertContractor = 0
            onCRCompany = 0
            odtCRClosureReceived = CDate("01/01/0001")
            odtCRClosureDate = CDate("01/01/0001")
            odtCRDateLastUsed = CDate("01/01/0001")
            obolClosureProcessed = False
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            onTecID = 0
            onTecType = 0
            Me.InitializeArray()
            Me.InitializeSamplesTable()
            Me.Reset()
        End Sub
        Private Sub InitializeArray()
            'obolArrCheckList = New SortedList
            'odtArrCheckList = New SortedList
            htBoolCheckList = New Hashtable
            htDateCheckList = New Hashtable
            ohtBoolCheckList = New Hashtable
            ohtDateCheckList = New Hashtable
        End Sub
        Private Sub InitializeSamplesTable()
            Try
                dtableSamples = New DataTable
                dtableSamples.Columns.Add("SAMPLE_ID") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("CLOSURE_ID") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("Sample #") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("Analysis Type") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("Analysis Level") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("Sample Media") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("Sample Location") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("Sample Value", GetType(System.Single)) ' , Type.GetType("System.Single")
                dtableSamples.Columns.Add("Sample Units") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("Sample Constituent") ', Type.GetType("System.Int32"))
                dtableSamples.Columns.Add("CREATED_BY")
                dtableSamples.Columns.Add("DATE_CREATED")
                dtableSamples.Columns.Add("LAST_EDITED_BY")
                dtableSamples.Columns.Add("DATE_LAST_EDITED")
                dtableSamples.Columns.Add("DELETED", Type.GetType("System.Boolean"))
                'dtableSamples.Columns("SAMPLE_ID").DefaultValue = 0
                'dtableSamples.Columns("CLOSURE_ID").DefaultValue = 0
                'dtableSamples.Columns("Sample #").DefaultValue = 0
                'dtableSamples.Columns("Analysis Type").DefaultValue = 0
                'dtableSamples.Columns("Analysis Level").DefaultValue = 0
                'dtableSamples.Columns("Sample Media").DefaultValue = 0
                'dtableSamples.Columns("Sample Location").DefaultValue = 0
                'dtableSamples.Columns("Sample Value").DefaultValue = 0.0
                'dtableSamples.Columns("Sample Units").DefaultValue = 0
                'dtableSamples.Columns("Sample Constituent").DefaultValue = 0
                dtableSamples.Columns("DELETED").DefaultValue = False

                odtableSamples = dtableSamples.Clone
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nClosureID
            End Get
            Set(ByVal Value As Integer)
                nClosureID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityID() As Integer
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Integer)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilitySequence() As Integer
            Get
                Return nFacilitySequence
            End Get
            Set(ByVal Value As Integer)
                nFacilitySequence = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ClosureType() As Integer
            Get
                Return nClosureType
            End Get
            Set(ByVal Value As Integer)
                nClosureType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ClosureStatus() As Integer
            Get
                Return nClosureStatus
            End Get
            Set(ByVal Value As Integer)
                nClosureStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property OldClosureStatus() As Integer
            Get
                Return onClosureStatus
            End Get
        End Property
        Public Property NOIReceived() As Integer
            Get
                Return nNOIReceived
            End Get
            Set(ByVal Value As Integer)
                nNOIReceived = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property NOI_Rcv_Date() As Date
            Get
                Return dtNOI_Rcv_Date
            End Get
            Set(ByVal Value As Date)
                dtNOI_Rcv_Date = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OwnerSign() As Boolean
            Get
                Return bolOwnerSign
            End Get
            Set(ByVal Value As Boolean)
                bolOwnerSign = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ScheduledDate() As Date
            Get
                Return dtScheduledDate
            End Get
            Set(ByVal Value As Date)
                dtScheduledDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CertContractor() As Integer
            Get
                Return nCertContractor
            End Get
            Set(ByVal Value As Integer)
                nCertContractor = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FillMaterial() As Integer
            Get
                Return nFillMaterial
            End Get
            Set(ByVal Value As Integer)
                nFillMaterial = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Company() As Integer
            Get
                Return nCompany
            End Get
            Set(ByVal Value As Integer)
                nCompany = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Contact() As Integer
            Get
                Return nContact
            End Get
            Set(ByVal Value As Integer)
                nContact = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property VerbalWaiver() As Boolean
            Get
                Return bolVerbalWaiver
            End Get
            Set(ByVal Value As Boolean)
                bolVerbalWaiver = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property NOIProcessed() As Boolean
            Get
                Return bolNOIProcessed
            End Get
            Set(ByVal Value As Boolean)
                bolNOIProcessed = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SentToTech() As Date
            Get
                Return dtSentToTech
            End Get
            Set(ByVal Value As Date)
                dtSentToTech = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property NFAbyClosure() As Date
            Get
                Return dtNFAByClosure
            End Get
            Set(ByVal Value As Date)
                dtNFAByClosure = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property NFAbyTech() As Date
            Get
                Return dtNFAByTech
            End Get
            Set(ByVal Value As Date)
                dtNFAByTech = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DueDate() As Date
            Get
                Return dtDueDate
            End Get
            Set(ByVal Value As Date)
                dtDueDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Location() As String
            Get
                Return strLocation
            End Get
            Set(ByVal Value As String)
                strLocation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankPipeID() As String
            Get
                Return strTankPipeID
            End Get
            Set(ByVal Value As String)
                strTankPipeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankPipeEntity() As String
            Get
                Return strTankPipeEntity
            End Get
            Set(ByVal Value As String)
                strTankPipeEntity = Value
                Me.CheckDirty()
            End Set
        End Property
        'Public Property Samples() As String
        '    Get
        '        Return strSamples
        '    End Get
        '    Set(ByVal Value As String)
        '        strSamples = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        Public Property SamplesTable() As DataTable
            Get
                Return dtableSamples
            End Get
            Set(ByVal Value As DataTable)
                dtableSamples = Value
                Me.CheckDirty()
            End Set
        End Property
        Public WriteOnly Property SamplesTableOriginal() As DataTable
            Set(ByVal Value As DataTable)
                odtableSamples = Value
                Me.CheckDirty()
            End Set
        End Property
        'Public Property BoolCheckList() As SortedList
        '    Get
        '        Return bolArrCheckList
        '    End Get
        '    Set(ByVal Value As SortedList)
        '        bolArrCheckList = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property DateCheckList() As SortedList
        '    Get
        '        Return dtArrCheckList
        '    End Get
        '    Set(ByVal Value As SortedList)
        '        dtArrCheckList = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        Public Property HashTableBoolCheckList() As Hashtable
            Get
                Return htBoolCheckList
            End Get
            Set(ByVal Value As Hashtable)
                htBoolCheckList = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HashTableDateCheckList() As Hashtable
            Get
                Return htDateCheckList
            End Get
            Set(ByVal Value As Hashtable)
                htDateCheckList = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HashTableBoolCheckListOriginal() As Hashtable
            Get
                Return ohtBoolCheckList
            End Get
            Set(ByVal Value As Hashtable)
                ohtBoolCheckList = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property HashTableDateCheckListOriginal() As Hashtable
            Get
                Return ohtDateCheckList
            End Get
            Set(ByVal Value As Hashtable)
                ohtDateCheckList = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CRCertContractor() As Integer
            Get
                Return nCRCertContractor
            End Get
            Set(ByVal Value As Integer)
                nCRCertContractor = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CRCompany() As Integer
            Get
                Return nCRCompany
            End Get
            Set(ByVal Value As Integer)
                nCRCompany = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CRClosureReceived() As Date
            Get
                Return dtCRClosureReceived
            End Get
            Set(ByVal Value As Date)
                dtCRClosureReceived = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CRClosureDate() As Date
            Get
                Return dtCRClosureDate
            End Get
            Set(ByVal Value As Date)
                dtCRClosureDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CRDateLastUsed() As Date
            Get
                Return dtCRDateLastUsed
            End Get
            Set(ByVal Value As Date)
                dtCRDateLastUsed = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ClosureProcessed() As Boolean
            Get
                Return bolClosureProcessed
            End Get
            Set(ByVal Value As Boolean)
                bolClosureProcessed = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get
            Set(ByVal value As Int16)
                nAgeThreshold = value
            End Set
        End Property
        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        Public Property CommentsCollection() As MUSTER.Info.CommentsCollection
            Get
                Return colComments
            End Get
            Set(ByVal Value As MUSTER.Info.CommentsCollection)
                colComments = Value
            End Set
        End Property
        Public Property TecID() As Integer
            Get
                Return nTecID
            End Get
            Set(ByVal Value As Integer)
                nTecID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TecType() As Integer
            Get
                Return nTecType
            End Get
            Set(ByVal Value As Integer)
                nTecType = Value
                Me.CheckDirty()
            End Set
        End Property
#Region "iAccessors"
        Public Property CreatedBy() As String
            Get
                If strCreatedBy = Nothing Then
                    Return String.Empty
                Else
                    Return strCreatedBy
                End If
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                If strModifiedBy = Nothing Then
                    Return String.Empty
                Else
                    Return strModifiedBy
                End If
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
            End Set
        End Property
#End Region
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
