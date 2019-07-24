'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.ClosureEvent
'   Provides the operations required to manipulate an ClosureEvent object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0     MNR         03/18/05    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pClosureEvent
#Region "Public Events"
        Public Event FlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer, ByVal eventID As Integer, ByVal eventType As Integer)
        Public Event evtClosureEventErr(ByVal MsgStr As String)
        Public Event evtClosureCRFillMaterial(ByVal bol As Boolean)
        Public Event evtClosureNOIReceived(ByVal bolEnable As Boolean)
        Public Event evtClosureLetter(ByVal letterType As LetterType)
        Public Event evtClosureEventInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oClosureEventInfo As MUSTER.Info.ClosureEventInfo
        Private oClosureEventDB As MUSTER.DataAccess.ClosureEventDB
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private WithEvents oComments As MUSTER.BusinessLogic.pComments
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Closure Event").ID
        Private oProperty As MUSTER.BusinessLogic.pProperty
        Private oFacilityInfo As MUSTER.Info.FacilityInfo
        Public Enum LetterType
            CIPApproval
            CIPDisApproval
            CIPNFA
            InfoNeeded
            RFGApproval
            RFGNFA
            SampleResultMemo
        End Enum
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing, Optional ByRef facInfo As MUSTER.Info.FacilityInfo = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            oClosureEventInfo = New MUSTER.Info.ClosureEventInfo
            oClosureEventDB = New MUSTER.DataAccess.ClosureEventDB
            oComments = New MUSTER.BusinessLogic.pComments
            oProperty = New MUSTER.BusinessLogic.pProperty
            If facInfo Is Nothing Then
                oFacilityInfo = New MUSTER.Info.FacilityInfo
            Else
                oFacilityInfo = facInfo
            End If
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property ID() As Integer
            Get
                Return oClosureEventInfo.ID
            End Get
        End Property
        Public Property FacilityID() As Integer
            Get
                Return oClosureEventInfo.FacilityID
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.FacilityID = Value
            End Set
        End Property
        Public Property FacilitySequence() As Integer
            Get
                Return oClosureEventInfo.FacilitySequence
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.FacilitySequence = Value
            End Set
        End Property
        Public Property ClosureType() As Integer
            Get
                Return oClosureEventInfo.ClosureType
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.ClosureType = Value
                If Value = 445 Then ' CIP
                    RaiseEvent evtClosureCRFillMaterial(True)
                Else
                    RaiseEvent evtClosureCRFillMaterial(False)
                End If
            End Set
        End Property
        Public Property ClosureStatus() As Integer
            Get
                Return oClosureEventInfo.ClosureStatus
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.ClosureStatus = Value
            End Set
        End Property
        Public ReadOnly Property ClosureStatusDesc() As String
            Get
                Return oProperty.GetPropertyNameByID(Me.ClosureStatus)
            End Get
        End Property
        Public Property NOIReceived() As Integer
            Get
                Return oClosureEventInfo.NOIReceived
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.NOIReceived = Value
                CheckNOIReceived()
            End Set
        End Property
        Public Property NOI_Rcv_Date() As Date
            Get
                Return oClosureEventInfo.NOI_Rcv_Date
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.NOI_Rcv_Date = Value
            End Set
        End Property
        Public Property OwnerSign() As Boolean
            Get
                Return oClosureEventInfo.OwnerSign
            End Get
            Set(ByVal Value As Boolean)
                oClosureEventInfo.OwnerSign = Value
            End Set
        End Property
        Public Property ScheduledDate() As Date
            Get
                Return oClosureEventInfo.ScheduledDate
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.ScheduledDate = Value
            End Set
        End Property
        Public Property CertContractor() As Integer
            Get
                Return oClosureEventInfo.CertContractor
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.CertContractor = Value
            End Set
        End Property
        Public Property FillMaterial() As Integer
            Get
                Return oClosureEventInfo.FillMaterial
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.FillMaterial = Value
            End Set
        End Property
        Public Property Company() As Integer
            Get
                Return oClosureEventInfo.Company
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.Company = Value
            End Set
        End Property
        Public Property Contact() As Integer
            Get
                Return oClosureEventInfo.Contact
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.Contact = Value
            End Set
        End Property
        Public Property VerbalWaiver() As Boolean
            Get
                Return oClosureEventInfo.VerbalWaiver
            End Get
            Set(ByVal Value As Boolean)
                oClosureEventInfo.VerbalWaiver = Value
            End Set
        End Property
        Public Property NOIProcessed() As Boolean
            Get
                Return oClosureEventInfo.NOIProcessed
            End Get
            Set(ByVal Value As Boolean)
                oClosureEventInfo.NOIProcessed = Value
            End Set
        End Property
        Public Property SentToTech() As Date
            Get
                Return oClosureEventInfo.SentToTech
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.SentToTech = Value
            End Set
        End Property
        Public Property NFAbyClosure() As Date
            Get
                Return oClosureEventInfo.NFAbyClosure
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.NFAbyClosure = Value
            End Set
        End Property
        Public Property NFAbyTech() As Date
            Get
                Return oClosureEventInfo.NFAbyTech
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.NFAbyTech = Value
            End Set
        End Property
        Public Property DueDate() As Date
            Get
                Return oClosureEventInfo.DueDate
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.DueDate = Value
            End Set
        End Property
        'Public Property EntityID() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        '    Set(ByVal Value As Integer)
        '        nEntityTypeID = Value
        '    End Set
        'End Property
        Public Property Location() As String
            Get
                Return oClosureEventInfo.Location
            End Get
            Set(ByVal Value As String)
                oClosureEventInfo.Location = Value
            End Set
        End Property
        Public Property HashTableBoolCheckList() As Hashtable
            Get
                Return oClosureEventInfo.HashTableBoolCheckList
            End Get
            Set(ByVal Value As Hashtable)
                oClosureEventInfo.HashTableBoolCheckList = Value
            End Set
        End Property
        Public Property HashTableDateCheckList() As Hashtable
            Get
                Return oClosureEventInfo.HashTableDateCheckList
            End Get
            Set(ByVal Value As Hashtable)
                oClosureEventInfo.HashTableDateCheckList = Value
            End Set
        End Property
        Public Property TankPipeID() As String
            Get
                Return oClosureEventInfo.TankPipeID
            End Get
            Set(ByVal Value As String)
                oClosureEventInfo.TankPipeID = Value
            End Set
        End Property
        Public Property TankPipeEntity() As String
            Get
                Return oClosureEventInfo.TankPipeEntity
            End Get
            Set(ByVal Value As String)
                oClosureEventInfo.TankPipeEntity = Value
            End Set
        End Property
        Public Property SamplesTable() As DataTable
            Get
                Return oClosureEventInfo.SamplesTable
            End Get
            Set(ByVal Value As DataTable)
                oClosureEventInfo.SamplesTable = Value
            End Set
        End Property
        Public WriteOnly Property SamplesTableOriginal() As DataTable
            Set(ByVal Value As DataTable)
                oClosureEventInfo.SamplesTableOriginal = Value
            End Set
        End Property
        Public Property CRCertContractor() As Integer
            Get
                Return oClosureEventInfo.CRCertContractor
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.CRCertContractor = Value
            End Set
        End Property
        Public Property CRCompany() As Integer
            Get
                Return oClosureEventInfo.CRCompany
            End Get
            Set(ByVal Value As Integer)
                oClosureEventInfo.CRCompany = Value
            End Set
        End Property
        Public Property CRClosureReceived() As Date
            Get
                Return oClosureEventInfo.CRClosureReceived
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.CRClosureReceived = Value
            End Set
        End Property
        Public Property CRClosureDate() As Date
            Get
                Return oClosureEventInfo.CRClosureDate
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.CRClosureDate = Value
            End Set
        End Property
        Public Property CRDateLastUsed() As Date
            Get
                Return oClosureEventInfo.CRDateLastUsed
            End Get
            Set(ByVal Value As Date)
                oClosureEventInfo.CRDateLastUsed = Value
            End Set
        End Property
        Public Property ClosureProcessed() As Boolean
            Get
                Return oClosureEventInfo.ClosureProcessed
            End Get
            Set(ByVal Value As Boolean)
                oClosureEventInfo.ClosureProcessed = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oClosureEventInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oClosureEventInfo.Deleted = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oClosureEventInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oClosureEventInfo.IsDirty = Value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim oClosureEventInfoLocal As MUSTER.Info.ClosureEventInfo
                For Each oClosureEventInfoLocal In oFacilityInfo.ClosureEventCollection.Values
                    If oClosureEventInfoLocal.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property FacilityInfo() As MUSTER.Info.FacilityInfo
            Get
                Return oFacilityInfo
            End Get
            Set(ByVal Value As MUSTER.Info.FacilityInfo)
                oFacilityInfo = Value
            End Set
        End Property
        Public Property Comments() As MUSTER.BusinessLogic.pComments
            Get
                Return oComments
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pComments)
                oComments = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oClosureEventInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oClosureEventInfo.CreatedBy = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oClosureEventInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oClosureEventInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oClosureEventInfo.CreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oClosureEventInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Function Retrieve(ByRef facInfo As MUSTER.Info.FacilityInfo, Optional ByVal closureID As Integer = 0, Optional ByVal facID As Integer = 0, Optional ByVal showDeleted As Boolean = False, Optional ByVal bolLoading As Boolean = False) As MUSTER.Info.ClosureEventInfo
            Dim bolDataAged As Boolean = False
            Dim ds As New DataSet
            oFacilityInfo = facInfo
            Try
                If Not (oClosureEventInfo.Deleted Or oClosureEventInfo.ID = 0 Or Not oClosureEventInfo.IsDirty Or bolLoading) Then
                    Me.ValidateData()
                End If
                Dim oClosureEventInfoLocal As MUSTER.Info.ClosureEventInfo
                If closureID = 0 And facID = 0 Then
                    Add(oFacilityInfo, 0)
                ElseIf closureID = 0 And facID <> 0 Then
                    ' get by facility id
                    ' check in collection
                    For Each oClosureEventInfoLocal In oFacilityInfo.ClosureEventCollection.Values
                        If oClosureEventInfoLocal.FacilityID = facID Then
                            oClosureEventInfo = oClosureEventInfoLocal
                            Exit Try
                        End If
                    Next
                    ' get from db
                    Dim colClosureEventsLocal As MUSTER.Info.ClosureEventCollection
                    colClosureEventsLocal = oClosureEventDB.DBGetByFacID(facID, showDeleted)
                    If colClosureEventsLocal.Count > 0 Then
                        For Each oClosureEventInfoLocal In colClosureEventsLocal.Values
                            oClosureEventInfo = oClosureEventInfoLocal
                            PopulateSampleTable(oClosureEventInfo.ID)
                            ds = oClosureEventDB.DBGetCheckList(oClosureEventInfo.ID, oClosureEventInfo.ClosureType)
                            PopulateCheckList(ds)
                            oFacilityInfo.ClosureEventCollection.Add(oClosureEventInfo)
                        Next
                    End If
                ElseIf closureID <> 0 And facID = 0 Then
                    ' get by closure id
                    oClosureEventInfo = oFacilityInfo.ClosureEventCollection.Item(closureID.ToString)
                    ' Check for Aged Data here.
                    If Not (oClosureEventInfo Is Nothing) Then
                        If oClosureEventInfo.IsAgedData = True And oClosureEventInfo.IsDirty = False Then
                            bolDataAged = True
                            oFacilityInfo.ClosureEventCollection.Remove(oClosureEventInfo)
                        End If
                    End If
                    If oClosureEventInfo Is Nothing Or bolDataAged Then
                        Add(oFacilityInfo, closureID, showDeleted)
                        PopulateSampleTable(oClosureEventInfo.ID)
                        If oClosureEventInfo.ID > 0 Then
                            ds = oClosureEventDB.DBGetCheckList(oClosureEventInfo.ID, oClosureEventInfo.ClosureType)
                            PopulateCheckList(ds)
                        End If
                    End If
                Else
                    ' get all
                End If

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Return oClosureEventInfo
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal SendMessageToCNE As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oClosureEventInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oClosureEventInfo.ID < 0 And oClosureEventInfo.Deleted) Then
                    oldID = oClosureEventInfo.ID
                    oClosureEventDB.Put(oClosureEventInfo, moduleID, staffID, returnVal, SendMessageToCNE)
                    If Not bolValidated Then
                        If oldID <> oClosureEventInfo.ID Then
                            oFacilityInfo.ClosureEventCollection.ChangeKey(oldID, oClosureEventInfo.ID)
                        End If
                    End If
                    oClosureEventInfo.Archive()
                    oClosureEventInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oClosureEventInfo.Deleted Then
                        ' check if other owners are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oClosureEventInfo.ID Then
                            If strPrev = oClosureEventInfo.ID Then
                                RaiseEvent evtClosureEventErr("ClosureEvent " + oClosureEventInfo.FacilitySequence.ToString + " deleted")
                                oFacilityInfo.ClosureEventCollection.Remove(oClosureEventInfo)
                                If bolDelete Then
                                    oClosureEventInfo = New MUSTER.Info.ClosureEventInfo
                                Else
                                    oClosureEventInfo = Me.Retrieve(oFacilityInfo, 0)
                                End If
                            Else
                                RaiseEvent evtClosureEventErr("ClosureEvent " + oClosureEventInfo.FacilitySequence.ToString + " deleted")
                                oFacilityInfo.ClosureEventCollection.Remove(oClosureEventInfo)
                                oClosureEventInfo = Me.Retrieve(oFacilityInfo, strPrev)
                            End If
                        Else
                            RaiseEvent evtClosureEventErr("ClosureEvent " + oClosureEventInfo.FacilitySequence.ToString + " deleted")
                            oFacilityInfo.ClosureEventCollection.Remove(oClosureEventInfo)
                            oClosureEventInfo = Me.Retrieve(oFacilityInfo, strNext)
                        End If
                    End If
                End If
                PopulateSampleTable(oClosureEventInfo.ID)
                RaiseEvent evtClosureEventInfoChanged(oClosureEventInfo.IsDirty)
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function ValidateData() As Boolean
            Dim strErr As String = String.Empty
            Dim validateSuccess As Boolean = True
            Try
                If oClosureEventInfo.ClosureType = 0 Then
                    strErr += "Closure Type is required" + vbCrLf
                    validateSuccess = False
                End If
                If oClosureEventInfo.NOIReceived = -1 Then
                    strErr += "NOI Received is required" + vbCrLf
                    validateSuccess = False
                End If
                ' if closure type = cip and noi received = true then fill material is required
                If oClosureEventInfo.ClosureType = 445 And oClosureEventInfo.NOIReceived = 1 Then
                    If oClosureEventInfo.FillMaterial = 0 Then
                        strErr += "Fill Material is required" + vbCrLf
                        validateSuccess = False
                    End If
                End If
                If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                    For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                        If htEntry.Value = True Then
                            If Date.Compare(CType(oClosureEventInfo.HashTableDateCheckList.Item(htEntry.Key), Date), CDate("01/01/1900")) = 0 Then
                                If Date.Compare(oClosureEventInfo.DueDate, CDate("01/01/1900")) = 0 Then
                                    strErr += "Due Date is required" + vbCrLf
                                    validateSuccess = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                If strErr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent evtClosureEventErr(strErr)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Return validateSuccess
        End Function
        Public Sub ProcessNOI(ByRef pCal As MUSTER.BusinessLogic.pCalendar, ByVal strUserID As String, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            ' Closure Status
            ' 865 - Open
            ' 866 - Pending
            ' 867 - Approved
            ' 868 - Closed
            ' 869 - Cancelled
            ' 870 - Disapproved
            Dim analysisType As New SortedList
            Dim analysisLevel As New SortedList
            Dim sampleMedia As New SortedList
            Dim sampleLocation As New SortedList
            Dim i As Integer
            Dim strErr As String = String.Empty
            Dim bolContinue As Boolean = True
            Dim ocalInfo As MUSTER.Info.CalendarInfo
            Dim oUser As New MUSTER.BusinessLogic.pUser
            Dim oUserInfo As New MUSTER.Info.UserInfo
            Try
                If oClosureEventInfo.IsDirty AndAlso Not oClosureEventInfo.NOIProcessed Then
                    RaiseEvent evtClosureEventErr("There are unsaved changes, Save before you can Process NOI")
                    Exit Sub
                ElseIf Not oClosureEventInfo.NOIProcessed AndAlso (oClosureEventInfo.ClosureStatus = 868 Or _
                        oClosureEventInfo.ClosureStatus = 867 Or _
                        oClosureEventInfo.ClosureStatus = 870) Then
                    RaiseEvent evtClosureEventErr("Closure Status cannot be Closed / Approved / Disapproved to Process NOI")
                    Exit Sub

                Else
                    For Each dr As DataRow In SamplesTable.Rows
                        If Not dr.Item("DELETED") Then
                            If Not dr.Item("Analysis Type") Is System.DBNull.Value Then
                                If Not analysisType.ContainsKey(CType(dr.Item("Analysis Type"), Integer)) Then
                                    analysisType.Add(CType(dr.Item("Analysis Type"), Integer), CType(dr.Item("Analysis Type"), Integer))
                                End If
                            End If
                            If Not dr.Item("Analysis Level") Is System.DBNull.Value Then
                                If Not analysisLevel.ContainsKey(CType(dr.Item("Analysis Level"), Integer)) Then
                                    analysisLevel.Add(CType(dr.Item("Analysis Level"), Integer), CType(dr.Item("Analysis Level"), Integer))
                                End If
                            End If
                            If Not dr.Item("Sample Media") Is System.DBNull.Value Then
                                If Not sampleMedia.ContainsKey(CType(dr.Item("Sample Media"), Integer)) Then
                                    sampleMedia.Add(CType(dr.Item("Sample Media"), Integer), CType(dr.Item("Sample Media"), Integer))
                                End If
                            End If
                            If Not dr.Item("Sample Location") Is System.DBNull.Value Then
                                If Not sampleLocation.ContainsKey(CType(dr.Item("Sample Location"), Integer)) Then
                                    sampleLocation.Add(CType(dr.Item("Sample Location"), Integer), CType(dr.Item("Sample Location"), Integer))
                                End If
                            End If
                        End If
                    Next

                    If oClosureEventInfo.ClosureType = 445 Then ' CIP
                        If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                            For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                                If CType(htEntry.Value, Boolean) = True Then
                                    If Date.Compare(CType(oClosureEventInfo.HashTableDateCheckList.Item(htEntry.Key), Date), CDate("01/01/1900")) = 0 Then
                                        If Date.Compare(oClosureEventInfo.DueDate, CDate("01/01/1900")) = 0 Then
                                            strErr += "Due Date is required" + vbCrLf
                                            RaiseEvent evtClosureEventErr(strErr)
                                            Exit Sub
                                        End If
                                        ' generate information needed letter (page 5)
                                        RaiseEvent evtClosureLetter(LetterType.InfoNeeded)
                                        oClosureEventInfo.ClosureStatus = 866 ' Pending
                                        Exit Try
                                    End If
                                End If
                            Next
                        End If
                    End If

                    ' samples are required for CIP
                    If oClosureEventInfo.ClosureType = 445 Then ' CIP
                        If analysisType.Count = 0 Then
                            strErr = "Samples required" + vbCrLf
                            RaiseEvent evtClosureEventErr(strErr)
                            Exit Sub
                        End If
                    End If

                    ' b
                    If Date.Compare(oClosureEventInfo.NOI_Rcv_Date, CDate("01/01/0001")) = 0 Then
                        strErr += "NOI Received Date is required" + vbCrLf
                    End If
                    If Date.Compare(oClosureEventInfo.ScheduledDate, CDate("01/01/0001")) = 0 Then
                        strErr += "Scheduled Date is required" + vbCrLf
                    End If
                    If oClosureEventInfo.CertContractor = 0 Then
                        strErr += "Certified Contractor is required" + vbCrLf
                    End If
                    If oClosureEventInfo.Company = 0 Then
                        strErr += "Company is required" + vbCrLf
                    End If
                    If oClosureEventInfo.ClosureType = 445 Then
                        If analysisLevel.Count = 0 Then
                            strErr += "Analysis Level is required" + vbCrLf
                        End If
                    End If
                    If strErr.Length > 0 Then
                        RaiseEvent evtClosureEventErr(strErr)
                        oClosureEventInfo.Reset()
                        Exit Sub
                    End If

                    ' c
                    If oClosureEventInfo.ClosureType = 444 Then ' RFG
                        ' generate RFG Approval letter
                        RaiseEvent evtClosureLetter(LetterType.RFGApproval)

                        If Not oClosureEventInfo.NOIProcessed Then
                            oClosureEventInfo.ClosureStatus = 867
                            oClosureEventInfo.NOIProcessed = True
                        End If

                        Exit Try
                    End If

                    ' d
                    If oClosureEventInfo.ClosureType = 445 Then ' CIP
                        If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                            For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                                If htEntry.Value = True Then
                                    If Date.Compare(CType(oClosureEventInfo.HashTableDateCheckList.Item(htEntry.Key), Date), CDate("01/01/0001")) = 0 Then
                                        bolContinue = False
                                    End If
                                End If
                            Next
                        End If
                        If bolContinue Then
                            If (Not analysisLevel.ContainsKey(890)) And _
                                (analysisLevel.ContainsKey(884) Or _
                                analysisLevel.ContainsKey(885)) Then
                                ' generate CIP Approval Letter
                                RaiseEvent evtClosureLetter(LetterType.CIPApproval)

                                If Not oClosureEventInfo.NOIProcessed Then
                                    oClosureEventInfo.ClosureStatus = 867
                                    oClosureEventInfo.NOIProcessed = True
                                End If

                                Exit Try
                            End If
                        End If
                    End If

                    ' e
                    If oClosureEventInfo.ClosureType = 445 Then ' CIP
                        If analysisLevel.ContainsKey(890) Then
                            ' generate CIP Disapproval letter
                            RaiseEvent evtClosureLetter(LetterType.CIPDisApproval)
                            ' generate Sample Results Memo
                            RaiseEvent evtClosureLetter(LetterType.SampleResultMemo)

                            If Not oClosureEventInfo.NOIProcessed Then


                                oClosureEventInfo.ClosureStatus = 870 ' Disapproved
                                Dim LEStatusCount As Integer = 0
                                Dim pLustEvent As New MUSTER.BusinessLogic.pLustEvent
                                Dim oLEInfoLocal As MUSTER.Info.LustEventInfo
                                Dim pLustEventActivity As New MUSTER.BusinessLogic.pLustEventActivity
                                Dim colLELocal As MUSTER.Info.LustEventCollection

                                oUserInfo = oUser.RetrievePMHead()
                                colLELocal = pLustEvent.GetAll(oClosureEventInfo.FacilityID)

                                For Each oLEInfoLocal In colLELocal.Values
                                    If oLEInfoLocal.EventStatus = 624 Then
                                        LEStatusCount += 1
                                    End If
                                Next

                                If LEStatusCount = 0 Then
                                    ' create lust event for current facility with info on page 6
                                    Dim soil As Boolean = sampleMedia.ContainsKey(878)
                                    Dim soilBTEX As Boolean
                                    Dim soilPAH As Boolean
                                    If soil Then
                                        soilBTEX = analysisType.Contains(875)
                                        soilPAH = analysisType.ContainsKey(876)
                                    End If
                                    Dim gw As Boolean = sampleMedia.ContainsKey(879)
                                    Dim gwBTEX As Boolean = False
                                    Dim gwPAH As Boolean = False
                                    If gw Then
                                        gwBTEX = analysisType.ContainsKey(875)
                                        gwPAH = analysisType.ContainsKey(876)
                                    End If
                                    Dim strTankAndPipe As String = String.Empty
                                    Dim strTankAndPipeID, strTankAndPipeType As String()
                                    strTankAndPipeID = oClosureEventInfo.TankPipeID.Split("|")
                                    strTankAndPipeType = oClosureEventInfo.TankPipeEntity.Split("|")
                                    For i = 0 To strTankAndPipeID.Length - 1
                                        If strTankAndPipeType(i) = 10 Then
                                            strTankAndPipe += "P"
                                        Else
                                            strTankAndPipe += "T"
                                        End If
                                        strTankAndPipe += strTankAndPipeID(i) + "|"
                                    Next
                                    strTankAndPipe = strTankAndPipe.Trim.TrimEnd("|")
                                    ' 624 - event status - open
                                    ' 617 - mgptfstatus - eud
                                    ' 623 - release status - confirmed
                                    ' 653 - how discovered - tank closure
                                    ' 655 - identified by - owner/operator
                                    ' 61 - event pm id
                                    oLEInfoLocal = New MUSTER.Info.LustEventInfo(0, CDate("01/01/0001"), CDate("01/01/0001"), CDate("01/01/0001"), CDate("01/01/0001"), oClosureEventInfo.NOI_Rcv_Date, oClosureEventInfo.NOI_Rcv_Date, _
                                    624, oClosureEventInfo.FacilityID, 0, 617, 0, oUserInfo.UserKey, 623, 0, String.Empty, strUserID, "01/01/0001", String.Empty, CDate("01/01/0001"), 0, _
                                    oClosureEventInfo.NOI_Rcv_Date, 655, 0, 0, oUserInfo.UserKey, Date.Now, CDate("01/01/0001"), soil, _
                                    soilBTEX, soilPAH, False, gw, gwBTEX, gwPAH, False, False, False, False, False, False, False, False, False, False, 0, String.Empty, strTankAndPipe, 0, _
                                    CDate("01/01/0001"), String.Empty, 0, CDate("01/01/0001"), String.Empty, 0, CDate("01/01/0001"), String.Empty, False, 0, CDate("01/01/0001"), String.Empty, _
                                    False, String.Empty, String.Empty, String.Empty, String.Empty, 0, 0, False, False, False, False, False, False, False, False, False, True, False, 0)
                                    pLustEvent.Add(oLEInfoLocal)
                                    pLustEvent.Save(moduleID, staffID, returnVal)
                                    oClosureEventInfo.SentToTech = Today.Date
                                    oClosureEventInfo.TecID = pLustEvent.ID
                                    oClosureEventInfo.TecType = 7

                                    ' create a tank closure activity for the lust event
                                    ' (new site created above or only one open existing site)
                                    ' of type tank closure if location includes tank,
                                    ' piping trench or pump island
                                    ' 
                                    If sampleLocation.ContainsKey(871) Or _
                                        sampleLocation.ContainsKey(872) Or _
                                        sampleLocation.ContainsKey(874) Then
                                        pLustEventActivity.Add(New MUSTER.Info.LustActivityInfo(0, _
                                            pLustEvent.ID, _
                                            Now.Date, _
                                            CDate("01/01/0001"), _
                                            CDate("01/01/0001"), _
                                            CDate("01/01/0001"), _
                                            CDate("01/01/0001"), _
                                            697, _
                                            strUserID, _
                                            Now.Date, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            0, 0))
                                        pLustEventActivity.Save(moduleID, staffID, returnVal)
                                        If Not returnVal = String.Empty Then
                                            Exit Sub
                                        End If
                                        ' override the id set above as activity will be closed before event is closed
                                        oClosureEventInfo.TecID = pLustEventActivity.ActivityID
                                        oClosureEventInfo.TecType = 23
                                    End If

                                    ' create a To Do calendar entry on the current date for
                                    ' the PM - Head user indicating
                                    ' "New Activities created from Closure"
                                    ocalInfo = New MUSTER.Info.CalendarInfo(0, _
                                                                        Now(), _
                                                                        Now(), _
                                                                        0, _
                                                                        "New Activities created from Closure : " + oClosureEventInfo.FacilitySequence.ToString + " belonging to Facility : " + oClosureEventInfo.FacilityID.ToString, _
                                                                        oUserInfo.ID, _
                                                                        "SYSTEM", _
                                                                        String.Empty, _
                                                                        False, _
                                                                        True, _
                                                                        False, _
                                                                        False, _
                                                                        strUserID, _
                                                                        CDate("01/01/0001"), _
                                                                        String.Empty, _
                                                                        CDate("01/01/0001"), _
                                                                        22, _
                                                                        oClosureEventInfo.ID)
                                    pCal.Add(ocalInfo)
                                    pCal.Save()
                                ElseIf LEStatusCount > 1 Then
                                    ' create a To Do Calendar entry on the current date
                                    ' for the PM-Head user indicating
                                    ' "A new Activity needs to be created for Dirty Samples"
                                    ocalInfo = New MUSTER.Info.CalendarInfo(0, _
                                                                        Now(), _
                                                                        Now(), _
                                                                        0, _
                                                                        "A new Activity needs to be created for Dirty Samples - Closure : " + oClosureEventInfo.FacilitySequence.ToString + " in Facility : " + oClosureEventInfo.FacilityID.ToString, _
                                                                        oUserInfo.ID, _
                                                                        "SYSTEM", _
                                                                        String.Empty, _
                                                                        False, _
                                                                        True, _
                                                                        False, _
                                                                        False, _
                                                                        strUserID, _
                                                                        CDate("01/01/0001"), _
                                                                        String.Empty, _
                                                                        CDate("01/01/0001"), _
                                                                        22, _
                                                                        oClosureEventInfo.ID)
                                    pCal.Add(ocalInfo)
                                    pCal.Save()
                                    oClosureEventInfo.SentToTech = Today.Date
                                Else
                                    ' create a tank closure activity for the lust event
                                    ' (new site created above or only one open existing site)
                                    ' of type tank closure if location includes tank,
                                    ' piping trench or pump island
                                    ' 
                                    If sampleLocation.ContainsKey(871) Or _
                                        sampleLocation.ContainsKey(872) Or _
                                        sampleLocation.ContainsKey(874) Then
                                        pLustEvent.Add(colLELocal.Item(colLELocal.GetKeys(0)))
                                        pLustEventActivity.Add(New MUSTER.Info.LustActivityInfo(0, _
                                            pLustEvent.ID, _
                                            Now.Date, _
                                            CDate("01/01/0001"), _
                                            CDate("01/01/0001"), _
                                            CDate("01/01/0001"), _
                                            CDate("01/01/0001"), _
                                            697, _
                                            strUserID, _
                                            Now.Date, _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            0, 0))
                                        pLustEventActivity.Save(moduleID, staffID, returnVal)
                                        If Not returnVal = String.Empty Then
                                            Exit Sub
                                        End If
                                        oClosureEventInfo.SentToTech = Today.Date
                                        oClosureEventInfo.TecID = pLustEventActivity.ActivityID
                                        oClosureEventInfo.TecType = 23
                                    End If

                                    ' create a To Do calendar entry on the current date for
                                    ' the PM - Head user indicating
                                    ' "New Activities created from Closure"
                                    ocalInfo = New MUSTER.Info.CalendarInfo(0, _
                                                                        Now(), _
                                                                        Now(), _
                                                                        0, _
                                                                        "New Activities created from Closure : " + oClosureEventInfo.FacilitySequence.ToString + " belonging to Facility : " + oClosureEventInfo.FacilityID.ToString, _
                                                                        oUserInfo.ID, _
                                                                        "SYSTEM", _
                                                                        String.Empty, _
                                                                        False, _
                                                                        True, _
                                                                        False, _
                                                                        False, _
                                                                        strUserID, _
                                                                        CDate("01/01/0001"), _
                                                                        String.Empty, _
                                                                        CDate("01/01/0001"), _
                                                                        22, _
                                                                        oClosureEventInfo.ID)
                                    pCal.Add(ocalInfo)
                                    pCal.Save()
                                    oClosureEventInfo.SentToTech = Today.Date
                                End If
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Me.Save(moduleID, staffID, returnVal)
        End Sub
        Public Function ProcessClosure(ByRef pCal As MUSTER.BusinessLogic.pCalendar, ByVal strUserID As String, ByVal ModuleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strGroupID As String = "", Optional ByVal BackFillADL As Boolean = True) As Boolean
            ' Closure Status
            ' 865 - Open
            ' 866 - Pending
            ' 867 - Approved
            ' 868 - Closed
            ' 869 - Cancelled
            ' 870 - Disapproved
            Dim dt As New Date
            dt = CDate("01/01/0001")
            Dim analysisLevel As New SortedList
            Dim sampleMedia As New SortedList
            Dim analysisType As New SortedList
            Dim sendmessagetoCne As Boolean = False
            Dim sampleLocation As New SortedList
            Dim i As Integer
            Dim strErr As String = String.Empty
            Dim bolContinue As Boolean = True
            Try
                If oClosureEventInfo.TankPipeID = String.Empty Then
                    RaiseEvent evtClosureEventErr("Atleast one Tank / Pipe has to be selected to Process Closure")
                    Return False
                ElseIf oClosureEventInfo.IsDirty Then
                    RaiseEvent evtClosureEventErr("There are unsaved changes, Save before you can Process Closure")
                    Return False
                ElseIf oClosureEventInfo.ClosureStatus = 868 Then

                    For Each dr As DataRow In SamplesTable.Rows

                        If Not dr.Item("Analysis Type") Is System.DBNull.Value Then
                            If Not analysisType.ContainsKey(CType(dr.Item("Analysis Type"), Integer)) Then
                                analysisType.Add(CType(dr.Item("Analysis Type"), Integer), CType(dr.Item("Analysis Type"), Integer))
                            End If
                        End If
                        If Not dr.Item("Analysis Level") Is System.DBNull.Value Then
                            If Not analysisLevel.ContainsKey(CType(dr.Item("Analysis Level"), Integer)) Then
                                analysisLevel.Add(CType(dr.Item("Analysis Level"), Integer), CType(dr.Item("Analysis Level"), Integer))
                            End If
                        End If
                        If Not dr.Item("Sample Media") Is System.DBNull.Value Then
                            If Not sampleMedia.ContainsKey(CType(dr.Item("Sample Media"), Integer)) Then
                                sampleMedia.Add(CType(dr.Item("Sample Media"), Integer), CType(dr.Item("Sample Media"), Integer))
                            End If
                        End If
                        If Not dr.Item("Sample Location") Is System.DBNull.Value Then
                            If Not sampleLocation.ContainsKey(CType(dr.Item("Sample Location"), Integer)) Then
                                sampleLocation.Add(CType(dr.Item("Sample Location"), Integer), CType(dr.Item("Sample Location"), Integer))
                            End If
                        End If
                    Next

                    'd

                    bolContinue = True

                    If oClosureEventInfo.ClosureType = 445 Then ' CIP
                        If oClosureEventInfo.NOIReceived = 1 Then

                            If (analysisLevel.ContainsKey(885) Or analysisLevel.ContainsKey(884) Or BackFillADL) And (Not analysisLevel.ContainsKey(890) Or BackFillADL) Then
                                ' Generate CIP NFA Letter
                                bolContinue = False
                                RaiseEvent evtClosureLetter(LetterType.CIPNFA)
                            End If

                        Else
                            ' Issue #3178    Need to send NFA letter even when a NOI has not been received
                            'code added on July 01, 2008 by Hua Cao
                            If (analysisLevel.ContainsKey(885) Or analysisLevel.ContainsKey(884) Or BackFillADL) And (Not analysisLevel.ContainsKey(890) Or BackFillADL) Then
                                ' Generate CIP NFA Letter
                                bolContinue = False

                                RaiseEvent evtClosureLetter(LetterType.CIPNFA)
                            End If
                        End If
                    End If


                    ' e
                    If oClosureEventInfo.ClosureType = 444 Then ' RFG
                        If (analysisLevel.ContainsKey(885) Or analysisLevel.ContainsKey(884) Or BackFillADL) And (Not analysisLevel.ContainsKey(890) Or BackFillADL) Then
                            ' Generate RFG NFA Letter
                            bolContinue = False

                            RaiseEvent evtClosureLetter(LetterType.RFGNFA)
                        End If
                    End If

                    ' f
                    If bolContinue AndAlso analysisLevel.ContainsKey(890) And Not BackFillADL Then
                        ' i
                        ' Generate Sample Result Memo
                        RaiseEvent evtClosureLetter(LetterType.SampleResultMemo)
                    End If

                    'RaiseEvent evtClosureEventErr("Closure Status cannot be Closed to Process Closure")

                    Return True

                Else

                    Dim strMsg As String
                    For Each dr As DataRow In SamplesTable.Rows
                        If Not dr.Item("Analysis Type") Is System.DBNull.Value Then
                            If Not analysisType.ContainsKey(CType(dr.Item("Analysis Type"), Integer)) Then
                                analysisType.Add(CType(dr.Item("Analysis Type"), Integer), CType(dr.Item("Analysis Type"), Integer))
                            End If
                        End If
                        If Not dr.Item("Analysis Level") Is System.DBNull.Value Then
                            If Not analysisLevel.ContainsKey(CType(dr.Item("Analysis Level"), Integer)) Then
                                analysisLevel.Add(CType(dr.Item("Analysis Level"), Integer), CType(dr.Item("Analysis Level"), Integer))
                            End If
                        End If
                        If Not dr.Item("Sample Media") Is System.DBNull.Value Then
                            If Not sampleMedia.ContainsKey(CType(dr.Item("Sample Media"), Integer)) Then
                                sampleMedia.Add(CType(dr.Item("Sample Media"), Integer), CType(dr.Item("Sample Media"), Integer))
                            End If
                        End If
                        If Not dr.Item("Sample Location") Is System.DBNull.Value Then
                            If Not sampleLocation.ContainsKey(CType(dr.Item("Sample Location"), Integer)) Then
                                sampleLocation.Add(CType(dr.Item("Sample Location"), Integer), CType(dr.Item("Sample Location"), Integer))
                            End If
                        End If
                    Next
                    ' a
                    If oClosureEventInfo.ClosureType = 444 Then ' RFG
                        If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                            For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                                If CType(htEntry.Value, Boolean) = True Then
                                    ' Due date required
                                    If Date.Compare(CType(oClosureEventInfo.HashTableDateCheckList.Item(htEntry.Key), Date), CDate("01/01/0001")) = 0 Then
                                        If Date.Compare(oClosureEventInfo.DueDate, CDate("01/01/0001")) = 0 Then
                                            strErr += "Due Date is required" + vbCrLf
                                            RaiseEvent evtClosureEventErr(strErr)
                                            Return False
                                        End If
                                        ' generate information needed letter (page 10)
                                        RaiseEvent evtClosureLetter(LetterType.InfoNeeded)
                                        oClosureEventInfo.ClosureStatus = 866 ' Pending
                                        Me.Save(ModuleID, staffID, returnVal)
                                        If Not returnVal = String.Empty Then
                                            Return False
                                        End If
                                        Return False
                                    End If
                                End If
                            Next
                        End If
                    End If

                    ' b
                    If oClosureEventInfo.ClosureType = 445 Then ' CIP
                        If oClosureEventInfo.FillMaterial = 0 Then
                            strErr += "Fill Material is required" + vbCrLf
                        End If
                    End If
                    If Date.Compare(oClosureEventInfo.CRClosureReceived, dt) = 0 Then
                        strErr += "Closure Received Date is required" + vbCrLf
                    End If
                    If Date.Compare(oClosureEventInfo.CRClosureDate, dt) = 0 Then
                        strErr += "Date Closed is required" + vbCrLf
                    End If
                    If Date.Compare(oClosureEventInfo.CRDateLastUsed, dt) = 0 Then
                        strErr += "Date Last Used is required" + vbCrLf
                    End If
                    If oClosureEventInfo.CRCertContractor = 0 Then
                        strErr += "Certified Contractor is required" + vbCrLf
                    End If
                    If analysisLevel.Count = 0 Then
                        strErr += "Analysis Level is required" + vbCrLf
                    End If
                    If sampleMedia.Count = 0 Then
                        strErr += "Sample Media is required" + vbCrLf
                    End If
                    If analysisType.Count = 0 Then
                        strErr += "Analysis Type is required" + vbCrLf
                    End If
                    If analysisLevel.ContainsKey(890) Then
                        If sampleLocation.Count = 0 Then
                            strErr += "Sample Location is required" + vbCrLf
                        End If
                    End If
                    If strErr.Length > 0 Then
                        RaiseEvent evtClosureEventErr(strErr)
                        oClosureEventInfo.Reset()
                        Return False
                    End If

                    ' c
                    If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                        For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                            If CType(htEntry.Value, Boolean) = True Then
                                If Date.Compare(CType(oClosureEventInfo.HashTableDateCheckList.Item(htEntry.Key), Date), CDate("01/01/0001")) = 0 Then
                                    bolContinue = False
                                    Exit For
                                End If
                            End If
                        Next
                    End If

                    ' #1804
                    ' If facility has fee balance, cannot close event
                    Dim ds As DataSet
                    ds = oClosureEventDB.DBGetDS("SELECT dbo.udfGetOwnerPastDueFees( (SELECT OWNER_ID FROM TBLREG_FACILITY WHERE FACILITY_ID = " + oClosureEventInfo.FacilityID.ToString + "),0," + oClosureEventInfo.FacilityID.ToString + ")")
                    If ds.Tables(0).Rows(0)(0) > 0 Then
                        ' #2937
                        ' if closure event only has pipes, continue without message
                        ' Only has pipes => it does not have any tanks
                        If oClosureEventInfo.TankPipeEntity.IndexOf("12") > -1 Then
                            Dim oUser As New MUSTER.BusinessLogic.pUser
                            Dim oUserInfo As New MUSTER.Info.UserInfo
                            oUserInfo = ouser.RetrieveClosureHead
                            If strUserID.Equals(ouserinfo.ID) Then
                                If MsgBox("Facility has Fee Balance. Do you want to continue?", MsgBoxStyle.YesNo, "Muster") = MsgBoxResult.No Then
                                    oClosureEventInfo.Reset()
                                    Return False
                                End If
                            Else
                                strMsg = "Facility has Fee Balance. Cannot Close Event"
                            End If
                        End If
                    End If

                    ' #3034 If an inspection for the facility is submitted but not processed by C&E cannot close closure event
                    '       Need to check for submitted inspection only if tank/pipe associated with the event
                    If oClosureEventInfo.TankPipeID <> String.Empty AndAlso (strMsg Is Nothing OrElse strMsg.Length = 0) Then
                        Dim strMsg1 As String = CheckIfFacHasSubmittedInspection()
                        If strMsg1 <> String.Empty Then

                            If MsgBox(String.Format("{0}. Do you want to continue?", strMsg1), MsgBoxStyle.YesNo, "Muster") = MsgBoxResult.No Then
                                oClosureEventInfo.Reset()
                                Return False
                            Else
                                strMsg1 = Nothing
                                sendmessagetoCne = True
                            End If

                            'strMsg += vbCrLf + strMsg1
                        End If
                    End If

                    If bolContinue Then
                        If strMsg <> String.Empty Then
                            oClosureEventInfo.Reset()
                            RaiseEvent evtClosureEventErr(strMsg)
                            Return False
                        End If
                        oClosureEventInfo.ClosureStatus = 868 ' Closed
                        oClosureEventInfo.ClosureProcessed = True
                        'oClosureEventInfo.NFAbyClosure = Today.Date
                        Me.Save(ModuleID, staffID, returnVal, , , sendmessagetoCne)
                        If Not returnVal = String.Empty Then
                            Return False
                        End If
                    Else
                        bolContinue = True
                        'added on 07/02/2008 by Hua Cao
                        If strMsg <> String.Empty Then
                            oClosureEventInfo.Reset()
                            RaiseEvent evtClosureEventErr(strMsg)
                            Return False
                        End If
                        oClosureEventInfo.ClosureStatus = 868 ' Closed
                        oClosureEventInfo.ClosureProcessed = True
                        Me.Save(ModuleID, staffID, returnVal, , , sendmessagetoCne)
                        If Not returnVal = String.Empty Then
                            Return False
                        End If
                    End If

                    'd
                    If oClosureEventInfo.ClosureType = 445 Then ' CIP
                        If oClosureEventInfo.NOIReceived = 1 Then
                            If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                                For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                                    If CType(htEntry.Value, Boolean) = True Then
                                        If Date.Compare(CType(oClosureEventInfo.HashTableDateCheckList.Item(htentry.Key), Date), dt) = 0 Then
                                            bolContinue = False
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If bolContinue Then
                                If (analysisLevel.ContainsKey(885) Or analysisLevel.ContainsKey(884) Or BackFillADL) And (Not analysisLevel.ContainsKey(890) Or BackFillADL) Then
                                    ' Generate CIP NFA Letter
                                    oClosureEventInfo.NFAbyClosure = Today.Date
                                    Me.Save(ModuleID, staffID, returnVal)
                                    If Not returnVal = String.Empty Then
                                        Return False
                                    End If
                                    RaiseEvent evtClosureLetter(LetterType.CIPNFA)
                                End If
                            End If
                        Else
                            ' Issue #3178    Need to send NFA letter even when a NOI has not been received
                            'code added on July 01, 2008 by Hua Cao
                            If (analysisLevel.ContainsKey(885) Or analysisLevel.ContainsKey(884) Or BackFillADL) And (Not analysisLevel.ContainsKey(890) Or BackFillADL) Then
                                ' Generate CIP NFA Letter
                                oClosureEventInfo.NFAbyClosure = Today.Date
                                Me.Save(ModuleID, staffID, returnVal)
                                If Not returnVal = String.Empty Then
                                    Return False
                                End If
                                RaiseEvent evtClosureLetter(LetterType.CIPNFA)
                            End If
                        End If
                    End If

                    bolContinue = True

                    ' e
                    If oClosureEventInfo.ClosureType = 444 Then ' RFG
                        If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                            For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                                If CType(htEntry.Value, Boolean) = True Then
                                    If Date.Compare(CType(oClosureEventInfo.HashTableDateCheckList.Item(htentry.Key), Date), dt) = 0 Then
                                        bolContinue = False
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                        If bolContinue Then
                            If (analysisLevel.ContainsKey(885) Or analysisLevel.ContainsKey(884) Or BackFillADL) And (Not analysisLevel.ContainsKey(890) Or BackFillADL) Then
                                ' Generate RFG NFA Letter
                                oClosureEventInfo.NFAbyClosure = Today.Date
                                Me.Save(ModuleID, staffID, returnVal)
                                If Not returnVal = String.Empty Then
                                    Return False
                                End If
                                RaiseEvent evtClosureLetter(LetterType.RFGNFA)
                            End If
                        End If
                    End If

                    bolContinue = True

                    ' f
                    If analysisLevel.ContainsKey(890) And Not BackFillADL Then
                        ' i
                        ' Generate Sample Result Memo
                        RaiseEvent evtClosureLetter(LetterType.SampleResultMemo)

                        Dim LEStatusCount As Integer = 0
                        Dim LEID As Integer = 0
                        Dim pLustEvent As New MUSTER.BusinessLogic.pLustEvent
                        Dim oLEInfoLocal As MUSTER.Info.LustEventInfo
                        Dim colLELocal As MUSTER.Info.LustEventCollection
                        Dim pLustEventActivity As New MUSTER.BusinessLogic.pLustEventActivity
                        Dim oLAInfoLocal As MUSTER.Info.LustActivityInfo
                        Dim colCal As New MUSTER.Info.CalendarCollection
                        Dim ocalInfo As MUSTER.Info.CalendarInfo
                        Dim oUser As New MUSTER.BusinessLogic.pUser
                        Dim oUserInfo As New MUSTER.Info.UserInfo

                        oUserInfo = oUser.RetrievePMHead()
                        colLELocal = pLustEvent.GetAll(oClosureEventInfo.FacilityID)

                        For Each oLEInfoLocal In colLELocal.Values
                            If oLEInfoLocal.EventStatus = 624 Then
                                LEStatusCount += 1
                                LEID = oLEInfoLocal.ID
                            End If
                        Next

                        ' ii
                        ' no lust event or no open lust event
                        If LEStatusCount = 0 Then
                            ' create lust event for current facility with info on page 11
                            Dim soil As Boolean = sampleMedia.ContainsKey(878)
                            Dim soilBTEX As Boolean = False
                            Dim soilPAH As Boolean = False
                            If soil Then
                                soilBTEX = analysisType.Contains(875)
                                soilPAH = analysisType.ContainsKey(876)
                            End If
                            Dim gw As Boolean = sampleMedia.ContainsKey(879)
                            Dim gwBTEX As Boolean = False
                            Dim gwPAH As Boolean = False
                            If gw Then
                                gwBTEX = analysisType.ContainsKey(875)
                                gwPAH = analysisType.ContainsKey(876)
                            End If
                            Dim strTankAndPipe As String = String.Empty
                            Dim strTankAndPipeID, strTankAndPipeType As String()
                            strTankAndPipeID = oClosureEventInfo.TankPipeID.Split("|")
                            strTankAndPipeType = oClosureEventInfo.TankPipeEntity.Split("|")
                            For i = 0 To strTankAndPipeID.Length - 1
                                If strTankAndPipeType(i) = 10 Then
                                    strTankAndPipe += "P"
                                Else
                                    strTankAndPipe += "T"
                                End If
                                strTankAndPipe += strTankAndPipeID(i) + "|"
                            Next
                            strTankAndPipe = strTankAndPipe.Trim.TrimEnd("|")
                            oLEInfoLocal = New MUSTER.Info.LustEventInfo(0, Nothing, Nothing, Nothing, Nothing, oClosureEventInfo.NOI_Rcv_Date, _
                                                oClosureEventInfo.NOI_Rcv_Date, 624, oClosureEventInfo.FacilityID, _
                                                0, 617, 0, oUserInfo.UserKey, 623, 0, String.Empty, strUserID, _
                                                CDate("01/01/1900"), String.Empty, CDate("01/01/1900"), 0, _
                                                oClosureEventInfo.NOI_Rcv_Date, 655, 0, 0, oUserInfo.UserKey, _
                                                Date.Now, Nothing, soil, soilBTEX, soilPAH, False, gw, gwBTEX, _
                                                gwPAH, False, False, False, False, False, False, False, False, _
                                                False, False, 0, String.Empty, strTankAndPipe, 0, Nothing, _
                                                String.Empty, 0, Nothing, String.Empty, 0, Nothing, String.Empty, _
                                                False, 0, Nothing, String.Empty, False, String.Empty, _
                                                String.Empty, String.Empty, String.Empty, 0, 0, _
                                                False, False, False, False, False, False, False, _
                                                False, False, True, False, 0)
                            pLustEvent.Add(oLEInfoLocal)
                            pLustEvent.Save(ModuleID, staffID, returnVal)
                            If Not returnVal = String.Empty Then
                                Return False
                            End If
                            LEStatusCount += 1
                            LEID = pLustEvent.ID
                            oClosureEventInfo.TecID = pLustEvent.ID
                            oClosureEventInfo.TecType = 7
                        End If

                        If LEStatusCount > 1 Then
                            ' iii
                            ' create a To Do Calendar entry on the current date
                            ' for the PM-Head user indicating
                            ' "A new Activity needs to be created for a Dirty Closure"
                            ocalInfo = New MUSTER.Info.CalendarInfo(0, _
                                        Now(), _
                                        Now(), _
                                        0, _
                                        "A new Activity needs to be created for a Dirty Closure : " + oClosureEventInfo.FacilitySequence.ToString + " belonging to Facility : " + oClosureEventInfo.FacilityID.ToString, _
                                        oUserInfo.ID, _
                                        "SYSTEM", _
                                        strGroupID, _
                                        False, _
                                        True, _
                                        False, _
                                        False, _
                                        String.Empty, _
                                        CDate("01/01/1900"), _
                                        String.Empty, _
                                        CDate("01/01/1900"), _
                                        22, _
                                        oClosureEventInfo.ID)
                            pCal.Add(ocalInfo)
                            pCal.Save(ModuleID, staffID, returnVal)
                            If Not returnVal = String.Empty Then
                                Return False
                            End If
                            oClosureEventInfo.SentToTech = Today.Date
                            Me.Save(ModuleID, staffID, returnVal)
                            If Not returnVal = String.Empty Then
                                Return False
                            End If
                        ElseIf LEStatusCount = 1 Then
                            ' if 1 open lust event
                            pLustEvent.Retrieve(LEID)

                            If (oClosureEventInfo.ClosureType = 445 And oClosureEventInfo.NOIReceived = 0) Or _
                                oClosureEventInfo.ClosureType = 444 Then
                                ' iv
                                ' (closure type = cip and noi received = no) or (closure type = rfg)
                                If sampleLocation.ContainsKey(871) Or _
                                    sampleLocation.ContainsKey(872) Or _
                                    sampleLocation.ContainsKey(874) Then
                                    ' create a tank closure activity for the lust event
                                    ' (new site created above or only one open existing site)
                                    ' of type tank closure if location includes tank or piping trench or pump island
                                    oLAInfoLocal = New MUSTER.Info.LustActivityInfo(0, pLustEvent.ID, Now, Nothing, Nothing, Nothing, Nothing, 697, strUserID, Now.ToString, String.Empty, CDate("01/01/1900"), 0, 0)
                                    pLustEventActivity.Add(oLAInfoLocal)
                                    pLustEventActivity.Save(ModuleID, staffID, returnVal)
                                    If Not returnVal = String.Empty Then
                                        Return False
                                    End If
                                    oClosureEventInfo.SentToTech = Today.Date
                                    oClosureEventInfo.TecID = pLustEventActivity.ActivityID
                                    oClosureEventInfo.TecType = 23
                                    Me.Save(ModuleID, staffID, returnVal)
                                    If Not returnVal = String.Empty Then
                                        Return False
                                    End If
                                ElseIf sampleLocation.ContainsKey(873) Then
                                    ' v
                                    ' (closure type = cip and noi received = no) or (closure type = rfg)
                                    ' create a tank closure activity for the lust event
                                    ' (new site created above or existing site)
                                    ' of type aerating backfill if location includes backfull
                                    oLAInfoLocal = New MUSTER.Info.LustActivityInfo(0, pLustEvent.ID, Now, Nothing, Nothing, Nothing, Nothing, 663, strUserID, Now.ToString, String.Empty, CDate("01/01/0001"), 0, 0)
                                    pLustEventActivity.Add(oLAInfoLocal)
                                    pLustEventActivity.Save(ModuleID, staffID, returnVal)
                                    If Not returnVal = String.Empty Then
                                        Return False
                                    End If
                                    oClosureEventInfo.SentToTech = Today.Date
                                    oClosureEventInfo.TecID = pLustEventActivity.ActivityID
                                    oClosureEventInfo.TecType = 23
                                    Me.Save(ModuleID, staffID, returnVal)
                                    If Not returnVal = String.Empty Then
                                        Return False
                                    End If
                                End If
                                ' create a To Do calendar entry on the current date for
                                ' the PM - Head user indicating
                                ' "New Activities created from Closure"
                                ocalInfo = New MUSTER.Info.CalendarInfo(0, _
                                            Now(), _
                                            Now(), _
                                            0, _
                                            "New Activities created from Closure : " + oClosureEventInfo.FacilitySequence.ToString + " belonging to Facility : " + oClosureEventInfo.FacilityID.ToString, _
                                            oUserInfo.ID, _
                                            "SYSTEM", _
                                            strGroupID, _
                                            False, _
                                            True, _
                                            False, _
                                            False, _
                                            String.Empty, _
                                            CDate("01/01/1900"), _
                                            String.Empty, _
                                            CDate("01/01/1900"), _
                                            22, _
                                            oClosureEventInfo.ID)
                                pCal.Add(ocalInfo)
                                pCal.Save()
                            End If

                        End If

                    End If

                    '''' g (a/h pg 12)
                    '''      remarked out   if not NOI received and closure is closed in place    '''
                    'If oClosureEventInfo.ClosureType = -445 And oClosureEventInfo.NOIReceived = 0 Then
                    ''' create citation in C & E
                    ''' manual fce with citation id 12 - got the citation id from kevin
                    'Dim oInspection As New MUSTER.BusinessLogic.pInspection
                    'Dim oFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
                    'Dim oInspectionCitation As New MUSTER.BusinessLogic.pInspectionCitation
                    'Dim nOwnerID As Integer = oFacilityInfo.OwnerID
                    'Try
                    '''' create inspection
                    'oInspection.Retrieve(0)
                    'oInspection.FacilityID = oClosureEventInfo.FacilityID
                    'ds = oClosureEventDB.DBGetDS("SELECT OWNER_ID FROM TBLREG_FACILITY WHERE FACILITY_ID = " + oClosureEventInfo.FacilityID.ToString)
                    'If ds.Tables(0).Rows(0)("OWNER_ID") Is DBNull.Value Then
                    'returnVal = "Invalid Owner ID for facility " + oClosureEventInfo.FacilityID.ToString
                    'Exit Function
                    'Else
                    '   nOwnerID = ds.Tables(0).Rows(0)("OWNER_ID")
                    'End If
                    'oInspection.OwnerID = nOwnerID
                    'oInspection.InspectionType = 1132
                    'oInspection.LetterGenerated = False
                    'oInspection.CreatedBy = oClosureEventInfo.ModifiedBy
                    'oInspection.Save(ModuleID, staffID, returnVal, , , , True)
                    'If Not returnVal = String.Empty Then
                    '  Exit Function
                    'End If
                    '''     create manual fce
                    'oFCE.Retrieve(0)
                    'oFCE.InspectionID = oInspection.ID
                    'oFCE.OwnerID = oInspection.OwnerID
                    'oFCE.FacilityID = oInspection.FacilityID
                    'oFCE.Source = "ADMIN"
                    'oFCE.FCEDate = Today.Date
                    'oInspection.CreatedBy = oInspection.CreatedBy
                    'oFCE.Save(ModuleID, staffID, returnVal, , , True)
                    'If Not returnVal = String.Empty Then
                    '   Exit Function
                    'End If
                    ''' create citation
                    'oInspectionCitation.Retrieve(oInspection.InspectionInfo, 0)
                    'oInspectionCitation.FacilityID = oFCE.FacilityID
                    'oInspectionCitation.FCEID = oFCE.ID
                    'oInspectionCitation.InspectionID = oInspection.ID
                    'oInspectionCitation.CitationID = 12
                    'oInspectionCitation.QuestionID = oInspection.CheckListMaster.RetrieveByCheckListItemNum("99998").ID
                    'oInspectionCitation.CreatedBy = oFCE.CreatedBy
                    'oInspectionCitation.Save(ModuleID, staffID, returnVal, , , True)
                    'If Not returnVal = String.Empty Then
                    ''''   Exit Function
                    'End If
                    'Catch ex As Exception
                    '           Throw ex
                    '          End Try
                    ' End If
                '
                If strMsg <> String.Empty Then
                    RaiseEvent evtClosureEventErr(strMsg)
                End If

                ' further functionalities are handles in the UI
                Return True
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub UpdateBoolCheckList(ByVal key As Integer, ByVal checked As Boolean)
            Try
                If oClosureEventInfo.HashTableBoolCheckList.Contains(key) Then
                    oClosureEventInfo.HashTableBoolCheckList.Item(key) = checked
                    oClosureEventInfo.CheckDirty()
                Else
                    oClosureEventInfo.HashTableBoolCheckList.Add(key, checked)
                    If Not oClosureEventInfo.HashTableDateCheckList.Contains(key) Then
                        oClosureEventInfo.HashTableDateCheckList.Add(key, CDate("01/01/1900"))
                    End If
                    oClosureEventInfo.CheckDirty()
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub UpdateDateCheckList(ByVal key As Integer, ByVal dt As Date)
            Try
                If oClosureEventInfo.HashTableDateCheckList.Contains(key) Then
                    oClosureEventInfo.HashTableDateCheckList.Item(key) = dt
                    oClosureEventInfo.CheckDirty()
                Else
                    oClosureEventInfo.HashTableDateCheckList.Add(key, dt)
                    If Not oClosureEventInfo.HashTableBoolCheckList.Contains(key) Then
                        oClosureEventInfo.HashTableBoolCheckList.Add(key, False)
                    End If
                    oClosureEventInfo.CheckDirty()
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub ClearChecklist()
            Try
                If oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                    oClosureEventInfo.HashTableBoolCheckList = New Hashtable
                Else
                    oClosureEventInfo.HashTableBoolCheckList.Clear()
                End If
                If oClosureEventInfo.HashTableDateCheckList Is Nothing Then
                    oClosureEventInfo.HashTableDateCheckList = New Hashtable
                Else
                    oClosureEventInfo.HashTableDateCheckList.Clear()
                End If
                oClosureEventInfo.CheckDirty()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub ManageChecklist(ByRef pCal As MUSTER.BusinessLogic.pCalendar, ByVal strUserID As String, ByVal dtDueDate As Date, ByVal dtNotificationDate As Date, Optional ByVal strTaskDesc As String = "", Optional ByVal bolToDo As Boolean = False, Optional ByVal bolDueToMe As Boolean = False, Optional ByVal strGroupID As String = "")
            Dim flags As New MUSTER.BusinessLogic.pFlag
            Dim flagsCol As MUSTER.Info.FlagsCollection
            Dim cal As New MUSTER.BusinessLogic.pCalendar
            Dim flagInfo As MUSTER.Info.FlagInfo
            Try
                Dim bolErr As Boolean = False
                If Date.Compare(oClosureEventInfo.DueDate, CDate("01/01/1900")) <> 0 Then
                    If Not oClosureEventInfo.HashTableBoolCheckList Is Nothing Then
                        For Each htEntry As DictionaryEntry In oClosureEventInfo.HashTableBoolCheckList
                            If htEntry.Value = True Then
                                If Date.Compare(oClosureEventInfo.HashTableDateCheckList.Item(htEntry.Key), CDate("01/01/1900")) = 0 Then
                                    bolErr = True
                                End If
                            End If
                        Next
                    End If
                    Dim colCal As New MUSTER.Info.CalendarCollection
                    Dim ocalInfo As MUSTER.Info.CalendarInfo
                    Dim bolExists As Boolean = False
                    If bolErr Then
                        colCal = pCal.RetrieveByOtherID(22, oClosureEventInfo.ID, strUserID, "USER")
                        For Each ocalInfo In colCal.Values
                            ' if exists
                            If ocalInfo.DueToMe Then
                                ocalInfo.NotificationDate = dtNotificationDate
                                ocalInfo.DateDue = dtDueDate
                                ocalInfo.CurrentColorCode = 0
                                ocalInfo.TaskDescription = strTaskDesc
                                ocalInfo.UserId = strUserID
                                ocalInfo.SourceUserId = "SYSTEM"
                                ocalInfo.GroupId = String.Empty
                                ocalInfo.DueToMe = bolDueToMe
                                ocalInfo.ToDo = bolToDo
                                ocalInfo.Completed = False
                                ocalInfo.Deleted = False
                                ocalInfo.OwningEntityType = 22
                                ocalInfo.OwningEntityID = oClosureEventInfo.ID
                                pCal.Add(ocalInfo)
                                pCal.Save()
                                bolExists = True
                                Exit For
                            End If
                            pCal.Add(ocalInfo)
                        Next
                        If Not bolExists Then
                            ' create a Due To Me calendar entry for
                            ' the Closure group on the Due Date
                            ' indicating Missing Information for the current Closure
                            ocalInfo = New MUSTER.Info.CalendarInfo(0, _
                                                dtNotificationDate, _
                                                dtDueDate, _
                                                0, _
                                                strTaskDesc, _
                                                strUserID, _
                                                "SYSTEM", _
                                                String.Empty, _
                                                bolDueToMe, _
                                                bolToDo, _
                                                False, _
                                                False, _
                                                strUserID, _
                                                Now(), _
                                                String.Empty, _
                                                CDate("01/01/1900"), _
                                                22, _
                                                oClosureEventInfo.ID)
                            pCal.Add(ocalInfo)
                            pCal.Save()
                        End If

                        ' Flag
                        flags.RetrieveFlags(oClosureEventInfo.FacilityID, 6, , , , , "SYSTEM", "Missing Information for Closure NOI # " + oClosureEventInfo.FacilitySequence.ToString)
                        If flags.FlagsCol.Count <= 0 Then
                            ' create flag
                            flags.Add(New MUSTER.Info.FlagInfo(0, _
                                 oClosureEventInfo.FacilityID, _
                                 6, _
                                 "Missing Information for Closure NOI # " + oClosureEventInfo.FacilitySequence.ToString + " on Facility ID " + oClosureEventInfo.FacilityID.ToString, _
                                 False, _
                                 DateAdd(DateInterval.Day, 30, Now.Date), _
                                 "CLOSURE", _
                                 0, _
                                 String.Empty, _
                                 CDate("01/01/1900"), _
                                 String.Empty, _
                                 CDate("01/01/1900"), _
                                 CDate("01/01/1900"), _
                                 "SYSTEM"))
                            flags.Save()
                            RaiseEvent FlagsChanged(oClosureEventInfo.FacilityID, 6, oClosureEventInfo.ID, 22)
                        End If
                    Else
                        ' Mark the associated due to me calendar entry as completed
                        colCal = pCal.RetrieveByOtherID(22, oClosureEventInfo.ID, strUserID, "USER")
                        For Each ocalInfo In colCal.Values
                            If ocalInfo.DueToMe Then
                                ocalInfo.Completed = True
                                pCal.Add(ocalInfo)
                                pCal.Save()
                            End If
                        Next

                        ' Delete associated Flag
                        flags.RetrieveFlags(oClosureEventInfo.FacilityID, 6, , , , , "SYSTEM", "Missing Information for Closure NOI # " + oClosureEventInfo.FacilitySequence.ToString)
                        For Each flagInfo In flags.FlagsCol.Values
                            flagInfo.Deleted = True
                        Next
                        If flags.FlagsCol.Count > 0 Then flags.Flush()
                        RaiseEvent FlagsChanged(oClosureEventInfo.FacilityID, 6, oClosureEventInfo.ID, 22)
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function DeleteClosure(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String) As Boolean
            Try
                If oClosureEventInfo.ID > 0 And (oClosureEventInfo.ClosureProcessed Or oClosureEventInfo.ClosureStatus = 868) Then
                    RaiseEvent evtClosureEventErr("Closure Event cannot be deleted after Processed")
                    Return False
                Else
                    oClosureEventInfo.Deleted = True
                    Return Me.Save(moduleID, staffID, returnVal, True, True)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetClosureSamples(Optional ByVal closureID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Try
                Return oClosureEventDB.DBGetClosureSamples(closureID, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub PutClosureSample(ByRef sampleID As Integer, _
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
            Try
                oClosureEventDB.DBPutClosureSamples(sampleID, _
                                                    closureID, _
                                                    sampleNumber, _
                                                    analysisType, _
                                                    analysisLevel, _
                                                    sampleMedia, _
                                                    sampleLocation, _
                                                    sampleValue, _
                                                    sampleUnits, _
                                                    sampleConstituent, _
                                                    deleted, _
                                                    usedID, _
                                                    moduleID, _
                                                    staffID, _
                                                    returnVal)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetClosureTankPipeList(ByVal closureID As Integer, Optional ByVal showDeleted As Boolean = False) As String
            Try
                Return oClosureEventDB.DBGetTankPipeList(closureID, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub PutClosureTankPipe(ByRef cloTankPipeID As Integer, _
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
            Try
                oClosureEventDB.DBPutTankPipe(cloTankPipeID, _
                                                tankPipeID, _
                                                tankPipeEntity, _
                                                closureID, _
                                                deleted, _
                                                moduleID, _
                                                staffID, _
                                                returnVal, _
                                                UserID, _
                                                analysisType, _
                                                analysisLevel, _
                                                sampleMedia, _
                                                sampleResultsID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Function CheckIfFacHasSubmittedInspection() As String
            Try
                Dim ds As DataSet
                Dim strSQL As String
                strSQL = "SELECT * FROM tblINS_INSPECTION WHERE SCHEDULED_BY IS NOT NULL AND DELETED = 0 " + _
                            "AND SUBMITTED_DATE IS NOT NULL AND COMPLETED IS NULL AND FACILITY_ID = " + oClosureEventInfo.FacilityID.ToString
                ds = oClosureEventDB.DBGetDS(strSQL)
                If ds.Tables(0).Rows.Count > 0 Then
                    Return "Facility has Submitted inspection which is not processed by C&E." + vbCrLf + "Cannot Process Closure Event"
                Else
                    ' check if facility is part of an open oce
                    ' check if non rescinded citations belonging to the facility are part of an open oce

                    strSQL = "sp_FacilityCitationsNotComplliantToClosure"


                    ds = oClosureEventDB.DBGetDS(String.Format("exec {0} {1}", strSQL, oClosureEventInfo.FacilityID))

                    If ds.Tables(0).Rows.Count > 0 Then
                        Return "This facility is part of an OCE that has a citation that has not been received."
                    Else
                        Return String.Empty
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Public Sub Add(ByRef facInfo As MUSTER.Info.FacilityInfo, ByVal id As Integer, Optional ByVal showDeleted As Boolean = False)
            Try
                oClosureEventInfo = oClosureEventDB.DBGetByID(id, showDeleted)
                If oClosureEventInfo.ID = 0 Then
                    oClosureEventInfo.FacilityID = facInfo.ID
                    oClosureEventInfo.ID = nID
                    nID -= 1
                End If
                facInfo.ClosureEventCollection.Add(oClosureEventInfo)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Add(ByRef facInfo As MUSTER.Info.FacilityInfo, ByRef oClosureEvent As MUSTER.Info.ClosureEventInfo)
            Try
                oClosureEventInfo = oClosureEvent
                If oClosureEventInfo.ID = 0 Then
                    oClosureEventInfo.FacilityID = facInfo.ID
                    oClosureEventInfo.ID = nID
                    nID -= 1
                End If
                facInfo.ClosureEventCollection.Add(oClosureEventInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Remove(ByVal id As Integer)
            Dim oClosureEventInfoLocal As MUSTER.Info.ClosureEventInfo
            Try
                oClosureEventInfo = oFacilityInfo.ClosureEventCollection.Item(id)
                If Not (oClosureEventInfoLocal Is Nothing) Then
                    oFacilityInfo.ClosureEventCollection.Remove(oClosureEventInfoLocal)
                    Exit Sub
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            'Throw New Exception("Closure Event " & id.ToString & " is not in the collection of Closure Events.")
        End Sub
        Public Sub Remove(ByVal oClosureEventInf As MUSTER.Info.ClosureEventInfo)
            Try
                oFacilityInfo.ClosureEventCollection.Remove(oClosureEventInf)
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim IDs As New Collection
            Dim delIDs As New Collection
            Dim index As Integer
            Dim oTempInfo As MUSTER.Info.ClosureEventInfo
            Try
                For Each oTempInfo In oFacilityInfo.ClosureEventCollection.Values
                    If oTempInfo.IsDirty Then
                        oClosureEventInfo = oTempInfo
                        If oClosureEventInfo.ID < 0 Then
                            If oClosureEventInfo.Deleted Then
                                delIDs.Add(oClosureEventInfo.ID)
                            Else
                                IDs.Add(oClosureEventInfo.ID)
                            End If
                        Else
                            If oClosureEventInfo.Deleted Then
                                delIDs.Add(oClosureEventInfo.ID)
                            End If
                            Me.Save(moduleID, staffID, returnVal, True)
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        oTempInfo = oFacilityInfo.ClosureEventCollection.Item(CType(delIDs.Item(index), String))
                        oFacilityInfo.ClosureEventCollection.Remove(oTempInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    Dim sortedList As New sortedList
                    For index = 0 To IDs.Count - 1
                        Dim str As String = CType(IDs(index + 1), String)
                        sortedList.Add(CType(IDs(index + 1), Integer), CType(IDs(index + 1), Integer))
                    Next
                    For index = sortedList.Count - 1 To 0 Step -1
                        oClosureEventInfo = oFacilityInfo.ClosureEventCollection.Item(CType(sortedList.GetByIndex(index), String))
                        Me.Save(moduleID, staffID, returnVal, True)
                    Next
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        oTempInfo = oFacilityInfo.ClosureEventCollection.Item(colKey)
                        oFacilityInfo.ClosureEventCollection.ChangeKey(colKey, oTempInfo.ID)
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub CancelAllOverdue(ByRef facInfo As MUSTER.Info.FacilityInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            oFacilityInfo = facInfo
            Dim cloInfoLocal As MUSTER.Info.ClosureEventInfo
            Dim bolCancelAllOverdue As Boolean = False
            Try
                ' if closure collection is empty, retrieve from db
                If oFacilityInfo.ClosureEventCollection.Count <= 0 Then
                    oFacilityInfo.ClosureEventCollection = oClosureEventDB.DBGetByFacID(oFacilityInfo.ID)
                End If
                For Each cloInfoLocal In oFacilityInfo.ClosureEventCollection.Values
                    If cloInfoLocal.ClosureStatus <> 868 Then ' not closed
                        ' scheduled date < current date - 90 days
                        If Date.Compare(cloInfoLocal.ScheduledDate, DateAdd(DateInterval.Day, -90, Today())) < 0 Then
                            cloInfoLocal.ClosureStatus = 869
                            Me.Save(moduleID, staffID, returnVal)
                            bolCancelAllOverdue = True
                        Else
                            Dim str As String = String.Empty
                            str = "Do you want to cancel Closure Event : " + _
                                cloInfoLocal.FacilitySequence.ToString + _
                                " on Facility ID : " + cloInfoLocal.FacilityID.ToString + vbCrLf
                            str += "Status : " + oProperty.GetPropertyNameByID(cloInfoLocal.ClosureStatus) + vbCrLf
                            str += "Type : " + oProperty.GetPropertyNameByID(cloInfoLocal.ClosureType) + vbCrLf
                            str += "Scheduled Date : " + cloInfoLocal.ScheduledDate.ToShortDateString
                            If MsgBox(str, MsgBoxStyle.YesNo, "Cancel Closure Event") = MsgBoxResult.Yes Then
                                cloInfoLocal.ClosureStatus = 869
                                Me.Save(moduleID, staffID, returnVal)
                                bolCancelAllOverdue = True
                            End If
                        End If
                    End If
                Next
                Me.Flush(moduleID, staffID, returnVal)
                If bolCancelAllOverdue Then
                    MsgBox("Cancelled overdue Closure Events")
                Else
                    MsgBox("There are no overdue Closure Events to cancel")
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub ReopenCancelled(ByVal closureID As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                If oClosureEventInfo.ID <> closureID Then
                    oClosureEventInfo = oFacilityInfo.ClosureEventCollection.Item(closureID)
                End If
                If Not (oClosureEventInfo Is Nothing) Then
                    If oClosureEventInfo.ClosureStatus = 869 Then
                        If oClosureEventInfo.OldClosureStatus <> 869 Then
                            ' check value in collection
                            oClosureEventInfo.ClosureStatus = oClosureEventInfo.OldClosureStatus
                        Else
                            ' get from db
                            Dim prevStat As Integer
                            prevStat = oClosureEventDB.GetPreviousStatus(oClosureEventInfo.ID, oClosureEventInfo.ClosureStatus)
                            If oClosureEventInfo.ClosureStatus <> prevStat Then
                                oClosureEventInfo.ClosureStatus = prevStat
                                Me.Save(moduleID, staffID, returnVal)
                            Else
                                RaiseEvent evtClosureEventErr("No Previous Status found")
                            End If
                        End If
                    Else
                        RaiseEvent evtClosureEventErr("Cannot Reopen Closure Event which are not Cancelled")
                    End If
                Else
                    ' if not in collection, retrieve from db
                    Add(oFacilityInfo, closureID)
                    ReopenCancelled(oClosureEventInfo.ID, moduleID, staffID, returnVal)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oClosureEventInfo = New MUSTER.Info.ClosureEventInfo
        End Sub
        Public Sub Reset()
            oClosureEventInfo.Reset()
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = oFacilityInfo.ClosureEventCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 Then
                If colIndex + direction <= nArr.GetUpperBound(0) Then
                    Return oFacilityInfo.ClosureEventCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
                Else
                    Return oFacilityInfo.ClosureEventCollection.Item(nArr.GetValue(0)).ID.ToString
                End If
            Else
                Return oFacilityInfo.ClosureEventCollection.Item(nArr.GetValue(nArr.GetUpperBound(0))).ID.ToString
            End If
        End Function
#End Region
#Region "Lookup Operations"
        Public Sub CheckNOIReceived()
            Try
                If oClosureEventInfo.NOIReceived = 0 Then
                    oClosureEventInfo.NOI_Rcv_Date = CDate("01/01/1900")
                    oClosureEventInfo.ScheduledDate = CDate("01/01/1900")
                    RaiseEvent evtClosureNOIReceived(False)
                Else
                    If Date.Compare(oClosureEventInfo.NOI_Rcv_Date, CDate("01/01/1900")) = 0 Then
                        oClosureEventInfo.NOI_Rcv_Date = Today.Date
                    End If
                    If Date.Compare(oClosureEventInfo.ScheduledDate, CDate("01/01/1900")) = 0 Then
                        oClosureEventInfo.ScheduledDate = Today.AddDays(30).Date
                    End If
                    RaiseEvent evtClosureNOIReceived(True)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function PopulateClosureStatus() As DataTable
            Return GetDataTable("vCLOSURESTATUS")
        End Function
        Public Function PopulateClosureType() As DataTable
            Return GetDataTable("vCLOSURETYPE")
        End Function
        Public Function PopulateFillMaterial() As DataTable
            Return GetDataTable_FillMaterial("vINNERTMATERIAL", oClosureEventInfo.ClosureType)
        End Function
        Public Function PopulateSampleMedia() As DataTable
            Return GetDataTable("vCLOSURESAMPLEMEDIA")
        End Function
        Public Function PopulateAnalysisType() As DataTable
            Return GetDataTable("vCLOSUREANALYSISTYPE")
        End Function
        Public Function PopulateAnalysisLevel(Optional ByVal analysisType As Integer = 0) As DataTable
            Return GetDataTable_AnalysisLevel("vCLOSUREANALYSISLEVEL", analysisType)
        End Function
        Public Function PopulateSampleLocation(Optional ByVal closureType As Integer = 0, Optional ByVal analysisLevel As Integer = 0) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            ' Closure Type
            ' 443 - Change in Service
            ' 444 - Removed from Ground
            ' 445 - Closed in Place

            ' Analysis Level
            ' 880 - AAL - BTEX
            ' 881 - AAL - PAH
            ' 882 - AAL - BTEX & PAH
            ' 884 - ADL
            ' 885 - BDL
            ' 890 - AAL

            ' Location
            ' 871 - Tank
            ' 872 - Piping Trench
            ' 873 - Backfill
            ' 874 - Pump Island
            Try
                ' #836
                If analysisLevel > 0 Then
                    strSQL = "SELECT * FROM vCLOSURELOCATION where property_id_parent = " + analysisLevel.ToString
                    If analysisLevel = 890 Then
                        If closureType = 445 Then ' CIP
                            strSQL += " and PROPERTY_ID <> 873"
                        End If
                    End If

                    dsReturn = oClosureEventDB.DBGetDS(strSQL)
                    If dsReturn.Tables(0).Rows.Count > 0 Then
                        dtReturn = dsReturn.Tables(0)
                    End If
                End If
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateAnalysisUnits()
            Return GetDataTable("vCLOSUREANALYSISUNITS")
        End Function
        Public Function PopulateCheckListItems(ByVal closureType As Integer) As Hashtable
            Dim ds As DataSet
            Dim dr As DataRow
            Dim ht As New Hashtable
            Try
                If closureType = 444 Or closureType = 445 Then
                    ds = oClosureEventDB.DBGetCheckListItems(closureType)
                    For Each dr In ds.Tables(0).Rows
                        ht.Add(dr("ITEM_ID"), dr("ITEM_TEXT"))
                    Next
                End If
                Return ht
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub PopulateCheckList(ByVal ds As DataSet)
            Dim dr As DataRow
            Try
                If ds.Tables(0).Rows.Count > 0 Then
                    ClearChecklist()
                    For Each dr In ds.Tables(0).Rows
                        oClosureEventInfo.HashTableBoolCheckList.Add(dr("ITEM_ID"), dr("ENTRY_OPEN"))
                        oClosureEventInfo.HashTableDateCheckList.Add(dr("ITEM_ID"), IIf(dr("DATE_CLOSED") Is DBNull.Value, CDate("01/01/1900"), dr("DATE_CLOSED")))
                        oClosureEventInfo.HashTableBoolCheckListOriginal.Add(dr("ITEM_ID"), dr("ENTRY_OPEN"))
                        oClosureEventInfo.HashTableDateCheckListOriginal.Add(dr("ITEM_ID"), IIf(dr("DATE_CLOSED") Is DBNull.Value, CDate("01/01/1900"), dr("DATE_CLOSED")))
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub PopulatechkBoxDate()
            Try
                Dim ds As DataSet
                Dim dr As DataRow
                Dim i As Integer = 0
                ds = oClosureEventDB.DBGetCheckList(oClosureEventInfo.ID, oClosureEventInfo.NOIProcessed)
                ClearChecklist()
                For Each dr In ds.Tables(0).Rows
                    oClosureEventInfo.HashTableBoolCheckList.Add(dr("ITEM_ID"), dr("ENTRY_OPEN"))
                    oClosureEventInfo.HashTableDateCheckList.Add(dr("ITEM_ID"), IIf(dr("DATE_CLOSED") Is DBNull.Value, CDate("01/01/1900"), dr("DATE_CLOSED")))
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub PopulateSampleTable(ByVal closureID As Integer)
            Dim ds As DataSet
            Dim drNew As DataRow
            Try
                ds = GetClosureSamples(closureID)
                SamplesTable.Rows.Clear()
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr As DataRow In ds.Tables(0).Rows
                            drNew = SamplesTable.NewRow
                            For Each col As DataColumn In ds.Tables(0).Columns
                                drNew(col.ColumnName) = dr(col.ColumnName)
                            Next
                            SamplesTable.Rows.Add(drNew)
                        Next
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal nVal As Integer = 0, Optional ByVal bolDistinct As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As New DataTable
            Dim strSQL As String
            If bolDistinct Then
                strSQL = "SELECT DISTINCT PROPERTY_ID, PROPERTY_NAME FROM " + strProperty
            Else
                strSQL = "SELECT * FROM " & strProperty
            End If
            If nVal <> 0 Then
                strSQL = strSQL + " WHERE PROPERTY_ID_PARENT = " + nVal.ToString()
            End If
            Try
                dsReturn = oClosureEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDataTable_FillMaterial(ByVal strProperty As String, Optional ByVal closureType As Integer = 0) As DataTable
            ' Fill Material
            '131 - Concrete
            '132 - Sand
            '133 - Water
            '134 - Other approved fill
            '135 - None
            '136 - Not Listed
            '137 - Drill Mud
            '138 - Grout
            '139 - Pea Gravel
            '140 - Sand/Slurry
            '141 - Drilling Mud
            '142 - Slurry
            '143 - Virgin Drilling
            '144 - Virgin Mud
            '883 - Approved("foam")

            ' Closure Type
            ' 443 - Change in Service
            ' 444 - Removed from Ground
            ' 445 - Closed in Place
            ' "PROPERTY_ID <> 139 AND " + _ Removed on May 1, 2008
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                If closureType = 445 Then
                    strSQL = "SELECT * FROM " & strProperty
                    strSQL += " WHERE PROPERTY_ID <> 133 AND " + _
                                "PROPERTY_ID <> 134 AND " + _
                                "PROPERTY_ID <> 135 AND " + _
                                "PROPERTY_ID <> 136 AND " + _
                                "PROPERTY_ID <> 137 AND " + _
                                "PROPERTY_ID <> 138 AND " + _
                                "PROPERTY_ID <> 140 AND " + _
                                "PROPERTY_ID <> 141 AND " + _
                                "PROPERTY_ID <> 142 AND " + _
                                "PROPERTY_ID <> 144 "
                Else
                    strSQL = "SELECT * FROM " & strProperty
                End If
                dsReturn = oClosureEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDataTable_AnalysisLevel(ByVal strProperty As String, ByVal analysisType As Integer) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                If analysisType > 0 Then
                    strSQL = "SELECT * FROM " & strProperty
                    strSQL += " WHERE PROPERTY_ID_PARENT = '" + analysisType.ToString + "'"
                    dsReturn = oClosureEventDB.DBGetDS(strSQL)
                    If dsReturn.Tables(0).Rows.Count > 0 Then
                        dtReturn = dsReturn.Tables(0)
                    End If
                End If
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateCertifiedContractor() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable_Company("vCOM_LICENSEENAME")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCompanyName(Optional ByVal LicenseeID As Integer = 0) As DataSet
            Dim dsReturn As New DataSet

            Try
                dsReturn = oClosureEventDB.DBGetCompanyDetails(LicenseeID)
                If Not dsReturn.Tables(0).Rows.Count > 0 Then
                    dsReturn = Nothing
                End If
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function GetDataTable_Company(ByVal DBViewName As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName

                dsReturn = oClosureEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                Else
                    dtReturn = Nothing
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oClosureEventInfoLocal As New MUSTER.Info.ClosureEventInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("ID")
                tbEntityTable.Columns.Add("Facility ID")
                tbEntityTable.Columns.Add("Facility Sequence")
                tbEntityTable.Columns.Add("Closure Type")
                tbEntityTable.Columns.Add("Closure Status")
                tbEntityTable.Columns.Add("NOI Received")
                tbEntityTable.Columns.Add("NOI Received Date")
                tbEntityTable.Columns.Add("Owner Sign")
                tbEntityTable.Columns.Add("Scheduled Date")
                tbEntityTable.Columns.Add("Certified Contractor")
                tbEntityTable.Columns.Add("Fill Material")
                tbEntityTable.Columns.Add("Company")
                tbEntityTable.Columns.Add("Contact")
                tbEntityTable.Columns.Add("Verbal Waiver")
                tbEntityTable.Columns.Add("NOI Processed")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Created On")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Last Edited On")
                For Each oClosureEventInfoLocal In oFacilityInfo.ClosureEventCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("ID") = oClosureEventInfoLocal.ID
                    dr("Facility ID") = oClosureEventInfoLocal.FacilityID
                    dr("Facility Sequence") = oClosureEventInfoLocal.FacilitySequence
                    dr("Closure Type") = oClosureEventInfoLocal.ClosureType
                    dr("Closure Status") = oClosureEventInfoLocal.ClosureStatus
                    dr("NOI Received") = oClosureEventInfoLocal.NOIReceived
                    dr("NOI Received Date") = oClosureEventInfoLocal.NOI_Rcv_Date
                    dr("Owner Sign") = oClosureEventInfoLocal.OwnerSign
                    dr("Scheduled Date") = oClosureEventInfoLocal.ScheduledDate
                    dr("Certified Contractor") = oClosureEventInfoLocal.CertContractor
                    dr("Fill Material") = oClosureEventInfoLocal.FillMaterial
                    dr("Company") = oClosureEventInfoLocal.Company
                    dr("Contact") = oClosureEventInfoLocal.Contact
                    dr("Verbal Waiver") = oClosureEventInfoLocal.VerbalWaiver
                    dr("NOI Processed") = oClosureEventInfoLocal.NOIProcessed
                    dr("Deleted") = oClosureEventInfoLocal.Deleted
                    dr("Created By") = oClosureEventInfoLocal.CreatedBy
                    dr("Created On") = oClosureEventInfoLocal.CreatedOn
                    dr("Last Edited By") = oClosureEventInfoLocal.ModifiedBy
                    dr("Last Edited On") = oClosureEventInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function SampleResultMemo() As DataTable
            Dim dr As DataRow
            Dim dtSampleResultMemo As New DataTable
            Try
                dtSampleResultMemo.Columns.Add("Sample#")
                dtSampleResultMemo.Columns.Add("Location")
                dtSampleResultMemo.Columns.Add("Results")
                dtSampleResultMemo.Columns.Add("Units")
                dtSampleResultMemo.Columns.Add("Constituent")

                For Each drow As DataRow In SamplesTable.Rows
                    dr = dtSampleResultMemo.NewRow
                    dr("Sample#") = drow("Sample #").ToString
                    If drow("Sample Location") Is System.DBNull.Value Then
                        dr("Location") = String.Empty
                    Else
                        dr("Location") = oProperty.GetPropertyNameByID(CType(drow("Sample Location"), Integer))
                    End If
                    dr("Results") = IIf(drow("Sample Value") Is System.DBNull.Value, String.Empty, drow("Sample Value").ToString)
                    If drow("Sample Units") Is System.DBNull.Value Then
                        dr("Units") = String.Empty
                    Else
                        dr("Units") = oProperty.GetPropertyNameByID(drow("Sample Units").ToString)
                    End If
                    If drow("Sample Constituent") Is System.DBNull.Value Then
                        dr("Constituent") = String.Empty
                    ElseIf drow("Sample Constituent") = String.Empty Then
                        dr("Constituent") = String.Empty
                    ElseIf drow("Sample Constituent") = "0" Then
                        dr("Constituent") = String.Empty
                    Else
                        dr("Constituent") = drow("Sample Constituent").ToString
                    End If
                    dtSampleResultMemo.Rows.Add(dr)
                Next
                Return dtSampleResultMemo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function ClosureTankPipeDataSet(ByVal facID As Integer, ByVal closureID As Integer, Optional ByVal showDeleted As Boolean = False) As DataSet
            Dim dsTankPipe As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim dsRel2 As DataRelation
            Dim strSQL As String
            Try
                dsTankPipe = oClosureEventDB.DBGetClosureTankPipe(facID, closureID, showDeleted)

                For Each oCol In dsTankPipe.Tables(0).Columns
                    If oCol.ColumnName <> "INCLUDED" And oCol.ColumnName <> "SUBSTANCE" Then
                        oCol.ReadOnly = True
                    End If
                Next
                For Each oCol In dsTankPipe.Tables(1).Columns
                    If oCol.ColumnName <> "INCLUDED" And oCol.ColumnName <> "SUBSTANCE" Then
                        oCol.ReadOnly = True
                    End If
                Next

                'dsTankPipe.Tables(0).DefaultView.Sort = "POSITION, TANK SITE ID"
                'dsTankPipe.Tables(1).DefaultView.Sort = "POSITION, PIPE SITE ID"

                Dim c1() As DataColumn = {dsTankPipe.Tables(1).Columns("TANK_ID"), dsTankPipe.Tables(1).Columns("PIPE_ID")}
                Dim c2() As DataColumn = {dsTankPipe.Tables(2).Columns("TANK_ID"), dsTankPipe.Tables(2).Columns("Parent_Pipe_ID")}


                dsRel = New DataRelation("TankToPipe", dsTankPipe.Tables(0).Columns("TANK_ID"), dsTankPipe.Tables(1).Columns("TANK_ID"), False)
                dsRel2 = New DataRelation("PipeToExt", c1, c2, False)


                dsTankPipe.Relations.Add(dsRel)
                dsTankPipe.Relations.Add(dsRel2)
                Return dsTankPipe
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "Event Handlers"
        Public Sub ClosureEventInfoChanged(ByVal bolValue As Boolean) Handles oClosureEventInfo.evtClosureEventInfoChanged
            RaiseEvent evtClosureEventInfoChanged(bolValue)
        End Sub
        'Public Sub CommentColClosureEvent(ByVal closureID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oComments.evtCommentColClosureEvent
        '    ' to be implemented
        'End Sub
#End Region
    End Class
End Namespace
