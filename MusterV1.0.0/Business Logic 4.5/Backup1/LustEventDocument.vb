'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.LustEventActivity
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0         AN       3/8/2005    Original class definition
'
' Function          Description
' Retrieve(Name)    Returns the Lust Event Activity requested by the string arg NAME
' Retrieve(ID)      Returns the Lust Event Activity requested by the int arg ID
' GetAll()          Returns an LustEventActivitiyCollection with all Lust Event Activity objects
' Add(ID)           Adds the Lust Event Activity identified by arg ID to the 
'                           internal LustEventActivityCollection
' Add(LustEventInfo)Adds the Lust Event Activity passed as the argument to the internal 
'                           LustEventActivityCollection
' Remove(ID)        Removes the Lust Event Activity identified by arg ID from the internal 
'                           LustEventActivityCollection
' Remove(NAME)      Removes the Lust Event Activity identified by arg NAME from the 
'                           internal LustEventActivityCollection
' Flush()           Saves all objects in the collection
' Clear()           Clears the current object and all objects in the collection
' Reset()           Resets the current object to its original state
' EntityTable()     Returns a datatable containing all columns for the Lust Event 
'                           objects in the internal LustEventsCollection.
'
'
' Attribute          Description
'-------------------------------------------------------------------------------
' AssocActivity     the Lust Activity that the document is associated with.
' EventId           Event that the Document is Associated to
' Comments          The collection of comments associated with the LUST event
' Deleted           Indicates the deleted state of the row
' DocClass          The "class" of the document associated with the LUST activity.  This is the "Document" as defined in the Technical DDD p 21.  Will be drawn from the tblProperty_Master table
' DocClosedDate     The date the document is closed.
' DocFinancialDate  The date the document was "sent to financial"
' DocRcvDate        The date the document was received.
' DocRevisionsDue   The date the revisions for the document are due.
' DocumentID        The system generated ID for the LUST document
' DocumentType      The document type of the LUST document
' DueDate           The date the task or reminder is to be completed by
' EntityID          The entity ID associated with a technical document.
' ID                The system ID for this LUST event
' IssueDate         The date the document was issued.
' IsAgedData        Returns a boolean indicating if the data has aged beyond its preset limit
'
' AgeThreshold       The maximum age the info object can attain before requiring a refresh
' CreatedBy          The ID of the user that created the row
' CreatedOn          The date on which the row was created
' Deleted            Indicates the deleted state of the row
' ModifiedBy         ID of the user that last made changes
' ModifiedOn         The date of the last changes made 
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLustEventDocument
#Region "Public Events"
        Public Event LustEventErr(ByVal MsgStr As String)
        Public Event LustEventChanged(ByVal bolValue As Boolean)
        Public Event LustEventStatusChange()
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oLustEventDocumentInfo As New MUSTER.Info.LustDocumentInfo
        'Private WithEvents colLustEventDocuments As MUSTER.Info.LustDocumentCollection
        Private WithEvents oLustActivityInfo As MUSTER.Info.LustActivityInfo
        Private oLustEventDocumentDB As New MUSTER.DataAccess.LustEventDocumentsDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private oCalendar As New MUSTER.BusinessLogic.pCalendar
        Private oLetterGen As New MUSTER.BusinessLogic.pLetterGen
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("LustDocument").ID
        Private onFacilityID As Integer = 0
        Private onUserID As Integer = 0
#End Region
#Region "Constructors"
        Public Sub New(Optional ByRef DocumentActivity As MUSTER.Info.LustActivityInfo = Nothing)
            oLustEventDocumentInfo = New MUSTER.Info.LustDocumentInfo
            'colLustEventDocuments = New MUSTER.Info.LustDocumentCollection
            If DocumentActivity Is Nothing Then
                oLustActivityInfo = New MUSTER.Info.LustActivityInfo
            Else
                oLustActivityInfo = DocumentActivity
                onFacilityID = DocumentActivity.FacilityID
                onUserID = DocumentActivity.UserID
            End If
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the LustEvent object with the provided ID.
        '
        '********************************************************
        Public Sub New(ByVal LustEventID As Integer, Optional ByRef DocumentActivity As MUSTER.Info.LustActivityInfo = Nothing)
            oLustEventDocumentInfo = New MUSTER.Info.LustDocumentInfo
            'colLustEventDocuments = New MUSTER.Info.LustDocumentCollection
            If DocumentActivity Is Nothing Then
                oLustActivityInfo = New MUSTER.Info.LustActivityInfo
            Else
                oLustActivityInfo = DocumentActivity
                onFacilityID = DocumentActivity.FacilityID
                onUserID = DocumentActivity.UserID
            End If
            Me.Retrieve(LustEventID)
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named LustEvent object.
        '
        '********************************************************
        Public Sub New(ByVal LustEventName As String, Optional ByRef DocumentActivity As MUSTER.Info.LustActivityInfo = Nothing)
            oLustEventDocumentInfo = New MUSTER.Info.LustDocumentInfo
            'colLustEventDocuments = New MUSTER.Info.LustDocumentCollection
            If DocumentActivity Is Nothing Then
                oLustActivityInfo = New MUSTER.Info.LustActivityInfo
            Else
                oLustActivityInfo = DocumentActivity
                onFacilityID = DocumentActivity.FacilityID
                onUserID = DocumentActivity.UserID
            End If
            Me.Retrieve(LustEventName)
        End Sub
#End Region
#Region "Exposed Attributes"
        ' the Lust Activity that the document is associated with.
        Public Property AssocActivity() As Long
            Get
                Return oLustEventDocumentInfo.AssocActivity
            End Get
            Set(ByVal Value As Long)
                oLustEventDocumentInfo.AssocActivity = Value
            End Set
        End Property
        'Event that the Document is Associated to
        Public Property EventId() As Long
            Get
                Return oLustEventDocumentInfo.EventID
            End Get
            Set(ByVal Value As Long)
                oLustEventDocumentInfo.EventID = Value
            End Set
        End Property
        ' The collection of comments associated with the LUST event
        Public Property Comments() As Object
            Get
                Return oLustEventDocumentInfo.Comments
            End Get
            Set(ByVal Value As Object)
                oLustEventDocumentInfo.Comments = Value
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oLustEventDocumentInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oLustEventDocumentInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLustEventDocumentInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oLustEventDocumentInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oLustEventDocumentInfo.Deleted = Value
            End Set
        End Property
        ' The "class" of the document associated with the LUST activity.  This is the "Document" as defined in the Technical DDD p 21.  Will be drawn from the tblProperty_Master table
        Public Property DocClass() As Long
            Get
                Return oLustEventDocumentInfo.DocClass
            End Get
            Set(ByVal Value As Long)
                oLustEventDocumentInfo.DocClass = Value
            End Set
        End Property
        ' The date the document is closed.
        Public Property DocClosedDate() As Date
            Get
                Return oLustEventDocumentInfo.DocClosedDate
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.DocClosedDate = Value
            End Set
        End Property
        ' The date the document was "sent to financial"
        Public Property DocFinancialDate() As Date
            Get
                Return oLustEventDocumentInfo.DocFinancialDate
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.DocFinancialDate = Value
            End Set
        End Property
        ' The date the document was received.
        Public Property DocRcvDate() As Date
            Get
                Return oLustEventDocumentInfo.DocRcvDate
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.DocRcvDate = Value
            End Set
        End Property
        ' The date the revisions for the document are due.
        Public Property DocRevisionsDue() As Date
            Get
                Return oLustEventDocumentInfo.DocRevisionsDue
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.DocRevisionsDue = Value
            End Set
        End Property
        ' The system generated ID for the LUST document
        Public Property DocumentID() As Long
            Get
                Return oLustEventDocumentInfo.DocumentID
            End Get
            Set(ByVal Value As Long)
                oLustEventDocumentInfo.DocumentID = Value
            End Set
        End Property
        ' The document type of the LUST document
        Public Property DocumentType() As Long
            Get
                Return oLustEventDocumentInfo.DocumentType
            End Get
            Set(ByVal Value As Long)
                oLustEventDocumentInfo.DocumentType = Value
            End Set
        End Property
        ' The date the task or reminder is to be completed by
        Public Property DueDate() As Date
            Get
                Return oLustEventDocumentInfo.DueDate
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.DueDate = Value
            End Set
        End Property
        ' The entity ID associated with a technical document.
        'Public ReadOnly Property EntityID() As Integer
        '    Get
        '        Return nEntityTypeID
        '        'Return oLustEventDocumentInfo.EntityID
        '    End Get
        'End Property
        ' The system ID for this LUST event
        Public ReadOnly Property ID() As Long
            Get
                Return oLustEventDocumentInfo.ID
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oLustEventDocumentInfo.IsDirty = Value
            End Set
        End Property
        ' The date the document was issued.
        Public Property IssueDate() As Date
            Get
                Return oLustEventDocumentInfo.IssueDate
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.IssueDate = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLustEventDocumentInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oLustEventDocumentInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLustEventDocumentInfo.ModifiedOn
            End Get
        End Property

        Public Property STARTDATE() As Date
            Get
                Return oLustEventDocumentInfo.STARTDATE
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.STARTDATE = Value
            End Set
        End Property

        Public Property EXTENSIONDATE() As Date
            Get
                Return oLustEventDocumentInfo.EXTENSIONDATE
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.EXTENSIONDATE = Value
            End Set
        End Property

        Public Property REV1EXTENSIONDATE() As Date
            Get
                Return oLustEventDocumentInfo.REV1EXTENSIONDATE
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.REV1EXTENSIONDATE = Value
            End Set
        End Property

        Public Property REV1RECEIVEDDATE() As Date
            Get
                Return oLustEventDocumentInfo.REV1RECEIVEDDATE
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.REV1RECEIVEDDATE = Value
            End Set
        End Property

        Public Property REV2RECEIVEDDATE() As Date
            Get
                Return oLustEventDocumentInfo.REV2RECEIVEDDATE
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.REV2RECEIVEDDATE = Value
            End Set
        End Property

        Public Property REV2EXTENSIONDATE() As Date
            Get
                Return oLustEventDocumentInfo.REV2EXTENSIONDATE
            End Get
            Set(ByVal Value As Date)
                oLustEventDocumentInfo.REV2EXTENSIONDATE = Value
            End Set
        End Property

        Public Property FacilityID() As Integer
            Get
                Return oLustEventDocumentInfo.FacilityID
            End Get
            Set(ByVal Value As Integer)
                oLustEventDocumentInfo.FacilityID = Value
            End Set
        End Property

        Public Property UserID() As Integer
            Get
                Return oLustEventDocumentInfo.UserID
            End Get
            Set(ByVal Value As Integer)
                oLustEventDocumentInfo.UserID = Value
            End Set
        End Property
        Public Property CommitmentID() As Long
            Get
                Return oLustEventDocumentInfo.CommitmentId
            End Get
            Set(ByVal Value As Long)
                oLustEventDocumentInfo.CommitmentId = Value
            End Set
        End Property
        Public Property Paid() As Boolean
            Get
                Return oLustEventDocumentInfo.Paid
            End Get
            Set(ByVal Value As Boolean)
                oLustEventDocumentInfo.Paid = Value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xLustEventinfo As MUSTER.Info.LustEventInfo
                For Each xLustEventinfo In oLustActivityInfo.Documents.Values 'colLustEventDocuments.Values
                    If xLustEventinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oLustEventDocumentInfo.IsDirty = Value
            End Set
        End Property


        Public ReadOnly Property IsDirtySentToFinancial() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtySentToFinancial
            End Get
        End Property
        Public ReadOnly Property IsDirtyClosedDate() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyClosedDate
            End Get
        End Property
        Public ReadOnly Property IsDirtyRecievedDate() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyRecievedDate
            End Get
        End Property

        Public ReadOnly Property IsDirtyDueDate() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyDueDate
            End Get
        End Property

        Public ReadOnly Property IsDirtyExtensionDate() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyExtensionDate
            End Get
        End Property

        Public ReadOnly Property IsDirtyREV1Date() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyREV1Date
            End Get
        End Property

        Public ReadOnly Property IsDirtyREV2Date() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyREV2Date
            End Get
        End Property
        Public ReadOnly Property IsDirtyREV1RecvdDate() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyREV1RecvdDate
            End Get
        End Property

        Public ReadOnly Property IsDirtyREV2RecvdDate() As Boolean
            Get
                Return oLustEventDocumentInfo.IsDirtyREV2RecvdDate
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.LustDocumentInfo
            Dim oLustEventDocumentInfoLocal As MUSTER.Info.LustDocumentInfo
            Try
                For Each oLustEventDocumentInfoLocal In oLustActivityInfo.Documents.Values
                    If oLustEventDocumentInfoLocal.ID = ID Then
                        oLustEventDocumentInfo = oLustEventDocumentInfoLocal
                        oLustEventDocumentInfo.FacilityID = onFacilityID
                        oLustEventDocumentInfo.UserID = onUserID
                        Return oLustEventDocumentInfo
                    End If
                Next
                oLustEventDocumentInfo = oLustEventDocumentDB.DBGetByID(ID)
                oLustEventDocumentInfo.FacilityID = onFacilityID
                oLustEventDocumentInfo.UserID = onUserID
                If oLustEventDocumentInfo.ID = 0 Then
                    'oLustEventDocumentInfo.ID = nID
                    nID -= 1
                End If
                oLustActivityInfo.Documents.Add(oLustEventDocumentInfo)
                Return oLustEventDocumentInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ActivityID As Integer, ByVal DocPropertyID As Integer) As MUSTER.Info.LustDocumentInfo
            Dim oLustEventDocumentInfoLocal As MUSTER.Info.LustDocumentInfo
            Try
                For Each oLustEventDocumentInfoLocal In oLustActivityInfo.Documents.Values
                    If oLustEventDocumentInfoLocal.AssocActivity = ActivityID And oLustEventDocumentInfoLocal.DocClass = DocPropertyID Then
                        oLustEventDocumentInfo = oLustEventDocumentInfoLocal
                        oLustEventDocumentInfo.FacilityID = onFacilityID
                        oLustEventDocumentInfo.UserID = onUserID
                        Return oLustEventDocumentInfo
                    End If
                Next
                oLustEventDocumentInfo = oLustEventDocumentDB.DBGetByActivityIDAndDocClass(ActivityID, DocPropertyID)
                oLustEventDocumentInfo.FacilityID = onFacilityID
                oLustEventDocumentInfo.UserID = onUserID
                If oLustEventDocumentInfo.ID = 0 Then
                    'oLustEventDocumentInfo.ID = nID
                    nID -= 1
                    oLustEventDocumentInfo.AssocActivity = ActivityID
                    oLustEventDocumentInfo.DocClass = DocPropertyID
                End If
                oLustActivityInfo.Documents.Add(oLustEventDocumentInfo)
                Return oLustEventDocumentInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strModuleName As String = String.Empty

            Try
                If Me.ValidateData(strModuleName) Then

                    oLustEventDocumentDB.Put(oLustEventDocumentInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oLustEventDocumentInfo.Archive()

                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True 'False
            '********************************************************
            '
            ' Sample validation code below.  Modify as necessary to
            '  handle rules for the object.  Note that errStr should
            '  be built with all failed validation reasons as it is
            '  raised to the consumer for display to the user.  The
            '  same goes for the boolean validateSuccess.  Since
            '  validateSuccess ASSUMES failure, it must be set to
            '  TRUE if all validations are passed successfully.
            '********************************************************

            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"
                        If oLustEventDocumentInfo.ID = 0 Then
                            'This is a new document
                            'Check to see if NFA can be generated...
                            '
                            'IF tank owner owes fees then
                            '    errStr += "The owner of the tank owes fees." + vbCrLf
                            '    validateSuccess = False
                            'End If
                            'IF Compliance and Enforcement violations due then
                            '    errStr += "Compliance and Enforcement has outstanding violations due to a site inspection." + vbCrLf
                            '    validateSuccess = False
                            'End If
                            'IF required Tank Closure pending then
                            '    errStr += "Closure of a tank is required but not complete." + vbCrLf
                            '    validateSuccess = False
                            'End If
                            'IF Agreed Order non compliance then
                            '    errStr += "The owner has not complied with an Agreed Order." + vbCrLf
                            '    validateSuccess = False
                            'End If
                            'IF invoices not submitted and processed then
                            '    errStr += "All invoices must be submitted and processed." + vbCrLf
                            '    validateSuccess = False
                            'End If
                            'IF documents not submitted and proicessed then
                            '    errStr += "All documents must be submitted and processed." + vbCrLf
                            '    validateSuccess = False
                            'End If
                        Else
                            'This is a modification to an existing document.

                        End If
                        Exit Select
                End Select
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent LustEventErr(errStr)
                End If

                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function CalendarEntries()
            oCalendar.Retrieve(24, ID, Nothing, Nothing)

            Dim bolAddCalendarEntry As Boolean = False

            Dim dtNotificationDate As Date = Now()
            Dim dtDueDate As Date
            Dim nColorCode
            Dim strTaskDesc = "Facility : " & onFacilityID & " - " & Me.DocumentID
            Dim strUserID As String = ""
            Dim strSourceUserID As String = "SYSTEM"
            Dim strGroupID As String = ""
            Dim bolDuetoMe As Boolean = False
            Dim bolToDo As Boolean = False
            Dim bolCompleted As Boolean = False
            Dim bolDeleted As Boolean = False

            If oLustEventDocumentInfo.DocClosedDate <> Nothing Then
                '•	Remove any existing associated To Do or Due to Me Calendar entries
                bolAddCalendarEntry = False
            ElseIf oLustEventDocumentInfo.DocFinancialDate <> Nothing Then
                '•	Remove any existing associated To Do or Due to Me Calendar entries
                '•	Create To Do Calendar entry for the Financial Group on the Sent To Financial date
                bolToDo = True
                dtDueDate = oLustEventDocumentInfo.DocFinancialDate

                'Set Boolean to add calendar entry
                bolAddCalendarEntry = True
            ElseIf oLustEventDocumentInfo.DocRcvDate <> Nothing Then
                '•	Remove any existing associated To Do or Due to Me Calendar entries
                '•	Create a To Do Calendar entry for the user on the Received date of the Document
                bolToDo = True
                dtDueDate = oLustEventDocumentInfo.DocRcvDate

                'Set Boolean to add calendar entry
                bolAddCalendarEntry = True
            ElseIf oLustEventDocumentInfo.DueDate <> Nothing Then
                '•	Create a Due To Me Calendar entry for the user on the Due date of the Document, unless it is a Task, 
                bolDuetoMe = True
                dtDueDate = oLustEventDocumentInfo.DueDate

                'Set Boolean to add calendar entry
                bolAddCalendarEntry = True
            ElseIf oLustEventDocumentInfo.DocRevisionsDue <> Nothing Then
                '•	Remove any existing associated To Do Calendar entries
                '•	Creates Due To Me Calendar entry for the user on the Revision Due date of the Document
                bolDuetoMe = True
                dtDueDate = oLustEventDocumentInfo.DocRevisionsDue

                'Set Boolean to add calendar entry
                bolAddCalendarEntry = True
            ElseIf oLustEventDocumentInfo.DocRcvDate <> Nothing Then
                '•	Remove any existing associated Due To Me Calendar entries
                '•	Create To Do Calendar entry for the user on the Revision Received date of the Documen
                bolToDo = True
                dtDueDate = oLustEventDocumentInfo.DocRcvDate

                'Set Boolean to add calendar entry
                bolAddCalendarEntry = True
            End If

            If bolAddCalendarEntry = True Then
                oCalendar.MarkToDoDeleted(24, ID)
                oCalendar.MarkDueToMeDeleted(24, ID)

                'Create a Calendar Info object 
                Dim oCalendarInfo As MUSTER.Info.CalendarInfo
                oCalendarInfo = New MUSTER.Info.CalendarInfo(0, _
                                                dtNotificationDate, _
                                                dtDueDate, _
                                                nColorCode, _
                                                strTaskDesc, _
                                                strUserID, _
                                                strSourceUserID, _
                                                strGroupID, _
                                                bolDuetoMe, _
                                                bolToDo, _
                                                bolCompleted, _
                                                bolDeleted, _
                                                "sdfsdf", _
                                                Now(), _
                                                "asdf", _
                                                Now())

                oCalendarInfo.OwningEntityID = oLustEventDocumentInfo.EntityID
                oCalendarInfo.OwningEntityType = 24
                oCalendarInfo.IsDirty = True
                oCalendar.Add(oCalendarInfo)
                oCalendar.Flush()
            End If
        End Function

#End Region
#Region "Collection Operations"
        'Gets all the info
        Function GetAll() As MUSTER.Info.LustDocumentCollection
            Try
                oLustActivityInfo.Documents.Clear()
                oLustActivityInfo.Documents = oLustEventDocumentDB.GetAllInfo
                Return oLustActivityInfo.Documents
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetAllbyEventId(ByVal EventID As Integer) As MUSTER.Info.LustDocumentCollection
            Try
                oLustActivityInfo.Documents.Clear()
                oLustActivityInfo.Documents = oLustEventDocumentDB.GetAllInfoByEventID(EventID)
                Return oLustActivityInfo.Documents
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetAllbyACTIVITYID(ByVal ActivityID As Integer) As MUSTER.Info.LustDocumentCollection
            Try
                oLustActivityInfo.Documents.Clear()
                oLustActivityInfo.Documents = oLustEventDocumentDB.GetAllInfoByActivityID(ActivityID)
                Return oLustActivityInfo.Documents
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetAllbyACTIVITYandEVENTID(ByVal ActivityId As Integer, ByVal EventID As Integer) As MUSTER.Info.LustDocumentCollection
            Try
                oLustActivityInfo.Documents.Clear()
                oLustActivityInfo.Documents = oLustEventDocumentDB.GetAllInfoByActivityANDEventID(EventID, ActivityId)
                Return oLustActivityInfo.Documents
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oLustEventDocumentInfo = oLustEventDocumentDB.DBGetByID(ID)
                oLustEventDocumentInfo.FacilityID = onFacilityID
                oLustEventDocumentInfo.UserID = onUserID
                If oLustEventDocumentInfo.ID = 0 Then
                    'oLustEventDocumentInfo.ID = nID
                    nID -= 1
                End If
                oLustActivityInfo.Documents.Add(oLustEventDocumentInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oLustEventActivity As MUSTER.Info.LustDocumentInfo)
            Try
                oLustEventDocumentInfo = oLustEventActivity
                oLustEventDocumentInfo.FacilityID = onFacilityID
                oLustEventDocumentInfo.UserID = onUserID
                oLustActivityInfo.Documents.Add(oLustEventDocumentInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oLustEventDocumentInfoLocal As MUSTER.Info.LustDocumentInfo

            Try
                For Each oLustEventDocumentInfoLocal In oLustActivityInfo.Documents.Values
                    If oLustEventDocumentInfoLocal.ID = ID Then
                        oLustActivityInfo.Documents.Remove(oLustEventDocumentInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Lust Document " & ID.ToString & " is not in the collection of Documents.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLustEventActivity As MUSTER.Info.LustDocumentInfo)
            Try
                oLustActivityInfo.Documents.Remove(oLustEventActivity)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEventDocument " & oLustEventActivity.ID & " is not in the collection of LustEventDocuments.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xLustEventInfo As MUSTER.Info.LustDocumentInfo
            For Each xLustEventInfo In oLustActivityInfo.Documents.Values
                If xLustEventInfo.IsDirty Then
                    oLustEventDocumentInfo = xLustEventInfo
                    '********************************************************
                    '
                    ' Note that if there are contained objects and the respective
                    '  contained object collections are dirty, then the contained
                    '  collections MUST BE FLUSHED before this object can be
                    '  saved.  Otherwise, there is a risk that an attempt will
                    '  be made to insert a new object to the repository without
                    '  corresponding contained information being present which 
                    '  may, in turn, cause a foreign key violation!
                    '
                    '********************************************************
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = oLustActivityInfo.Documents.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return oLustActivityInfo.Documents.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return oLustActivityInfo.Documents.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLustEventDocumentInfo = New MUSTER.Info.LustDocumentInfo
        End Sub
        Public Sub Reset()
            oLustEventDocumentInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oLustEventDocumentInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("LustEvent ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oLustEventDocumentInfoLocal In oLustActivityInfo.Documents.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oLustEventDocumentInfoLocal.ID
                    dr("Deleted") = oLustEventDocumentInfoLocal.Deleted
                    dr("Created By") = oLustEventDocumentInfoLocal.CreatedBy
                    dr("Date Created") = oLustEventDocumentInfoLocal.CreatedOn
                    dr("Last Edited By") = oLustEventDocumentInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oLustEventDocumentInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub LustEventInfoChanged() Handles oLustEventDocumentInfo.LustDocInfoChanged
            RaiseEvent LustEventChanged(True)
        End Sub
        'Private Sub TamplateColChanged() Handles colLustEventDocuments.LustDocumentColChanged
        '    RaiseEvent ColChanged(True)
        'End Sub
#End Region

        Public Function MarkDueToMeCompleted(ByVal DocID As Int64) As String
            Dim strReturn As String
            Dim strSQL As String
            Try
                strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_ID = " & DocID & " AND Owning_Entity_Type = 24 AND Due_To_Me = 1"
                oLustEventDocumentDB.DBExeNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function MarkToDoCompleted(ByVal DocID As Int64) As String
            Dim strReturn As String
            Dim strSQL As String
            Try
                strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_ID = " & DocID & " AND Owning_Entity_Type = 24 AND To_Do = 1"
                oLustEventDocumentDB.DBExeNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function MarkDueToMeCompleted_ByDesc(ByVal DocID As Int64, ByVal strDesc As String) As String
            Dim strReturn As String
            Dim strSQL As String
            Try
                strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_ID = " & DocID & " AND Owning_Entity_Type = 24 AND Due_To_Me = 1 and Task_Description like '" & strDesc & "%'"
                oLustEventDocumentDB.DBExeNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function CloseParentDocument(ByVal ActivityID As Int64, ByVal DocumentID As Int64) As String
            Dim strReturn As String
            Dim strSQL As String
            Try

                strSQL = " UPDATE dbo.tblTEC_EVENT_ACTIVITY_DOCUMENT "
                strSQL &= " SET Date_Closed = GetDate() "
                strSQL &= " WHERE Event_Activity_ID = " & ActivityID & " "
                strSQL &= " and Document_Property_ID in (Select Document_ID from tblTEC_DOCUMENT where Auto_Doc_1 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_2 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_3 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_4 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_5 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_6 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_7 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_8 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_9 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & ")"
                strSQL &= " or Auto_Doc_10 in (Select Document_Property_ID from tblTEC_EVENT_ACTIVITY_DOCUMENT where Event_Activity_Document_ID  = " & DocumentID & "))"
                strSQL &= " and Date_Closed is NULL"

                oLustEventDocumentDB.DBExeNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Public Function GetDocumentTrigger(ByVal DocumentID As Int64) As String
        '    Dim dsReturn As New DataSet
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "SELECT [Trigger Field] as DocTrigger FROM tblTEC_ACTIVITY_DOCUMENT_RELATION where Document_Property_ID = " & DocumentID
        '        dsReturn = oLustEventDocumentDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            strReturn = dsReturn.Tables(0).Rows(0)("DocTrigger")
        '        Else
        '            strReturn = ""
        '        End If
        '        Return strReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function GetDocumentTemplate(ByVal DocumentID As Int64) As String
        '    Dim dsReturn As New DataSet
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "SELECT [Word Document] as DocTrigger FROM tblTEC_ACTIVITY_DOCUMENT_RELATION where Document_Property_ID = " & DocumentID
        '        dsReturn = oLustEventDocumentDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            strReturn = dsReturn.Tables(0).Rows(0)("DocTrigger")
        '        Else
        '            strReturn = "Default"
        '        End If
        '        Return strReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function GetDocumentTitle(ByVal DocumentID As Int64) As String
        '    Dim dsReturn As New DataSet
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "SELECT [Document] as DocName FROM tblTEC_ACTIVITY_DOCUMENT_RELATION where Document_Property_ID = " & DocumentID
        '        dsReturn = oLustEventDocumentDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            strReturn = dsReturn.Tables(0).Rows(0)("DocName")
        '        Else
        '            strReturn = "Default"
        '        End If
        '        Return strReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function GetAutomaticallyCreatedDocuments(ByVal DocumentID As Int64) As DataTable
        '    Dim dsReturn As New DataSet
        '    Dim dtReturn As New DataTable
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "SELECT [Automatically Created Document(s)] as DocID FROM tblTEC_ACTIVITY_DOCUMENT_RELATION where Document_Property_ID = " & DocumentID
        '        dsReturn = oLustEventDocumentDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            dtReturn = dsReturn.Tables(0)
        '        Else
        '            dtReturn = Nothing
        '        End If
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
    End Class
End Namespace
