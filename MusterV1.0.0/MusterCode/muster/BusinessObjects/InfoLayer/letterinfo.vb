'-------------------------------------------------------------------------------
' MUSTER.Info.letterinfo
' Provides the container to persist MUSTER Owner state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN       12/04/04    Original class definition.
'  1.1        AN       12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        AB       02/18/05    Added AgeThreshold and IsAgedData Attributes
'  1.3        MR       03/27/05    Removed current date assigned to Created On and Modified On.
'
'
' Function          Description
' New()             Instantiates an empty letter object
' New(ByVal DocumentId As Integer, ByVal DocumentName As String, ByVal TypeofDocument As String, ByVal DocumentLocation As String, ByVal EntityType As Integer, ByVal EntityId As Integer, ByVal DocumentDescription As String, ByVal WorkFlow As Integer, ByVal DatePrinted As Date, ByVal Deleted As Boolean, ByVal CREATED_BY As String, ByVal DATE_CREATED As Date, ByVal LAST_EDITED_BY As String, ByVal DATE_LAST_EDITED As Date)
'                   Instantiates a populated Letter object
'Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'Archive            Sets the object state to the old state when loaded from or
'                   last saved to the repository
'CheckDirty         'Check for dirty....
'Init               Intialise the object attributes...

'Attribute          Description
' ID                The unique key identifier associated with the letter in the Collection.
' Name               The name of the Document in the repository
' COMPARTMENTNumber  The unique identifier associated with the letter in the repository
'TypeofDocument      The type of document.. 
'DocumentLocation    location of document.. 
'EntityType          The type of the Entity..
'EntityId            The unique identifier associated with the Entity Object.
'DocumentDescription Description of the Document
'WorkFlow            
'DatePrinted       Indicates Letter is  already printed or not.
'IsDirty           Indicates if the Facility state has been altered since it was
'Deleted           'Flag to make the record as deleted.. 
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    <Serializable()> _
    Public Class LetterInfo
#Region "Private member variables"

        Private nDocumentId As Integer
        Private strDocumentName As String
        Private strTypeofDocument As String
        Private strDocumentLocation As String
        Private nEntityType As Integer
        Private nEntityId As Integer
        Private strDocumentDescription As String
        Private nWorkFlow As Integer
        Private dtPrinted As Date
        Private strCreatedBy As String
        Private dtCreatedOn As Date
        Private strModifiedBy As String
        Private dtModifiedOn As Date
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private strOwningUser As String
        Private nModuleID As Integer
        Private nEventID As Int64
        Private nEventSequence As Integer
        Private nEventType As Integer

        'Current values
        Private onDocumentId As Integer
        Private ostrDocumentName As String
        Private ostrTypeofDocument As String
        Private ostrDocumentLocation As String
        Private onEntityType As Integer
        Private onEntityId As Integer
        Private ostrDocumentDescription As String
        Private onWorkFlow As Integer
        Private odtPrinted As Date
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date
        Private obolDeleted As Boolean
        Private obolIsDirty As Boolean
        Private dtDataAge As DateTime
        Private ostrOwningUser As String
        Private onModuleID As Integer
        Private onEventID As Int64
        Private onEventSequence As Integer
        Private onEventType As Integer

        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
        End Sub
        Public Sub New(ByVal DocumentId As Integer, ByVal DocumentName As String, ByVal TypeofDocument As String, ByVal DocumentLocation As String, ByVal EntityType As Integer, ByVal EntityId As Integer, ByVal DocumentDescription As String, ByVal WorkFlow As Integer, ByVal DatePrinted As Date, ByVal Deleted As Boolean, ByVal CREATED_BY As String, ByVal DATE_CREATED As Date, ByVal LAST_EDITED_BY As String, ByVal DATE_LAST_EDITED As Date, ByVal OwningUser As String, ByVal ModuleID As String, ByVal EventID As Int64, ByVal EventSequence As Integer, ByVal EventType As Integer)

            onDocumentId = DocumentId
            ostrDocumentName = DocumentName
            ostrTypeofDocument = TypeofDocument
            ostrDocumentLocation = DocumentLocation
            onEntityType = EntityType
            onEntityId = EntityId
            ostrDocumentDescription = DocumentDescription
            onWorkFlow = WorkFlow
            odtPrinted = DatePrinted
            obolDeleted = Deleted
            ostrCreatedBy = CREATED_BY
            odtCreatedOn = DATE_CREATED
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            ostrOwningUser = OwningUser
            onModuleID = ModuleID
            onEventID = EventID
            onEventSequence = EventSequence
            onEventType = EventType
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nDocumentId = onDocumentId
            strDocumentName = ostrDocumentName
            strTypeofDocument = ostrTypeofDocument
            strDocumentLocation = ostrDocumentLocation
            nEntityType = onEntityType
            nEntityId = onEntityId
            strDocumentDescription = ostrDocumentDescription
            nWorkFlow = onWorkFlow
            dtPrinted = odtPrinted
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            strOwningUser = ostrOwningUser
            nModuleID = onModuleID
            nEventID = onEventID
            nEventSequence = onEventSequence
            nEventType = onEventType
        End Sub
        Public Sub Archive()
            onDocumentId = nDocumentId
            ostrDocumentName = strDocumentName
            ostrTypeofDocument = strTypeofDocument
            ostrDocumentLocation = strDocumentLocation
            onEntityType = nEntityType
            onEntityId = nEntityId
            ostrDocumentDescription = strDocumentDescription
            onWorkFlow = nWorkFlow
            odtPrinted = dtPrinted
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            ostrOwningUser = strOwningUser
            onModuleID = nModuleID
            onEventID = nEventID
            onEventSequence = nEventSequence
            onEventType = nEventType
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (onDocumentId <> nDocumentId) Or _
            (ostrDocumentName <> strDocumentName) Or _
            (ostrTypeofDocument <> strTypeofDocument) Or _
            (ostrDocumentLocation <> strDocumentLocation) Or _
            (onEntityType <> nEntityType) Or _
            (onEntityId <> nEntityId) Or _
            (ostrDocumentDescription <> strDocumentDescription) Or _
            (onWorkFlow <> nWorkFlow) Or _
            (odtPrinted <> dtPrinted) Or _
            (obolDeleted <> bolDeleted) Or _
            (ostrCreatedBy <> strCreatedBy) Or _
            (odtCreatedOn <> dtCreatedOn) Or _
            (ostrModifiedBy <> strModifiedBy) Or _
            (odtModifiedOn <> dtModifiedOn) Or _
            (ostrOwningUser <> strOwningUser) Or _
            (onModuleID <> nModuleID) Or _
            (onEventID <> nEventID) Or _
            (onEventSequence <> nEventSequence) Or _
            (onEventType <> nEventType)
        End Sub
        Private Sub Init()

            onDocumentId = 0
            ostrDocumentName = String.Empty
            ostrTypeofDocument = String.Empty
            ostrDocumentLocation = String.Empty
            onEntityType = 0
            onEntityId = 0
            ostrDocumentDescription = String.Empty
            onWorkFlow = 0
            odtPrinted = CDate("01/01/0001")
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = System.DateTime.Now
            ostrModifiedBy = String.Empty
            odtModifiedOn = System.DateTime.Now
            obolIsDirty = False
            ostrOwningUser = String.Empty
            onModuleID = 0
            onEventID = 0
            onEventSequence = 0
            onEventType = 0
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return Me.nDocumentId
            End Get
            Set(ByVal value As Integer)
                Me.nDocumentId = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Name() As String
            Get
                Return Me.strDocumentName
            End Get
            Set(ByVal value As String)
                Me.strDocumentName = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TypeofDocument() As String
            Get
                Return Me.strTypeofDocument
            End Get
            Set(ByVal value As String)
                Me.strTypeofDocument = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DocumentLocation() As String
            Get
                Return Me.strDocumentLocation
            End Get
            Set(ByVal value As String)
                Me.strDocumentLocation = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EntityType() As Integer
            Get
                Return Me.nEntityType
            End Get
            Set(ByVal value As Integer)
                Me.nEntityType = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EntityId() As Integer

            Get
                Return Me.nEntityId
            End Get
            Set(ByVal value As Integer)
                Me.nEntityId = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DocumentDescription() As String
            Get
                Return Me.strDocumentDescription
            End Get
            Set(ByVal value As String)
                Me.strDocumentDescription = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DatePrinted() As Date
            Get
                Return Me.dtPrinted
            End Get
            Set(ByVal value As Date)
                Me.dtPrinted = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WorkFlow() As Integer
            Get
                Return Me.nWorkFlow
            End Get
            Set(ByVal value As Integer)
                Me.nWorkFlow = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get

            Set(ByVal value As Boolean)
                bolDeleted = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get

            Set(ByVal value As Int16)
                nAgeThreshold = Int16.Parse(value)
            End Set
        End Property

        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get

        End Property
        Public Property OwningUser() As String
            Get
                Return Me.strOwningUser
            End Get
            Set(ByVal value As String)
                Me.strOwningUser = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModuleID() As Integer
            Get
                Return nModuleID
            End Get
            Set(ByVal value As Integer)
                nModuleID = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EventID() As Int64
            Get
                Return nEventID
            End Get
            Set(ByVal value As Int64)
                nEventID = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EventSequence() As Integer
            Get
                Return nEventSequence
            End Get
            Set(ByVal Value As Integer)
                nEventSequence = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EventType() As Integer
            Get
                Return nEventType
            End Get
            Set(ByVal value As Integer)
                nEventType = value
                Me.CheckDirty()
            End Set
        End Property
#End Region
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
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace






