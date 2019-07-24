'-------------------------------------------------------------------------------
' MUSTER.Info.LustDocumentInfo
'   Provides the container to persist MUSTER LustDocument state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC       03/16/05    Original class definition.
'  1.1        JVC       03/24/05    Added CalendarID to the object.
'
' Function          Description
'-------------------------------------------------------------------------------
' New()             Instantiates an empty LustDocumentInfo object.
' New()             Instantiates a populated LustDocumentInfo object.
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
' Save()            Saves the object state to the repository.
' Archive()         Replaces the current value of the object to the one in collection
' CheckDirty()      Checks if the values are different in the collection and the current object
' Init()            Initializes the member variables to their default values
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
' IsDirty            Returns a Boolean if the object has changed from its original status
' ModifiedBy         ID of the user that last made changes
' ModifiedOn         The date of the last changes made 
'

Namespace MUSTER.Info
    Public Class LustDocumentInfo
#Region "Public Events"
        Public Delegate Sub LustDocInfoChangedEventHandler()
        ' Raised when any of the LustEventInfo attributes are modified
        Public Event LustDocInfoChanged As LustDocInfoChangedEventHandler
#End Region
#Region "Private Member Variables"
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private colComments As MUSTER.Info.CommentsCollection
        Private WithEvents colLustDocs As LustDocumentCollection
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtDocClosedDate As Date
        Private dtDocFinancialDate As Date
        Private dtDocRcvDate As Date
        Private dtDocRevisionsDue As Date
        Private dtDueDate As Date
        Private dtIssueDate As Date
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private nAgeThreshold As Int16 = 5
        Private nAssocActivity As Long
        Private nDocClass As Long
        Private nDocumentID As Long
        Private nDocumentType As Long
        Private nEntityID As Integer
        Private nEventID As Integer
        Private onId As Integer
        Private onEventID As Integer
        Private obolDeleted As Boolean
        Private obolIsDirty As Boolean
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtDocClosedDate As Date
        Private odtDocFinancialDate As Date
        Private odtDocRcvDate As Date
        Private odtDocRevisionsDue As Date
        Private odtDueDate As Date
        Private odtIssueDate As Date
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private onAssocActivity As Long
        Private onDocClass As Long
        Private onDocumentID As Long
        Private onDocumentType As Long
        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty
        Private nCommitmentID As Long
        Private bolPaid As Boolean
        Private onCommitmentID As Long
        Private obolPaid As Boolean

        Private dtSTARTDATE As Date
        Private dtEXTENSIONDATE As Date
        Private dtREV1RECEIVEDDATE As Date
        Private dtREV1EXTENSIONDATE As Date
        Private dtREV2RECEIVEDDATE As Date
        Private dtREV2EXTENSIONDATE As Date

        Private odtSTARTDATE As Date
        Private odtEXTENSIONDATE As Date
        Private odtREV1RECEIVEDDATE As Date
        Private odtREV1EXTENSIONDATE As Date
        Private odtREV2RECEIVEDDATE As Date
        Private odtREV2EXTENSIONDATE As Date



        Private onUserID As Integer = 0
        Private onFacilityID As Integer = 0
#End Region
#Region "Constructors"
        Public Sub New()
            Me.Init()
            dtDataAge = Now()
        End Sub

        Public Sub New(ByVal Id As Long, _
            ByVal DocClosedDate As Date, _
            ByVal DocFinancialDate As Date, _
            ByVal DocRcvDate As Date, _
            ByVal REV1RECEIVEDDATE As Date, _
            ByVal REV1EXTENSIONDATE As Date, _
            ByVal REV2RECEIVEDDATE As Date, _
            ByVal REV2EXTENSIONDATE As Date, _
            ByVal STARTDATE As Date, _
            ByVal EXTENSIONDATE As Date, _
            ByVal DueDate As Date, _
            ByVal IssueDate As Date, _
            ByVal EventID As Long, _
            ByVal AssocActivity As Long, _
            ByVal DocClass As Long, _
            ByVal DocumentID As Long, _
            ByVal DocumentType As Long, _
            ByVal CommitmentID As Long, _
            ByVal Paid As Boolean, _
            ByVal CREATED_BY As String, _
            ByVal CREATE_DATE As String, _
            ByVal LAST_EDITED_BY As String, _
            ByVal DATE_LAST_EDITED As Date, _
            ByVal Deleted As Boolean)

            onId = Id
            obolDeleted = Deleted
            odtDocClosedDate = DocClosedDate
            odtDocFinancialDate = DocFinancialDate
            odtDocRcvDate = DocRcvDate
            odtDocRevisionsDue = DocRevisionsDue
            odtDueDate = DueDate
            odtIssueDate = IssueDate
            onAssocActivity = AssocActivity
            onEventID = EventID
            onDocClass = DocClass
            onDocumentID = DocumentID
            onDocumentType = DocumentType
            onCommitmentID = CommitmentID
            obolPaid = Paid

            odtSTARTDATE = STARTDATE
            odtEXTENSIONDATE = EXTENSIONDATE
            odtREV1RECEIVEDDATE = REV1RECEIVEDDATE
            odtREV1EXTENSIONDATE = REV1EXTENSIONDATE
            odtREV2RECEIVEDDATE = REV2RECEIVEDDATE
            odtREV2EXTENSIONDATE = REV2EXTENSIONDATE

            ostrCreatedBy = CREATED_BY
            odtCreatedOn = CREATE_DATE
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Methods"
        Public Sub Reset()
            bolDeleted = obolDeleted
            dtCreatedOn = odtCreatedOn
            dtDocClosedDate = odtDocClosedDate
            dtDocFinancialDate = odtDocFinancialDate
            dtDocRcvDate = odtDocRcvDate
            dtDocRevisionsDue = odtDocRevisionsDue
            dtDueDate = odtDueDate
            dtIssueDate = odtIssueDate
            dtModifiedOn = odtModifiedOn
            nAssocActivity = onAssocActivity
            nEventID = onEventID
            nDocClass = onDocClass
            nDocumentID = onDocumentID
            nDocumentType = onDocumentType
            nCommitmentID = onCommitmentID
            bolPaid = obolPaid

            dtSTARTDATE = odtSTARTDATE
            dtEXTENSIONDATE = odtEXTENSIONDATE
            dtREV1RECEIVEDDATE = odtREV1RECEIVEDDATE
            dtREV1EXTENSIONDATE = odtREV1EXTENSIONDATE
            dtREV2RECEIVEDDATE = odtREV2RECEIVEDDATE
            dtREV2EXTENSIONDATE = odtREV2EXTENSIONDATE

            strCreatedBy = ostrCreatedBy
            strModifiedBy = ostrModifiedBy

            IsDirty = False
        End Sub
        Public Sub Archive()
            obolDeleted = bolDeleted
            odtCreatedOn = dtCreatedOn
            odtDocClosedDate = dtDocClosedDate
            odtDocFinancialDate = dtDocFinancialDate
            odtDocRcvDate = dtDocRcvDate
            odtDocRevisionsDue = dtDocRevisionsDue
            odtDueDate = dtDueDate
            odtIssueDate = dtIssueDate
            odtModifiedOn = dtModifiedOn
            onAssocActivity = nAssocActivity
            onEventID = nEventID
            onDocClass = nDocClass
            onDocumentID = nDocumentID
            onDocumentType = nDocumentType
            onCommitmentID = nCommitmentID
            obolPaid = bolPaid

            odtSTARTDATE = dtSTARTDATE
            odtEXTENSIONDATE = dtEXTENSIONDATE
            odtREV1RECEIVEDDATE = dtREV1RECEIVEDDATE
            odtREV1EXTENSIONDATE = dtREV1EXTENSIONDATE
            odtREV2RECEIVEDDATE = dtREV2RECEIVEDDATE
            odtREV2EXTENSIONDATE = dtREV2EXTENSIONDATE

            ostrCreatedBy = strCreatedBy
            ostrModifiedBy = strModifiedBy

            IsDirty = False
        End Sub
#End Region
#Region "Private Methods"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                (dtCreatedOn <> odtCreatedOn) Or _
                (dtDocClosedDate <> odtDocClosedDate) Or _
                (dtDocFinancialDate <> odtDocFinancialDate) Or _
                (dtDocRcvDate <> odtDocRcvDate) Or _
                (dtDocRevisionsDue <> odtDocRevisionsDue) Or _
                (dtDueDate <> odtDueDate) Or _
                (dtIssueDate <> odtIssueDate) Or _
                (dtModifiedOn <> odtModifiedOn) Or _
                (nAssocActivity <> onAssocActivity) Or _
                (nEventID <> onEventID) Or _
                (nDocClass <> onDocClass) Or _
                (nDocumentID <> onDocumentID) Or _
                (nDocumentType <> onDocumentType) Or _
                (dtSTARTDATE <> odtSTARTDATE) Or _
                (dtEXTENSIONDATE <> odtEXTENSIONDATE) Or _
                (dtREV1RECEIVEDDATE <> odtREV1RECEIVEDDATE) Or _
                (dtREV1EXTENSIONDATE <> odtREV1EXTENSIONDATE) Or _
                (dtREV2RECEIVEDDATE <> odtREV2RECEIVEDDATE) Or _
                (onCommitmentID <> nCommitmentID) Or _
                (obolPaid <> bolPaid) Or _
                (dtREV2EXTENSIONDATE <> odtREV2EXTENSIONDATE)

            If bolOldState <> bolIsDirty Then
                RaiseEvent LustDocInfoChanged()
            End If
        End Sub
        Private Sub Init()
            obolDeleted = False
            odtCreatedOn = System.DateTime.Now
            'odtDocClosedDate = System.DateTime.Now
            'odtDocFinancialDate = System.DateTime.Now
            'odtDocRcvDate = System.DateTime.Now
            'odtDocRevisionsDue = System.DateTime.Now
            'odtDueDate = System.DateTime.Now
            'odtIssueDate = System.DateTime.Now
            odtModifiedOn = System.DateTime.Now
            onAssocActivity = 0
            onEventID = 0
            onDocClass = 0
            onDocumentID = 0
            onDocumentType = 0
            onCommitmentID = 0
            obolPaid = False
            ostrCreatedBy = String.Empty
            ostrModifiedBy = String.Empty
            'odtSTARTDATE = System.DateTime.Now
            'odtEXTENSIONDATE = System.DateTime.Now
            'odtREV1RECEIVEDDATE = System.DateTime.Now
            'odtREV1EXTENSIONDATE = System.DateTime.Now
            'odtREV2RECEIVEDDATE = System.DateTime.Now
            'odtREV2EXTENSIONDATE = System.DateTime.Now
        End Sub
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
            '
            ' Need to fill this in
            '
        End Sub
#End Region
#Region "Exposed Attributes"
        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{4A717F21-3341-4208-A483-9A8BACA51984}
                Return dtDataAge
                ' #End Region ' XDEOperation End Template Expansion{4A717F21-3341-4208-A483-9A8BACA51984}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{81C833E2-220F-421E-9A7A-3B2CC3B9FB39}
                dtDataAge = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{81C833E2-220F-421E-9A7A-3B2CC3B9FB39}
            End Set
        End Property
        ' the Lust Activity that the document is associated with.
        Public Property AssocActivity() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{4557DB37-4CBF-4B13-A613-805797565E37}
                Return nAssocActivity
                ' #End Region ' XDEOperation End Template Expansion{4557DB37-4CBF-4B13-A613-805797565E37}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{EAD07A5A-C202-4764-8578-9C3AA879B4DA}
                nAssocActivity = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{EAD07A5A-C202-4764-8578-9C3AA879B4DA}
            End Set
        End Property
        'Event that the Document is Associated to
        Public Property EventId() As Long
            Get
                Return nEventID
            End Get
            Set(ByVal Value As Long)
                nEventID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CommitmentId() As Long
            Get
                Return nCommitmentID
            End Get
            Set(ByVal Value As Long)
                nCommitmentID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Paid() As Boolean
            Get
                Return bolPaid
            End Get
            Set(ByVal Value As Boolean)
                bolPaid = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The collection of comments associated with the LUST event
        Public Property Comments() As Object
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{EA582CE6-0A21-4E39-BDE4-B792F61F0702}
                Return colComments
                ' #End Region ' XDEOperation End Template Expansion{EA582CE6-0A21-4E39-BDE4-B792F61F0702}
            End Get
            Set(ByVal Value As Object)
                ' #Region "XDEOperation" ' Begin Template Expansion{1C712BAA-74E0-4E56-9B37-D2182D031E0A}
                colComments = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{1C712BAA-74E0-4E56-9B37-D2182D031E0A}
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{3C051DD3-2D74-4F71-9C3C-30F06ACC7E78}
                Return strCreatedBy
                ' #End Region ' XDEOperation End Template Expansion{3C051DD3-2D74-4F71-9C3C-30F06ACC7E78}
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{F0EC3ABE-53C5-4ECB-9ED0-7AACB21427B7}
                Return dtCreatedOn
                ' #End Region ' XDEOperation End Template Expansion{F0EC3ABE-53C5-4ECB-9ED0-7AACB21427B7}
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{62C7C52B-FE4E-49F3-AB97-7A51ADFAB009}
                Return bolDeleted
                ' #End Region ' XDEOperation End Template Expansion{62C7C52B-FE4E-49F3-AB97-7A51ADFAB009}
            End Get
            Set(ByVal Value As Boolean)
                ' #Region "XDEOperation" ' Begin Template Expansion{F45D08F9-6F85-4F50-8CC1-7F38CED752CF}
                bolDeleted = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{F45D08F9-6F85-4F50-8CC1-7F38CED752CF}
            End Set
        End Property
        ' The "class" of the document associated with the LUST activity.  This is the "Document" as defined in the Technical DDD p 21.  Will be drawn from the tblProperty_Master table
        Public Property DocClass() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{D27D4EDE-FB47-4681-85B5-856B8C41F38F}
                Return nDocClass
                ' #End Region ' XDEOperation End Template Expansion{D27D4EDE-FB47-4681-85B5-856B8C41F38F}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{F7CDCE65-51A3-44EA-96BE-C37E08737AA8}
                nDocClass = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{F7CDCE65-51A3-44EA-96BE-C37E08737AA8}
            End Set
        End Property
        ' The date the document is closed.
        Public Property DocClosedDate() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{3E65FD5B-AECA-4C47-B4A9-77B9F88795E0}
                Return dtDocClosedDate
                ' #End Region ' XDEOperation End Template Expansion{3E65FD5B-AECA-4C47-B4A9-77B9F88795E0}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{2DCA670E-705A-4DAA-8D14-74501F48953D}
                dtDocClosedDate = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{2DCA670E-705A-4DAA-8D14-74501F48953D}
            End Set
        End Property
        ' The date the document was "sent to financial"
        Public Property DocFinancialDate() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{656900B3-0783-4491-AEB1-08B6EFB9E5C6}
                Return dtDocFinancialDate
                ' #End Region ' XDEOperation End Template Expansion{656900B3-0783-4491-AEB1-08B6EFB9E5C6}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{343C3AEB-7418-40A7-BF01-46DFCE1A7AC4}
                dtDocFinancialDate = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{343C3AEB-7418-40A7-BF01-46DFCE1A7AC4}
            End Set
        End Property
        ' The date the document was received.
        Public Property DocRcvDate() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{C028FCBF-F743-4535-BC6D-522977804FDE}
                Return dtDocRcvDate
                ' #End Region ' XDEOperation End Template Expansion{C028FCBF-F743-4535-BC6D-522977804FDE}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{3A324E6E-D98F-406D-A73C-42579919214F}
                dtDocRcvDate = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{3A324E6E-D98F-406D-A73C-42579919214F}
            End Set
        End Property
        ' The date the revisions for the document are due.
        Public Property DocRevisionsDue() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{EEC0F267-0B24-4635-87E3-A2AF26100E3F}
                Return dtDocRevisionsDue
                ' #End Region ' XDEOperation End Template Expansion{EEC0F267-0B24-4635-87E3-A2AF26100E3F}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{0E10685F-C1F0-4C65-BD9F-7F57EF7A08AC}
                dtDocRevisionsDue = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{0E10685F-C1F0-4C65-BD9F-7F57EF7A08AC}
            End Set
        End Property
        ' The system generated ID for the LUST document
        Public Property DocumentID() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{6FCD7429-803E-49C3-8FB9-535A7D53E336}
                Return nDocumentID
                ' #End Region ' XDEOperation End Template Expansion{6FCD7429-803E-49C3-8FB9-535A7D53E336}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{3E62DCA1-1941-4242-9075-BBE4ED5E2727}
                nDocumentID = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{3E62DCA1-1941-4242-9075-BBE4ED5E2727}
            End Set
        End Property
        ' The document type of the LUST document
        Public Property DocumentType() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{7C259F1C-BAD9-47D7-B0FC-63B51CC706D0}
                Return nDocumentType
                ' #End Region ' XDEOperation End Template Expansion{7C259F1C-BAD9-47D7-B0FC-63B51CC706D0}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{A2F698E6-56AC-49F0-A109-1CF85A54A270}
                nDocumentType = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{A2F698E6-56AC-49F0-A109-1CF85A54A270}
            End Set
        End Property
        ' The date the task or reminder is to be completed by
        Public Property DueDate() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{3541EC9C-6DF4-4CEB-87D8-FCF9909AFA5F}
                Return dtDueDate
                ' #End Region ' XDEOperation End Template Expansion{3541EC9C-6DF4-4CEB-87D8-FCF9909AFA5F}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{3245BA3D-0354-4456-A9AC-64D4F9EBF226}
                dtDueDate = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{3245BA3D-0354-4456-A9AC-64D4F9EBF226}
            End Set
        End Property
        ' The entity ID associated with a technical document.
        Public ReadOnly Property EntityID() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{960B6F22-0703-4492-B891-8DB6468B000B}
                Return nEntityID
                ' #End Region ' XDEOperation End Template Expansion{960B6F22-0703-4492-B891-8DB6468B000B}
            End Get

        End Property
        ' The system ID for this LUST event
        Public Property ID() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{A0B81136-12BB-4719-8E07-49B06333F62F}

                Return onId
                ' #End Region ' XDEOperation End Template Expansion{A0B81136-12BB-4719-8E07-49B06333F62F}

            End Get
            Set(ByVal Value As Long)
                onId = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
            End Set
        End Property
        ' The date the document was issued.
        Public Property IssueDate() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{03B60188-7306-48CD-9820-82AEBBAE0D02}

                Return dtIssueDate
                ' #End Region ' XDEOperation End Template Expansion{03B60188-7306-48CD-9820-82AEBBAE0D02}

            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{54FB1033-F896-4BC6-A756-47467BF4DBFF}
                dtIssueDate = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{54FB1033-F896-4BC6-A756-47467BF4DBFF}
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{07B7B8F4-0B24-4EA0-B546-4C0D4AB74F19}      
                Return dtDataAge < AgeThreshold
                ' #End Region ' XDEOperation End Template Expansion{07B7B8F4-0B24-4EA0-B546-4C0D4AB74F19}      
            End Get
        End Property

        Public Property STARTDATE() As Date
            Get
                Return dtSTARTDATE
            End Get
            Set(ByVal Value As Date)
                dtSTARTDATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EXTENSIONDATE() As Date
            Get
                Return dtEXTENSIONDATE
            End Get
            Set(ByVal Value As Date)
                dtEXTENSIONDATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property REV1EXTENSIONDATE() As Date
            Get
                Return dtREV1EXTENSIONDATE
            End Get
            Set(ByVal Value As Date)
                dtREV1EXTENSIONDATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property REV1RECEIVEDDATE() As Date
            Get
                Return dtREV1RECEIVEDDATE
            End Get
            Set(ByVal Value As Date)
                dtREV1RECEIVEDDATE = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property REV2RECEIVEDDATE() As Date
            Get
                Return dtREV2RECEIVEDDATE
            End Get
            Set(ByVal Value As Date)
                dtREV2RECEIVEDDATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property REV2EXTENSIONDATE() As Date
            Get
                Return dtREV2EXTENSIONDATE
            End Get
            Set(ByVal Value As Date)
                dtREV2EXTENSIONDATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property UserID() As Integer
            Get
                Return onUserID
            End Get
            Set(ByVal Value As Integer)
                onUserID = Value
            End Set
        End Property

        Public Property FacilityID() As Integer
            Get
                Return onFacilityID
            End Get
            Set(ByVal Value As Integer)
                onFacilityID = Value
            End Set
        End Property


        Public ReadOnly Property IsDirtyClosedDate() As Boolean
            Get
                If odtDocClosedDate = dtDocClosedDate Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public ReadOnly Property IsDirtySentToFinancial() As Boolean
            Get
                If odtDocFinancialDate = dtDocFinancialDate Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public ReadOnly Property IsDirtyRecievedDate() As Boolean
            Get
                If odtDocRcvDate = dtDocRcvDate Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public ReadOnly Property IsDirtyDueDate() As Boolean
            Get
                If dtDueDate = odtDueDate Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property


        Public ReadOnly Property IsDirtyExtensionDate() As Boolean
            Get
                If dtEXTENSIONDATE = odtEXTENSIONDATE Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public ReadOnly Property IsDirtyREV1Date() As Boolean
            Get

                If dtREV1EXTENSIONDATE = odtREV1EXTENSIONDATE Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property
        Public ReadOnly Property IsDirtyREV2Date() As Boolean
            Get

                If dtREV2EXTENSIONDATE = odtREV2EXTENSIONDATE Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public ReadOnly Property IsDirtyREV1RecvdDate() As Boolean
            Get

                If dtREV1RECEIVEDDATE = odtREV1RECEIVEDDATE Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public ReadOnly Property IsDirtyREV2RecvdDate() As Boolean
            Get

                If dtREV2RECEIVEDDATE = odtREV2RECEIVEDDATE Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

#End Region
    End Class
End Namespace