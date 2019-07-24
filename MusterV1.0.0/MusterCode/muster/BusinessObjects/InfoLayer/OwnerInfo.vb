'-------------------------------------------------------------------------------
' MUSTER.Info.OwnerInfo
'   Provides the container to persist MUSTER Owner state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR       12/03/04    Original class definition.
'  1.1        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        MNR       01/11/05    Added Events
'  1.3        JVCII     02/11/05    Added  
'                                   and nAddressID <> 0 and (nPersonID <> 0 Or nOrganizationID <> 0) to
'                                       InfoChanged calls in order to enable/disable
'                                       Save/Cancel buttons.
'                                   Modified CheckDirty() to trigger event if either dirty
'                                       state changes or address or persona changes.
'  1.4        AB        02/21/05    Added AgeThreshold and IsAgedData Attributes
'  1.5        JVC2      03/07/05    Change Modified By and Modified On to read/write.
'  1.6        MR        03/14/05    Changed Created By and Created On to read/write.
'  1.7        MNR       03/15/05    Updated Constructor New(ByVal drOwner As DataRow) to check for System.DBNull.Value
'  1.8        MNR       03/16/05    Removed strSrc from events
'  1.9        KKM       03/17/05    FacilityCollection and commentsCollection properties are added
'
' Function          Description
' New()             Instantiates an empty OwnerInfo object
' New(OwnerID, OrganizationID, PersonID, PhoneNumberOne, PhoneNumberTwo, FaxNumber,
'       EmailAddress, EmailPersonal, AddressID, DateCapSignup, CapCurrentStatus,
'       OwnerType, BP2KOwnerType, FeesProfileID, FeesStatus, ComplianceProfileID,
'       CompliaceStatus, Active, FeeActive, EnsiteOrganizationID, EnsitePersonID,
'       EnsiteAgencyInterestID, OwnerDesc, CustEntityCode, CustTypeCode, Deleted,
'       CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated OwnerInfo object
' New(dr)           Instantiates a populated OwnerInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' TODO - Address, Name
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
      Public Class OwnerInfo
#Region "Public Events"
        Public Event OwnerInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"

        Private nOwnerID As Int64                  'The internal ID number associated with the user group
        Private nOrganizationID As Int64
        Private nPersonID As Int64
        Private strPhoneNumberOne As String
        Private strPhoneNumberTwo As String
        Private strFaxNumber As String
        Private strEmailAddress As String
        Private strEmailPersonal As String
        Private nAddressID As Int64
        'Private strAddressLine1 As String
        'Private strAddressLine2 As String
        'Private strState As String
        'Private strCity As String
        'Private strZip As String
        'Private strFIPSCode As String
        Private dtDateCapSignup As Date
        Private bolCapCurrentStatus As Boolean
        Private nOwnerType As Int64
        Private nBP2KOwnerType As Int64
        Private nFeesProfileID As Int64
        Private bolFeesStatus As Boolean
        Private nComplianceProfileID As Int64
        Private bolCompliaceStatus As Boolean
        Private bolActive As Boolean
        Private bolFeeActive As Boolean
        Private nEnsiteOrganizationID As Int64
        Private nEnsitePersonID As Int64
        Private bolEnsiteAgencyInterestID As Boolean
        '   Private nOwnerDesc As Int64
        Private nCustEntityCode As Int64
        Private nCustTypeCode As Int64
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private bolOwnerL2CSnippet As Boolean
        Private strBP2KOwnerID As String
        Private strCapParticipationLevel As String

        ' TODO - Address, Name
        'Private OwnerPM As New Address
        'Private strFirst_Name As String = ""
        'Private strLast_Name As String = ""
        'Private strOrg_Name As String = ""

        'Private colAssociatedGroups As UserGroups   'The list of user groups associated with the form
        'Private dtOwnerNames As DataTable         'Lists the Owner names for the client
        'Private bolFavoriteOwner As Boolean = False

        Private onOwnerID As Int64
        Private onOrganizationID As Int64
        Private onPersonID As Int64
        Private ostrPhoneNumberOne As String
        Private ostrPhoneNumberTwo As String
        Private ostrFaxNumber As String
        Private ostrEmailAddress As String
        Private ostrEmailPersonal As String
        Private onAddressID As Int64
        'Private ostrAddressLine1 As String
        'Private ostrAddressLine2 As String
        'Private ostrState As String
        'Private ostrCity As String
        'Private ostrZip As String
        'Private ostrFIPSCode As String
        Private odtDateCapSignup As Date
        Private obolCapCurrentStatus As Boolean
        Private onOwnerType As Int64
        Private onBP2KOwnerType As Int64
        Private onFeesProfileID As Int64
        Private obolFeesStatus As Boolean
        Private onComplianceProfileID As Int64
        Private obolCompliaceStatus As Boolean
        Private obolActive As Boolean
        Private obolFeeActive As Boolean
        Private onEnsiteOrganizationID As Int64
        Private onEnsitePersonID As Int64
        Private obolEnsiteAgencyInterestID As Boolean
        Private onOwnerDesc As Int64
        Private onCustEntityCode As Int64
        Private onCustTypeCode As Int64
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date
        Private obolOwnerL2CSnippet As Boolean
        Private ostrBP2KOwnerID As String
        Private ostrCapParticipationLevel As String

        Private bolShowDeleted As Boolean = False
        Private dtDataAge As Date
        Private nAgeThreshold As Int16 = 5

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        'added by kiran
        Private colFacility As MUSTER.Info.FacilityCollection
        Private colComments As MUSTER.Info.CommentsCollection
        'end changes
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
            'added by kiran
            colFacility = New MUSTER.Info.FacilityCollection
            colComments = New MUSTER.Info.CommentsCollection
            'end changes
        End Sub
        Sub New(ByVal OwnerID As Integer, _
            ByVal OrganizationID As Integer, _
            ByVal PersonID As Integer, _
            ByVal PhoneNumberOne As String, _
            ByVal PhoneNumberTwo As String, _
            ByVal FaxNumber As String, _
            ByVal EmailAddress As String, _
            ByVal EmailPersonal As String, _
            ByVal AddressID As Int64, _
            ByVal DateCapSignup As Date, _
            ByVal CapCurrentStatus As Boolean, _
            ByVal OwnerType As Integer, _
            ByVal BP2KOwnerType As Integer, _
            ByVal FeesProfileID As Integer, _
            ByVal FeesStatus As Boolean, _
            ByVal ComplianceProfileID As Integer, _
            ByVal CompliaceStatus As Boolean, _
            ByVal Active As Boolean, _
            ByVal FeeActive As Boolean, _
            ByVal EnsiteOrganizationID As Integer, _
            ByVal EnsitePersonID As Integer, _
            ByVal EnsiteAgencyInterestID As Boolean, _
            ByVal CustEntityCode As Integer, _
            ByVal CustTypeCode As Integer, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date, _
            ByVal OwnerL2CSnippet As Boolean, _
            ByVal bp2kOwnID As String, _
            ByVal capParticipationLevel As String)
            ' ByVal OwnerDesc As Integer, _
            'ByVal AddressLineOne As String, _
            'ByVal AddressTwo As String, _
            'ByVal City As String, _
            'ByVal State As String, _
            'ByVal Zip As String, _
            'ByVal FIPSCode As String, _

            onOwnerID = OwnerID
            onOrganizationID = OrganizationID
            onPersonID = PersonID
            ostrPhoneNumberOne = PhoneNumberOne
            ostrPhoneNumberTwo = PhoneNumberTwo
            ostrFaxNumber = FaxNumber
            ostrEmailAddress = EmailAddress
            ostrEmailPersonal = EmailPersonal
            onAddressID = AddressID
            'ostrAddressLine1 = AddressLineOne
            'ostrAddressLine2 = AddressTwo
            'ostrCity = City
            'ostrState = State
            'ostrZip = Zip
            'ostrFIPSCode = FIPSCode
            odtDateCapSignup = DateCapSignup.Date
            obolCapCurrentStatus = CapCurrentStatus
            onOwnerType = OwnerType
            onBP2KOwnerType = BP2KOwnerType
            onFeesProfileID = FeesProfileID
            obolFeesStatus = FeesStatus
            onComplianceProfileID = ComplianceProfileID
            obolCompliaceStatus = CompliaceStatus
            obolActive = Active
            obolFeeActive = FeeActive
            onEnsiteOrganizationID = EnsiteOrganizationID
            onEnsitePersonID = EnsitePersonID
            obolEnsiteAgencyInterestID = EnsiteAgencyInterestID
            ' onOwnerDesc = OwnerDesc
            onCustEntityCode = CustEntityCode
            onCustTypeCode = CustTypeCode
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            obolOwnerL2CSnippet = OwnerL2CSnippet
            ostrBP2KOwnerID = bp2kOwnID
            ostrCapParticipationLevel = CapParticipationLevel
            dtDataAge = Now()
            'added by kiran
            colFacility = New MUSTER.Info.FacilityCollection
            colComments = New MUSTER.Info.CommentsCollection
            'end changes
            Me.Reset()
        End Sub
        Sub New(ByVal drOwner As DataRow)
            Try
                onOwnerID = drOwner.Item("OWNER_ID")
                onOrganizationID = IIf(drOwner.Item("ORGANIZATION_ID") Is System.DBNull.Value, 0, drOwner.Item("ORGANIZATION_ID")) 'drOwner.Item("ORGANIZATION_ID")
                onPersonID = IIf(drOwner.Item("PERSON_ID") Is System.DBNull.Value, 0, drOwner.Item("PERSON_ID")) 'drOwner.Item("PERSON_ID")
                ostrPhoneNumberOne = drOwner.Item("PHONE_NUMBER_ONE")
                ostrPhoneNumberTwo = drOwner.Item("PHONE_NUMBER_TWO")
                ostrFaxNumber = drOwner.Item("FAX_NUMBER")
                ostrEmailAddress = drOwner.Item("EMAIL_ADDRESS")
                ostrEmailPersonal = drOwner.Item("EMAIL_ADDRESS_PERSONAL")
                onAddressID = drOwner.Item("ADDRESS_ID")
                'ostrAddressLine1 = drOwner.Item("ADDRESS_LINE_ONE")
                'ostrAddressLine2 = IIf(drOwner.Item("ADDRESS_TWO") Is System.DBNull.Value, String.Empty, drOwner.Item("ADDRESS_TWO"))
                'ostrCity = IIf(drOwner.Item("CITY") Is System.DBNull.Value, String.Empty, drOwner.Item("CITY"))
                'ostrState = IIf(drOwner.Item("STATE") Is System.DBNull.Value, String.Empty, drOwner.Item("STATE"))
                'ostrZip = IIf(drOwner.Item("ZIP") Is System.DBNull.Value, String.Empty, drOwner.Item("ZIP"))
                'ostrFIPSCode = IIf(drOwner.Item("FIPS_CODE") Is System.DBNull.Value, String.Empty, drOwner.Item("FIPS_CODE"))
                odtDateCapSignup = IIf(drOwner.Item("DATE_CAP_SIGNUP") Is System.DBNull.Value, CDate("01/01/0001"), drOwner.Item("DATE_CAP_SIGNUP")) 'drOwner.Item("DATE_CAP_SIGNUP")
                odtDateCapSignup = odtDateCapSignup.Date
                obolCapCurrentStatus = IIf(drOwner.Item("CAP_CURRENT_STATUS") Is System.DBNull.Value, False, drOwner.Item("CAP_CURRENT_STATUS")) 'drOwner.Item("CAP_CURRENT_STATUS")
                onOwnerType = drOwner.Item("OWNER_TYPE")
                onBP2KOwnerType = IIf(drOwner.Item("BP2K_OWNER_TYPE") Is System.DBNull.Value, 0, drOwner.Item("BP2K_OWNER_TYPE")) 'drOwner.Item("BP2K_OWNER_TYPE")
                onFeesProfileID = IIf(drOwner.Item("FEES_PROFILE_ID") Is System.DBNull.Value, 0, drOwner.Item("FEES_PROFILE_ID")) 'drOwner.Item("FEES_PROFILE_ID")
                obolFeesStatus = IIf(drOwner.Item("FEES_STATUS") Is System.DBNull.Value, False, drOwner.Item("FEES_STATUS")) 'drOwner.Item("FEES_STATUS")
                onComplianceProfileID = IIf(drOwner.Item("COMPLIANCE_PROFILE_ID") Is System.DBNull.Value, 0, drOwner.Item("COMPLIANCE_PROFILE_ID")) 'drOwner.Item("COMPLIANCE_PROFILE_ID")
                obolCompliaceStatus = IIf(drOwner.Item("COMPLIACE_STATUS") Is System.DBNull.Value, False, drOwner.Item("COMPLIACE_STATUS")) 'drOwner.Item("COMPLIACE_STATUS")
                obolActive = drOwner.Item("ACTIVE")
                obolFeeActive = IIf(drOwner.Item("FEE_ACTIVE") Is System.DBNull.Value, False, drOwner.Item("FEE_ACTIVE")) 'drOwner.Item("FEE_ACTIVE")
                onEnsiteOrganizationID = IIf(drOwner.Item("ENSITE_ORGANIZATION_ID") Is System.DBNull.Value, 0, drOwner.Item("ENSITE_ORGANIZATION_ID")) 'drOwner.Item("ENSITE_ORGANIZATION_ID")
                onEnsitePersonID = IIf(drOwner.Item("ENSITE_PERSON_ID") Is System.DBNull.Value, 0, drOwner.Item("ENSITE_PERSON_ID")) 'drOwner.Item("ENSITE_PERSON_ID")
                obolEnsiteAgencyInterestID = IIf(drOwner.Item("ENSITE_AGENCY_INTEREST_ID") Is System.DBNull.Value, False, drOwner.Item("ENSITE_AGENCY_INTEREST_ID")) 'drOwner.Item("ENSITE_AGENCY_INTEREST_ID")
                onCustEntityCode = drOwner.Item("CUST_ENTITY_CODE")
                onCustTypeCode = drOwner.Item("CUST_TYPE_CODE")
                obolDeleted = drOwner.Item("DELETED")
                ostrCreatedBy = drOwner.Item("CREATED_BY")
                odtCreatedOn = drOwner.Item("DATE_CREATED")
                ostrModifiedBy = IIf(drOwner.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drOwner.Item("LAST_EDITED_BY")) 'drOwner.Item("LAST_EDITED_BY")
                odtModifiedOn = IIf(drOwner.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drOwner.Item("DATE_LAST_EDITED")) 'drOwner.Item("DATE_LAST_EDITED")
                obolOwnerL2CSnippet = drOwner.Item("OWNER_L2C_SNIPPET")
                ostrBP2KOwnerID = IIf(drOwner.Item("BP2K_OWNER_ID") Is DBNull.Value, String.Empty, drOwner.Item("BP2K_OWNER_ID"))
                ostrCapParticipationLevel = IIf(drOwner.Item("CAP_PARTICIPATION_LEVEL") Is DBNull.Value, String.Empty, drOwner.Item("CAP_PARTICIPATION_LEVEL"))
                dtDataAge = Now()
                'added by kiran
                colFacility = New MUSTER.Info.FacilityCollection
                colComments = New MUSTER.Info.CommentsCollection
                'end changes
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nOwnerID >= 0 Then
                nOwnerID = onOwnerID
            End If
            nOrganizationID = onOrganizationID
            nPersonID = onPersonID
            strPhoneNumberOne = ostrPhoneNumberOne
            strPhoneNumberTwo = ostrPhoneNumberTwo
            strFaxNumber = ostrFaxNumber
            strEmailAddress = ostrEmailAddress
            strEmailPersonal = ostrEmailPersonal
            nAddressID = onAddressID
            'strAddressLine1 = ostrAddressLine1
            'strAddressLine2 = ostrAddressLine2
            'strState = ostrState
            'strCity = ostrCity
            'strZip = ostrZip
            'strFIPSCode = ostrFIPSCode
            dtDateCapSignup = odtDateCapSignup
            bolCapCurrentStatus = obolCapCurrentStatus
            nOwnerType = onOwnerType
            nBP2KOwnerType = onBP2KOwnerType
            nFeesProfileID = onFeesProfileID
            bolFeesStatus = obolFeesStatus
            nComplianceProfileID = onComplianceProfileID
            bolCompliaceStatus = obolCompliaceStatus
            bolActive = obolActive
            bolFeeActive = obolFeeActive
            nEnsiteOrganizationID = onEnsiteOrganizationID
            nEnsitePersonID = onEnsitePersonID
            bolEnsiteAgencyInterestID = obolEnsiteAgencyInterestID
            '    nOwnerDesc = onOwnerDesc
            nCustEntityCode = onCustEntityCode
            nCustTypeCode = onCustTypeCode
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolOwnerL2CSnippet = obolOwnerL2CSnippet
            strBP2KOwnerID = ostrBP2KOwnerID
            strCapParticipationLevel = ostrCapParticipationLevel
            bolIsDirty = False
            'RaiseEvent OwnerInfoChanged(bolIsDirty And nAddressID <> 0 And (nPersonID <> 0 Or nOrganizationID <> 0))
            RaiseEvent OwnerInfoChanged(bolIsDirty And (nPersonID <> 0 Or nOrganizationID <> 0))
        End Sub
        Public Sub Archive()
            onOwnerID = nOwnerID
            onOrganizationID = nOrganizationID
            onPersonID = nPersonID
            ostrPhoneNumberOne = strPhoneNumberOne
            ostrPhoneNumberTwo = strPhoneNumberTwo
            ostrFaxNumber = strFaxNumber
            ostrEmailAddress = strEmailAddress
            ostrEmailPersonal = strEmailPersonal
            onAddressID = nAddressID
            'ostrAddressLine1 = strAddressLine1
            'ostrAddressLine2 = strAddressLine2
            'ostrState = strState
            'ostrCity = strCity
            'ostrZip = strZip
            'ostrFIPSCode = strFIPSCode
            odtDateCapSignup = dtDateCapSignup
            obolCapCurrentStatus = bolCapCurrentStatus
            onOwnerType = nOwnerType
            onBP2KOwnerType = nBP2KOwnerType
            onFeesProfileID = nFeesProfileID
            obolFeesStatus = bolFeesStatus
            onComplianceProfileID = nComplianceProfileID
            obolCompliaceStatus = bolCompliaceStatus
            obolActive = bolActive
            obolFeeActive = bolFeeActive
            onEnsiteOrganizationID = nEnsiteOrganizationID
            onEnsitePersonID = nEnsitePersonID
            obolEnsiteAgencyInterestID = bolEnsiteAgencyInterestID
            '    onOwnerDesc = nOwnerDesc
            onCustEntityCode = nCustEntityCode
            onCustTypeCode = nCustTypeCode
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolOwnerL2CSnippet = bolOwnerL2CSnippet
            ostrBP2KOwnerID = strBP2KOwnerID
            ostrCapParticipationLevel = strCapParticipationLevel
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty
            '(strAddressLine1 <> ostrAddressLine1) Or _
            '(strAddressLine2 <> ostrAddressLine2) Or _
            '(strState <> ostrState) Or _
            '(strCity <> ostrCity) Or _
            '(strZip <> ostrZip) Or _
            '(strFIPSCode <> ostrFIPSCode) Or _
            bolIsDirty = (nOrganizationID <> onOrganizationID) Or _
            (nPersonID <> onPersonID) Or _
            (nAddressID <> onAddressID) Or _
            (strPhoneNumberOne <> ostrPhoneNumberOne) Or _
            (strPhoneNumberTwo <> ostrPhoneNumberTwo) Or _
            (strFaxNumber <> ostrFaxNumber) Or _
            (strEmailAddress <> ostrEmailAddress) Or _
            (strEmailPersonal <> ostrEmailPersonal) Or _
            (dtDateCapSignup <> odtDateCapSignup) Or _
            (bolCapCurrentStatus <> obolCapCurrentStatus) Or _
            (nOwnerType <> onOwnerType) Or _
            (nBP2KOwnerType <> onBP2KOwnerType) Or _
            (nFeesProfileID <> onFeesProfileID) Or _
            (bolFeesStatus <> obolFeesStatus) Or _
            (nComplianceProfileID <> onComplianceProfileID) Or _
            (bolCompliaceStatus <> obolCompliaceStatus) Or _
            (bolActive <> obolActive) Or _
            (bolFeeActive <> obolFeeActive) Or _
            (nEnsiteOrganizationID <> onEnsiteOrganizationID) Or _
            (nEnsitePersonID <> onEnsitePersonID) Or _
            (bolEnsiteAgencyInterestID <> obolEnsiteAgencyInterestID) Or _
            (nCustEntityCode <> onCustEntityCode) Or _
            (nCustTypeCode <> onCustTypeCode) Or _
            (bolDeleted <> obolDeleted) Or _
            (bolOwnerL2CSnippet <> obolOwnerL2CSnippet) Or _
            (strBP2KOwnerID <> ostrBP2KOwnerID)

            If obolIsDirty <> bolIsDirty And (nPersonID <> 0 Or nOrganizationID <> 0) Then
                RaiseEvent OwnerInfoChanged(bolIsDirty And (nPersonID <> 0 Or nOrganizationID <> 0))
            End If
        End Sub
        Private Sub Init()
            onOwnerID = 0
            onOrganizationID = 0
            onPersonID = 0
            ostrPhoneNumberOne = String.Empty
            ostrPhoneNumberTwo = String.Empty
            ostrFaxNumber = String.Empty
            ostrEmailAddress = String.Empty
            ostrEmailPersonal = String.Empty
            onAddressID = 0
            'ostrAddressLine1 = String.Empty
            'ostrAddressLine2 = String.Empty
            'ostrCity = String.Empty
            'ostrState = String.Empty
            'ostrZip = String.Empty
            'ostrFIPSCode = String.Empty
            odtDateCapSignup = System.DateTime.Now
            obolCapCurrentStatus = False
            onOwnerType = 0
            onBP2KOwnerType = 0
            onFeesProfileID = 0
            obolFeesStatus = False
            onComplianceProfileID = 0
            obolCompliaceStatus = False
            obolActive = False
            obolFeeActive = False
            onEnsiteOrganizationID = 0
            onEnsitePersonID = 0
            obolEnsiteAgencyInterestID = False
            onOwnerDesc = 0
            onCustEntityCode = 0
            onCustTypeCode = 0
            obolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = CDate("01/01/0001")
            strModifiedBy = String.Empty
            dtModifiedOn = CDate("01/01/0001")
            obolOwnerL2CSnippet = True
            ostrBP2KOwnerID = String.Empty
            ostrCapParticipationLevel = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        'property added by kiran
        Public Property facilityCollection() As MUSTER.Info.FacilityCollection
            Get
                Return colFacility
            End Get
            Set(ByVal Value As MUSTER.Info.FacilityCollection)
                colFacility = Value
            End Set
        End Property
        Public Property commentsCollection() As MUSTER.Info.CommentsCollection
            Get
                Return colComments
            End Get
            Set(ByVal Value As MUSTER.Info.CommentsCollection)
                colComments = Value
            End Set
        End Property
        'end changes

        Public Property ID() As Integer
            Get
                Return nOwnerID
            End Get

            Set(ByVal value As Integer)
                nOwnerID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property OrganizationID() As Integer
            Get
                Return nOrganizationID
            End Get

            Set(ByVal value As Integer)
                nOrganizationID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property PersonID() As Integer
            Get
                Return nPersonID
            End Get

            Set(ByVal value As Integer)
                nPersonID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property PhoneNumberOne() As String
            Get
                Return strPhoneNumberOne
            End Get

            Set(ByVal value As String)
                strPhoneNumberOne = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PhoneNumberTwo() As String
            Get
                Return strPhoneNumberTwo
            End Get

            Set(ByVal value As String)
                strPhoneNumberTwo = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Fax() As String
            Get
                Return strFaxNumber
            End Get

            Set(ByVal value As String)
                strFaxNumber = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EmailAddress() As String
            Get
                Return strEmailAddress
            End Get

            Set(ByVal value As String)
                strEmailAddress = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EmailAddressPersonal() As String
            Get
                Return strEmailPersonal
            End Get

            Set(ByVal value As String)
                strEmailPersonal = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property AddressId() As Long
            Get
                Return nAddressID
            End Get

            Set(ByVal value As Long)
                nAddressID = Long.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        'Public Property AddressLine1() As String
        '    Get
        '        Return strAddressLine1
        '    End Get

        '    Set(ByVal value As String)
        '        strAddressLine1 = value
        '        Me.CheckDirty()
        '    End Set
        'End Property

        'Public Property AddressLine2() As String
        '    Get
        '        Return strAddressLine2
        '    End Get

        '    Set(ByVal value As String)
        '        strAddressLine2 = value
        '        Me.CheckDirty()
        '    End Set
        'End Property

        'Public Property City() As String
        '    Get
        '        Return strCity
        '    End Get

        '    Set(ByVal value As String)
        '        strCity = value
        '        Me.CheckDirty()
        '    End Set
        'End Property

        'Public Property State() As String
        '    Get
        '        Return strState
        '    End Get

        '    Set(ByVal value As String)
        '        strState = value
        '        Me.CheckDirty()
        '    End Set
        'End Property

        'Public Property Zip() As String
        '    Get
        '        Return strZip
        '    End Get

        '    Set(ByVal value As String)
        '        strZip = value
        '        Me.CheckDirty()
        '    End Set
        'End Property

        'Public Property FIPSCode() As String
        '    Get
        '        Return strFIPSCode
        '    End Get

        '    Set(ByVal value As String)
        '        strFIPSCode = value
        '        Me.CheckDirty()
        '    End Set
        'End Property

        Public Property DateCapSignUp() As Date
            Get
                Return dtDateCapSignup.Date
            End Get

            Set(ByVal value As Date)
                dtDateCapSignup = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property CapCurrentStatus() As Boolean
            Get
                Return bolCapCurrentStatus
            End Get

            Set(ByVal value As Boolean)
                bolCapCurrentStatus = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OwnerType() As Integer
            Get
                Return nOwnerType
            End Get

            Set(ByVal value As Integer)
                nOwnerType = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property BP2KType() As Integer
            Get
                Return nBP2KOwnerType
            End Get

            Set(ByVal value As Integer)
                nBP2KOwnerType = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property FeesProfileID() As Integer
            Get
                Return nFeesProfileID
            End Get

            Set(ByVal value As Integer)
                nFeesProfileID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property FeesStatus() As Boolean
            Get
                Return bolFeesStatus
            End Get

            Set(ByVal value As Boolean)
                bolFeesStatus = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ComplianceProfileID() As Integer
            Get
                Return nComplianceProfileID
            End Get

            Set(ByVal value As Integer)
                nComplianceProfileID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property ComplianceStatus() As Boolean
            Get
                Return bolCompliaceStatus
            End Get

            Set(ByVal value As Boolean)
                bolCompliaceStatus = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get

            Set(ByVal value As Boolean)
                bolActive = value
                Me.CheckDirty()
            End Set
        End Property

        Public WriteOnly Property ActiveOriginal() As Boolean
            Set(ByVal value As Boolean)
                obolActive = value
            End Set
        End Property

        Public Property FeeActive() As Boolean
            Get
                Return bolFeeActive
            End Get

            Set(ByVal value As Boolean)
                bolFeeActive = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EnsiteOrganizationID() As Integer
            Get
                Return nEnsiteOrganizationID
            End Get

            Set(ByVal value As Integer)
                nEnsiteOrganizationID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property EnsitePersonID() As Integer
            Get
                Return nEnsitePersonID
            End Get

            Set(ByVal value As Integer)
                nEnsitePersonID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property EnsiteAgencyInterestID() As Boolean
            Get
                Return bolEnsiteAgencyInterestID
            End Get

            Set(ByVal value As Boolean)
                bolEnsiteAgencyInterestID = value
                Me.CheckDirty()
            End Set
        End Property

        ' Public Property Description() As Integer
        '    Get
        '       Return nOwnerDesc
        '  End Get

        '  Set(ByVal value As Integer)
        '     nOwnerDesc = Integer.Parse(value)
        '    Me.CheckDirty()
        '  End Set
        '  End Property

        Public Property CustEntityCode() As Integer
            Get
                Return nCustEntityCode
            End Get

            Set(ByVal value As Integer)
                nCustEntityCode = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property CustTypeCode() As Integer
            Get
                Return nCustTypeCode
            End Get

            Set(ByVal value As Integer)
                nCustTypeCode = Integer.Parse(value)
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

        Public Property OwnerL2CSnippet() As Boolean
            Get
                Return bolOwnerL2CSnippet
            End Get

            Set(ByVal value As Boolean)
                bolOwnerL2CSnippet = value
                Me.CheckDirty()
            End Set
        End Property

        Public ReadOnly Property BP2KOwnerID() As String
            Get
                Return strBP2KOwnerID
            End Get
        End Property

        Public Property CapParticipationLevel() As String
            Get
                Return strCapParticipationLevel
            End Get
            Set(ByVal Value As String)
                ostrCapParticipationLevel = Value
                strCapParticipationLevel = Value
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                If bolIsDirty Then Return bolIsDirty
                For Each fac As MUSTER.Info.FacilityInfo In facilityCollection.Values
                    If fac.IsDirty Then
                        Return True
                    End If
                Next
                Return False
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
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
        Public ReadOnly Property ChildrenDirty() As Boolean
            Get
                For Each fac As MUSTER.Info.FacilityInfo In facilityCollection.Values
                    If fac.IsDirty Or fac.ChildrenDirty Then Return True
                Next
                Return False
            End Get
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
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace

