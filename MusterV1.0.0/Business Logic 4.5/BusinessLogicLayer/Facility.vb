'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Facility
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         EN      12/04/04    Original class definition.
'   1.1         AN      12/16/04    Added Address Object
'   1.2         EN      12/29/04    Changed from ArrayList to Datatable. changed the properties Datum,Method,LocationType.
'   1.3         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.4         KJ      01/07/05    Changed the GetAddress to function Retrive as I have removed that function.
'   1.5         EN      01/10/05    Modified the Add method. 
'   1.6         MNR     01/12/05    Modified Retrieve function to handle hirearchy
'   1.7         MNR     01/13/05    Added Events
'   1.8         MNR     01/14/05    Added ValidateData()
'   1.9         JVC2    01/19/05    Added PreviousOwners()
'   2.0         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   2.1         MNR     01/26/05    Added ValidateEmail(..) ValidatePhone(..) Functions
'   2.2         MNR     01/28/05    Implemented ChangeKey functionality
'                                   Modified Retrieve to get FacilityStatus from DB
'   2.3         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Facility" type.
'                                       Also added EntityType attribute to expose the typeID.
'   2.4         AN      02/02/05    Added Comments object
'   2.5         EN      02/10/05    Added GetCAPSTATUS,getCAPTanksandPipesByFacility,FacilityCAPTable  functions.. 
'   2.6         AB      03/04/05    Changed FacilityCAPTable to get its info from a SPROC
'   2.7         MNR     03/10/05    After tanks/pipes are retrieved, setting FacilityPowerOff to true in Pipes realted to the facility if DatePowerOff is not null
'   2.8         AB      03/15/05    Corrected an error in GetCAPStatus
'   2.9         MNR     03/15/05    Added Load Sub
'   3.0         MNR     03/16/05    Removed strSrc from events
'   3.1         KKM     03/18/05    Events for handling local tanksCollection and CommentsCollection are added
'   3.2         MR      04/12/05    Modified Lat Long Validation in Validate Function().
'   3.3         MNR     07/25/05    Added events to handle address changed
'
' Function          Description
' Retrieve (NAME)   Returns the Facility requested by the string arg NAME
' Retrieve(ID)     Returns the Facility requested by the int arg ID
' GetAllInfo(optional owner id, optional ShowDeleted)    Returns an FacilityCollection with all Facility objects
' Add(ID)           Adds the Facility identified by arg ID to the 
'                           internal ReportsCollection
' Add(Name)         Adds the Facility identified by arg NAME to the internal 
'                   FacilityCollection            
' Add(Entity)       Adds the Facility passed as the argument to the internal 
'                           FacilityCollection
' Remove(ID)        Removes the Facility identified by arg ID from the internal 
'                           FacilityCollection
' Remove(NAME)      Removes the Facility identified by arg NAME from the 
'                           internal FacilityCollection
' Flush()            Marshalls all modified/new Onwer Info objects in the
'                    Facility Collection to the repository
' FacilityTable()     Returns a datatable containing all columns for the Facility 
'                           objects in the internal FacilityCollection.
' FacilityCombo()     Returns a two-column datatable containing Name and ID for 
'                           the Facility objects in the internal FacilityCollection.
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
   Public Class pFacility
#Region "Public Events"
        Public Event FlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer)
        Public Event FacilityExists(ByVal MsgStr As String)
        Public Event evtFacilityErr(ByVal MsgStr As String)
        Public Event evtFacilityChanged(ByVal bolValue As Boolean)
        'Public Event evtFacilityCommentsChanged(ByVal bolValue As Boolean)
        Public Event evtFacilitiesChanged(ByVal bolValue As Boolean)
        Public Event evtFacilityValidationErr(ByVal FacID As Integer, ByVal MsgStr As String)
        'Public Event evtFacilitySaved(ByVal bolStatus As Boolean)

        'Public Event evtTankCommentsChanged(ByVal bolValue As Boolean)
        'Public Event evtPipeCommentsChanged(ByVal bolValue As Boolean)
        'Public Event evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String)
        Public Event evtFacilityCAPStatus(ByVal facID As Integer)
        'Public Event evtTankStatusChanged(ByVal oldStat As Integer, ByVal newStat As Integer, ByVal facID As Integer)
        Public Event evtFacilityIDChanged(ByVal oldID As Integer, ByVal newID As Integer)
        'added by kiran
        'Public Event evtFacColOwner(ByVal OwnerID As Integer, ByVal facilityCol As MUSTER.Info.FacilityCollection)
        'Public Event evtTankColFacOwner(ByVal FacId As Integer, ByVal tankCol As MUSTER.Info.TankCollection)
        'Public Event evtCompartmentCol(ByVal TankID As Integer, ByVal CompartmentCol As MUSTER.Info.CompartmentCollection, ByVal FacId As Integer)
        'Public Event evtCommentsCol(ByVal FacId As Integer, ByVal CommentsCol As MUSTER.Info.CommentsCollection)
        'Public Event evtCommentsColTank(ByVal tankID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection, ByVal facID As Integer)
        'Public Event evtPipeCommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal tankID As Integer, ByVal facID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        'end changes
        'Public Event evtFacInfoOwner(ByVal facInfo As MUSTER.Info.FacilityInfo, ByVal strDesc As String)
        'Public Event evtFacInfoFacID(ByVal facID As Integer)
        'Public Event evtTankInfoFac(ByVal tankInfo As MUSTER.Info.TankInfo, ByVal strDesc As String)
        'Public Event evtTankInfoTankID(ByVal tnkID As Integer)
        'Public Event evtCommentInfoFac(ByVal Facid As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo)
        'Public Event evtCompInfoTank(ByVal compartmentInfo As MUSTER.Info.CompartmentInfo, ByVal strDesc As String)
        'Public Event evtPipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String)
        'Public Event evtOwnerInfoFacCol(ByRef colFac As MUSTER.Info.FacilityCollection)
        'Public Event evtOwnerInfoFacColByOwnerID(ByVal ownerID As Integer, ByRef colFac As MUSTER.Info.FacilityCollection)
        'Public Event evtFacilityInfoTankColByFacilityID(ByVal facID As Integer, ByRef colTnk As MUSTER.Info.TankCollection)
        'Public Event evtFacilityChangeKey(ByVal oldID As Integer, ByVal newID As Integer)
        Public Event evtAddressChanged(ByVal bolValue As Boolean)
        Public Event evtAddressesChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private oOwnerInfo As MUSTER.Info.OwnerInfo
        'Private colFacility As MUSTER.Info.FacilityCollection
        Private WithEvents oFacilityInfo As MUSTER.Info.FacilityInfo
        Private WithEvents oFacAddress As MUSTER.BusinessLogic.pAddress
        Private WithEvents oFacTanks As MUSTER.BusinessLogic.pTank
        Private WithEvents oFacClosureEvent As MUSTER.BusinessLogic.pClosureEvent
        Private WithEvents oComments As MUSTER.BusinessLogic.pComments
        Private WithEvents oLustEvents As MUSTER.BusinessLogic.pLustEvent
        Private oFacilityDB As MUSTER.DataAccess.FacilityDB
        Private oProperty As MUSTER.BusinessLogic.pProperty
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private bolShowDeleted As Boolean
        Private oDatum As LookupProperty
        Private oMethod As LookupProperty
        Private oLocationType As LookupProperty
        Private nID As Int64 = -1
        Private MusterException As MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Facility").ID
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing, Optional ByRef FacOwner As MUSTER.Info.OwnerInfo = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            If FacOwner Is Nothing Then
                'colFacility = New MUSTER.Info.FacilityCollection
                oOwnerInfo = New MUSTER.Info.OwnerInfo
            Else
                'colFacility = FacOwner.facilityCollection
                oOwnerInfo = FacOwner
            End If
            oFacilityInfo = New MUSTER.Info.FacilityInfo
            'colFacility = New MUSTER.Info.FacilityCollection
            oFacAddress = New MUSTER.BusinessLogic.pAddress
            oFacilityDB = New MUSTER.DataAccess.FacilityDB
            oFacTanks = New MUSTER.BusinessLogic.pTank(oFacilityInfo)
            oComments = New MUSTER.BusinessLogic.pComments
            oFacClosureEvent = New MUSTER.BusinessLogic.pClosureEvent
            oProperty = New MUSTER.BusinessLogic.pProperty
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Enum FacilityModule

            ALL = -1
            Registration = 0
            Financial = 1
            Compliance = 2
            Technical = 3
            Inspection = 4
            Fees = 5
            Closure = 6

        End Enum



        Public facPics As Collections.ArrayList

        Public ReadOnly Property ModuleModifiedBy(ByVal type As FacilityModule, ByVal id As Integer) As String
            Get
                Return GetModifyByModule(type, id)
            End Get
        End Property

        Public ReadOnly Property ModuleModifiedOn(ByVal type As FacilityModule, ByVal id As Integer) As String
            Get
                Return GetModifyOnModule(type, id)
            End Get
        End Property

        Public Property DesignatedOperator() As String
            Get
                Return oFacilityInfo.DesignatedOperator
            End Get

            Set(ByVal Value As String)
                oFacilityInfo.DesignatedOperator = Value
            End Set
        End Property
        Public Property DesignatedManager() As String
            Get
                Return oFacilityInfo.DesignatedManager
            End Get

            Set(ByVal Value As String)
                oFacilityInfo.DesignatedManager = Value
            End Set
        End Property

        Public Property ID() As Integer
            Get
                Return oFacilityInfo.ID
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.ID = Integer.Parse(value)
            End Set
        End Property
        Public Property AIID() As Integer
            Get
                Return oFacilityInfo.AIID
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.AIID = Integer.Parse(value)
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oFacilityInfo.Name
            End Get
            Set(ByVal value As String)
                oFacilityInfo.Name = value
            End Set
        End Property
        Public Property NameForEnsite() As String
            Get
                Return oFacilityInfo.NameForEnsite
            End Get
            Set(ByVal Value As String)
                oFacilityInfo.NameForEnsite = Value
            End Set
        End Property
        Public Property OwnerID() As Integer
            Get
                Return oFacilityInfo.OwnerID
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.OwnerID = Integer.Parse(value)
            End Set
        End Property
        Public Property AddressID() As Integer
            Get
                Return oFacilityInfo.AddressID
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.AddressID = Integer.Parse(value)

            End Set
        End Property
        'Public Property AddressLine1() As String
        '    Get
        '        Return oFacilityInfo.AddressLine1
        '    End Get

        '    Set(ByVal value As String)
        '        oFacilityInfo.AddressLine1 = value
        '    End Set
        'End Property
        'Public Property AddressLine2() As String
        '    Get
        '        Return oFacilityInfo.AddressLine2
        '    End Get

        '    Set(ByVal value As String)
        '        oFacilityInfo.AddressLine2 = value
        '    End Set
        'End Property
        'Public Property City() As String
        '    Get
        '        Return oFacilityInfo.City
        '    End Get

        '    Set(ByVal value As String)
        '        oFacilityInfo.City = value
        '    End Set
        'End Property
        'Public Property State() As String
        '    Get
        '        Return oFacilityInfo.State
        '    End Get

        '    Set(ByVal value As String)
        '        oFacilityInfo.State = value
        '    End Set
        'End Property
        'Public Property Zip() As String
        '    Get
        '        Return oFacilityInfo.Zip
        '    End Get

        '    Set(ByVal value As String)
        '        oFacilityInfo.Zip = value
        '    End Set
        'End Property
        'Public Property FIPSCode() As String
        '    Get
        '        Return oFacilityInfo.FIPSCode
        '    End Get

        '    Set(ByVal value As String)
        '        oFacilityInfo.FIPSCode = value
        '    End Set
        'End Property
        Public ReadOnly Property Current_MGPTF_Status() As String

            Get
                Return oFacilityInfo.CurrentMGPTFStatus
            End Get

        End Property

        Public Property BillingAddressID() As Integer
            Get
                Return oFacilityInfo.BillingAddressID
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.BillingAddressID = Integer.Parse(value)

            End Set
        End Property
        Public Property LatitudeDegree() As Single

            Get
                Return oFacilityInfo.LatitudeDegree
            End Get
            Set(ByVal value As Single)
                oFacilityInfo.LatitudeDegree = value

            End Set
        End Property
        Public Property LatitudeMinutes() As Single
            Get
                Return oFacilityInfo.LatitudeMinutes
            End Get
            Set(ByVal value As Single)
                oFacilityInfo.LatitudeMinutes = value

            End Set
        End Property
        Public Property LatitudeSeconds() As Double
            Get
                Return oFacilityInfo.LatitudeSeconds
            End Get
            Set(ByVal value As Double)
                oFacilityInfo.LatitudeSeconds = value

            End Set
        End Property
        Public Property LongitudeDegree() As Single
            Get
                Return oFacilityInfo.LongitudeDegree
            End Get
            Set(ByVal value As Single)
                oFacilityInfo.LongitudeDegree = value

            End Set
        End Property
        Public Property LongitudeMinutes() As Single

            Get
                Return oFacilityInfo.LongitudeMinutes
            End Get
            Set(ByVal value As Single)
                oFacilityInfo.LongitudeMinutes = value

            End Set
        End Property
        Public Property LongitudeSeconds() As Double

            Get
                Return oFacilityInfo.LongitudeSeconds
            End Get
            Set(ByVal value As Double)
                oFacilityInfo.LongitudeSeconds = value

            End Set
        End Property
        Public Property Phone() As String

            Get
                Return oFacilityInfo.Phone

            End Get
            Set(ByVal value As String)
                oFacilityInfo.Phone = value

            End Set
        End Property
        Public Property Datum() As Integer
            Get
                Return oFacilityInfo.Datum
            End Get
            Set(ByVal Value As Integer)
                oFacilityInfo.Datum = Value

            End Set
        End Property
        Public Property Method() As Integer
            Get
                Return oFacilityInfo.Method
            End Get
            Set(ByVal Value As Integer)
                oFacilityInfo.Method = Value

            End Set
        End Property
        Public Property Fax() As String

            Get
                Return oFacilityInfo.Fax
            End Get
            Set(ByVal value As String)
                oFacilityInfo.Fax = value

            End Set
        End Property
        Public Property FeesProfileId() As Integer

            Get
                Return oFacilityInfo.FeesProfileId
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.FeesProfileId = Integer.Parse(value)

            End Set
        End Property
        Public Property FacilityType() As Integer

            Get
                Return oFacilityInfo.FacilityType
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.FacilityType = Integer.Parse(value)

            End Set
        End Property
        Public Property FeesStatus() As Integer

            Get
                Return oFacilityInfo.FeesStatus
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.FeesStatus = value

            End Set
        End Property
        Public Property CurrentCIUNumber() As Integer
            Get
                Return oFacilityInfo.CurrentCIUNumber
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.CurrentCIUNumber = Integer.Parse(value)

            End Set
        End Property
        Public Property CapStatus() As Integer
            Get
                Return oFacilityInfo.CapStatus
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.CapStatus = value
            End Set
        End Property
        Public Property CapStatusOriginal() As Integer
            Get
                Return oFacilityInfo.CapStatusOriginal
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.CapStatusOriginal = value
            End Set
        End Property
        Public Property CAPCandidate() As Boolean
            Get
                Return oFacilityInfo.CAPCandidate
            End Get
            Set(ByVal Value As Boolean)
                oFacilityInfo.CAPCandidate = Value

            End Set
        End Property
        Public ReadOnly Property CAPCandidateOriginal() As Boolean
            Get
                Return oFacilityInfo.CAPCandidateOriginal
            End Get
        End Property
        Public Property CitationProfileID() As Integer

            Get
                Return oFacilityInfo.CitationProfileID
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.CitationProfileID = Integer.Parse(value)

            End Set
        End Property
        Public Property CurrentLUSTStatus() As Integer
            Get
                Return oFacilityInfo.CurrentLUSTStatus
            End Get
            Set(ByVal value As Integer)
                oFacilityInfo.CurrentLUSTStatus = value

            End Set
        End Property
        Public Property FuelBrand() As String
            Get
                Return oFacilityInfo.FuelBrand
            End Get
            Set(ByVal Value As String)
                oFacilityInfo.FuelBrand = Value

            End Set
        End Property
        Public Property FacilityDescription() As String
            Get
                Return oFacilityInfo.FacilityDescription
            End Get
            Set(ByVal value As String)
                oFacilityInfo.FacilityDescription = value
            End Set
        End Property
        Public Property SignatureOnNF() As Boolean
            Get
                Return oFacilityInfo.SignatureOnNF
            End Get
            Set(ByVal Value As Boolean)
                oFacilityInfo.SignatureOnNF = Value

            End Set
        End Property
        Public ReadOnly Property SignatureOnNFOriginal() As Boolean
            Get
                Return oFacilityInfo.SignatureOnNFOriginal
            End Get
        End Property
        Public Property DateReceived() As Date
            Get
                Return oFacilityInfo.DateReceived
            End Get
            Set(ByVal Value As Date)
                oFacilityInfo.DateReceived = Value
            End Set
        End Property
        Public Property DateTransferred() As Date
            Get
                Return oFacilityInfo.DateTransferred
            End Get
            Set(ByVal Value As Date)
                oFacilityInfo.DateTransferred = Value
            End Set
        End Property
        Public Property FacilityStatus() As Integer
            Get
                Return oFacilityInfo.FacilityStatus
            End Get
            Set(ByVal Value As Integer)
                oFacilityInfo.FacilityStatus = Value
            End Set
        End Property
        Public ReadOnly Property FacilityStatusDescription() As String
            Get
                Return oProperty.GetPropertyNameByID(FacilityStatus) ' oFacilityInfo.FacilityStatusDescription
            End Get
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oFacilityInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oFacilityInfo.Deleted = value
            End Set
        End Property
        Public Property DatePowerOff() As Date
            Get
                Return oFacilityInfo.DatePowerOff
            End Get
            Set(ByVal Value As Date)
                oFacilityInfo.DatePowerOff = Value
                CheckDatePowerOff()
            End Set
        End Property
        Public Property LocationType() As Integer
            Get
                Return oFacilityInfo.LocationType
            End Get
            Set(ByVal Value As Integer)
                oFacilityInfo.LocationType = Value
            End Set
        End Property
        Public Property UpcomingInstallation() As Boolean
            Get
                Return oFacilityInfo.UpcomingInstallation
            End Get
            Set(ByVal Value As Boolean)
                oFacilityInfo.UpcomingInstallation = Value
            End Set
        End Property
        Public ReadOnly Property UpcomingInstallationOriginal() As Boolean
            Get
                Return oFacilityInfo.UpcomingInstallationOriginal
            End Get
        End Property
        Public Property UpcomingInstallationDate() As Date
            Get
                Return oFacilityInfo.UpcomingInstallationDate
            End Get
            Set(ByVal Value As Date)
                oFacilityInfo.UpcomingInstallationDate = Value
            End Set
        End Property
        Public ReadOnly Property UpcomingInstallationDateOriginal() As Date
            Get
                Return oFacilityInfo.UpcomingInstallationDateOriginal
            End Get
        End Property


        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property
        Public Property LicenseeID() As Integer
            Get
                Return oFacilityInfo.LicenseeID
            End Get
            Set(ByVal Value As Integer)
                oFacilityInfo.LicenseeID = Value
            End Set
        End Property
        Public Property ContractorID() As Integer
            Get
                Return oFacilityInfo.ContractorID
            End Get
            Set(ByVal Value As Integer)
                oFacilityInfo.ContractorID = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFacilityInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oFacilityInfo.IsDirty = value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xFacinfo As MUSTER.Info.FacilityInfo
                For Each xFacinfo In oOwnerInfo.facilityCollection.Values
                    If xFacinfo.IsDirty Then
                        'MsgBox("BLL F:" + xFacinfo.ID.ToString)
                        Return True
                        Exit Property
                    End If
                Next
                'If oFacAddress.colIsDirty Then
                '    Return True
                '    Exit Property
                'End If
                If oFacTanks.colIsDirty Or oFacClosureEvent.colIsDirty Then
                    Return True
                    Exit Property
                End If
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property ShowDeleted() As Boolean
            Get
                Return bolShowDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolShowDeleted = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oFacilityInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFacilityInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFacilityInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFacilityInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFacilityInfo.ModifiedBy = Value
            End Set
        End Property



        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFacilityInfo.ModifiedOn
            End Get
        End Property
        Public Property Facility() As MUSTER.Info.FacilityInfo
            Get
                Return oFacilityInfo
            End Get
            Set(ByVal value As MUSTER.Info.FacilityInfo)
                oFacilityInfo = value
            End Set
        End Property
        Public ReadOnly Property FacilityAddress() As MUSTER.Info.AddressInfo
            Get
                If oFacilityInfo.AddressID > 0 Then
                    Return oFacAddress.Retrieve(oFacilityInfo.AddressID)
                ElseIf oFacAddress.AddressId > 0 Then
                    Return oFacAddress.Retrieve(oFacAddress.AddressId)
                ElseIf Me.AddressID > 0 Then
                    Return oFacAddress.Retrieve(AddressID)
                End If

            End Get
        End Property
        Public ReadOnly Property FacilityBillingAddress() As MUSTER.Info.AddressInfo
            Get
                Return oFacAddress.Retrieve(oFacilityInfo.BillingAddressID)
            End Get
        End Property
        Public ReadOnly Property FacilityAddresses() As MUSTER.BusinessLogic.pAddress
            Get
                Return oFacAddress
            End Get
        End Property
        Public Property FacilityTanks() As MUSTER.BusinessLogic.pTank
            Get
                Return oFacTanks
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pTank)
                oFacTanks = Value
            End Set
        End Property
        Public Property FacilityLustEvents() As MUSTER.BusinessLogic.pLustEvent
            Get
                Return oLustEvents
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pLustEvent)
                oLustEvents = Value
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
        Public Property ClosureEvent() As MUSTER.BusinessLogic.pClosureEvent
            Get
                Return oFacClosureEvent
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pClosureEvent)
                oFacClosureEvent = Value
            End Set
        End Property
        Public Property OwnerInfo() As MUSTER.Info.OwnerInfo
            Get
                Return oOwnerInfo
            End Get
            Set(ByVal Value As MUSTER.Info.OwnerInfo)
                oOwnerInfo = Value
            End Set
        End Property
        Public ReadOnly Property FacilityCollection() As MUSTER.Info.FacilityCollection
            Get
                Return oOwnerInfo.facilityCollection
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub RetrieveAll(ByVal ownerID As Integer, Optional ByVal [Module] As String = "", _
                                Optional ByVal showDeleted As Boolean = False, Optional ByVal facID As Int64 = 0, _
                                Optional ByVal tankID As Int64 = 0, _
                                Optional ByVal intInspectionID As Integer = Integer.MinValue)
            Try
                Dim ds As New DataSet
                ds = oFacilityDB.DBGetDS(ownerID, [Module], showDeleted, facID, tankID, intInspectionID)
                If facID = 0 Then
                    ' 0 - Owner+address
                    ' 1 - Person / Organization
                    ' 2 - Facilities + addresses
                    ' 3 - Tanks
                    ' 4 - Compartments
                    ' 5 - Pipes
                    ds.Tables(0).TableName = "Owner"
                    ds.Tables(1).TableName = "OrgPerson"
                    ds.Tables(2).TableName = "Facilities"
                    ds.Tables(3).TableName = "Tanks"
                    ds.Tables(4).TableName = "Compartments"
                    ds.Tables(5).TableName = "Pipes"
                    ' Facility
                    Load(oOwnerInfo, ds, [Module])
                    ds = Nothing
                Else
                    ' 0 - Facility + address
                    ' 1 - Tanks
                    ' 2 - Compartments
                    ' 3 - Pipes
                    ds.Tables(0).TableName = "Facilities"
                    ds.Tables(1).TableName = "Tanks"
                    ds.Tables(2).TableName = "Compartments"
                    ds.Tables(3).TableName = "Pipes"
                    ' Facility
                    If ds.Tables("Facilities").Rows.Count > 0 Then
                        oFacilityInfo = New MUSTER.Info.FacilityInfo(ds.Tables("Facilities").Rows(0))
                        oOwnerInfo.facilityCollection.Add(oFacilityInfo)
                        oFacAddress.Load(ds.Tables("Facilities").Rows(0))
                        ds.Tables.Remove("Facilities")
                        ' Tank, Comp, Pipe
                        oFacTanks.Load(oFacilityInfo, ds, [Module])
                        ds = Nothing
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Load(ByRef OwnerInfo As MUSTER.Info.OwnerInfo, ByRef ds As DataSet, ByVal [Module] As String)
            Dim dr As DataRow
            oOwnerInfo = OwnerInfo
            Try
                If ds.Tables("Facilities").Rows.Count > 0 Then
                    'oComments.Clear()
                    For Each dr In ds.Tables("Facilities").Rows
                        oFacilityInfo = New MUSTER.Info.FacilityInfo(dr)
                        oFacAddress.Load(dr)
                        oOwnerInfo.facilityCollection.Add(oFacilityInfo)
                        'oComments.Load(ds, [Module], nEntityTypeID, oFacilityInfo.ID)
                        oFacTanks.Load(oFacilityInfo, ds, [Module])
                    Next
                    ds.Tables.Remove("Facilities")
                    ds.Tables.Remove("Tanks")
                    ds.Tables.Remove("Compartments")
                    ds.Tables.Remove("Pipes")
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function SaveTANKCAPData(ByVal isInspection As Boolean, ByVal facID As Integer, ByVal tankID As Integer, ByVal dateSpillTested As Object, _
                                        ByVal dateOverfillTested As Object, ByVal dateTankElecInsp As Object, ByVal dateLastTCP As Object, _
                                        ByVal dateLineInteriorInspect As Object, ByVal dateATGInsp As Object, ByVal dateTTT As Object, ByVal userID As String, ByVal dateLIInstalled As Object, ByVal dateSpillInstalled As Object, ByVal dateOverfillInstalled As Object)
            Try

                Return oFacilityDB.DBSaveTANKCAPData(isInspection, facID, tankID, dateSpillTested, dateOverfillTested, dateTankElecInsp, dateLastTCP, dateLineInteriorInspect, dateATGInsp, dateTTT, userID, dateLIInstalled, dateSpillInstalled, dateOverfillInstalled)

            Catch ex As Exception

                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try


        End Function



        Public Function SavePIPECAPData(ByVal isInspection As Boolean, ByVal facID As Integer, ByVal pipeID As Integer, ByVal tankID As Integer, ByVal dateALLD_Test As Object, _
                                        ByVal dateLTT As Object, ByVal datePipeCPTest As Object, ByVal dateTermCPTest As Object, _
                                        ByVal datePipeShearTest As Object, ByVal datePipeSecInsp As Object, ByVal datePipeElecInsp As Object, ByVal userID As String)

            Try

                Return oFacilityDB.DBSavePIPECAPData(isInspection, facID, pipeID, tankID, dateALLD_Test, dateLTT, datePipeCPTest, dateTermCPTest, datePipeShearTest, datePipeSecInsp, datePipeElecInsp, userID)

            Catch ex As Exception

                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal bolSaveAsInspection As Boolean = False, Optional ByRef AdviceIdUseForTransfers As Integer = 0) As Boolean
            ' have to optimize the code - Manju (1/20/05)
            Dim nNewFacilityId As Integer
            Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oFacilityInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Exit Function
                    End If
                End If
                ' if id is negative (new facility) and deleted is true,
                ' there is no need to enter in the DB
                If Not (oFacilityInfo.ID < 0 And oFacilityInfo.Deleted) Then
                    oldID = oFacilityInfo.ID

                    'Facility Status is dependent on Tank Status (bug 848)
                    'If oFacilityInfo.CurrentLUSTStatus = 0 Then
                    '    oFacilityInfo.FacilityStatus = GetFacilityStatus()
                    'End If
                    'oFacilityInfo.FacilityStatusDescription = GetFacilityStatusDesc()
                    'If oFacilityInfo.Deleted Then
                    '    RaiseEvent evtFacilityIDChanged(oFacilityInfo.ID, 0)
                    'End If
                    Dim bolSigFlagDel As Boolean = False
                    Dim bolSigFlagAdd As Boolean = False

                    Dim strModule As String = String.Empty

                    If moduleID = 612 Then
                        strModule = "REGISTRATION"
                    ElseIf moduleID = 891 Then
                        strModule = "CLOSURE"
                    End If
                    If strModule.ToUpper = "REGISTRATION" Or strModule.ToUpper = "CLOSURE" Then
                        If oFacilityInfo.SignatureOnNFOriginal = False And oFacilityInfo.SignatureOnNF = True Then
                            bolSigFlagDel = True
                        ElseIf (oFacilityInfo.SignatureOnNFOriginal = True And oFacilityInfo.SignatureOnNF = False) Or _
                                (oFacilityInfo.SignatureOnNFOriginal = False And oFacilityInfo.SignatureOnNF = False) Then
                            bolSigFlagAdd = True
                        End If

                        ' Upcoming Install
                    End If
                    If oFacilityInfo.IsDirty Then
                        AdviceIdUseForTransfers = oFacilityDB.Put(oFacilityInfo, moduleID, staffID, returnVal, strUser, AdviceIdUseForTransfers)
                        If Not returnVal = String.Empty Then
                            Exit Function
                        End If
                    End If

                    If bolSigFlagDel Or bolSigFlagAdd Then
                        Dim flags As New MUSTER.BusinessLogic.pFlag
                        Dim flagsCol As MUSTER.Info.FlagsCollection
                        Dim cal As New MUSTER.BusinessLogic.pCalendar
                        Dim flagInfo As MUSTER.Info.FlagInfo

                        If bolSigFlagDel Then
                            flags.RetrieveFlags(oFacilityInfo.ID, 6, , , , , "SYSTEM", "Signature Required Letter due for facility")
                            For Each flagInfo In flags.FlagsCol.Values
                                If flagInfo.CalendarInfoID <> 0 Then
                                    cal.Retrieve(flagInfo.CalendarInfoID)
                                    cal.Completed = True
                                    cal.Deleted = True
                                    cal.Save()
                                End If
                                flagInfo.Deleted = True
                            Next
                            If flags.FlagsCol.Count > 0 Then flags.Flush()
                            RaiseEvent FlagsChanged(oFacilityInfo.ID, 6)
                        End If

                        If bolSigFlagAdd Then
                            flags.RetrieveFlags(oFacilityInfo.ID, 6, , , , , "SYSTEM", "Signature Required Letter due for facility")
                            If flags.FlagsCol.Count <= 0 Then
                                ' create cal entry
                                cal.Add(New MUSTER.Info.CalendarInfo(0, _
                                                Now.AddDays(30).Date, _
                                                Now.AddDays(60).Date, _
                                                0, _
                                                "Facility " + oFacilityInfo.ID.ToString + ": Signature Required Letter due for facility " + oFacilityInfo.Name + " that belongs to owner " & oFacilityInfo.OwnerID.ToString, _
                                                strUser, _
                                                "SYSTEM", _
                                                "C&E", _
                                                True, _
                                                False, _
                                                False, _
                                                False, _
                                                String.Empty, _
                                                CDate("01/01/0001"), _
                                                String.Empty, _
                                                CDate("01/01/0001"), _
                                                6, _
                                                oFacilityInfo.ID))
                                cal.Save()
                                ' create flag
                                flags.Add(New MUSTER.Info.FlagInfo(0, _
                                     oFacilityInfo.ID, _
                                     6, _
                                     "Signature Required Letter due for facility " + oFacilityInfo.Name + " that belongs to owner " & oFacilityInfo.OwnerID.ToString, _
                                     False, _
                                     DateAdd(DateInterval.Day, 30, Now.Date), _
                                     strModule, _
                                     cal.CalendarId, _
                                     String.Empty, _
                                     CDate("01/01/0001"), _
                                     String.Empty, _
                                     CDate("01/01/0001"), _
                                     DateAdd(DateInterval.Day, 60, DateAdd(DateInterval.Day, 30, Now.Date)), _
                                     "SYSTEM"))
                                flags.Save()
                                RaiseEvent FlagsChanged(oFacilityInfo.ID, 6)
                            End If
                        End If
                    End If

                    If Not bolValidated Then
                        If oldID <> oFacilityInfo.ID Then
                            'RaiseEvent evtFacilityIDChanged(oldID, oFacilityInfo.ID)
                            oOwnerInfo.facilityCollection.ChangeKey(oldID, oFacilityInfo.ID)
                            'RaiseEvent evtFacilityChangeKey(oldID, oFacilityInfo.ID)
                        End If
                        'oFacAddress.Flush()
                        'oComments.Flush()
                        'oFacTanks.Flush()
                    End If
                    SetInfoInChild()
                    oFacTanks.Flush(moduleID, staffID, returnVal, strUser, bolSaveAsInspection)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                    oFacilityInfo.Archive()
                    oFacilityInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oFacilityInfo.Deleted Then
                        'If Not bolValidated Then
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oFacilityInfo.ID Then
                            If strPrev = oFacilityInfo.ID Then
                                RaiseEvent evtFacilityErr("Facility " + oFacilityInfo.ID.ToString + " deleted")
                                'RaiseEvent evtFacilityIDChanged(oFacilityInfo.ID, 0)
                                oOwnerInfo.facilityCollection.Remove(oFacilityInfo)
                                'RaiseEvent evtFacInfoOwner(oFacilityInfo, "REMOVE")
                                'If bolDelete Then
                                '    oFacilityInfo = New MUSTER.Info.FacilityInfo
                                'Else
                                '    oFacilityInfo = Me.Retrieve(0, , "FACILITY")
                                'End If
                                If Not bolDelete Then
                                    oFacilityInfo = Me.Retrieve(oOwnerInfo, 0, , "FACILITY")
                                End If
                                RaiseEvent evtFacilityChanged(False)
                            Else
                                RaiseEvent evtFacilityErr("Facility " + oFacilityInfo.ID.ToString + " deleted")
                                'RaiseEvent evtFacilityIDChanged(oFacilityInfo.ID, 0)
                                oOwnerInfo.facilityCollection.Remove(oFacilityInfo)
                                'RaiseEvent evtFacInfoOwner(oFacilityInfo, "REMOVE")
                                oFacilityInfo = Me.Retrieve(oOwnerInfo, strPrev, , "FACILITY")
                            End If
                        Else
                            RaiseEvent evtFacilityErr("Facility " + oFacilityInfo.ID.ToString + " deleted")
                            'RaiseEvent evtFacilityIDChanged(oFacilityInfo.ID, 0)
                            oOwnerInfo.facilityCollection.Remove(oFacilityInfo)
                            'RaiseEvent evtFacInfoOwner(oFacilityInfo, "REMOVE")
                            oFacilityInfo = Me.Retrieve(oOwnerInfo, strNext, , "FACILITY")
                        End If
                        'Else
                        'RaiseEvent evtFacilityErr("Facility " + oFacilityInfo.ID.ToString + " deleted")
                        'colFacility.Remove(oFacilityInfo)
                        'End If
                    End If
                End If
                RaiseEvent evtFacilityChanged(Me.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
                Return False
            End Try
        End Function
        Public Function Retrieve(ByRef OwnerInfo As MUSTER.Info.OwnerInfo, ByVal id As Integer, Optional ByVal strDepth As String = "SELF", Optional ByVal strIDInfo As String = "OWNER", Optional ByVal showDeleted As Boolean = False, Optional ByVal bolLoading As Boolean = False) As MUSTER.Info.FacilityInfo
            Dim bolDataAged As Boolean = False
            'Dim oTankInfoLocal As MUSTER.Info.TankInfo
            'Dim oCompInfoLocal As MUSTER.Info.CompartmentInfo
            'Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
            Dim ownerID As Integer
            Try
                oOwnerInfo = OwnerInfo
                If oFacilityInfo.ID < 0 And _
                    Not oFacilityInfo.IsDirty And _
                    id = 0 Then
                    Exit Try
                End If
                If Not bolLoading Then ' Not bolValidationErrorOccurred Or 
                    If Not oFacilityInfo.Deleted And oFacilityInfo.IsDirty Then
                        Me.ValidateData()
                    End If
                End If

                Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
                ' Choose how deep you want to retrieve the info
                'Dim dtAddressID As DataTable = GetDataTable("tblReg_Facility where Deleted = 0 and Facility_ID = " + oFacilityInfo.ID)
                Dim dsAddressID As DataSet
                Dim strAddressID As String
                If oFacilityInfo.ID > 0 And oFacilityInfo.AddressID > 0 Then
                    strAddressID = "Select Address_ID from tblReg_Facility where Deleted = 0 and Facility_ID = " + oFacilityInfo.ID.ToString
                    dsAddressID = oFacAddress.GetDataSet(strAddressID)
                    oFacilityInfo.AddressID = dsAddressID.Tables(0).Rows(0).Item("Address_ID")
                End If
                Select Case UCase(strIDInfo).Trim
                    Case "FACILITY"
                        ownerID = oOwnerInfo.ID
                        Select Case UCase(strDepth).Trim
                            Case "SELF"
                                If id = 0 Then
                                    Add(New MUSTER.Info.FacilityInfo)
                                Else
                                    oFacilityInfo = oOwnerInfo.facilityCollection.Item(id)
                                    ' Check to see if data is old.  
                                    ' If yes, remove from collection and get new
                                    If Not (oFacilityInfo Is Nothing) Then
                                        If oFacilityInfo.IsAgedData = True And oFacilityInfo.IsDirty = False Then
                                            bolDataAged = True
                                            oOwnerInfo.facilityCollection.Remove(oFacilityInfo)
                                            'RaiseEvent evtFacInfoOwner(oFacilityInfo, "REMOVE")
                                        End If
                                    End If
                                    If bolDataAged Then
                                        oOwnerInfo.facilityCollection.Remove(oFacilityInfo)
                                        RetrieveAll(ownerID, , showDeleted, id, )
                                    ElseIf oFacilityInfo Is Nothing Then
                                        Add(id, showDeleted)
                                        'Else
                                        'RetrieveAll(ownerID, , showDeleted, id, )
                                    End If
                                    'If oFacilityInfo Is Nothing Or bolDataAged Then
                                    '    Add(id, showDeleted)
                                    'End If
                                End If
                                oFacAddress.Retrieve(oFacilityInfo.AddressID, "SELF", showDeleted, bolLoading)
                            Case Else
                                oFacilityInfo = oOwnerInfo.facilityCollection.Item(id)
                                ' Check to see if data is old.  
                                ' If yes, remove from collection and get new
                                If Not (oFacilityInfo Is Nothing) Then
                                    If oFacilityInfo.IsAgedData = True And oFacilityInfo.IsDirty = False Then
                                        bolDataAged = True
                                        oOwnerInfo.facilityCollection.Remove(oFacilityInfo)
                                        ' RaiseEvent evtFacInfoOwner(oFacilityInfo, "REMOVE")
                                    End If
                                End If

                                If bolDataAged Then
                                    RetrieveAll(ownerID, , showDeleted, id, )
                                ElseIf oFacilityInfo Is Nothing Then
                                    Add(id, showDeleted)
                                    'Else
                                    'RetrieveAll(ownerID, , showDeleted, id, )
                                End If
                                'If oFacilityInfo Is Nothing Or bolDataAged Then
                                '    Add(id, showDeleted)
                                'End If
                                oFacAddress.Retrieve(oFacilityInfo.AddressID, "SELF", showDeleted, bolLoading)
                                If Not bolDataAged Then
                                    oFacTanks.Retrieve(oFacilityInfo, oFacilityInfo.ID, , showDeleted)
                                End If
                        End Select
                    Case "OWNER"
                        ' added by kumar on sep 28th
                        ownerID = OwnerInfo.ID
                        'Dim bolFacilityRetrieved As Boolean = False
                        ' check in collection
                        'oFacilityInfo = oOwnerInfo.facilityCollection.Item(id)
                        'Dim colFacilityContained As New MUSTER.Info.FacilityCollection
                        'RaiseEvent evtOwnerInfoFacCol(colFacilityContained)
                        For Each oFacilityInfoLocal In oOwnerInfo.facilityCollection.Values
                            If oFacilityInfoLocal.OwnerID = id Then
                                ' Check to see if data is old.  If yes, remove from collection and get new
                                If oFacilityInfoLocal.IsAgedData = True And oFacilityInfoLocal.IsDirty = False Then
                                    bolDataAged = True
                                    Exit For
                                Else
                                    oFacilityInfo = oFacilityInfoLocal
                                    'oFacilityInfo.FacilityStatus = GetFacilityStatus()
                                    'oFacilityInfo.FacilityStatusDescription = GetFacilityStatusDesc()
                                    oFacAddress.Retrieve(oFacilityInfo.AddressID, "SELF", showDeleted)
                                    If UCase(strDepth).Trim <> "SELF" Then
                                        oFacTanks.Retrieve(oFacilityInfo, oFacilityInfo.ID, , showDeleted)
                                        'oFacilityInfo.FacilityStatus = GetFacilityStatus()
                                        'oFacilityInfo.FacilityStatusDescription = GetFacilityStatusDesc()
                                    End If
                                    Exit Select
                                End If
                            End If
                        Next
                        ' if the data is old, remove from the collection and get new
                        If bolDataAged = True Then
                            oOwnerInfo.facilityCollection.Remove(oFacilityInfo)
                        End If
                        RetrieveAll(ownerID, , showDeleted, , )
                        oFacAddress.Retrieve(oFacilityInfo.AddressID, "SELF", showDeleted)
                        If Not bolDataAged Then
                            oFacTanks.Retrieve(oFacilityInfo, oFacilityInfo.ID, , showDeleted)
                        End If
                        ' get from DB
                        'If Not bolFacilityRetrieved Then
                        'Dim colFacilityLocal As New MUSTER.Info.FacilityCollection
                        'oOwnerInfo.facilityCollection = oFacilityDB.DBGetByOwnerID(id, showDeleted)
                        'If oOwnerInfo.facilityCollection.Count > 0 Then
                        '    'added by kiran
                        '    'RaiseEvent evtFacColOwner(id, colFacilityLocal)
                        '    'end changes
                        '    For Each oFacilityInfoLocal In oOwnerInfo.facilityCollection.Values
                        '        'oFacilityInfo = oFacilityInfoLocal
                        '        'oOwnerInfo.facilityCollection.Add(oFacilityInfo)
                        '        oFacAddress.Retrieve(oFacilityInfo.AddressID, "SELF", showDeleted)
                        '        If UCase(strDepth).Trim <> "SELF" Then
                        '            oFacTanks.Retrieve(oFacilityInfo, oFacilityInfo.ID, , showDeleted)

                        '            For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                        '                If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                        '                    oTankInfoLocal.FacilityPowerOff = True
                        '                Else
                        '                    oTankInfoLocal.FacilityPowerOff = False
                        '                End If
                        '                For Each oCompInfoLocal In oTankInfoLocal.CompartmentCollection.Values
                        '                    For Each oPipeInfoLocal In oTankInfoLocal.pipesCollection.Values
                        '                        If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                        '                            oPipeInfoLocal.FacilityPowerOff = True
                        '                        Else
                        '                            oPipeInfoLocal.FacilityPowerOff = False
                        '                        End If
                        '                    Next
                        '                Next
                        '            Next
                        '        End If
                        '    Next
                        'End If
                    Case Else
                        RaiseEvent evtFacilityErr("Pass correct param for Facility retrieve")
                End Select
                Dim facStatus = GetFacilityStatus()

                If oFacilityInfo.FacilityStatus <> facStatus Then
                    oFacilityInfo.FacilityStatusOriginal = facStatus
                    oFacilityInfo.FacilityStatus = facStatus
                    oFacilityDB.PutFacStatus(oFacilityInfo.ID, oFacilityInfo.FacilityStatus)
                End If
                oFacilityInfo.CurrentLUSTStatus = GetFacilityLustStatus()
                'oFacilityInfo.CapStatus = GetCAPSTATUS(oFacilityInfo.ID, False)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'oComments.Clear()
            'oComments.GetByModule("", nEntityTypeID, oFacilityInfo.ID)
            RaiseEvent evtFacilityChanged(oFacilityInfo.IsDirty)
            SetInfoInChild()
            Return oFacilityInfo
        End Function
        Public Function DeleteFacility(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String) As Boolean
            ' check if facility has any tanks in coll
            Try
                'Dim oTankLocal As MUSTER.Info.TankInfo
                'If (oFacilityInfo.TankCollection.Count > 0) Then
                '    RaiseEvent evtFacilityErr("The Specified facility has associated Tank(s). Delete Tank(s) before deleting the facility")
                '    Return False
                'Else
                '    For Each oTankLocal In oFacilityInfo.TankCollection.Values
                '        If oTankLocal.TankId > 0 Then
                '            RaiseEvent evtFacilityErr("The Specified facility has associated Tank(s). Delete Tank(s) before deleting the facility")
                '            Return False
                '        End If
                '    Next
                'End If

                If oFacilityInfo.ID > 0 Then
                    Dim ds As DataSet
                    ds = oFacilityDB.DBGetDS("EXEC spCheckDependancy NULL," + oFacilityInfo.ID.ToString + ",NULL,0,NULL")
                    If ds.Tables(0).Rows(0)("EXISTS") Then
                        RaiseEvent evtFacilityErr(IIf(ds.Tables(0).Rows(0)("MSG") Is DBNull.Value, "Facility has dependants", ds.Tables(0).Rows(0)("MSG")))
                        Return False
                    End If
                End If

                ' facility does not have dependents, delete facility
                oFacilityInfo.Deleted = True
                Return Me.Save(moduleID, staffID, returnVal, "", True, True)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return False
            End Try
        End Function
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            ' to be completed according to DDD specs for registration / technical ... - Manju 01/14/05
            Try
                Dim errStr As String = ""
                Dim validateSuccess As Boolean = True
                Select Case [module]
                    Case "Registration"
                        'req info
                        ' Name, address, city, zip, county
                        ' datum (if latitude & longitude specified)
                        ' method (if latitude & longitude specified)
                        ' type (if latitude & longitude specified)
                        ' SIC (if facility AI ID is specified)
                        ' signature due by date (if signature received is not checked)
                        If oFacilityInfo.ID <> 0 Then
                            If oFacilityInfo.ID < 0 And _
                                Not oFacilityInfo.IsDirty Then
                                'bolValidationErrorOccurred = False
                                validateSuccess = True
                                Exit Select
                            End If
                            If oFacilityInfo.AddressID <= 0 Then
                                errStr += "Facility Address is required" + vbCrLf
                                validateSuccess = False
                            End If
                            If oFacilityInfo.ID < 0 And _
                                oFacilityInfo.Name = String.Empty And _
                                oFacilityInfo.AddressID = 0 And _
                                oFacilityInfo.IsDirty Then
                                errStr += "Required fields cannot be empty" + vbCrLf
                                'bolValidationErrorOccurred = True
                                validateSuccess = False
                                Exit Select
                            End If
                            If oFacilityInfo.Phone <> String.Empty Then
                                If Not ValidatePhone(oFacilityInfo.Phone) Then
                                    errStr += "Facility Phone Validation Failed" + vbCrLf
                                    'bolValidationErrorOccurred = True
                                    validateSuccess = False
                                End If
                            End If
                            If oFacilityInfo.Fax <> String.Empty Then
                                If Not ValidatePhone(oFacilityInfo.Fax) Then
                                    errStr += "Facility Fax Validation Failed" + vbCrLf
                                    'bolValidationErrorOccurred = True
                                    validateSuccess = False
                                End If
                            End If
                            If oFacilityInfo.UpcomingInstallation Then
                                If Date.Compare(oFacilityInfo.UpcomingInstallationDate, CDate("01/01/0001")) = 0 Then
                                    errStr += "Facility Upcoming Installation Date is required" + vbCrLf
                                    'bolValidationErrorOccurred = True
                                    validateSuccess = False
                                End If
                            End If
                            If Not validateSuccess Then
                                Exit Select
                            End If
                            If oFacilityInfo.Name <> String.Empty Then
                                If oFacilityInfo.AddressID <> 0 Then 'If oFacAddress.ValidateData() Then 
                                    If (oFacilityInfo.LatitudeDegree <> -1.0 Or _
                                        oFacilityInfo.LatitudeMinutes <> -1.0 Or _
                                        oFacilityInfo.LatitudeSeconds <> -1.0) Or _
                                        (oFacilityInfo.LongitudeDegree <> -1.0 Or _
                                        oFacilityInfo.LongitudeMinutes <> -1.0 Or _
                                        oFacilityInfo.LongitudeSeconds <> -1.0) Then

                                        If oFacilityInfo.LatitudeSeconds <= 99.0 And oFacilityInfo.LongitudeSeconds <= 99.0 Then
                                            ' if latitude and longitude is specified
                                            If oFacilityInfo.Datum <> 0 Then
                                                If oFacilityInfo.Method <> 0 Then
                                                    If oFacilityInfo.LocationType <> 0 Then
                                                        If Date.Compare(oFacilityInfo.DateReceived, CDate("01/01/0001")) = 0 Then
                                                            errStr += "Date Received cannot be empty" + vbCrLf
                                                            'bolValidationErrorOccurred = True
                                                            validateSuccess = False
                                                        Else
                                                            validateSuccess = True
                                                        End If
                                                    Else
                                                        errStr += "Location Type cannot be empty" + vbCrLf
                                                        'bolValidationErrorOccurred = True
                                                        validateSuccess = False
                                                    End If
                                                Else
                                                    errStr += "Method cannot be empty" + vbCrLf
                                                    'bolValidationErrorOccurred = True
                                                    validateSuccess = False
                                                End If
                                            Else
                                                errStr += "Datum cannot be empty" + vbCrLf
                                                'bolValidationErrorOccurred = True
                                                validateSuccess = False
                                            End If
                                        Else
                                            errStr += "Lat Long Seconds must be less than or equal to 99." + vbCrLf
                                            validateSuccess = False
                                        End If
                                    End If
                                Else
                                    errStr += "Facility Address Validate Failed" + vbCrLf
                                    'bolValidationErrorOccurred = True
                                    validateSuccess = False
                                End If
                            Else
                                errStr += "Facility Name cannot be empty" + vbCrLf
                                'bolValidationErrorOccurred = True
                                validateSuccess = False
                            End If
                        End If
                        Exit Select
                        'Case "Technical"
                End Select
                If errStr.Length > 0 And Not validateSuccess Then
                    RaiseEvent evtFacilityValidationErr(oFacilityInfo.ID, errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub SetInfoInChild()
            If Not oFacTanks Is Nothing Then
                oFacTanks.FacilityInfo = oFacilityInfo
            End If
            If Not oFacClosureEvent Is Nothing Then
                oFacClosureEvent.FacilityInfo = oFacilityInfo
            End If
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAllInfo(Optional ByVal OwnerID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FacilityCollection
            Try
                oOwnerInfo.facilityCollection.Clear()
                oOwnerInfo.facilityCollection = oFacilityDB.DBGetAllInfo(OwnerID, showDeleted)
                'RaiseEvent evtFacColOwner(OwnerID, oOwnerInfo.facilityCollection)
                If oOwnerInfo.facilityCollection.Count > 0 Then
                    oFacilityInfo = oOwnerInfo.facilityCollection(oOwnerInfo.facilityCollection.GetKeys(0))
                End If
                Return oOwnerInfo.facilityCollection
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Int64, Optional ByVal showDeleted As Boolean = False)
            Try
                oFacilityInfo = oFacilityDB.DBGetByID(ID, showDeleted)
                If oFacilityInfo.ID = 0 Then
                    oFacilityInfo.ID = nID
                    oFacilityInfo.OwnerID = oOwnerInfo.ID
                    nID -= 1
                End If
                oOwnerInfo.facilityCollection.Add(oFacilityInfo)
                SetInfoInChild()
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as called for by Name
        Public Sub Add(ByVal Name As String)
            Try
                oFacilityInfo = oFacilityDB.DBGetByName(Name)
                If oFacilityInfo.ID = 0 Then
                    oFacilityInfo.ID = nID
                    oFacilityInfo.OwnerID = oOwnerInfo.ID
                    nID -= 1
                End If
                oOwnerInfo.facilityCollection.Add(oFacilityInfo)
                SetInfoInChild()
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oFacility As MUSTER.Info.FacilityInfo)
            Try
                oFacilityInfo = oFacility
                If oFacilityInfo.ID = 0 Then
                    oFacilityInfo.ID = nID
                    oFacilityInfo.OwnerID = oOwnerInfo.ID
                    nID -= 1
                End If
                oOwnerInfo.facilityCollection.Add(oFacilityInfo)
                SetInfoInChild()
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Int64)
            Dim myIndex As Int16 = 1
            Dim oFacInf As MUSTER.Info.FacilityInfo
            Try
                oFacInf = oOwnerInfo.facilityCollection.Item(ID)
                If Not (oFacInf Is Nothing) Then
                    oOwnerInfo.facilityCollection.Remove(oFacInf)
                    'RaiseEvent evtFacInfoOwner(oFacInf, "REMOVE")
                    Exit Sub
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Facility " & ID.ToString & " is not in the collection of Facilities.")
        End Sub
        'Removes the entity called for by Name from the collection
        Public Sub Remove(ByVal Name As String)
            Dim myIndex As Int16 = 1
            Dim oFacInf As MUSTER.Info.FacilityInfo
            Try
                oFacInf = oOwnerInfo.facilityCollection.Item(Name)
                If Not (oFacInf Is Nothing) Then
                    oOwnerInfo.facilityCollection.Remove(oFacInf)
                    'RaiseEvent evtFacInfoOwner(oFacInf, "REMOVE")
                    Exit Sub
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Facility " & ID.ToString & " is not in the collection of Facilities.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFacInf As MUSTER.Info.FacilityInfo)
            Try
                oOwnerInfo.facilityCollection.Remove(oFacInf)
                'RaiseEvent evtFacInfoOwner(oFacInf, "REMOVE")
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Facility " & oFacInf.Name & " is not in the collection of Facilities.")
        End Sub
        Public Function Items() As MUSTER.Info.FacilityCollection
            Return oOwnerInfo.facilityCollection
        End Function
        Public Function Values() As ICollection
            Return oOwnerInfo.facilityCollection.Values
        End Function
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolSaveAsInspection As Boolean = False)
            Dim IDs As New Collection
            Dim delIDs As New Collection
            Dim index As Integer
            Dim xFacInf As MUSTER.Info.FacilityInfo
            Try
                For Each xFacInf In oOwnerInfo.facilityCollection.Values
                    If xFacInf.IsDirty Then
                        oFacilityInfo = xFacInf
                        If oFacilityInfo.Deleted Then
                            If oFacilityInfo.ID < 0 Then
                                delIDs.Add(oFacilityInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, strUser, , , bolSaveAsInspection)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Me.ValidateData Then
                                If oFacilityInfo.ID < 0 Then
                                    IDs.Add(oFacilityInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, strUser, True, , bolSaveAsInspection)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            Else : Exit For
                            End If
                        End If
                    ElseIf xFacInf.ID > 0 And xFacInf.ChildrenDirty Then
                        oFacilityInfo = xFacInf
                        SetInfoInChild()
                        oFacTanks.Flush(moduleID, staffID, returnVal, strUser, bolSaveAsInspection)
                        If Not returnVal = String.Empty Then
                            Exit Sub
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        xFacInf = oOwnerInfo.facilityCollection.Item(CType(delIDs.Item(index), String))
                        oOwnerInfo.facilityCollection.Remove(xFacInf)
                        'RaiseEvent evtFacInfoOwner(xFacInf, "REMOVE")
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xFacInf = oOwnerInfo.facilityCollection.Item(colKey)
                        oOwnerInfo.facilityCollection.ChangeKey(colKey, xFacInf.ID)
                        'RaiseEvent evtFacInfoFacID(colKey)
                        'RaiseEvent evtFacInfoOwner(xFacInf, "ADD")
                        oOwnerInfo.facilityCollection.Remove(colKey)
                        oOwnerInfo.facilityCollection.Add(xFacInf)
                    Next
                End If
                'oFacAddress.Flush()
                'oComments.Flush()
                'oFacTanks.Flush(strUser, strModule)
                RaiseEvent evtFacilitiesChanged(oFacilityInfo.IsDirty)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Function GetFacilityStatus() As Integer
            'Return oFacTanks.GetTanksStatus(oFacilityInfo)
            Return oFacilityDB.DBGetFacStatus(oFacilityInfo.ID)
        End Function
        Private Function GetFacilityLustStatus() As Integer
            'Return oFacTanks.GetTanksStatus(oFacilityInfo)
            Return oFacilityDB.DBGetLustStatus(oFacilityInfo.ID)
        End Function
        Public Function GetOwnerFacilityStatus(ByRef OwnerInfo As MUSTER.Info.OwnerInfo) As Boolean
            ' 514   Active()
            ' 424   CIU
            Dim bolStatus As Boolean = False
            oOwnerInfo = OwnerInfo
            Try
                ' if atleast one facility is active, return active is true
                Dim xFacInf As MUSTER.Info.FacilityInfo
                Dim xTnkInf As MUSTER.Info.TankInfo
                'Dim colFacilityContained As MUSTER.Info.FacilityCollection
                'RaiseEvent evtOwnerInfoFacCol(colFacilityContained)
                For Each xFacInf In oOwnerInfo.facilityCollection.Values
                    If xFacInf.FacilityStatus = 514 Then
                        bolStatus = True
                        Exit Try
                    End If
                    For Each xTnkInf In xFacInf.TankCollection.Values
                        If xTnkInf.TankStatus = 424 Then
                            bolStatus = True
                            Exit Try
                        End If
                    Next
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Return bolStatus
        End Function
        'Public Function GetCAPSTATUS(ByVal nFacId As Integer, Optional ByVal BolEvent As Boolean = True) As Integer
        '    'Added By Elango on Feb 6 2005 
        '    ' 510	Confirmed Release
        '    ' 511   Unconfirmed(Release)
        '    ' 512   Closed(Release)
        '    ' 513   Unregulated()
        '    ' 514   Active()
        '    ' 515   Closed()
        '    ' 516   Pre(88)
        '    ' 424   Currently In User (CIU)
        '    ' 425   Temporarily Out of Service (TOS)
        '    ' 426   Permanently Out of Use (POU)
        '    ' 429   Temporarily Out of Service Indefinitely (TOSI)

        '    'Dim nTankCount As Integer = 0
        '    'Dim nPipeCount As Integer = 0
        '    'Dim nTotalTankCount As Integer = 0

        '    Dim nStatus As Integer
        '    Dim bolCapStatusValid As Boolean = False
        '    Dim oTankInfoLocal As MUSTER.Info.TankInfo
        '    Dim oPipeInfoLocal As MUSTER.Info.PipeInfo

        '    Try
        '        If oFacilityInfo.ID <> nFacId Then
        '            oFacilityInfo = Me.Retrieve(oOwnerInfo, nFacId, "CHILD", "FACILITY")
        '        End If
        '        If oFacilityInfo.CAPCandidate = False Then
        '            Return 0
        '        End If
        '        If oFacilityInfo.TankCollection.Count > 0 Then
        '            For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
        '                If bolCapStatusValid Then Exit For
        '                If oTankInfoLocal.TankIndex <> 0 Then
        '                    ' condition 1
        '                    If (oTankInfoLocal.TankStatus = 424 Or oTankInfoLocal.TankStatus = 429) And _
        '                        (oTankInfoLocal.TankModDesc = 412 Or oTankInfoLocal.TankModDesc = 415 Or oTankInfoLocal.TankModDesc = 475) Then
        '                        If (Date.Compare(Now.Date, DateAdd(DateInterval.Year, 3, oTankInfoLocal.LastTCPDate)) > 0 And _
        '                            Date.Compare(DateAdd(DateInterval.Day, 90, Now.Date), DateAdd(DateInterval.Year, 3, oTankInfoLocal.LastTCPDate)) <= 0) Or _
        '                            Date.Compare(oTankInfoLocal.LastTCPDate, CDate("01/01/0001")) <> 0 Then
        '                            bolCapStatusValid = True
        '                            Exit For
        '                        End If
        '                    End If
        '                    ' condition 2
        '                    If oTankInfoLocal.TankStatus = 424 And oTankInfoLocal.TankModDesc = 476 Then
        '                        If (Date.Compare(Now.Date, DateAdd(DateInterval.Year, 5, oTankInfoLocal.LinedInteriorInspectDate)) < 0 And _
        '                            Date.Compare(DateAdd(DateInterval.Day, 90, Now.Date), DateAdd(DateInterval.Year, 5, oTankInfoLocal.LinedInteriorInspectDate)) <= 0) Or _
        '                            Date.Compare(oTankInfoLocal.LinedInteriorInspectDate, CDate("01/01/0001")) <> 0 Or _
        '                            Date.Compare(oTankInfoLocal.LinedInteriorInstallDate, CDate("01/01/0001")) <> 0 Then
        '                            bolCapStatusValid = True
        '                            Exit For
        '                        End If
        '                    End If
        '                    ' condition 3
        '                    If oTankInfoLocal.TankStatus = 424 And oTankInfoLocal.TankLD = 338 Then
        '                        If (Date.Compare(Now.Date, DateAdd(DateInterval.Year, 5, oTankInfoLocal.TTTDate)) < 0 And _
        '                            Date.Compare(DateAdd(DateInterval.Day, 90, Now.Date), DateAdd(DateInterval.Year, 5, oTankInfoLocal.TTTDate)) <= 0) Or _
        '                            Date.Compare(oTankInfoLocal.TTTDate, CDate("01/01/0001")) <> 0 Then
        '                            bolCapStatusValid = True
        '                            Exit For
        '                        End If
        '                    End If
        '                    ' pipe
        '                    For Each oPipeInfoLocal In oTankInfoLocal.pipesCollection.Values
        '                        If bolCapStatusValid Then Exit For
        '                        If oPipeInfoLocal.Index <> 0 Then
        '                            ' condition 1
        '                            If (oPipeInfoLocal.PipeStatusDesc = 424 Or oPipeInfoLocal.PipeStatusDesc = 429) And _
        '                                oPipeInfoLocal.PipeModDesc = 260 Then
        '                                If (Date.Compare(Now.Date, DateAdd(DateInterval.Year, 3, oPipeInfoLocal.PipeCPTest)) < 0 And _
        '                                    Date.Compare(DateAdd(DateInterval.Day, 90, Now.Date), DateAdd(DateInterval.Year, 3, oPipeInfoLocal.PipeCPTest)) <= 0) Or _
        '                                    Date.Compare(oPipeInfoLocal.PipeCPTest, CDate("01/01/0001")) <> 0 Then
        '                                    bolCapStatusValid = True
        '                                    Exit For
        '                                End If
        '                            End If
        '                            ' condition 2
        '                            If oPipeInfoLocal.PipeStatusDesc = 424 And oPipeInfoLocal.ALLDType = 496 Then
        '                                If (Date.Compare(Now.Date, DateAdd(DateInterval.Year, 1, oPipeInfoLocal.ALLDTestDate)) < 0 And _
        '                                    Date.Compare(DateAdd(DateInterval.Day, 90, Now.Date), DateAdd(DateInterval.Year, 1, oPipeInfoLocal.ALLDTestDate)) <= 0) Or _
        '                                    Date.Compare(oPipeInfoLocal.ALLDTestDate, CDate("01/01/0001")) <> 0 Then
        '                                    bolCapStatusValid = True
        '                                    Exit For
        '                                End If
        '                            End If
        '                            ' condition 3
        '                            If oPipeInfoLocal.PipeStatusDesc = 424 And oPipeInfoLocal.PipeLD = 245 And oPipeInfoLocal.PipeTypeDesc = 268 Then
        '                                If (Date.Compare(Now.Date, DateAdd(DateInterval.Year, 3, oPipeInfoLocal.LTTDate)) < 0 And _
        '                                    Date.Compare(DateAdd(DateInterval.Day, 90, Now.Date), DateAdd(DateInterval.Year, 3, oPipeInfoLocal.LTTDate)) <= 0) Or _
        '                                    Date.Compare(oPipeInfoLocal.LTTDate, CDate("01/01/0001")) <> 0 Then
        '                                    bolCapStatusValid = True
        '                                    Exit For
        '                                End If
        '                            End If
        '                            ' condition 4
        '                            If (oPipeInfoLocal.PipeStatusDesc = 424 Or oPipeInfoLocal.PipeStatusDesc = 429) And _
        '                                (oPipeInfoLocal.TermCPTypeDisp = 611 Or oPipeInfoLocal.TermCPTypeDisp = 488) And _
        '                                (oPipeInfoLocal.TermCPTypeTank = 610 Or oPipeInfoLocal.TermCPTypeTank = 481) Then
        '                                If (Date.Compare(Now.Date, DateAdd(DateInterval.Year, 3, oPipeInfoLocal.TermCPLastTested)) < 0 And _
        '                                    Date.Compare(DateAdd(DateInterval.Day, 90, Now.Date), DateAdd(DateInterval.Year, 3, oPipeInfoLocal.TermCPLastTested)) <= 0) Or _
        '                                    Date.Compare(oPipeInfoLocal.TermCPLastTested, CDate("01/01/0001")) <> 0 Then
        '                                    bolCapStatusValid = True
        '                                    Exit For
        '                                End If
        '                            End If

        '                        End If

        '                    Next

        '                End If

        '            Next

        '        End If

        '        nStatus = IIf(bolCapStatusValid, 1, 0)
        '        If oFacilityInfo.CapStatus <> nStatus Then
        '            If BolEvent Then
        '                oFacilityInfo.CapStatus = nStatus
        '                RaiseEvent evtFacilityCAPStatus(oFacilityInfo.OwnerID, oFacilityInfo.ID)
        '            Else
        '                If oFacilityInfo.ID > 0 Then
        '                    oFacilityInfo.CapStatusOriginal = nStatus
        '                    oFacilityInfo.CapStatus = nStatus
        '                    oFacilityDB.PutCAPStatus(oFacilityInfo.ID, oFacilityInfo.CapStatus)
        '                End If
        '            End If
        '        End If
        '        Return nStatus

        '        'For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
        '        '    nTotalTankCount += 1
        '        '    If (Not IsDBNull(oTankInfoLocal.TankIndex) And (oTankInfoLocal.TankStatus = 424 Or oTankInfoLocal.TankStatus = 429) And (oTankInfoLocal.TankModDesc = 412 Or oTankInfoLocal.TankModDesc = 415 Or oTankInfoLocal.TankModDesc = 475) And _
        '        '         (DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 3, oTankInfoLocal.LastTCPDate)) < 0 Or IsDBNull(oTankInfoLocal.LastTCPDate)) _
        '        '        Or (Not IsDBNull(oTankInfoLocal.TankIndex) And oTankInfoLocal.TankStatus = 424 And (oTankInfoLocal.TankModDesc = 476 Or oTankInfoLocal.TankModDesc = 475)) And _
        '        '           (DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 5, oTankInfoLocal.LinedInteriorInspectDate)) < 0 Or IsDBNull(oTankInfoLocal.LinedInteriorInspectDate)) _
        '        '        Or (Not IsDBNull(oTankInfoLocal.TankIndex) And oTankInfoLocal.TankStatus = 424 And oTankInfoLocal.TankLD = 338) And _
        '        '           (DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 5, oTankInfoLocal.TTTDate)) < 0 Or IsDBNull(oTankInfoLocal.TTTDate))) Then
        '        '        nTankCount += 1
        '        '    End If
        '        '    For Each oPipeInfoLocal In oTankInfoLocal.pipesCollection.Values
        '        '        If (Not IsDBNull(oPipeInfoLocal.Index) And (oPipeInfoLocal.PipeStatusDesc = 424 Or oPipeInfoLocal.PipeStatusDesc = 429) And (oPipeInfoLocal.PipeModDesc = 260 Or oPipeInfoLocal.PipeModDesc = 263)) And _
        '        '             (DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 3, oPipeInfoLocal.PipeCPTest)) < 0 Or IsDBNull(oPipeInfoLocal.PipeCPTest)) _
        '        '            Or (Not IsDBNull(oPipeInfoLocal.Index) And oPipeInfoLocal.PipeStatusDesc = 424 And oPipeInfoLocal.ALLDType = 496) And _
        '        '               (DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 1, oPipeInfoLocal.ALLDTestDate)) < 0 Or IsDBNull(oPipeInfoLocal.ALLDTestDate)) _
        '        '             Or (Not IsDBNull(oPipeInfoLocal.Index) And oPipeInfoLocal.PipeStatusDesc = 424 And oPipeInfoLocal.PipeLD = 245) And _
        '        '               (DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 3, oPipeInfoLocal.LTTDate)) < 0 Or IsDBNull(oPipeInfoLocal.LTTDate) _
        '        '               Or (Not IsDBNull(oPipeInfoLocal.Index) And oPipeInfoLocal.PipeStatusDesc = 424 And oPipeInfoLocal.PipeTypeDesc = 268) And _
        '        '               (DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 3, oPipeInfoLocal.LTTDate)) < 0 Or IsDBNull(oPipeInfoLocal.LTTDate) _
        '        '                Or (Not IsDBNull(oPipeInfoLocal.Index) And oPipeInfoLocal.PipeStatusDesc = 424 Or oPipeInfoLocal.PipeStatusDesc = 429) And (oPipeInfoLocal.TermCPTypeDisp = 611 Or oPipeInfoLocal.TermCPTypeDisp = 488) And (oPipeInfoLocal.TermCPTypeTank = 610 Or oPipeInfoLocal.TermCPTypeTank = 481) And _
        '        '                    DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Now), DateAdd(DateInterval.Year, 3, oPipeInfoLocal.TermCPLastTested)) < 0 Or IsDBNull(oPipeInfoLocal.TermCPLastTested))) Then
        '        '            nPipeCount += 1
        '        '        End If
        '        '    Next
        '        'Next
        '        'If (nTankCount + nPipeCount) > 0 Then
        '        '    nStatus = 0 ' -- CAP Status InValid
        '        'ElseIf (nTankCount + nPipeCount) = 0 And nTotalTankCount = 0 Then
        '        '    nStatus = 0 ' -- CAP Status InValid
        '        'ElseIf (nTankCount + nPipeCount) = 0 And nTotalTankCount <> 0 Then
        '        '    nStatus = 1 '-- CAP Status valid 
        '        'End If
        '        'If oFacilityInfo.CapStatus <> nStatus Then
        '        '    If BolEvent Then
        '        '        oFacilityInfo.CapStatus = nStatus
        '        '        RaiseEvent evtFacilityCAPStatus(Me.OwnerID, Me.ID)
        '        '    Else
        '        '        If oFacilityInfo.ID > 0 Then
        '        '            oFacilityInfo.CapStatusOriginal = nStatus
        '        '            oFacilityInfo.CapStatus = nStatus
        '        '            oFacilityDB.PutCAPStatus(oFacilityInfo.ID, oFacilityInfo.CapStatus)
        '        '        End If
        '        '    End If
        '        'End If
        '        'Return nStatus
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Public Sub GetCapStatus(ByVal facID As Integer)
            Dim capStat As Integer = oFacilityDB.GetCAPStatus(facID)
            If capStat <> oFacilityInfo.CapStatus Then
                oFacilityInfo.CapStatusOriginal = capStat
                oFacilityInfo.CapStatus = capStat
            End If
        End Sub
        Private Function GetFacilityStatusDesc() As String
            Try
                Select Case oFacilityInfo.FacilityStatus
                    Case 513
                        Return "Unregulated"
                    Case 514
                        Return "Active"
                    Case 515
                        Return "Closed"
                    Case 516
                        Return "Pre(88)"
                    Case Else
                        Return ""
                End Select
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function HasCIUTOSITanks() As Boolean
            Try
                Return oFacilityDB.DBHasCIUTOSITanks(Me.ID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Sub LoadFacilityPics(ByVal path As String)


            Dim dirpath As IO.DirectoryInfo

            Try


                If facPics Is Nothing Then
                    facPics = New Collections.ArrayList
                End If

                facPics.Clear()

                dirpath = New IO.DirectoryInfo(String.Format("{0}", path))
                For Each f As IO.FileInfo In dirpath.GetFiles(String.Format("*{0}*.jpg", Me.ID))

                    Dim name As String = f.Name.ToUpper

                    If name.Substring(name.IndexOf("_") + 1).StartsWith("F_") Then
                        name = name.Substring(name.IndexOf("_") + 1)
                    End If



                    Dim add As Integer = 0

                    If name.StartsWith("F_") OrElse name.StartsWith("_") Then
                        add += 1
                    End If

                    With name.Replace("_", "")

                        If .StartsWith(String.Format("F{0}", ID)) AndAlso .Length > (ID.ToString.Length + 1 + add) AndAlso Not IsNumeric(name.Substring(ID.ToString.Length + 1 + add, 1)) Then
                            Me.facPics.Add(f)
                        ElseIf .StartsWith(String.Format("{0}", ID)) AndAlso .Length > (ID.ToString.Length + add) AndAlso Not IsNumeric(name.Substring(ID.ToString.Length + add, 1)) Then
                            Me.facPics.Add(f)
                        End If
                    End With
                Next


            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                dirpath = Nothing
            End Try

        End Sub

        Public Function FacilityCAPTable(ByVal nOwerID As Integer) As DataSet
            Dim dSet As New DataSet
            Dim strSQL As String
            Try
                'tbFacilityTable.Columns.Add("Facility ID", Type.GetType("System.Int64"))
                'tbFacilityTable.Columns.Add("Cap Status", Type.GetType("System.Boolean"))
                'tbFacilityTable.Columns.Add("Facility Name", Type.GetType("System.String"))
                'tbFacilityTable.Columns.Add("Address", Type.GetType("System.String"))
                'tbFacilityTable.Columns.Add("City", Type.GetType("System.String"))

                'For Each oFacilityInfoLocal In colFacility.Values
                '    If nOwerID = oFacilityInfoLocal.OwnerID And Not (oFacilityInfoLocal.Deleted) Then
                '        oFacilityInfo = oFacilityInfoLocal
                '        dr = tbFacilityTable.NewRow()
                '        dr("Facility ID") = oFacilityInfoLocal.ID
                '        dr("Cap Status") = oFacilityInfoLocal.CapStatus
                '        dr("Facility Name") = oFacilityInfoLocal.Name
                '        dr("Address") = FacilityAddress.AddressLine1 & "" & FacilityAddress.AddressLine2
                '        dr("City") = FacilityAddress.City
                '        tbFacilityTable.Rows.Add(dr)
                '    End If
                'Next

                strSQL = "exec spGetCAPFacilities_ByOwner " & nOwerID

                dSet = oFacilityDB.DBGetDS(strSQL)
                Return dSet
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function getCAPTanksandPipesByFacility(ByVal nFacID As Integer, Optional ByVal Isinspection As Boolean = False, Optional ByVal clearFields As Boolean = False) As DataSet
            Dim dsTankPipe As New DataSet
            '   Dim dsRel As DataRelation
            Dim strSQL As String

            Try
                If Not Isinspection Then
                    strSQL = "SELECT * FROM V_CAP_TANK_DATA WHERE FACILITY_ID = " + nFacID.ToString + _
                                " AND ([TANK ID] IN (SELECT [TANK ID] FROM V_CAP_PIPE_DATA WHERE FACILITY_ID = " + nFacID.ToString + " AND STATUS <> 'Permanently Out of Use') OR STATUS = 'Currently In Use' OR STATUS = 'Temporarily Out of Service Indefinitely');" + _
                        "SELECT * FROM V_CAP_PIPE_DATA WHERE FACILITY_ID = " + nFacID.ToString + " AND STATUS <> 'Permanently Out of Use'"
                ElseIf Not clearFields Then
                    strSQL = "SELECT * FROM V_CAP_TANK_DATA_INSPECTED WHERE FACILITY_ID = " + nFacID.ToString + _
                                " AND ([TANK ID] IN (SELECT [TANK ID] FROM V_CAP_PIPE_DATA_INSPECTED WHERE FACILITY_ID = " + nFacID.ToString + " AND STATUS <> 'Permanently Out of Use') OR STATUS = 'Currently In Use' OR STATUS = 'Temporarily Out of Service Indefinitely');" + _
                        "SELECT * FROM V_CAP_PIPE_DATA_INSPECTED WHERE FACILITY_ID = " + nFacID.ToString + " AND STATUS <> 'Permanently Out of Use'"
                Else
                    strSQL = "SELECT * FROM V_CAP_TANK_DATA_INSPECTED_ClEARED WHERE FACILITY_ID = " + nFacID.ToString + _
                                " AND ([TANK ID] IN (SELECT [TANK ID] FROM V_CAP_PIPE_DATA_INSPECTED_CLEARED WHERE FACILITY_ID = " + nFacID.ToString + " AND STATUS <> 'Permanently Out of Use') OR STATUS = 'Currently In Use' OR STATUS = 'Temporarily Out of Service Indefinitely');" + _
                        "SELECT * FROM V_CAP_PIPE_DATA_INSPECTED_CLEARED WHERE FACILITY_ID = " + nFacID.ToString + " AND STATUS <> 'Permanently Out of Use'"
                End If

                dsTankPipe = oFacilityDB.DBGetDS(strSQL)

                If Not (dsTankPipe Is Nothing) Then

                    ' Create Tanks/Pipes relationship and add to DataSet
                    Dim relTankPipe As New DataRelation("TankPipe" _
                        , dsTankPipe.Tables(0).Columns("TANK ID") _
                        , dsTankPipe.Tables(1).Columns("TANK ID"))
                    dsTankPipe.Relations.Add(relTankPipe)

                End If

                Return dsTankPipe

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear(Optional ByVal strDepth As String = "ALL")
            Try
                oFacilityInfo = New MUSTER.Info.FacilityInfo
                oOwnerInfo.facilityCollection.Clear()
                'oFacAddress.Clear(strDepth)
                oFacTanks.Clear()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Reset(Optional ByVal strDepth As String = "ALL")
            Try
                oFacilityInfo.Reset()
                'Dim xFacInf As Muster.Info.FacilityInfo
                'If Not colFacility.Values Is Nothing Then
                '    For Each xFacInf In colFacility.Values
                '        If xFacInf.IsDirty Then
                '            xFacInf.Reset()
                '        End If
                '    Next
                'Else
                '    xFacInf.Reset()
                'End If
                'oFacAddress.Reset(strDepth)
                'oFacTanks.Reset()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetTable(ByVal strSQL As String) As DataTable
            Try
                Return oFacilityDB.DBGetDS(strSQL).Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Try
                Dim oFacInfoLocal As MUSTER.Info.FacilityInfo
                Dim colstrArr As New Collection
                Dim i As Integer
                'Dim colFacilityContained As MUSTER.Info.FacilityCollection
                'RaiseEvent evtOwnerInfoFacCol(colFacilityContained)

                For Each oFacInfoLocal In oOwnerInfo.facilityCollection.Values
                    colstrArr.Add(oFacInfoLocal.ID)
                Next
                Dim strArr(colstrArr.Count - 1) As String
                For i = 0 To colstrArr.Count - 1
                    strArr(i) = CType(colstrArr(i + 1), String)
                Next
                Dim nArr(strArr.GetUpperBound(0)) As Integer
                Dim y As String
                For Each y In strArr
                    nArr(strArr.IndexOf(strArr, y)) = CInt(y)
                Next
                nArr.Sort(nArr)
                colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))

                If colIndex + direction > -1 Then
                    If colIndex + direction <= nArr.GetUpperBound(0) Then
                        Return oOwnerInfo.facilityCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
                    Else
                        Return oOwnerInfo.facilityCollection.Item(nArr.GetValue(0)).ID.ToString
                    End If
                Else
                    Return oOwnerInfo.facilityCollection.Item(nArr.GetValue(nArr.GetUpperBound(0))).ID.ToString
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Look Up Operations"


        Public Function PopulateFacilityType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFACILITYTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateFacilityDatum() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFACILITY_DATUM")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateFacilityMethod() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFACILITY_METHOD")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateFacilityLocationType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFACILITY_LOCATION_TYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PreviousOwners() As DataSet
            Try
                Return oFacilityDB.DBGetPreviousOwners(oFacilityInfo.ID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Function GetDataTable(ByVal DBViewName As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName
                dsReturn = oFacilityDB.DBGetDS(strSQL)
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


        Public Function getFacilityCAPFieldAllowed(ByVal tankID As Integer)

            Dim HasCapRestrictions As Boolean = False
            Dim dt As DataTable

            dt = GetDataTable(String.Format("vTankGuideLineVariables where HasPressurizedpipes = 1 and IsEmergencyTank = 1 and tank_id = {0}", _
                                  tankID))

            Try
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    HasCapRestrictions = True

                    dt.Dispose()

                End If

            Catch ex As Exception

                Throw New Exception("Error retrieving Cap Capability Status")

            Finally

                If Not dt Is Nothing Then
                    dt.Dispose()
                End If

            End Try

            Return HasCapRestrictions







        End Function


        Private Function getModuleModification(ByVal type As FacilityModule, ByVal ID As Integer) As DataRow

            Try

                Dim moduleStr As String = String.Empty

                Select Case type
                    Case FacilityModule.Registration
                        moduleStr = " , 'Registration'"
                    Case FacilityModule.Closure
                        moduleStr = ", 'Closure'"
                    Case FacilityModule.Compliance
                        moduleStr = ", 'Compliance'"
                    Case FacilityModule.Fees
                        moduleStr = ", 'Fee Admin'"
                    Case FacilityModule.Financial
                        moduleStr = ", 'Financial'"
                    Case FacilityModule.Technical
                        moduleStr = ", 'Technical'"
                    Case FacilityModule.Inspection
                        moduleStr = ", 'INSPECTION'"
                    Case Else
                        moduleStr = String.Empty
                End Select

                Dim drReturn As DataRow = oFacilityDB.GetDataRow(String.Format(" sp_facilityGetLatestUserEdited {0}{1}", ID, moduleStr))

                Return drReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.Info")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Private Function GetModifyByModule(ByVal type As FacilityModule, ByVal id As Integer) As String

            Dim dr As DataRow = Me.getModuleModification(type, id)

            If Not dr Is Nothing Then
                Return dr("userID")
            Else
                Return Nothing
            End If

        End Function

        Private Function GetModifyOnModule(ByVal type As FacilityModule, ByVal id As Integer) As String

            Dim dr As DataRow = Me.getModuleModification(type, id)

            If Not dr Is Nothing AndAlso Not TypeOf dr("date_Last_edited") Is DBNull Then
                Return String.Format("{0:d}", Convert.ToDateTime(dr("date_Last_edited")))
            Else
                Return Nothing
            End If

        End Function


        Private Sub CheckDatePowerOff()
            'TankStatus 425 is Temporarily out of service
            'TankStatus 429 is Temporarily out of service indefinetly
            'TankCPType 418 is Impression Current
            Dim dateLocal As Date
            Try
                If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                    Dim oTankInfoLocal As MUSTER.Info.TankInfo
                    For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                        If oTankInfoLocal.TankStatus = 429 And _
                            oTankInfoLocal.TankCPType = 418 And _
                            oTankInfoLocal.FacilityId = oFacilityInfo.ID Then
                            oTankInfoLocal.TankStatus = 425
                        End If
                        If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                            oTankInfoLocal.FacilityPowerOff = True
                        Else
                            oTankInfoLocal.FacilityPowerOff = False
                        End If
                        For Each oPipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfoLocal.pipesCollection.Values
                            oPipeInfoLocal.FacilityPowerOff = oTankInfoLocal.FacilityPowerOff
                        Next
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Miscellaneous Operations"
        '''Public Function DeActivateFacility(ByVal nFacilityID As Integer) As Boolean
        '''    Dim BolSuccess As Boolean
        '''    BolSuccess = oFacilityDB.DBDeActivateFacility(nFacilityID)
        '''    If BolSuccess Then
        '''        Me.Remove(nFacilityID) ' Remove it from collections 
        '''    End If
        '''    Return BolSuccess
        '''End Function
        'Public Function EntityTablewithAddressDetails(ByVal nOwnrID As Integer) As DataSet
        '    Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Dim dr As DataRow
        '    Dim ds As New DataSet
        '    Dim tbFacilityTable As New DataTable
        '    Dim sList As New SortedList
        '    Dim i As Integer
        '    Try
        '        tbFacilityTable.Columns.Add("OWNER_ID")
        '        tbFacilityTable.Columns.Add("FACILITY_ID")
        '        tbFacilityTable.Columns.Add("CAP_PARTICIPANT")
        '        tbFacilityTable.Columns.Add("FACILITYNAME")
        '        tbFacilityTable.Columns.Add("ADDRESS")
        '        tbFacilityTable.Columns.Add("CITY")

        '        'Dim colFacilityContained As MUSTER.Info.FacilityCollection
        '        'RaiseEvent evtOwnerInfoFacColByOwnerID(nOwnrID, colFacilityContained)
        '        For Each oFacilityInfoLocal In oOwnerInfo.facilityCollection.Values
        '            dr = tbFacilityTable.NewRow()
        '            dr("OWNER_ID") = oFacilityInfoLocal.OwnerID
        '            dr("FACILITY_ID") = oFacilityInfoLocal.ID
        '            dr("CAP_PARTICIPANT") = oFacilityInfoLocal.CapStatus
        '            dr("FACILITYNAME") = oFacilityInfoLocal.Name
        '            oFacAddress.Retrieve(oFacilityInfoLocal.AddressID)
        '            dr("ADDRESS") = oFacAddress.AddressLine1
        '            dr("CITY") = oFacAddress.City
        '            sList.Add(oFacilityInfoLocal.ID, dr)
        '        Next
        '        For i = 0 To sList.Count - 1
        '            tbFacilityTable.Rows.Add(CType(sList.GetByIndex(i), DataRow))
        '        Next
        '        ds.Tables.Add(tbFacilityTable)
        '        Return ds
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Public Function EntityTable() As DataTable
            Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
            Dim dr As DataRow
            Dim tbFacilityTable As New DataTable
            Try
                tbFacilityTable.Columns.Add("Facility_ID")
                tbFacilityTable.Columns.Add("Facility AIID")
                tbFacilityTable.Columns.Add("Facility Name")
                tbFacilityTable.Columns.Add("OWNER ID")
                tbFacilityTable.Columns.Add("ADDRESS ID")
                tbFacilityTable.Columns.Add("BILLING ADDRESS ID")
                tbFacilityTable.Columns.Add("LATITUDE DEGREE")
                tbFacilityTable.Columns.Add("LATITUDE MINUTES")
                tbFacilityTable.Columns.Add("LATITUDE SECONDS")
                tbFacilityTable.Columns.Add("LONGITUDE DEGREE")
                tbFacilityTable.Columns.Add("LONGITUDE MINUTES")
                tbFacilityTable.Columns.Add("LONGITUDE SECONDS")
                tbFacilityTable.Columns.Add("PHONE")
                tbFacilityTable.Columns.Add("DATUM")
                tbFacilityTable.Columns.Add("METHOD")
                tbFacilityTable.Columns.Add("FAX")
                tbFacilityTable.Columns.Add("FEES PROFILE ID")
                tbFacilityTable.Columns.Add("FACILITY TYPE")
                tbFacilityTable.Columns.Add("FEES STATUS")
                tbFacilityTable.Columns.Add("CURRENT CIU NUMBER")
                tbFacilityTable.Columns.Add("CAP STATUS")
                tbFacilityTable.Columns.Add("CAP CANDIDATE")
                tbFacilityTable.Columns.Add("CITATION PROFILE ID")
                tbFacilityTable.Columns.Add("CURRENT LUST STATUS")
                tbFacilityTable.Columns.Add("FUEL BRAND")
                tbFacilityTable.Columns.Add("FACILITY DESCRIPTION")
                tbFacilityTable.Columns.Add("SIGNATURE NEEDED")
                tbFacilityTable.Columns.Add("DATE RECD")
                tbFacilityTable.Columns.Add("DATE TRANSFERRED")
                tbFacilityTable.Columns.Add("FACILITY STATUS")
                tbFacilityTable.Columns.Add("DELETED")
                tbFacilityTable.Columns.Add("CREATED BY")
                tbFacilityTable.Columns.Add("DATE CREATED")
                tbFacilityTable.Columns.Add("LAST EDITED BY")
                tbFacilityTable.Columns.Add("DATE LAST EDITED")
                tbFacilityTable.Columns.Add("DATE POWEROFF")
                tbFacilityTable.Columns.Add("LOCATION TYPE")
                tbFacilityTable.Columns.Add("UPCOMING INSTALLATION")
                tbFacilityTable.Columns.Add("UPCOMING INSTALLATION DATE")

                'Dim colFacilityContained As MUSTER.Info.FacilityCollection
                'RaiseEvent evtOwnerInfoFacCol(colFacilityContained)
                For Each oFacilityInfoLocal In oOwnerInfo.facilityCollection.Values
                    dr = tbFacilityTable.NewRow()
                    dr("Facility_ID") = oFacilityInfoLocal.ID
                    dr("Facility AIID") = oFacilityInfoLocal.AIID
                    dr("Facility Name") = oFacilityInfoLocal.Name
                    dr("OWNER ID") = oFacilityInfoLocal.OwnerID
                    dr("ADDRESS ID") = oFacilityInfoLocal.AddressID
                    dr("BILLING ADDRESS ID") = oFacilityInfoLocal.BillingAddressID
                    dr("LATITUDE DEGREE") = oFacilityInfoLocal.LatitudeDegree
                    dr("LATITUDE MINUTES") = oFacilityInfoLocal.LatitudeMinutes
                    dr("LATITUDE SECONDS") = oFacilityInfoLocal.LatitudeSeconds
                    dr("LONGITUDE DEGREE") = oFacilityInfoLocal.LongitudeDegree
                    dr("LONGITUDE MINUTES") = oFacilityInfoLocal.LongitudeMinutes
                    dr("LONGITUDE SECONDS") = oFacilityInfoLocal.LongitudeSeconds
                    dr("PHONE") = oFacilityInfoLocal.Phone
                    dr("DATUM") = oFacilityInfoLocal.Datum
                    dr("METHOD") = oFacilityInfoLocal.Method
                    dr("FAX") = oFacilityInfoLocal.Fax
                    dr("FEES PROFILE ID") = oFacilityInfoLocal.FeesProfileId
                    dr("FACILITY TYPE") = oFacilityInfoLocal.FacilityType
                    dr("FEES STATUS") = oFacilityInfoLocal.FeesStatus
                    dr("CURRENT CIU NUMBER") = oFacilityInfoLocal.CurrentCIUNumber
                    dr("CAP STATUS") = oFacilityInfoLocal.CapStatus
                    dr("CAP CANDIDATE") = oFacilityInfoLocal.CAPCandidate
                    dr("CITATION PROFILE ID") = oFacilityInfoLocal.CitationProfileID
                    dr("CURRENT LUST STATUS") = oFacilityInfoLocal.CurrentLUSTStatus
                    dr("FUEL BRAND") = oFacilityInfoLocal.FuelBrand
                    dr("FACILITY DESCRIPTION") = oFacilityInfoLocal.FacilityDescription
                    dr("SIGNATURE NEEDED") = oFacilityInfoLocal.SignatureOnNF
                    dr("DATE RECD") = oFacilityInfoLocal.DateReceived.Date
                    dr("DATE TRANSFERRED") = oFacilityInfoLocal.DateTransferred.Date
                    dr("FACILITY STATUS") = oFacilityInfoLocal.FacilityStatus
                    dr("DELETED") = oFacilityInfoLocal.Deleted
                    dr("CREATED BY") = oFacilityInfoLocal.CreatedBy
                    dr("DATE CREATED") = oFacilityInfoLocal.CreatedOn
                    dr("LAST EDITED BY") = oFacilityInfoLocal.ModifiedBy
                    dr("DATE LAST EDITED") = oFacilityInfoLocal.ModifiedOn
                    dr("DATE POWEROFF") = oFacilityInfoLocal.DatePowerOff.Date
                    dr("LOCATION TYPE") = oFacilityInfoLocal.LocationType
                    dr("UPCOMING INSTALLATION") = oFacilityInfoLocal.UpcomingInstallation
                    dr("UPCOMING INSTALLATION DATE") = oFacilityInfoLocal.UpcomingInstallationDate.Date
                    tbFacilityTable.Rows.Add(dr)
                Next
                Return tbFacilityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function FacilityCombo() As DataTable
            Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
            Dim dr As DataRow
            Dim tbFacilityTable As New DataTable
            Try
                tbFacilityTable.Columns.Add("Facility ID")
                tbFacilityTable.Columns.Add("Facility Name")

                'Dim colFacilityContained As MUSTER.Info.FacilityCollection
                'RaiseEvent evtOwnerInfoFacCol(colFacilityContained)
                For Each oFacilityInfoLocal In oOwnerInfo.facilityCollection.Values
                    dr = tbFacilityTable.NewRow()
                    dr("Facility ID") = oFacilityInfoLocal.ID
                    dr("Facility Name") = oFacilityInfoLocal.Name
                    tbFacilityTable.Rows.Add(dr)
                Next
                Return tbFacilityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function ValidatePhone(ByVal strPhone As String) As Boolean
            Try
                Dim strRegex As String = "(\(\d\d\d\))?\s*(\d\d\d)\s*[\-]?\s*(\d\d\d\d)"
                '"^\(?\d{3}\)?\s|-\d{3}-\d{4}$" -  matches (555) 555-5555, or 555-555-5555
                Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
                If rx.IsMatch(strPhone) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function TankPipeDataset(ByVal facID As Integer) As DataSet
            Dim dsTankPipe As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim dsRel2 As DataRelation
            Dim strSQL As String
            Try
                'dsTankPipe.Tables.Add(oFacTanks.TanksTable(oFacilityInfo.ID))
                'dsTankPipe.Tables.Add(oFacTanks.Compartments.Pipes.PipesTable(oFacilityInfo.ID))
                'For Each oCol In dsTankPipe.Tables(0).Columns
                '    oCol.ReadOnly = True
                'Next
                'For Each oCol In dsTankPipe.Tables(1).Columns
                '    oCol.ReadOnly = True
                'Next
                'dsRel = New DataRelation("TankToPipe", dsTankPipe.Tables(0).Columns("Tank_ID"), dsTankPipe.Tables(1).Columns("Tank_ID"), False)

                strSQL = "SELECT * FROM V_TANK_DISPLAY_DATA WHERE FACILITY_ID = '" + facID.ToString + "' ORDER BY POSITION, [TANK SITE ID];" & _
                          "SELECT * FROM V_PIPES_DISPLAY_DATA WHERE FACILITY_ID = '" + facID.ToString + "' AND PARENT_PIPE_ID = 0 ORDER BY POSITION, [PIPE SITE ID];" & _
                          "SELECT * FROM V_PIPES_DISPLAY_DATA WHERE FACILITY_ID = '" + facID.ToString + "' AND PARENT_PIPE_ID > 0 ORDER BY POSITION, [PIPE SITE ID];"

                dsTankPipe = oFacilityDB.DBGetDS(strSQL)

                'For Each drRow In dsTankPipe.Tables(2).Rows
                '    dsTankPipe.Tables(0).ImportRow(drRow)
                'Next
                'For Each drRow In dsTankPipe.Tables(3).Rows
                '    dsTankPipe.Tables(1).ImportRow(drRow)
                'Next
                'For Each oCol In dsTankPipe.Tables(0).Columns
                '    oCol.ReadOnly = True
                'Next
                'For Each oCol In dsTankPipe.Tables(1).Columns
                '    oCol.ReadOnly = True
                'Next
                'oCol = dsTankPipe.Tables(0).Columns("COMPARTMENT")
                'dsTankPipe.Tables(0).Columns.Remove(oCol)
                'oCol = dsTankPipe.Tables(1).Columns("FILLER2")
                'dsTankPipe.Tables(1).Columns.Remove(oCol)
                'dsTankPipe.Tables(0).DefaultView.Sort = "TANK SITE ID, STATUS ASC"
                'dsTankPipe.Tables(1).DefaultView.Sort = "PIPE SITE ID, PIPE STATUS ASC"

                dsRel = New DataRelation("TankToPipe", dsTankPipe.Tables(0).Columns("TANK ID"), dsTankPipe.Tables(1).Columns("TANK ID"), False)

                Dim c1 As DataColumn() = {dsTankPipe.Tables(1).Columns("TANK ID"), dsTankPipe.Tables(1).Columns("PIPE ID")}
                Dim c2 As DataColumn() = {dsTankPipe.Tables(2).Columns("TANK ID"), dsTankPipe.Tables(2).Columns("PARENT_PIPE_ID")}




                dsRel2 = New DataRelation("PipeToChild", c1, c2, False)

                dsTankPipe.Relations.Add(dsRel)
                dsTankPipe.Relations.Add(dsRel2)

                Return dsTankPipe
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ClosureEventDataSet() As DataSet
            Dim dsClosureEvents As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try
                strSQL = "select * from vCLOSURE_FACILITY_DISPLAY_DATA"
                strSQL += " where facility_id = " + oFacilityInfo.ID.ToString + " order by [NOI ID]"
                dsClosureEvents = oFacilityDB.DBGetDS(strSQL)
                Return dsClosureEvents
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function LustEventDataset(Optional ByVal bolForTransfer As Boolean = False) As DataSet
            Dim dsLustEvents As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try

                strSQL = "select * from VLUSTEVENTS"
                strSQL = strSQL & " where Facility_ID = " & oFacilityInfo.ID.ToString & " "
                If bolForTransfer Then
                    strSQL = strSQL & " and Event_ID not in (Select Tec_Event_ID from tblFIN_Event where deleted = 0 and Fin_Event_ID in (Select Fin_Event_ID from tblFIN_Commitment where PONumber > '0')) "
                End If

                dsLustEvents = oFacilityDB.DBGetDS(strSQL)

                Return dsLustEvents

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function FinancialEventDataset() As DataSet
            Dim dsLustEvents As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try

                'strSQL = "select * from vFinancialEvents"
                'strSQL = strSQL & " where Facility_ID = " & oFacilityInfo.ID.ToString & " "
                strSQL = "select FinancialEvent,TechnicalEvent,Event_date,Project_Manager,Financial_event_status,MGPTF_status,facility_id,"
                strSQL += "Convert(varchar,CommitmentTotal,1) as CommitmentTotal,Convert(varchar,RequestedTotal,1) as RequestedTotal,"
                strSQL += "Convert(varchar,TotalPaid,1) as TotalPaid,FIN_EVENT_ID from vFinancialEvents where facility_id=" & oFacilityInfo.ID.ToString & " "

                dsLustEvents = oFacilityDB.DBGetDS(strSQL)

                Return dsLustEvents

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function LustTankPipeDataset(ByVal nEventID As Int64) As DataSet
            Dim dsTankPipe As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try

                'strSQL = "SELECT * FROM V_LUST_TANK_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [TANK SITE ID];" & _
                '        "SELECT * FROM V_LUST_PIPE_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [PIPE SITE ID] "


                strSQL = " Select cast((case IncludedDet when 0 then 0 else 1 end) as bit) as Included, * from ("
                strSQL = strSQL & " select (select Count(*) from dbo.tblTEC_EVENT_TANK_PIPE where Event_ID = " & nEventID & " and Tank_ID = v_LUST_TANK_DISPLAY_DATA.[Tank ID]) as IncludedDet, * "
                strSQL = strSQL & " from v_LUST_TANK_DISPLAY_DATA "
                strSQL = strSQL & " where Facility_ID = " & oFacilityInfo.ID.ToString & ") as tempView;"
                strSQL = strSQL & " "
                strSQL = strSQL & " Select cast((case IncludedDet when 0 then 0 else 1 end) as bit) as Included, * from ("
                strSQL = strSQL & " select (select Count(*) from dbo.tblTEC_EVENT_TANK_PIPE where Event_ID = " & nEventID & " and PIPE_ID = v_LUST_PIPE_DISPLAY_DATA.[PIPE_ID]) as IncludedDet, * "
                strSQL = strSQL & " from v_LUST_PIPE_DISPLAY_DATA "
                strSQL = strSQL & " where Facility_ID = " & oFacilityInfo.ID.ToString & ") as tempView"


                dsTankPipe = oFacilityDB.DBGetDS(strSQL)

                'For Each oCol In dsTankPipe.Tables(0).Columns
                '    oCol.ReadOnly = True
                'Next
                'For Each oCol In dsTankPipe.Tables(1).Columns
                '    oCol.ReadOnly = True
                'Next
                dsTankPipe.Tables(0).DefaultView.Sort = "POSITION, TANK SITE ID"
                dsTankPipe.Tables(1).DefaultView.Sort = "POSITION, PIPE SITE ID"

                For Each oCol In dsTankPipe.Tables(0).Columns
                    If oCol.Caption <> "Included" Then
                        oCol.ReadOnly = True
                    Else
                        oCol.ReadOnly = False
                    End If
                Next

                For Each oCol In dsTankPipe.Tables(1).Columns
                    If oCol.Caption <> "Included" Then
                        oCol.ReadOnly = True
                    Else
                        oCol.ReadOnly = False
                    End If
                Next
                dsRel = New DataRelation("TankToPipe", dsTankPipe.Tables(0).Columns("TANK ID"), dsTankPipe.Tables(1).Columns("TANK ID"), False)

                dsTankPipe.Relations.Add(dsRel)
                Return dsTankPipe
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function CheckTankPlacedInService(ByVal tnkIDs As String) As DataSet
            Dim ds As New DataSet
            Dim strSQL As String = String.Empty
            Try
                strSQL = "SELECT FACILITY_ID, dbo.udfGetTankWithNoPlacedInServiceDate(FACILITY_ID) AS TANK_INDEX " + _
                            "FROM TBLREG_TANK WHERE TANK_ID IN (" + tnkIDs + ") GROUP BY FACILITY_ID"
                ds = oFacilityDB.DBGetDS(strSQL)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "Event Handlers"
        'Private Sub FacilitiesChanged(ByVal strSrc As String) Handles colFacility.FacilityColChanged
        'RaiseEvent evtFacilitiesChanged(Me.colIsDirty)
        'End Sub
        Private Sub FacilityChanged(ByVal bolValue As Boolean) Handles oFacilityInfo.FacilityInfoChanged
            RaiseEvent evtFacilityChanged(bolValue)
        End Sub
        Private Sub AddressErr(ByVal MsgStr As String) Handles oFacAddress.evtAddressErr
            RaiseEvent evtFacilityErr(MsgStr)
        End Sub
        'Private Sub evtCAPStatusfromTank(ByVal facID As Integer) Handles oFacTanks.evtCAPStatusfromTank
        '    RaiseEvent evtFacilityCAPStatus(facID)
        '    'GetCAPSTATUS(nFacId)
        'End Sub
        'Private Sub evtCAPStatusfromPipe(ByVal nOwnerID As Integer, ByVal nFacId As Integer) Handles oFacTanks.evtCAPStatusfromPipe
        '    GetCAPSTATUS(nFacId)
        'End Sub
        'Private Sub FacilityCommentsChanged(ByVal bolValue As Boolean) Handles oComments.InfoBecameDirty
        '    RaiseEvent evtFacilityCommentsChanged(bolValue)
        'End Sub

        'Private Sub TankCommentsChanged(ByVal bolValue As Boolean) Handles oFacTanks.evtTankCommentsChanged
        '    RaiseEvent evtTankCommentsChanged(bolValue)
        'End Sub

        'Private Sub PipeCommentsChanged(ByVal bolValue As Boolean) Handles oFacTanks.evtPipeCommentsChanged
        '    RaiseEvent evtPipeCommentsChanged(bolValue)
        'End Sub
        'Public Event evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String)
        'Private Sub TankValidationError(ByVal tnkID As Integer, ByVal strMessage As String) Handles oFacTanks.evtTankValidationErr
        '    RaiseEvent evtTankValidationErr(tnkID, strMessage)
        'End Sub
        'Private Sub TankStatusChanged(ByVal oldStat As Integer, ByVal newStat As Integer, ByVal facID As Integer) Handles oFacTanks.evtTankStatusChanged
        '    RaiseEvent evtTankStatusChanged(oldStat, newStat, facID)
        'End Sub
        'Events added by kiran
        'Private Sub FacTankCol(ByVal facID As Integer, ByVal tankCol As MUSTER.Info.TankCollection) Handles oFacTanks.evtTankColFac
        '    'Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    'Try
        '    '    oFacilityInfoLocal = colFacility.Item(facID)
        '    '    If Not (oFacilityInfoLocal Is Nothing) Then
        '    '        oFacilityInfoLocal.TankCollection = tankCol
        '    '    End If
        '    'Catch ex As Exception
        '    '    If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '    '    Throw ex
        '    'End Try
        '    RaiseEvent evtTankColFacOwner(facID, tankCol)
        'End Sub
        'Private Sub CommentsCol(ByVal fACID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oComments.evtCommentColFac
        '    'Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    'Try
        '    '    oFacilityInfoLocal = colFacility.Item(fACID)
        '    '    If Not (oFacilityInfoLocal Is Nothing) Then
        '    '        oFacilityInfoLocal.CommentsCollection = commentsCol
        '    '    End If
        '    'Catch ex As Exception
        '    '    If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '    '    Throw ex
        '    'End Try
        '    RaiseEvent evtCommentsCol(fACID, commentsCol)
        'End Sub
        'Private Sub compartmentCol(ByVal TankID As Integer, ByVal CompartmentCol As MUSTER.Info.CompartmentCollection, ByVal FacId As Integer) Handles oFacTanks.evtCompartmentCol
        '    RaiseEvent evtCompartmentCol(TankID, CompartmentCol, FacId)
        'End Sub
        'Private Sub commentsCol(ByVal TankID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection, ByVal FacId As Integer) Handles oFacTanks.evtCommentsCol
        '    RaiseEvent evtCommentsColTank(TankID, commentsCol, FacId)
        'End Sub
        'Private Sub pipeCommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal TankID As Integer, ByVal facID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oFacTanks.evtPipeCommentsCol
        '    RaiseEvent evtPipeCommentsCol(pipeID, compID, TankID, facID, commentsCol)
        'End Sub
        'End changes
        'Private Sub CommentInfoFac(ByVal Facid As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo) Handles oComments.evtCommentInfoFac
        '    RaiseEvent evtCommentInfoFac(Facid, commentsInfo)
        'End Sub
        'Private Sub TankInfoFac(ByVal tankInfo As MUSTER.Info.TankInfo, ByVal strDesc As String) Handles oFacTanks.evtTankInfoFac
        '    RaiseEvent evtTankInfoFac(tankInfo, strDesc)
        'End Sub
        'Public Sub TankInfoTankID(ByVal tnkID As Integer) Handles oFacTanks.evtTankInfoTankID
        '    Try
        '        oFacilityInfo.TankCollection.Remove(tnkID)
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CompInfoTank(ByVal compartmentInfo As MUSTER.Info.CompartmentInfo, ByVal strDesc As String) Handles oFacTanks.evtCompInfoTank
        '    RaiseEvent evtCompInfoTank(compartmentInfo, strDesc)
        'End Sub
        'Private Sub PipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String) Handles oFacTanks.evtPipeInfoCompartment
        '    RaiseEvent evtPipeInfoCompartment(pipeInfo, strDesc)
        'End Sub
        'Private Sub FacilityInfoTankCol(ByRef colTnk As MUSTER.Info.TankCollection) Handles oFacTanks.evtFacilityInfoTankCol
        '    colTnk = oFacilityInfo.TankCollection
        'End Sub
        'Private Sub FacilityInfoTankColByFacilityID(ByVal facID As Integer, ByRef colTnk As MUSTER.Info.TankCollection) Handles oFacTanks.evtFacilityInfoTankColByFacilityID
        '    RaiseEvent evtFacilityInfoTankColByFacilityID(facID, colTnk)
        'End Sub
        'Private Sub TankChangeKey(ByVal oldID As Integer, ByVal newID As Integer) Handles oFacTanks.evtTankChangeKey
        '    Try
        '        If oFacilityInfo.TankCollection.Contains(oldID) Then
        '            oFacilityInfo.TankCollection.ChangeKey(oldID, newID)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub

        Private Sub FacilityAddressChanged(ByVal bolValue As Boolean) Handles oFacAddress.evtAddressChanged
            RaiseEvent evtAddressChanged(bolValue)

            oFacilityInfo.AddressID = oFacAddress.AddressId

        End Sub
        Private Sub FacilityAddressesChanged(ByVal bolValue As Boolean) Handles oFacAddress.evtAddressesChanged
            RaiseEvent evtAddressesChanged(bolValue)
            oFacilityInfo.AddressID = oFacAddress.AddressId
        End Sub
#End Region
    End Class
End Namespace
